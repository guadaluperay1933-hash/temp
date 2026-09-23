"""MRG-Net：缺失感知重构与可靠性门控融合网络（Missing-aware Reconstruction & reliability-Gated fusion Network）。

前向只接收"一个视图"：{x_m, avail_m, valid_m}（m ∈ 启用的模态）。训练时同一 batch 分别送入
完整视图与缺失视图（common.torch_data.apply_drop_tensors），两次前向共享全部参数。

结构（d = d_model）：
    1. 单模态编码  H_m = Enc_m(x_m)                     common.layers.ModalityEncoder，缺失位置是 [MISS] 查询
    2. 跨模态重构  Ĥ_m = MHA(q = LN(H_m)+P(τ_m),       k = [null; LN(H_o)+P(τ_o)]_{o≠m},  v = [null; H_o]_{o≠m})
                   g   = σ(W_g [H_m; Ĥ_m; 1_miss])     H'_m = H_m + g ⊙ Ĥ_m
                   x̂_m = W_rec,m H'_m                  （只在训练增广置零的位置计重构损失）
    3. 池化        z_m = MaskedAttentionPool(H'_m | valid_m)，  ρ_m = |avail∩valid| / |valid|
    4. 可靠性门控  α = softmax_m( MLP_m([z_m; ρ_m]) )
       融合        [CLS, z_m + e_m + W_ρ ρ_m] → Transformer → c；  h = c + Σ_m α_m z_m
    5. 输出        f = MLP(h)；ŷ = 3·tanh(w_r^T f)；ℓ = W_c f（3 类 logits）
"""
from __future__ import annotations

import torch
import torch.nn as nn

from common.data import FEAT_DIMS, MODALITIES
from common.layers import MaskedAttentionPool, ModalityEncoder, TimeEncoding, masked_mean, relative_time


class CrossModalReconstruction(nn.Module):
    """对每个模态 m：以本模态各位置为 query，从其余模态的"可用"位置取信息（时间编码让非对齐序列可比）。

    改进：K/V 前拼一个可学习的 null 槽位且永不屏蔽 —— 其余模态整段不可用时，注意力可以"什么都不取"，
    而不是被迫落在填充位置上（safe_key_mask 的兜底会读到全零向量）。
    """

    def __init__(self, modalities, d_model: int, n_heads: int, dropout: float):
        super().__init__()
        self.modalities = tuple(modalities)
        self.time_enc = TimeEncoding(d_model)
        self.time_proj = nn.Linear(d_model, d_model)
        self.q_norm = nn.ModuleDict({m: nn.LayerNorm(d_model) for m in self.modalities})
        self.k_norm = nn.ModuleDict({m: nn.LayerNorm(d_model) for m in self.modalities})
        self.attn = nn.ModuleDict({m: nn.MultiheadAttention(d_model, n_heads, dropout=dropout, batch_first=True)
                                   for m in self.modalities})
        self.null_k = nn.ParameterDict({m: nn.Parameter(torch.randn(1, 1, d_model) * 0.02) for m in self.modalities})
        self.null_v = nn.ParameterDict({m: nn.Parameter(torch.zeros(1, 1, d_model)) for m in self.modalities})
        self.gate = nn.ModuleDict({m: nn.Linear(2 * d_model + 1, d_model) for m in self.modalities})
        self.drop = nn.Dropout(dropout)

    def _time(self, valid: torch.Tensor) -> torch.Tensor:
        tau, _ = relative_time(valid)
        # 只用相对时间 τ（序号 rank 在非对齐模态间不可比，置 0）
        te = self.time_enc(tau, torch.zeros_like(tau, dtype=torch.long))
        return self.time_proj(te) * valid.unsqueeze(-1).float()

    def forward(self, H: dict, avail: dict, valid: dict) -> tuple[dict, dict]:
        te = {m: self._time(valid[m]) for m in self.modalities}
        out, gates = {}, {}
        for m in self.modalities:
            others = [o for o in self.modalities if o != m]
            B = H[m].shape[0]
            keys = [self.null_k[m].expand(B, 1, -1)] + [self.k_norm[m](H[o]) + te[o] for o in others]
            vals = [self.null_v[m].expand(B, 1, -1)] + [H[o] for o in others]
            kpm = [torch.zeros(B, 1, dtype=torch.bool, device=H[m].device)] + [~avail[o] for o in others]
            q = self.q_norm[m](H[m]) + te[m]
            h_hat, _ = self.attn[m](q, torch.cat(keys, 1), torch.cat(vals, 1), key_padding_mask=torch.cat(kpm, 1),
                                    need_weights=False)
            h_hat = self.drop(h_hat)
            miss = (valid[m] & ~avail[m]).unsqueeze(-1).float()
            g = torch.sigmoid(self.gate[m](torch.cat([H[m], h_hat, miss], dim=-1)))
            out[m] = (H[m] + g * h_hat) * valid[m].unsqueeze(-1).float()
            gates[m] = g
        return out, gates


class MRGNet(nn.Module):
    def __init__(self, cfg_model: dict, feat_dims: dict | None = None):
        super().__init__()
        c = cfg_model
        self.modalities = tuple(m for m in MODALITIES if m in c.get("modalities", MODALITIES))
        assert self.modalities, "至少启用一个模态"
        feat_dims = feat_dims or FEAT_DIMS
        d = int(c.get("d_model", 128))
        self.d_model = d
        self.use_rec = bool(c.get("use_reconstruction", True)) and len(self.modalities) >= 2
        self.use_gate = bool(c.get("use_gate", True))
        drop = float(c.get("dropout", 0.2))
        self.encoders = nn.ModuleDict({m: ModalityEncoder(feat_dims[m], d, int(c.get("n_layers", 2)),
                                                          int(c.get("n_heads", 4)), int(c.get("d_ff", 256)), drop,
                                                          int(c.get("conv_kernel", 3)),
                                                          float(c.get("input_dropout", 0.1)))
                                       for m in self.modalities})
        if self.use_rec:
            self.cmr = CrossModalReconstruction(self.modalities, d, int(c.get("n_heads", 4)), drop)
            self.rec_head = nn.ModuleDict({m: nn.Linear(d, feat_dims[m]) for m in self.modalities})
        self.pool = nn.ModuleDict({m: MaskedAttentionPool(d) for m in self.modalities})
        gh = int(c.get("gate_hidden", 64))
        self.gate_mlp = nn.ModuleDict({m: nn.Sequential(nn.Linear(d + 1, gh), nn.GELU(), nn.Linear(gh, 1))
                                       for m in self.modalities})
        self.mod_emb = nn.Parameter(torch.randn(len(self.modalities), d) * 0.02)
        self.rho_emb = nn.Linear(1, d)
        self.cls = nn.Parameter(torch.randn(1, 1, d) * 0.02)
        layer = nn.TransformerEncoderLayer(d, int(c.get("n_heads", 4)), int(c.get("d_ff", 256)), drop,
                                           batch_first=True, norm_first=True)
        self.fusion = nn.TransformerEncoder(layer, int(c.get("fusion_layers", 1)), enable_nested_tensor=False)
        self.fusion_norm = nn.LayerNorm(d)
        hh = int(c.get("head_hidden", 128))
        self.head = nn.Sequential(nn.Linear(d, hh), nn.GELU(), nn.Dropout(drop))
        self.reg_out = nn.Linear(hh, 1)
        self.cls_out = nn.Linear(hh, 3)

    def forward(self, view: dict) -> dict:
        mods = self.modalities
        avail = {m: view[f"avail_{m}"] for m in mods}
        valid = {m: view[f"valid_{m}"] for m in mods}
        # avail 可能超出 valid（数据里极少数异常行在 common.data 已并入 valid）；这里保险起见取交集
        avail = {m: avail[m] & valid[m] for m in mods}
        H = {m: self.encoders[m](view[f"x_{m}"], avail[m], valid[m]) for m in mods}
        rec, cm_gate = None, None
        if self.use_rec:
            H, cm_gate = self.cmr(H, avail, valid)
            rec = {m: self.rec_head[m](H[m]) for m in mods}
        z, tatt = {}, {}
        for m in mods:
            z[m], tatt[m] = self.pool[m](H[m], valid[m])
        nvalid = {m: valid[m].sum(1).float() for m in mods}
        rho = torch.stack([(avail[m].sum(1).float() / nvalid[m].clamp(min=1.0)) for m in mods], dim=1)  # (B,M)
        Z = torch.stack([z[m] for m in mods], dim=1)                                                    # (B,M,d)
        if self.use_gate:
            s = torch.cat([self.gate_mlp[m](torch.cat([z[m], rho[:, i:i + 1]], -1)) for i, m in enumerate(mods)], 1)
            alpha = torch.softmax(s, dim=1)
        else:
            alpha = torch.full_like(rho, 1.0 / len(mods))
        tokens = Z + self.mod_emb.unsqueeze(0) + self.rho_emb(rho.unsqueeze(-1))
        seq = torch.cat([self.cls.expand(Z.shape[0], 1, -1), tokens], dim=1)
        c = self.fusion_norm(self.fusion(seq)[:, 0])
        h = c + torch.einsum("bm,bmd->bd", alpha, Z)
        f = self.head(h)
        y = 3.0 * torch.tanh(self.reg_out(f).squeeze(-1))
        logits = self.cls_out(f)
        return {"y": y, "logits": logits, "h": h, "alpha": alpha, "rho": rho, "rec": rec, "cm_gate": cm_gate,
                "time_attn": tatt}


class LateFusionBaseline(nn.Module):
    """晚期融合基线：各模态在"可用位置"上做掩码均值（原始标准化特征空间），拼接后过 MLP。

    缺失位置天然被排除在均值之外，是最朴素的"忽略缺失"做法；不含任何缺失建模。
    """

    def __init__(self, cfg_model: dict, feat_dims: dict | None = None):
        super().__init__()
        c = cfg_model
        self.modalities = tuple(m for m in MODALITIES if m in c.get("modalities", MODALITIES))
        feat_dims = feat_dims or FEAT_DIMS
        hh = int(c.get("head_hidden", 128))
        d = int(c.get("d_model", 128))
        drop = float(c.get("dropout", 0.2))
        d_in = sum(feat_dims[m] for m in self.modalities)
        self.body = nn.Sequential(nn.Dropout(float(c.get("input_dropout", 0.1))), nn.Linear(d_in, d), nn.GELU(),
                                  nn.Dropout(drop), nn.Linear(d, d), nn.GELU())
        self.head = nn.Sequential(nn.Dropout(drop), nn.Linear(d, hh), nn.GELU(), nn.Dropout(drop))
        self.reg_out = nn.Linear(hh, 1)
        self.cls_out = nn.Linear(hh, 3)

    def forward(self, view: dict) -> dict:
        mods = self.modalities
        avail = {m: view[f"avail_{m}"] & view[f"valid_{m}"] for m in mods}
        u = torch.cat([masked_mean(view[f"x_{m}"], avail[m]) for m in mods], dim=-1)
        rho = torch.stack([avail[m].sum(1).float() / view[f"valid_{m}"].sum(1).float().clamp(min=1.0)
                           for m in mods], dim=1)
        h = self.body(u)
        f = self.head(h)
        return {"y": 3.0 * torch.tanh(self.reg_out(f).squeeze(-1)), "logits": self.cls_out(f), "h": h,
                "alpha": None, "rho": rho, "rec": None, "cm_gate": None, "time_attn": None}


def build_model(cfg: dict, feat_dims: dict | None = None) -> nn.Module:
    arch = cfg["model"].get("arch", "mrgnet")
    if arch == "mrgnet":
        return MRGNet(cfg["model"], feat_dims)
    if arch == "late_fusion":
        return LateFusionBaseline(cfg["model"], feat_dims)
    raise ValueError(f"未知结构 {arch}")


def count_params(model: nn.Module) -> int:
    return int(sum(p.numel() for p in model.parameters() if p.requires_grad))
