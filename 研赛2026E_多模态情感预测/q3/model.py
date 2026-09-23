"""AMEE：可加性模态专家与证据注意力网络（Additive Modality Experts with Evidence attention）。

设计目标：预测值可以**精确地**逐级分解 —— 先分解到"专家"（文本 / 语音 / 视觉 / 跨模态交互），
再把每个模态专家的份额分解到该模态的序列位置上。任何一级分解都是等式，而不是事后近似。

记模态 m ∈ {T, A, V}，d = d_model，B = intensity_bound（=3）。
    1. 单模态编码  H_m = Enc_m(x_m)                      common.layers.ModalityEncoder，只在本模态内部混合
                                                          （不做跨模态交互 → 位置证据只属于本模态）
    2. 证据注意力  a_{m,t} = softmax_{t∈C_m}( w_m^T tanh(W_m H_{m,t}) )     C_m = 本模态可用位置
                   （文本再去掉 [CLS]/[SEP]，使文本证据都能落回到词）；C_m = ∅（模态缺席）时 a_m ≡ 0
       位置值      v_{m,t} = B·tanh(u_m^T H_{m,t} + d_m)（强度）,   ℓ_{m,t} = U_m H_{m,t} + e_m（3 类 logits）
       模态专家    s_m = Σ_t a_{m,t} v_{m,t},   ℓ_m = Σ_t a_{m,t} ℓ_{m,t},   z_m = Σ_t a_{m,t} H_{m,t}
                   （缺席模态 z_m = 可学习的"缺席"向量 z_m^∅）
    3. 交互专家    [CLS; z_T+p_T; z_A+p_A; z_V+p_V] → Transformer → h_I；  s_I = B·tanh(u_I^T h_I + d_I)，ℓ_I = U_I h_I + e_I
    4. 门控        w = softmax( G([z_T; z_A; z_V; h_I]) + log 1[专家在场] )   —— 缺席模态的专家权重恰为 0
       输出        ŷ = Σ_k w_k s_k ∈ [−B, B]（凸组合，天然有界，裁剪不起作用）,   ℓ = Σ_k w_k ℓ_k
    精确分解：c_k = w_k s_k，Σ_k c_k = ŷ；模态 m 的位置贡献 e_{m,t} = w_m a_{m,t} v_{m,t}，Σ_t e_{m,t} = c_m。

前置的时间下采样（与问题2 相同的做法，只对长序列）：L_m > max_positions 时按步长 s_m = ⌈L_m / max_positions⌉
把相邻 s_m 行做"可用行掩码均值"合并；解释时下采样后的位置 k 对应原始位置 [k·s_m, (k+1)·s_m)。
"""
from __future__ import annotations

import math

import torch
import torch.nn as nn
import torch.nn.functional as F

from common.data import FEAT_DIMS, MODALITIES
from common.layers import MaskedAttentionPool, ModalityEncoder

EXPERTS = ("text", "audio", "vision", "interaction")


def pool_time(x: torch.Tensor, avail: torch.Tensor, valid: torch.Tensor, s: int):
    """按步长 s 合并相邻位置：x 取窗口内可用行的均值；avail = 窗口内有可用行；valid = 窗口内有有效行。"""
    if s <= 1:
        return x, avail, valid
    B, L, D = x.shape
    pad = (-L) % s
    if pad:
        x = F.pad(x, (0, 0, 0, pad))
        avail = F.pad(avail, (0, pad))
        valid = F.pad(valid, (0, pad))
    Lp = x.shape[1] // s
    a = (avail & valid).view(B, Lp, s).float()
    xs = (x.view(B, Lp, s, D) * a.unsqueeze(-1)).sum(2) / a.sum(2).clamp(min=1.0).unsqueeze(-1)
    return xs, a.sum(2) > 0, valid.view(B, Lp, s).any(2)


def resolve_strides(cfg: dict, seq_lens: dict) -> dict:
    """把 model.temporal_stride 解析成 {m: s_m} 写回配置（存进 model.pt，推理时按同样步长）。"""
    mc = cfg["model"]
    st = mc.get("temporal_stride")
    if not isinstance(st, dict):
        mp = int(mc.get("max_positions", 100) or 0)
        st = {m: (int(math.ceil(seq_lens[m] / mp)) if mp and seq_lens[m] > mp else 1) for m in MODALITIES}
    mc["temporal_stride"] = {m: int(st.get(m, 1)) for m in MODALITIES}
    return cfg


def special_positions(valid: torch.Tensor) -> torch.Tensor:
    """文本有效范围的首、尾位置（BERT 的 [CLS] / [SEP]）；有效长度 < 3 时不认定特殊位置。"""
    B, L = valid.shape
    vi = valid.int()
    idx = torch.arange(L, device=valid.device).unsqueeze(0)
    first = vi.argmax(1, keepdim=True)
    last = L - 1 - vi.flip(1).argmax(1, keepdim=True)
    sp = (idx == first) | (idx == last)
    return sp & (vi.sum(1, keepdim=True) >= 3) & valid


def evidence_candidates(avail: torch.Tensor, valid: torch.Tensor, exclude_special: bool) -> torch.Tensor:
    """证据候选位置 C_m：可用且在有效范围内；文本可选去掉 [CLS]/[SEP]（去掉后一个不剩则保留原样）。"""
    cand = avail & valid
    if not exclude_special:
        return cand
    c2 = cand & ~special_positions(valid)
    keep = c2.any(1, keepdim=True)
    return torch.where(keep, c2, cand)


class _Base(nn.Module):
    """两种结构共用：模态编码器、证据注意力池化、缺席向量、时间下采样。"""

    def __init__(self, c: dict, feat_dims: dict | None):
        super().__init__()
        feat_dims = feat_dims or FEAT_DIMS
        self.modalities = MODALITIES
        d = int(c.get("d_model", 128))
        self.d_model = d
        self.bound = float(c.get("intensity_bound", 3.0))
        self.exclude_special = bool(c.get("text_exclude_special", True))
        st = c.get("temporal_stride") if isinstance(c.get("temporal_stride"), dict) else {}
        self.strides = {m: int(st.get(m, 1)) for m in MODALITIES}
        drop = float(c.get("dropout", 0.2))
        self.encoders = nn.ModuleDict({m: ModalityEncoder(feat_dims[m], d, int(c.get("n_layers", 2)),
                                                          int(c.get("n_heads", 4)), int(c.get("d_ff", 256)), drop,
                                                          int(c.get("conv_kernel", 3)),
                                                          float(c.get("input_dropout", 0.1)))
                                       for m in MODALITIES})
        self.evid = nn.ModuleDict({m: MaskedAttentionPool(d, int(c.get("att_hidden", 64))) for m in MODALITIES})
        self.absent = nn.Parameter(torch.randn(len(MODALITIES), d) * 0.02)

    def prepare(self, view: dict) -> dict:
        """时间下采样（步长 1 时原样返回）；avail 与 valid 取交集。"""
        out = {}
        for m in MODALITIES:
            x, a, v = pool_time(view[f"x_{m}"], view[f"avail_{m}"] & view[f"valid_{m}"], view[f"valid_{m}"],
                                self.strides[m])
            out[f"x_{m}"], out[f"avail_{m}"], out[f"valid_{m}"] = x, a & v, v
        return out

    def encode(self, view: dict) -> dict:
        """逐模态：编码 → 证据注意力。返回 H、a（缺席时全 0）、候选 C、在场标志、池化向量 z。"""
        res = {"H": {}, "attn": {}, "cand": {}, "z": {}}
        present = []
        for i, m in enumerate(MODALITIES):
            av, va = view[f"avail_{m}"], view[f"valid_{m}"]
            h = self.encoders[m](view[f"x_{m}"], av, va)
            cand = evidence_candidates(av, va, self.exclude_special and m == "text")
            z, a = self.evid[m](h, cand)
            pres = cand.any(1)
            a = a * pres.unsqueeze(1).float()
            z = torch.where(pres.unsqueeze(1), z, self.absent[i].unsqueeze(0).expand_as(z))
            res["H"][m], res["attn"][m], res["cand"][m], res["z"][m] = h, a, cand, z
            present.append(pres)
        res["present"] = torch.stack(present, dim=1)                                   # (B, 3)
        return res


class AMEE(_Base):
    def __init__(self, cfg_model: dict, feat_dims: dict | None = None):
        super().__init__(cfg_model, feat_dims)
        c = cfg_model
        d = self.d_model
        drop = float(c.get("dropout", 0.2))
        self.use_interaction = bool(c.get("use_interaction", True))
        self.val_head = nn.ModuleDict({m: nn.Linear(d, 1) for m in MODALITIES})
        self.cls_head = nn.ModuleDict({m: nn.Linear(d, 3) for m in MODALITIES})
        self.experts = EXPERTS if self.use_interaction else EXPERTS[:3]
        if self.use_interaction:
            self.mod_emb = nn.Parameter(torch.randn(len(MODALITIES), d) * 0.02)
            self.cls_tok = nn.Parameter(torch.randn(1, 1, d) * 0.02)
            layer = nn.TransformerEncoderLayer(d, int(c.get("n_heads", 4)), int(c.get("d_ff", 256)), drop,
                                               batch_first=True, norm_first=True)
            self.inter = nn.TransformerEncoder(layer, int(c.get("inter_layers", 1)), enable_nested_tensor=False)
            self.inter_norm = nn.LayerNorm(d)
            self.int_reg = nn.Linear(d, 1)
            self.int_cls = nn.Linear(d, 3)
        K = len(self.experts)
        gh = int(c.get("gate_hidden", 64))
        self.gate_mode = str(c.get("gate", "joint"))
        if self.gate_mode == "joint":        # g = G([z_T; z_A; z_V; h_I])：每个专家的打分看得到全部模态
            self.gate = nn.Sequential(nn.Linear(K * d, gh), nn.GELU(), nn.Dropout(drop), nn.Linear(gh, K))
        elif self.gate_mode == "per_expert":  # g_k = G_k(z_k)：每个专家只按自己的证据给出置信度
            self.gate = nn.ModuleDict({e: nn.Sequential(nn.Linear(d, gh), nn.GELU(), nn.Dropout(drop), nn.Linear(gh, 1))
                                       for e in self.experts})
        else:
            raise ValueError(f"未知门控方式 {self.gate_mode}")

    def _interaction(self, Z: torch.Tensor):
        seq = torch.cat([self.cls_tok.expand(Z.shape[0], 1, -1), Z + self.mod_emb.unsqueeze(0)], dim=1)
        h = self.inter_norm(self.inter(seq)[:, 0])
        return h, self.bound * torch.tanh(self.int_reg(h).squeeze(-1)), self.int_cls(h)

    def forward(self, view: dict, prepared: bool = False) -> dict:
        if not prepared:
            view = self.prepare(view)
        enc = self.encode(view)
        s_list, l_list, value, pos_logits = [], [], {}, {}
        for m in MODALITIES:
            h, a = enc["H"][m], enc["attn"][m]
            v = self.bound * torch.tanh(self.val_head[m](h).squeeze(-1))              # (B, L)
            lt = self.cls_head[m](h)                                                   # (B, L, 3)
            s_list.append((a * v).sum(1))
            l_list.append(torch.einsum("bl,blc->bc", a, lt))
            value[m], pos_logits[m] = v, lt
        Z = torch.stack([enc["z"][m] for m in MODALITIES], dim=1)                     # (B, 3, d)
        present = enc["present"]
        mask = present
        gate_in = [Z.flatten(1)]
        h_I = None
        if self.use_interaction:
            h_I, s_I, l_I = self._interaction(Z)
            s_list.append(s_I)
            l_list.append(l_I)
            gate_in.append(h_I)
            mask = torch.cat([present, torch.ones_like(present[:, :1])], dim=1)
        else:
            none = ~present.any(1, keepdim=True)          # 三个模态都缺席（只在空联盟出现）：均匀权重，输出恒为 0
            mask = present | none
        if self.gate_mode == "joint":
            g = self.gate(torch.cat(gate_in, dim=1))
        else:
            vecs = [Z[:, i] for i in range(len(MODALITIES))] + ([h_I] if h_I is not None else [])
            g = torch.cat([self.gate[e](v) for e, v in zip(self.experts, vecs)], dim=1)
        g = g.masked_fill(~mask, float("-inf"))
        w = torch.softmax(g, dim=1)                                                    # (B, K)
        S = torch.stack(s_list, dim=1)                                                 # (B, K)
        Lg = torch.stack(l_list, dim=1)                                                # (B, K, 3)
        y = (w * S).sum(1)
        logits = torch.einsum("bk,bkc->bc", w, Lg)
        return {"y": y, "logits": logits, "w": w, "s": S, "l": Lg, "attn": enc["attn"], "value": value,
                "pos_logits": pos_logits, "cand": enc["cand"], "present": present, "z": enc["z"], "h_I": h_I,
                "view": view}

    def null_output(self, device=None) -> tuple[torch.Tensor, torch.Tensor]:
        """空联盟（三模态全部缺席）的输出 (ŷ(∅), ℓ(∅))：与输入无关的常数，= 交互专家在三个"缺席"向量上的输出。
        无交互专家时为 (0, 0)。"""
        device = device or self.absent.device
        if not self.use_interaction:
            return torch.zeros(1, device=device), torch.zeros(1, 3, device=device)
        _, s, l = self._interaction(self.absent.unsqueeze(0))
        return s, l


class ConcatNet(_Base):
    """不可解释对照：同样的编码器与证据注意力池化，三模态向量拼接后过 MLP（没有可加分解）。"""

    def __init__(self, cfg_model: dict, feat_dims: dict | None = None):
        super().__init__(cfg_model, feat_dims)
        c = cfg_model
        d = self.d_model
        hh = int(c.get("head_hidden", 128))
        drop = float(c.get("dropout", 0.2))
        self.experts = ()
        self.use_interaction = False
        self.head = nn.Sequential(nn.Linear(len(MODALITIES) * d, hh), nn.GELU(), nn.Dropout(drop))
        self.reg_out = nn.Linear(hh, 1)
        self.cls_out = nn.Linear(hh, 3)

    def forward(self, view: dict, prepared: bool = False) -> dict:
        if not prepared:
            view = self.prepare(view)
        enc = self.encode(view)
        f = self.head(torch.cat([enc["z"][m] for m in MODALITIES], dim=1))
        return {"y": self.bound * torch.tanh(self.reg_out(f).squeeze(-1)), "logits": self.cls_out(f), "w": None,
                "s": None, "l": None, "attn": enc["attn"], "value": None, "pos_logits": None, "cand": enc["cand"],
                "present": enc["present"], "z": enc["z"], "h_I": None, "view": view}


def build_model(cfg: dict, feat_dims: dict | None = None) -> nn.Module:
    arch = cfg["model"].get("arch", "amee")
    if arch == "amee":
        return AMEE(cfg["model"], feat_dims)
    if arch == "concat":
        return ConcatNet(cfg["model"], feat_dims)
    raise ValueError(f"未知结构 {arch}")


def count_params(model: nn.Module) -> int:
    return int(sum(p.numel() for p in model.parameters() if p.requires_grad))
