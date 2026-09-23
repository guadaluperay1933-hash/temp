"""问题2 / 问题3 共用的 PyTorch 网络层。

掩码语义（与 common.data 一致）：
    avail (B, L) bool  该位置特征可用（非零）
    valid (B, L) bool  名义有效范围（可用 + 缺失）；其外为填充
    缺失位置 = valid & ~avail，用可学习的 [MISS] 向量占位，只作为 query 从可用位置"取"信息，
    从不作为 key/value 被别的位置读取 —— 保证置零区间不会把"零特征"当作真实证据传播出去。
"""
from __future__ import annotations

import math

import torch
import torch.nn as nn


def relative_time(valid: torch.Tensor) -> tuple[torch.Tensor, torch.Tensor]:
    """valid (B,L) → (tau, rank)：tau 为有效范围内的相对时间 ∈ [0,1]，rank 为有效范围内的序号。"""
    v = valid.long()
    rank = (v.cumsum(dim=1) - 1).clamp(min=0) * v
    length = v.sum(dim=1, keepdim=True)
    tau = rank.float() / (length - 1).clamp(min=1).float()
    return tau * valid.float(), rank


def safe_key_mask(avail: torch.Tensor) -> torch.Tensor:
    """注意力 key_padding_mask（True=屏蔽）。整行都不可用时放开第 0 个位置，避免 softmax 全 -inf 得 NaN。"""
    kpm = ~avail
    empty = kpm.all(dim=1)
    if empty.any():
        kpm = kpm.clone()
        kpm[empty, 0] = False
    return kpm


class TimeEncoding(nn.Module):
    """正弦时间编码：一半维度编码相对时间 tau（跨模态可比，非对齐版本靠它粗对齐），一半编码序号 rank。"""

    def __init__(self, d_model: int, tau_scale: float = 100.0):
        super().__init__()
        assert d_model % 4 == 0, "d_model 需为 4 的倍数"
        self.d_half = d_model // 2
        self.tau_scale = tau_scale
        freqs = torch.exp(-math.log(10000.0) * torch.arange(0, self.d_half, 2).float() / self.d_half)
        self.register_buffer("freqs", freqs, persistent=False)

    def _sin(self, x: torch.Tensor) -> torch.Tensor:
        ang = x.unsqueeze(-1) * self.freqs
        return torch.cat([torch.sin(ang), torch.cos(ang)], dim=-1)

    def forward(self, tau: torch.Tensor, rank: torch.Tensor) -> torch.Tensor:
        return torch.cat([self._sin(tau * self.tau_scale), self._sin(rank.float())], dim=-1)


class ModalityEncoder(nn.Module):
    """单模态时序编码器：线性投影 → [MISS] 占位 → 时间编码 → 局部卷积 → Transformer（只以可用位置为 key）。

    输出 (B, L, d)，填充位置输出为 0。
    """

    def __init__(self, d_in: int, d_model: int = 128, n_layers: int = 2, n_heads: int = 4, d_ff: int = 256,
                 dropout: float = 0.2, conv_kernel: int = 3, input_dropout: float = 0.1):
        super().__init__()
        self.proj = nn.Sequential(nn.Dropout(input_dropout), nn.Linear(d_in, d_model), nn.LayerNorm(d_model))
        self.miss_token = nn.Parameter(torch.randn(d_model) * 0.02)
        self.time_enc = TimeEncoding(d_model)
        self.conv = nn.Conv1d(d_model, d_model, conv_kernel, padding=conv_kernel // 2)
        layer = nn.TransformerEncoderLayer(d_model, n_heads, d_ff, dropout, batch_first=True, norm_first=True)
        self.encoder = nn.TransformerEncoder(layer, n_layers, enable_nested_tensor=False)
        self.norm = nn.LayerNorm(d_model)

    def forward(self, x: torch.Tensor, avail: torch.Tensor, valid: torch.Tensor) -> torch.Tensor:
        vf = valid.unsqueeze(-1).float()
        h = self.proj(x)
        h = torch.where(avail.unsqueeze(-1), h, self.miss_token.expand_as(h)) * vf
        tau, rank = relative_time(valid)
        h = h + self.time_enc(tau, rank) * vf
        h = h + self.conv(h.transpose(1, 2)).transpose(1, 2) * vf
        h = self.encoder(h, src_key_padding_mask=safe_key_mask(avail))
        return self.norm(h) * vf


class MaskedAttentionPool(nn.Module):
    """加性注意力池化：a_t = softmax_t( w^T tanh(W h_t) )，只在 mask 为真的位置上归一化。返回 (z, a)。"""

    def __init__(self, d_model: int, d_hidden: int | None = None):
        super().__init__()
        d_hidden = d_hidden or d_model
        self.score = nn.Sequential(nn.Linear(d_model, d_hidden), nn.Tanh(), nn.Linear(d_hidden, 1))

    def forward(self, h: torch.Tensor, mask: torch.Tensor) -> tuple[torch.Tensor, torch.Tensor]:
        s = self.score(h).squeeze(-1)
        m = ~safe_key_mask(mask)
        s = s.masked_fill(~m, float("-inf"))
        a = torch.softmax(s, dim=1)
        z = torch.einsum("bl,bld->bd", a, h)
        return z, a


def masked_mean(h: torch.Tensor, mask: torch.Tensor) -> torch.Tensor:
    m = mask.unsqueeze(-1).float()
    return (h * m).sum(dim=1) / m.sum(dim=1).clamp(min=1.0)
