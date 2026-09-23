"""SplitData → PyTorch Dataset / DataLoader。

每个 batch 是一个字典：
    x_text / x_audio / x_vision      float (B, L_m, D_m)   —— 原始（完整）特征
    avail_* / valid_*                bool  (B, L_m)
    drop_*                           bool  (B, L_m)        —— 训练增广要额外置零的位置（无增广时全 False）
    y_reg float (B,)   y_cls long (B,)（无标签时为 NaN / -1）   index long (B,)
增广只给出 drop 掩码，由模型前向时施加：这样同一个 batch 既能得到"完整视图"又能得到"缺失视图"
（问题2 的一致性蒸馏需要两者），缺失位置的原始特征还能作为重构目标。
collate 时把序列裁到本 batch 实际用到的范围，非对齐版本（L=500）能省很多计算。
"""
from __future__ import annotations

import numpy as np
import torch
from torch.utils.data import DataLoader, Dataset

from .data import MODALITIES, SplitData
from .missing import AugmentConfig, random_drop_for_sample


class MSADataset(Dataset):
    def __init__(self, sd: SplitData, augment: AugmentConfig | None = None, seed: int = 0):
        self.sd = sd
        self.augment = augment
        self.seed = seed
        self.epoch = 0

    def set_epoch(self, epoch: int) -> None:
        self.epoch = epoch

    def __len__(self) -> int:
        return len(self.sd)

    def __getitem__(self, i: int) -> dict:
        sd = self.sd
        item = {"index": i}
        for m in MODALITIES:
            item[f"x_{m}"] = sd.feats[m][i]
            item[f"avail_{m}"] = sd.avail[m][i]
            item[f"valid_{m}"] = sd.valid[m][i]
        if self.augment is not None:
            rng = np.random.default_rng((self.seed, self.epoch, i))
            drop = random_drop_for_sample({m: sd.valid[m][i] for m in MODALITIES}, self.augment, rng)
        else:
            drop = {m: np.zeros_like(sd.avail[m][i]) for m in MODALITIES}
        for m in MODALITIES:
            item[f"drop_{m}"] = drop[m]
        item["y_reg"] = float(sd.y_reg[i]) if sd.y_reg is not None else float("nan")
        item["y_cls"] = int(sd.y_cls[i]) if sd.y_cls is not None else -1
        return item


def collate(items: list[dict]) -> dict:
    out = {"index": torch.tensor([it["index"] for it in items], dtype=torch.long),
           "y_reg": torch.tensor([it["y_reg"] for it in items], dtype=torch.float32),
           "y_cls": torch.tensor([it["y_cls"] for it in items], dtype=torch.long)}
    for m in MODALITIES:
        valid = np.stack([it[f"valid_{m}"] for it in items])
        used = np.where(valid.any(axis=0))[0]
        lo, hi = (int(used[0]), int(used[-1]) + 1) if used.size else (0, 1)
        out[f"x_{m}"] = torch.from_numpy(np.stack([it[f"x_{m}"][lo:hi] for it in items])).float()
        out[f"avail_{m}"] = torch.from_numpy(np.stack([it[f"avail_{m}"][lo:hi] for it in items]))
        out[f"valid_{m}"] = torch.from_numpy(valid[:, lo:hi].copy())
        out[f"drop_{m}"] = torch.from_numpy(np.stack([it[f"drop_{m}"][lo:hi] for it in items]))
        out[f"offset_{m}"] = lo  # 裁剪起点：把位置映射回原始序列坐标时要加回去
    return out


def make_loader(sd: SplitData, batch_size: int = 64, shuffle: bool = False, augment: AugmentConfig | None = None,
                seed: int = 0, num_workers: int = 0) -> DataLoader:
    ds = MSADataset(sd, augment=augment, seed=seed)
    g = torch.Generator()
    g.manual_seed(seed)
    return DataLoader(ds, batch_size=batch_size, shuffle=shuffle, collate_fn=collate, num_workers=num_workers,
                      generator=g if shuffle else None)


def to_device(batch: dict, device) -> dict:
    return {k: (v.to(device) if isinstance(v, torch.Tensor) else v) for k, v in batch.items()}


def apply_drop_tensors(batch: dict) -> dict:
    """返回施加了 drop 的特征与可用掩码：{'x_m':…, 'avail_m':…}（valid 不变）。"""
    out = {}
    for m in MODALITIES:
        d = batch[f"drop_{m}"]
        out[f"x_{m}"] = batch[f"x_{m}"].masked_fill(d.unsqueeze(-1), 0.0)
        out[f"avail_{m}"] = batch[f"avail_{m}"] & ~d
        out[f"valid_{m}"] = batch[f"valid_{m}"]
    return out
