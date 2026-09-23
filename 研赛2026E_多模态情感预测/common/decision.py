"""决策层：在验证集上选分类阈值（中性类 logit 偏置），并把强度投影成与极性自洽的输出。

题面："在验证集上选择模型结构、超参数与决策阈值"，"0 仅属于中性"。
    极性 ĉ = argmax_k (ℓ_k + b_k)，b = (0, β, 0)，β 在验证集上网格搜索使加权 F1 最大；
    强度一致性投影 P(ŷ, ĉ)：Negative → min(ŷ, -ε)；Neutral → 0；Positive → max(ŷ, ε)。
"""
from __future__ import annotations

import numpy as np

from .metrics import evaluate_all

PROJ_EPS = 0.01


def project_intensity(pred_cls, y_raw, eps: float = PROJ_EPS) -> np.ndarray:
    c = np.asarray(pred_cls).reshape(-1)
    y = np.asarray(y_raw, np.float64).reshape(-1).copy()
    y = np.clip(y, -3.0, 3.0)
    y[c == 0] = np.minimum(y[c == 0], -eps)
    y[c == 1] = 0.0
    y[c == 2] = np.maximum(y[c == 2], eps)
    return y


def apply_bias(logits, bias) -> np.ndarray:
    return np.argmax(np.asarray(logits) + np.asarray(bias)[None, :], axis=1)


def tune_neutral_bias(logits, y_cls_true, grid=None, metric: str = "f1_weighted") -> tuple[np.ndarray, dict]:
    """网格搜索中性类偏置 β。返回 (bias(3,), 搜索记录)。并列时取 |β| 最小者（最保守）。"""
    from sklearn.metrics import f1_score

    grid = np.round(np.linspace(-2.0, 2.0, 41), 3) if grid is None else np.asarray(grid)
    y = np.asarray(y_cls_true).reshape(-1)
    records = []
    for b in grid:
        pred = apply_bias(logits, (0.0, b, 0.0))
        if metric == "f1_macro":
            s = f1_score(y, pred, average="macro", zero_division=0)
        elif metric == "acc3":
            s = float((pred == y).mean())
        else:
            s = f1_score(y, pred, average="weighted", zero_division=0)
        records.append((float(b), float(s)))
    best = max(records, key=lambda r: (round(r[1], 6), -abs(r[0])))
    return np.array([0.0, best[0], 0.0], np.float32), {"metric": metric, "grid": records, "best_bias": best[0],
                                                          "best_score": best[1]}


def decide(logits, y_raw, bias=(0.0, 0.0, 0.0), project: bool = True) -> tuple[np.ndarray, np.ndarray]:
    c = apply_bias(logits, bias)
    y = project_intensity(c, y_raw) if project else np.clip(np.asarray(y_raw, np.float64), -3, 3)
    return c, y


def evaluate_decisions(y_true, logits, y_raw, bias) -> dict:
    """同时报告：投影前（分类头 + 原始回归）与投影后（自洽输出）两套指标，以及不自洽比例。"""
    c = apply_bias(logits, bias)
    y_raw = np.clip(np.asarray(y_raw, np.float64), -3, 3)
    before = evaluate_all(y_true, y_raw, c)
    after = evaluate_all(y_true, project_intensity(c, y_raw), c)
    sign_cls = np.where(y_raw < 0, 0, 2)
    inconsistent = float(((c != 1) & (c != sign_cls)).mean() + ((c == 1) & (np.abs(y_raw) >= 0.5)).mean())
    return {"before_projection": before, "after_projection": after, "inconsistency_rate_before": inconsistent}
