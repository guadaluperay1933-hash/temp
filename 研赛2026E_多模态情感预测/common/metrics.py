"""评价指标：题目规定的分类 Accuracy / F1、回归 MAE / Pearson 相关系数，外加 MOSEI 文献常用的补充指标。

主指标（写进论文主表）：
    acc3        三分类（Negative/Neutral/Positive）准确率
    f1_weighted 三分类加权 F1（按类别样本数加权，MOSEI 文献惯例）
    f1_macro    三分类宏平均 F1（类别不均衡时更严格，一并报告）
    mae         情感强度平均绝对误差
    corr        情感强度皮尔逊相关系数
补充指标（便于和文献对比）：
    acc2_has0 / f1_has0   负向 vs 非负向（y<0 vs y>=0）
    acc2_non0 / f1_non0   去掉 y==0 后的负向 vs 正向
    acc7                  强度四舍五入到 {-3..3} 后的 7 类准确率
"""
from __future__ import annotations

import numpy as np
from sklearn.metrics import accuracy_score, confusion_matrix, f1_score

from .data import CLASS_NAMES, NEUTRAL_EPS, reg_to_cls


def pearson(a, b) -> float:
    a = np.asarray(a, np.float64).reshape(-1)
    b = np.asarray(b, np.float64).reshape(-1)
    if a.size < 2 or a.std() < 1e-12 or b.std() < 1e-12:
        return 0.0
    return float(np.corrcoef(a, b)[0, 1])


def classification_metrics(y_true, y_pred, n_classes: int = 3) -> dict:
    y_true = np.asarray(y_true).reshape(-1).astype(int)
    y_pred = np.asarray(y_pred).reshape(-1).astype(int)
    labels = list(range(n_classes))
    per = f1_score(y_true, y_pred, labels=labels, average=None, zero_division=0)
    return {
        "acc3": float(accuracy_score(y_true, y_pred)),
        "f1_weighted": float(f1_score(y_true, y_pred, labels=labels, average="weighted", zero_division=0)),
        "f1_macro": float(f1_score(y_true, y_pred, labels=labels, average="macro", zero_division=0)),
        "f1_per_class": {CLASS_NAMES[i]: float(per[i]) for i in labels},
        "confusion": confusion_matrix(y_true, y_pred, labels=labels).tolist(),  # 行=真实，列=预测
    }


def regression_metrics(y_true, y_pred) -> dict:
    y = np.asarray(y_true, np.float64).reshape(-1)
    p = np.asarray(y_pred, np.float64).reshape(-1)
    out = {"mae": float(np.mean(np.abs(y - p))), "corr": pearson(y, p)}
    # 负向 vs 非负向
    t2, p2 = (y >= 0), (p >= 0)
    out["acc2_has0"] = float((t2 == p2).mean())
    out["f1_has0"] = float(f1_score(t2, p2, average="weighted", zero_division=0))
    # 去零后负向 vs 正向
    nz = np.abs(y) >= NEUTRAL_EPS
    if nz.any():
        out["acc2_non0"] = float(((y[nz] > 0) == (p[nz] > 0)).mean())
        out["f1_non0"] = float(f1_score(y[nz] > 0, p[nz] > 0, average="weighted", zero_division=0))
    else:
        out["acc2_non0"] = out["f1_non0"] = float("nan")
    out["acc7"] = float((np.round(np.clip(y, -3, 3)) == np.round(np.clip(p, -3, 3))).mean())
    return out


def evaluate_all(y_reg_true, y_reg_pred, y_cls_pred=None, y_cls_true=None) -> dict:
    """一次性算全部指标。y_cls_true 缺省时按题面定义由 y_reg_true 推出；
    y_cls_pred 缺省时由回归输出的符号推出（仅作参考）。"""
    y_reg_true = np.asarray(y_reg_true, np.float64).reshape(-1)
    if y_cls_true is None:
        y_cls_true = reg_to_cls(y_reg_true)
    if y_cls_pred is None:
        y_cls_pred = reg_to_cls(y_reg_pred)
    out = {"n": int(y_reg_true.size)}
    out.update(classification_metrics(y_cls_true, y_cls_pred))
    out.update(regression_metrics(y_reg_true, y_reg_pred))
    return out


def composite_score(m: dict) -> float:
    """验证集模型选择用的综合分：兼顾分类与回归（越大越好）。"""
    return float(m["f1_weighted"] + m["acc3"] + m["corr"] - m["mae"])


MAIN_KEYS = ("acc3", "f1_weighted", "f1_macro", "mae", "corr")
EXTRA_KEYS = ("acc2_has0", "f1_has0", "acc2_non0", "f1_non0", "acc7")


def flat_row(m: dict, keys=MAIN_KEYS + EXTRA_KEYS) -> dict:
    return {k: m.get(k) for k in keys}
