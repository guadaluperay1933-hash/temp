"""错误归因：把验证/测试集上的错误按可解释的因素分桶统计，问题2、问题3 共用。

归因维度：
    1. 混淆矩阵与各类召回（哪类最容易错成哪类）
    2. 真实强度分桶（强负/弱负/中性/弱正/强正）上的 MAE 与准确率 —— 弱情感、中性附近最难
    3. 文本长度分桶（词数）—— 极短片段信息不足
    4. 语言现象：否定词、转折词（not / never / but / however …）—— 文本语义反转
    5. 可选协变量（如缺失率、主要参考模态）分桶
    6. 误差最大的样本清单（附原文），供论文举例
"""
from __future__ import annotations

import re

import numpy as np

from .data import CLASS_NAMES, reg_to_cls

NEGATION = {"not", "no", "never", "nothing", "none", "nobody", "neither", "nor", "cannot", "without"}
CONTRAST = {"but", "however", "although", "though", "yet", "except", "whereas", "unless"}
_TOKEN = re.compile(r"[a-z']+")

STRENGTH_BINS = [(-3.01, -1.5, "强负[-3,-1.5)"), (-1.5, -1e-6, "弱负[-1.5,0)"), (-1e-6, 1e-6, "中性0"),
                 (1e-6, 1.5, "弱正(0,1.5)"), (1.5, 3.01, "强正[1.5,3]")]


def _words(t: str) -> list[str]:
    return _TOKEN.findall((t or "").lower())


def _has_negation(ws) -> bool:
    return any(w in NEGATION or w.endswith("n't") for w in ws)


def _bucket_stats(mask, y_true, y_pred, c_true, c_pred) -> dict:
    n = int(mask.sum())
    if n == 0:
        return {"n": 0}
    return {"n": n, "mae": float(np.abs(y_true[mask] - y_pred[mask]).mean()),
            "acc3": float((c_true[mask] == c_pred[mask]).mean())}


def error_attribution(ids, raw_text, y_true, y_pred, c_pred, covariates: dict | None = None,
                      top_k: int = 20) -> dict:
    y_true = np.asarray(y_true, np.float64)
    y_pred = np.asarray(y_pred, np.float64)
    c_true = reg_to_cls(y_true)
    c_pred = np.asarray(c_pred).astype(int)
    out: dict = {"n": int(len(y_true))}

    cm = np.zeros((3, 3), int)
    for a, b in zip(c_true, c_pred):
        cm[a, b] += 1
    out["confusion"] = cm.tolist()
    out["recall"] = {CLASS_NAMES[k]: float(cm[k, k] / max(cm[k].sum(), 1)) for k in range(3)}
    off = [(CLASS_NAMES[i], CLASS_NAMES[j], int(cm[i, j])) for i in range(3) for j in range(3) if i != j]
    out["top_confusions"] = sorted(off, key=lambda t: -t[2])[:3]

    out["by_strength"] = {name: _bucket_stats((y_true >= lo) & (y_true < hi), y_true, y_pred, c_true, c_pred)
                          for lo, hi, name in STRENGTH_BINS}

    texts = list(raw_text) if raw_text is not None else [""] * len(y_true)
    words = [_words(t) for t in texts]
    nw = np.array([len(w) for w in words])
    if raw_text is not None:
        edges = [0, 5, 10, 20, 40, 10 ** 6]
        out["by_text_length"] = {f"{edges[i]}-{edges[i + 1] - 1}词" if edges[i + 1] < 10 ** 6 else f">={edges[i]}词":
                                 _bucket_stats((nw >= edges[i]) & (nw < edges[i + 1]), y_true, y_pred, c_true, c_pred)
                                 for i in range(len(edges) - 1)}
        neg = np.array([_has_negation(w) for w in words])
        con = np.array([any(x in CONTRAST for x in w) for w in words])
        out["by_linguistic"] = {"含否定词": _bucket_stats(neg, y_true, y_pred, c_true, c_pred),
                                "含转折词": _bucket_stats(con, y_true, y_pred, c_true, c_pred),
                                "都不含": _bucket_stats(~neg & ~con, y_true, y_pred, c_true, c_pred)}

    if covariates:
        out["by_covariate"] = {}
        for name, vals in covariates.items():
            vals = np.asarray(vals)
            if vals.dtype.kind in "fi" and len(np.unique(vals)) > 6:
                qs = np.unique(np.quantile(vals, [0, 0.25, 0.5, 0.75, 1.0]))
                res = {}
                for i in range(len(qs) - 1):
                    lo, hi = qs[i], qs[i + 1]
                    m = (vals >= lo) & ((vals < hi) if i < len(qs) - 2 else (vals <= hi))
                    res[f"[{lo:.3g},{hi:.3g}{']' if i == len(qs) - 2 else ')'}"] = _bucket_stats(m, y_true, y_pred, c_true, c_pred)
            else:
                res = {str(v): _bucket_stats(vals == v, y_true, y_pred, c_true, c_pred) for v in np.unique(vals)}
            out["by_covariate"][name] = res

    err = np.abs(y_true - y_pred)
    order = np.argsort(-err)[:top_k]
    out["worst_samples"] = [{"id": ids[i], "text": texts[i], "y_true": float(y_true[i]), "y_pred": float(y_pred[i]),
                             "cls_true": CLASS_NAMES[c_true[i]], "cls_pred": CLASS_NAMES[c_pred[i]],
                             "abs_err": float(err[i])} for i in order]
    return out


def plot_error_analysis(ea: dict, path, title: str = "") -> None:
    """一张 2×2 图：混淆矩阵、强度分桶 MAE、文本长度分桶 MAE、语言现象分桶 MAE。"""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    _setup_cjk_font(plt)
    fig, axes = plt.subplots(2, 2, figsize=(11, 8))
    cm = np.array(ea["confusion"])
    ax = axes[0, 0]
    ax.imshow(cm, cmap="Blues")
    for i in range(3):
        for j in range(3):
            ax.text(j, i, str(cm[i, j]), ha="center", va="center",
                    color="white" if cm[i, j] > cm.max() * 0.6 else "black")
    ax.set_xticks(range(3), ["负向", "中性", "正向"])
    ax.set_yticks(range(3), ["负向", "中性", "正向"])
    ax.set_xlabel("预测")
    ax.set_ylabel("真实")
    ax.set_title("混淆矩阵")

    def _bar(ax, d, t):
        keys = [k for k, v in d.items() if v.get("n", 0) > 0]
        ax.bar(range(len(keys)), [d[k]["mae"] for k in keys], color="#4C78A8")
        for i, k in enumerate(keys):
            ax.text(i, d[k]["mae"], f"n={d[k]['n']}", ha="center", va="bottom", fontsize=8)
        ax.set_xticks(range(len(keys)), keys, rotation=20, fontsize=8)
        ax.set_ylabel("MAE")
        ax.set_title(t)

    _bar(axes[0, 1], ea["by_strength"], "按真实情感强度分桶")
    if "by_text_length" in ea:
        _bar(axes[1, 0], ea["by_text_length"], "按文本长度分桶")
    if "by_linguistic" in ea:
        _bar(axes[1, 1], ea["by_linguistic"], "按语言现象分桶")
    fig.suptitle(title)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def _setup_cjk_font(plt) -> None:
    """尽量找一个能显示中文的字体；找不到也不报错（图里中文会显示成方框，但数值正确）。"""
    from matplotlib import font_manager

    cands = ["Noto Sans CJK SC", "Source Han Sans SC", "SimHei", "Microsoft YaHei", "WenQuanYi Zen Hei",
             "PingFang SC", "Heiti SC", "Arial Unicode MS"]
    have = {f.name for f in font_manager.fontManager.ttflist}
    for c in cands:
        if c in have:
            plt.rcParams["font.sans-serif"] = [c] + plt.rcParams["font.sans-serif"]
            break
    plt.rcParams["axes.unicode_minus"] = False
