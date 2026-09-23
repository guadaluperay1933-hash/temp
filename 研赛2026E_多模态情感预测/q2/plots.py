"""问题2 作图（matplotlib，Agg 后端，中文字体沿用 common.error_analysis 的查找逻辑）。

配色：类别色按固定顺序分配（同一实体在所有图里同色）；极性用 红(负)/灰(中)/蓝(正) 发散配色。
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from common.data import MOD_CN, MODALITIES  # noqa: E402
from common.error_analysis import _setup_cjk_font  # noqa: E402

_setup_cjk_font(plt)
plt.rcParams.update({"axes.grid": True, "grid.alpha": 0.25, "grid.linewidth": 0.6, "axes.spines.top": False,
                     "axes.spines.right": False, "lines.linewidth": 1.8, "lines.markersize": 5,
                     "legend.frameon": False, "figure.dpi": 100})

PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*"]
COMBO_ORDER = ["T", "A", "V", "TA", "TV", "AV", "TAV"]
COMBO_CN = {"T": "文本", "A": "语音", "V": "视觉", "TA": "文本+语音", "TV": "文本+视觉", "AV": "语音+视觉",
            "TAV": "三模态"}
MODEL_ORDER = ["full", "no_aug", "naive_impute", "no_rec", "no_kd", "no_gate", "late_fusion",
               "uni_text", "uni_audio", "uni_vision"]
MODEL_CN = {"full": "MRG-Net", "no_aug": "w/o 增广", "naive_impute": "线性插补", "no_rec": "w/o 重构",
            "no_kd": "w/o 蒸馏", "no_gate": "w/o 门控", "late_fusion": "晚期融合", "uni_text": "仅文本",
            "uni_audio": "仅语音", "uni_vision": "仅视觉"}
POL_COLORS = {"Negative": "#e34948", "Neutral": "#9a9892", "Positive": "#2a78d6"}
METRIC_CN = {"mae": "MAE（越低越好）", "f1_weighted": "加权 F1", "corr": "Pearson 相关", "acc3": "三分类准确率",
             "f1_macro": "宏 F1"}


def _color(key: str, order: list[str]) -> tuple[str, str]:
    i = order.index(key) if key in order else len(order)
    return PALETTE[i % len(PALETTE)], MARKERS[i % len(MARKERS)]


def _save(fig, path) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(path, dpi=200)
    plt.close(fig)


def plot_train_curves(log: pd.DataFrame, path, title: str = "") -> None:
    fig, axes = plt.subplots(1, 3, figsize=(15, 4))
    for i, (seed, g) in enumerate(log.groupby("seed")):
        c = PALETTE[i % len(PALETTE)]
        axes[0].plot(g["epoch"], g["train_loss"], color=c, marker="o", label=f"种子 {seed}")
        axes[1].plot(g["epoch"], g["val_score_complete"], color=c, marker="o", label=f"种子 {seed} 完整")
        axes[1].plot(g["epoch"], g["val_score_mixed"], color=c, marker="s", ls="--", label=f"种子 {seed} 混合缺失")
        axes[2].plot(g["epoch"], g["val_complete_mae"], color=c, marker="o", label=f"种子 {seed} 完整")
        axes[2].plot(g["epoch"], g["val_mixed_mae"], color=c, marker="s", ls="--", label=f"种子 {seed} 混合缺失")
    axes[0].set_title("训练总损失")
    axes[1].set_title("验证综合分 F1w+Acc+r−MAE")
    axes[2].set_title("验证 MAE")
    for ax in axes:
        ax.set_xlabel("epoch")
        ax.legend(fontsize=7)
    fig.suptitle(title)
    _save(fig, path)


def plot_type_ratio(df: pd.DataFrame, model: str, split: str, path, metrics=("mae", "f1_weighted", "corr")) -> None:
    d = df[(df["model"] == model) & (df["split"] == split)]
    if d.empty:
        return
    fig, axes = plt.subplots(1, len(metrics), figsize=(5 * len(metrics), 4))
    for ax, met in zip(np.atleast_1d(axes), metrics):
        for combo in [c for c in COMBO_ORDER if c in set(d["combo"])]:
            g = d[d["combo"] == combo].sort_values("ratio")
            c, mk = _color(combo, COMBO_ORDER)
            ax.plot(g["ratio"], g[met], color=c, marker=mk, label=f"缺失 {COMBO_CN[combo]}")
        ax.set_xlabel("缺失率 r")
        ax.set_title(METRIC_CN.get(met, met))
    np.atleast_1d(axes)[-1].legend(fontsize=8, loc="best")
    fig.suptitle(f"{MODEL_CN.get(model, model)}：缺失模态类型 × 缺失率（{split}，单段随机位置，掩码种子平均）")
    _save(fig, path)


def plot_models_vs_ratio(df: pd.DataFrame, split: str, combos, path, metrics=("mae", "f1_weighted", "corr")) -> None:
    d = df[df["split"] == split]
    combos = [c for c in combos if c in set(d["combo"])]
    if d.empty or not combos:
        return
    fig, axes = plt.subplots(len(combos), len(metrics), figsize=(5 * len(metrics), 3.6 * len(combos)), squeeze=False)
    models = [m for m in MODEL_ORDER if m in set(d["model"])]
    for i, combo in enumerate(combos):
        for j, met in enumerate(metrics):
            ax = axes[i, j]
            for mdl in models:
                g = d[(d["combo"] == combo) & (d["model"] == mdl)].sort_values("ratio")
                c, mk = _color(mdl, MODEL_ORDER)
                ax.plot(g["ratio"], g[met], color=c, marker=mk, lw=2.4 if mdl == "full" else 1.4,
                        label=MODEL_CN.get(mdl, mdl))
            ax.set_title(f"缺失 {COMBO_CN[combo]}：{METRIC_CN.get(met, met)}")
            ax.set_xlabel("缺失率 r")
    axes[0, -1].legend(fontsize=8)
    fig.suptitle(f"完整模型与消融/基线随缺失率的退化（{split}，相同掩码）")
    _save(fig, path)


def plot_position(df: pd.DataFrame, model: str, split: str, path, metrics=("mae", "f1_weighted")) -> None:
    d = df[(df["model"] == model) & (df["split"] == split)]
    if d.empty:
        return
    groups = [(c, r) for c in COMBO_ORDER for r in sorted(set(d["ratio"])) if ((d["combo"] == c) & (d["ratio"] == r)).any()]
    fig, axes = plt.subplots(1, len(metrics), figsize=(6.5 * len(metrics), 4))
    pos_order = ["head", "middle", "tail"]
    pos_cn = {"head": "开头", "middle": "中间", "tail": "结尾"}
    w = 0.26
    for ax, met in zip(np.atleast_1d(axes), metrics):
        x = np.arange(len(groups))
        for k, p in enumerate(pos_order):
            vals = [d[(d["combo"] == c) & (d["ratio"] == r) & (d["position"] == p)][met].mean() for c, r in groups]
            ax.bar(x + (k - 1) * w, vals, width=w - 0.03, color=PALETTE[k], label=pos_cn[p])
        ax.set_xticks(x, [f"{COMBO_CN[c]}\nr={r:g}" for c, r in groups], fontsize=8)
        ax.set_title(METRIC_CN.get(met, met))
        lo = np.nanmin(d[met]) if len(d) else 0
        hi = np.nanmax(d[met]) if len(d) else 1
        ax.set_ylim(max(0.0, lo - 0.15 * (hi - lo + 1e-6)), hi + 0.1 * (hi - lo + 1e-6))
    np.atleast_1d(axes)[0].legend(fontsize=8)
    fig.suptitle(f"{MODEL_CN.get(model, model)}：缺失位置的影响（{split}）")
    _save(fig, path)


def plot_fragmentation(df: pd.DataFrame, model: str, split: str, path, metrics=("mae", "f1_weighted")) -> None:
    d = df[(df["model"] == model) & (df["split"] == split)]
    if d.empty:
        return
    fig, axes = plt.subplots(1, len(metrics), figsize=(6 * len(metrics), 4))
    styles = ["-", "--", ":", "-."]
    for ax, met in zip(np.atleast_1d(axes), metrics):
        for combo in [c for c in COMBO_ORDER if c in set(d["combo"])]:
            c, mk = _color(combo, COMBO_ORDER)
            for k, r in enumerate(sorted(set(d["ratio"]))):
                g = d[(d["combo"] == combo) & (d["ratio"] == r)].sort_values("n_spans")
                ax.plot(g["n_spans"], g[met], color=c, marker=mk, ls=styles[k % 4],
                        label=f"{COMBO_CN[combo]} r={r:g}")
        ax.set_xlabel("缺失段数（总缺失率固定）")
        ax.set_title(METRIC_CN.get(met, met))
    np.atleast_1d(axes)[-1].legend(fontsize=7, ncol=2)
    fig.suptitle(f"{MODEL_CN.get(model, model)}：缺失时长的碎片化程度（{split}）")
    _save(fig, path)


def plot_gate_alpha(df: pd.DataFrame, split: str, path, model: str = "full") -> None:
    d = df[(df["model"] == model) & (df["split"] == split)]
    if d.empty or "alpha_text" not in d or d["alpha_text"].isna().all():
        return
    fig, axes = plt.subplots(1, 3, figsize=(15, 4), sharey=True)
    for ax, (combo, m) in zip(axes, [("T", "text"), ("A", "audio"), ("V", "vision")]):
        g = d[d["combo"] == combo].sort_values("ratio")
        for k, mm in enumerate(MODALITIES):
            ax.plot(g["ratio"], g[f"alpha_{mm}"], color=PALETTE[k], marker=MARKERS[k], lw=2.4 if mm == m else 1.4,
                    label=f"α_{MOD_CN[mm]}")
        ax.set_title(f"仅{MOD_CN[m]}缺失时的门控权重")
        ax.set_xlabel(f"{MOD_CN[m]}缺失率 r")
    axes[0].set_ylabel("平均门控权重 α")
    axes[-1].legend(fontsize=8)
    fig.suptitle(f"可靠性门控随模态受损程度的迁移（{split}）")
    _save(fig, path)


def plot_degradation_fit(df: pd.DataFrame, fits: dict, model: str, split: str, path, metric: str = "mae") -> None:
    d = df[(df["model"] == model) & (df["split"] == split)]
    if d.empty:
        return
    fig, ax = plt.subplots(figsize=(7, 4.5))
    rr = np.linspace(0, 1, 101)
    for combo in [c for c in COMBO_ORDER if c in set(d["combo"])]:
        g = d[d["combo"] == combo].sort_values("ratio")
        base = g[g["ratio"] == 0][metric].mean()
        c, mk = _color(combo, COMBO_ORDER)
        ax.plot(g["ratio"], g[metric] - base, ls="none", color=c, marker=mk)
        f = fits.get(model, {}).get(split, {}).get(combo, {}).get("power_law", {}).get(metric)
        if f and f.get("beta") is not None:
            ax.plot(rr, f["beta"] * rr ** f["gamma"], color=c, lw=1.4,
                    label=f"{COMBO_CN[combo]}: β={f['beta']:.3f}, γ={f['gamma']:.2f}")
    ax.axhline(0, color="#9a9892", lw=0.8)
    ax.set_xlabel("缺失率 r")
    ax.set_ylabel(f"Δ{metric.upper()}(r) = {metric.upper()}(r) − {metric.upper()}(0)")
    ax.legend(fontsize=8)
    ax.set_title(f"{MODEL_CN.get(model, model)}：退化律拟合 ΔMAE = β·r^γ（{split}）")
    _save(fig, path)


def plot_ablation(abl: pd.DataFrame, path, split: str = "test", stage: str = "projected") -> None:
    d = abl[(abl["split"] == split) & (abl["decision"] == stage)]
    if d.empty:
        return
    models = [m for m in MODEL_ORDER if m in set(d["variant"])]
    fig, axes = plt.subplots(1, 3, figsize=(16, 4.2))
    x = np.arange(len(models))
    w = 0.38
    for ax, met in zip(axes, ("f1_weighted", "mae", "corr")):
        for k, (proto, cn) in enumerate((("complete", "完整输入"), ("mixed", "混合缺失"))):
            g = d[d["protocol"] == proto].set_index("variant")
            vals = [g.loc[m, met] if m in g.index else np.nan for m in models]
            err = [g.loc[m, f"{met}_std"] if m in g.index and f"{met}_std" in g else 0 for m in models]
            ax.bar(x + (k - 0.5) * w, vals, width=w - 0.04, yerr=err, color=PALETTE[k], label=cn, capsize=2)
        ax.set_xticks(x, [MODEL_CN.get(m, m) for m in models], rotation=30, fontsize=8)
        ax.set_title(METRIC_CN.get(met, met))
        vals = d[met].to_numpy(dtype=float)
        if np.isfinite(vals).any():
            lo, hi = np.nanmin(vals), np.nanmax(vals)
            ax.set_ylim(max(0.0, lo - 0.2 * (hi - lo + 1e-6)) if met != "corr" else min(0.0, lo), hi + 0.12 * (hi - lo + 1e-6))
    axes[0].legend(fontsize=8)
    fig.suptitle(f"消融与基线（{split}，投影后输出；误差棒 = 种子标准差）")
    _save(fig, path)


def plot_error_covariates(ea: dict, path, title: str = "") -> None:
    cov = ea.get("by_covariate", {})
    keys = [k for k in cov if k.startswith("miss_ratio_")]
    if not keys:
        return
    fig, axes = plt.subplots(1, len(keys) + 1, figsize=(4.6 * (len(keys) + 1), 4))
    for ax, k in zip(axes, keys + ["missing_combo"]):
        d = cov.get(k, {})
        names = [n for n, v in d.items() if v.get("n", 0) > 0]
        if k == "missing_combo":
            names = sorted(names, key=lambda n: (n != "none", len(n), n))
        else:
            order = ["0", "(0,0.3]", "(0.3,0.6]", ">0.6"]
            names = sorted(names, key=lambda n: order.index(n) if n in order else len(order))
        ax.bar(range(len(names)), [d[n]["mae"] for n in names], color=PALETTE[0], width=0.6, label="MAE")
        # 不用双轴：准确率直接标在柱顶
        for i, n in enumerate(names):
            ax.text(i, d[n]["mae"], f"Acc {d[n]['acc3']:.2f}\nn={d[n]['n']}", ha="center", va="bottom", fontsize=7)
        ax.set_xticks(range(len(names)), names, rotation=25, fontsize=8)
        ax.set_ylabel("MAE")
        ax.set_title({"missing_combo": "按缺失模态组合"}.get(k, f"按{MOD_CN.get(k.replace('miss_ratio_', ''), k)}缺失率"))
    fig.suptitle(title)
    _save(fig, path)


def plot_scatter(panels: list[tuple[str, np.ndarray, np.ndarray, np.ndarray]], path, title: str = "") -> None:
    """panels: [(标题, y_true, y_pred, c_pred)]；点按预测极性着色。"""
    fig, axes = plt.subplots(1, len(panels), figsize=(5 * len(panels), 4.6), squeeze=False)
    names = ["Negative", "Neutral", "Positive"]
    for ax, (t, y, p, c) in zip(axes[0], panels):
        for k, n in enumerate(names):
            s = c == k
            ax.scatter(y[s], p[s], s=10, alpha=0.6, color=POL_COLORS[n], label=f"预测{['负向', '中性', '正向'][k]}")
        ax.plot([-3, 3], [-3, 3], color="#52514e", lw=0.8)
        ax.set_xlim(-3.2, 3.2)
        ax.set_ylim(-3.2, 3.2)
        ax.set_xlabel("真实强度")
        ax.set_ylabel("预测强度（投影后）")
        ax.set_title(t)
    axes[0, -1].legend(fontsize=8)
    fig.suptitle(title)
    _save(fig, path)
