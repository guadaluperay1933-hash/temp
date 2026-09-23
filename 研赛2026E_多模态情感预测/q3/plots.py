"""问题3 论文用图（全部中文标注，Noto Sans CJK SC）。"""
from __future__ import annotations

from pathlib import Path

import numpy as np

from common.data import CLASS_CN, CLASS_NAMES, MOD_CN, MODALITIES

from .cards import YH, esc, setup_font
from .records import EXPERT_CN

MOD_COLORS = {"text": "#4C78A8", "audio": "#F58518", "vision": "#54A24B", "interaction": "#B279A2"}
SPLIT_CN = {"valid": "验证集", "test": "测试集", "att4": "附件4"}


def _plt():
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    setup_font(plt)
    return plt


def _save(fig, path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=200, bbox_inches="tight")
    import matplotlib.pyplot as plt

    plt.close(fig)


def plot_training(hist, path, title=""):
    """hist: DataFrame（train_log.csv，可含多个 seed）。"""
    plt = _plt()
    fig, axes = plt.subplots(1, 3, figsize=(15, 3.8))
    for seed, g in hist.groupby("seed"):
        axes[0].plot(g["epoch"], g["train_loss"], label=f"seed {seed}")
        axes[1].plot(g["epoch"], g["val_score"], label=f"seed {seed}")
    axes[0].set_title("训练总损失")
    axes[1].set_title("验证综合分 F1_w + Acc3 + r − MAE")
    for ax in axes[:2]:
        ax.set_xlabel("轮次")
        ax.legend(fontsize=8)
    g = hist[hist["seed"] == hist["seed"].iloc[0]]
    for e in ("text", "audio", "vision", "interaction"):
        col = f"val_gate_{e}"
        if col in g:
            axes[2].plot(g["epoch"], g[col], label=EXPERT_CN[e], color=MOD_COLORS[e])
    axes[2].set_title("验证集平均门控权重（第一个种子）")
    axes[2].set_xlabel("轮次")
    axes[2].legend(fontsize=8)
    fig.suptitle(title)
    fig.tight_layout()
    _save(fig, path)


def plot_share_distribution(bundles: dict, path):
    """bundles: {split_name: bundle}。左：Shapley 份额箱线图；右：主要参考模态占比。"""
    plt = _plt()
    names = list(bundles)
    fig, axes = plt.subplots(1, 2, figsize=(14, 4.6), gridspec_kw={"width_ratios": [1.6, 1]})
    ax = axes[0]
    w = 0.8 / max(len(names), 1)
    for j, n in enumerate(names):
        sh = bundles[n]["share"] * 100
        pos = np.arange(3) + (j - (len(names) - 1) / 2) * w
        bp = ax.boxplot([sh[:, i] for i in range(3)], positions=pos, widths=w * 0.85, patch_artist=True,
                        showfliers=False)
        for patch in bp["boxes"]:
            patch.set_facecolor(plt.cm.Set2(j))
            patch.set_alpha(0.85)
        ax.plot([], [], color=plt.cm.Set2(j), lw=8, label=f"{SPLIT_CN.get(n, n)}（n={len(sh)}）")
    ax.set_xticks(range(3), [MOD_CN[m] for m in MODALITIES])
    ax.set_ylabel("Shapley 作用份额 (%)")
    ax.set_title("逐样本模态作用份额分布")
    ax.legend(fontsize=8)
    ax = axes[1]
    for j, n in enumerate(names):
        main = bundles[n]["main"]
        rate = [100 * (main == i).mean() for i in range(3)]
        pos = np.arange(3) + (j - (len(names) - 1) / 2) * w
        ax.bar(pos, rate, w * 0.9, color=plt.cm.Set2(j), label=SPLIT_CN.get(n, n))
        for p, r, i in zip(pos, rate, range(3)):
            ax.text(p, r + 1, f"{int((main == i).sum())}", ha="center", fontsize=7)
    ax.set_xticks(range(3), [MOD_CN[m] for m in MODALITIES])
    ax.set_ylabel("作为主要参考模态的样本占比 (%)（柱上为条数）")
    ax.set_title("主要参考模态分布")
    ax.legend(fontsize=8)
    fig.tight_layout()
    _save(fig, path)


def plot_position_profile(profiles: dict, path):
    """profiles: {split: explain.position_profile(...)}。每个模态一幅：归一化 |e| 随相对时间的平均分布。"""
    plt = _plt()
    fig, axes = plt.subplots(1, 3, figsize=(15, 3.8), sharey=False)
    for i, m in enumerate(MODALITIES):
        ax = axes[i]
        for j, (n, pr) in enumerate(profiles.items()):
            x = np.asarray(pr["bin_centers"])
            ax.plot(x, pr["all"][m], marker="o", ms=3, label=f"{SPLIT_CN.get(n, n)}·全部", color=plt.cm.Set2(j))
            bm = pr["by_main"][m].get(m)
            if bm is not None:
                ax.plot(x, bm, ls="--", color=plt.cm.Set2(j), label=f"{SPLIT_CN.get(n, n)}·主模态={MOD_CN[m]}")
        ax.axhline(1.0 / len(profiles[next(iter(profiles))]["bin_centers"]), color="gray", lw=0.8, ls=":",
                   label="均匀分布")
        ax.set_title(f"{MOD_CN[m]}：局部重要性随相对时间分布")
        ax.set_xlabel("有效范围内的相对时间 τ")
        ax.set_ylabel("平均归一化 |e| 占比")
        ax.legend(fontsize=7)
    fig.tight_layout()
    _save(fig, path)


def plot_role_by_polarity(b: dict, y_cls_true, path, title=""):
    """按极性分组的模态作用：上排平均份额（堆叠），下排平均带符号 φ。左按预测类别，右按真实类别。"""
    plt = _plt()
    fig, axes = plt.subplots(2, 2, figsize=(12, 7))
    groups = [("预测类别", b["pred_cls"])] + ([("真实类别", np.asarray(y_cls_true))] if y_cls_true is not None else [])
    for col, (gname, cls) in enumerate(groups):
        ax = axes[0, col]
        bottom = np.zeros(3)
        for i, m in enumerate(MODALITIES):
            vals = np.array([b["share"][cls == c, i].mean() * 100 if (cls == c).any() else 0 for c in range(3)])
            ax.bar(range(3), vals, bottom=bottom, color=MOD_COLORS[m], label=MOD_CN[m])
            for c in range(3):
                if vals[c] > 6:
                    ax.text(c, bottom[c] + vals[c] / 2, f"{vals[c]:.0f}%", ha="center", va="center", fontsize=8,
                            color="white")
            bottom += vals
        ax.set_xticks(range(3), [f"{CLASS_CN[n]}\n(n={(cls == c).sum()})" for c, n in enumerate(CLASS_NAMES)])
        ax.set_ylabel("平均 Shapley 份额 (%)")
        ax.set_title(f"按{gname}：三模态平均作用份额")
        ax.legend(fontsize=8, loc="lower right")
        ax = axes[1, col]
        w = 0.25
        for i, m in enumerate(MODALITIES):
            vals = [b["phi"][cls == c, i].mean() if (cls == c).any() else 0 for c in range(3)]
            ax.bar(np.arange(3) + (i - 1) * w, vals, w, color=MOD_COLORS[m], label=MOD_CN[m])
        ax.axhline(0, color="black", lw=0.6)
        ax.set_xticks(range(3), [CLASS_CN[n] for n in CLASS_NAMES])
        ax.set_ylabel("平均带符号 φ_m")
        ax.set_title(f"按{gname}：各模态把强度推向哪个方向")
        ax.legend(fontsize=8)
    if len(groups) == 1:
        axes[0, 1].set_axis_off()
        axes[1, 1].set_axis_off()
    fig.suptitle(title)
    fig.tight_layout()
    _save(fig, path)


def plot_faithfulness(faith: dict, path):
    """上排：逐步删除曲线（证据顺序 vs 随机）；下排：插入曲线。四列 = 文本 / 语音 / 视觉 / 主要参考模态。"""
    plt = _plt()
    curves = faith.get("curves", {})
    targets = [t for t in ("text", "audio", "vision", "main") if t in curves]
    if not targets:
        return
    fig, axes = plt.subplots(2, len(targets), figsize=(4 * len(targets), 7), squeeze=False)
    for j, t in enumerate(targets):
        c = curves[t]
        x = np.asarray(c["fracs"]) * 100
        name = "主要参考模态" if t == "main" else MOD_CN[t]
        ax = axes[0, j]
        ax.plot(x, c["deletion_evidence"], "o-", color="#C0392B", label=f"按贡献删除 AOPC={c['aopc_deletion_evidence']:.3f}")
        ax.plot(x, c["deletion_random"], "s--", color="gray", label=f"随机删除 AOPC={c['aopc_deletion_random']:.3f}")
        ax.set_title(f"{name}：逐步删除")
        ax.set_xlabel("删除的候选位置比例 (%)")
        ax.set_ylabel(f"平均 |Δ{YH}|（越大越忠实）")
        ax.legend(fontsize=7.5)
        ax = axes[1, j]
        ax.plot(x, c["insertion_evidence"], "o-", color="#2166AC", label=f"按贡献保留 AOPC={c['aopc_insertion_evidence']:.3f}")
        ax.plot(x, c["insertion_random"], "s--", color="gray", label=f"随机保留 AOPC={c['aopc_insertion_random']:.3f}")
        ax.set_title(f"{name}：只保留前 f 比例")
        ax.set_xlabel("保留的候选位置比例 (%)")
        ax.set_ylabel(f"与原预测的平均 |Δ{YH}|（越小越忠实）")
        ax.legend(fontsize=7.5)
    fig.tight_layout()
    _save(fig, path)


def plot_comprehensiveness(faith: dict, path):
    """删除证据 vs 删除同样数量随机位置的 |Δŷ|（各模态、主模态、全部），以及只保留证据的充分性。"""
    plt = _plt()
    comp = faith["comprehensiveness"]
    keys = [k for k in ("text", "audio", "vision", "main_modality", "all_modalities") if comp.get(k, {}).get("abs_dy", {}).get("n")]
    lab = {"text": "文本", "audio": "语音", "vision": "视觉", "main_modality": "主要参考模态", "all_modalities": "三模态全部"}
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.2), gridspec_kw={"width_ratios": [2.2, 1]})
    ax = axes[0]
    x = np.arange(len(keys))
    ev = [comp[k]["abs_dy"]["evidence_mean"] for k in keys]
    rd = [comp[k]["abs_dy"]["random_mean"] for k in keys]
    ax.bar(x - 0.2, ev, 0.4, color="#C0392B", label="删除关键证据")
    ax.bar(x + 0.2, rd, 0.4, color="gray", label="删除同样数量随机位置")
    for i, k in enumerate(keys):
        a = comp[k]["abs_dy"]
        ax.text(i, max(ev[i], rd[i]) * 1.03, f"胜率 {a['win_rate']:.2f}\np={a.get('wilcoxon_p_one_sided', float('nan')):.1e}",
                ha="center", fontsize=7.5)
    ax.set_xticks(x, [lab[k] for k in keys])
    ax.set_ylabel(f"平均 |Δ{YH}|")
    ax.set_title("删除检验（comprehensiveness）")
    ax.legend(fontsize=8)
    ax.set_ylim(0, max(ev + rd) * 1.3 + 1e-6)
    ax = axes[1]
    s = faith["sufficiency"]["abs_dy"]
    ax.bar([0, 1], [s["evidence_mean"], s["random_mean"]], color=["#2166AC", "gray"])
    ax.set_xticks([0, 1], ["只保留关键证据", "只保留随机位置"])
    ax.set_ylabel(f"与原预测的平均 |Δ{YH}|（越小越好）")
    ax.set_title(f"保留检验（sufficiency），胜率 {s['win_rate']:.2f}")
    fig.tight_layout()
    _save(fig, path)


def plot_gate_vs_shapley(b: dict, agree: dict, path):
    plt = _plt()
    fig, axes = plt.subplots(1, 4, figsize=(17, 4))
    c = np.abs(b["c"][:, :3])
    c_share = c / np.maximum(c.sum(1, keepdims=True), 1e-12)
    for i, m in enumerate(MODALITIES):
        ax = axes[i]
        ax.scatter(b["share"][:, i] * 100, b["w"][:, i] * 100, s=8, alpha=0.5, color=MOD_COLORS[m])
        ax.set_xlabel("Shapley 份额 (%)")
        ax.set_ylabel("门控权重 w (%)")
        rho = agree.get("spearman_gate_vs_shapley_share", {}).get(m, float("nan"))
        ax.set_title(f"{MOD_CN[m]}：门控 vs Shapley（Spearman {rho:.2f}）", fontsize=10)
        ax.plot([0, 100], [0, 100], color="gray", lw=0.6, ls=":")
    ax = axes[3]
    for i, m in enumerate(MODALITIES):
        ax.scatter(b["share"][:, i] * 100, c_share[:, i] * 100, s=6, alpha=0.4, color=MOD_COLORS[m], label=MOD_CN[m])
    ax.plot([0, 100], [0, 100], color="gray", lw=0.6, ls=":")
    ax.set_xlabel("Shapley 份额 (%)")
    ax.set_ylabel("内置贡献份额 |c_m|/Σ|c| (%)")
    ax.set_title(f"内置贡献 vs Shapley（主模态一致率 {agree.get('main_modality_agreement_contrib_vs_shapley', float('nan')):.2f}）",
                 fontsize=10)
    ax.legend(fontsize=8)
    fig.tight_layout()
    _save(fig, path)


def plot_ig_consistency(ig_raw: dict, ig_cons: dict, path):
    """ig_raw: {m: 逐样本 Spearman 列表}。"""
    plt = _plt()
    fig, axes = plt.subplots(1, 3, figsize=(14, 3.6))
    for i, m in enumerate(MODALITIES):
        ax = axes[i]
        v = np.asarray([x for x in ig_raw.get(m, []) if np.isfinite(x)])
        if v.size:
            ax.hist(v, bins=np.linspace(-1, 1, 21), color=MOD_COLORS[m], alpha=0.85)
            ax.axvline(np.median(v), color="black", ls="--", lw=1, label=f"中位数 {np.median(v):.2f}")
            ax.legend(fontsize=8)
        ov = ig_cons["position_level"][m]["topk_overlap"]
        rnd = ig_cons["position_level"][m]["topk_overlap_random_expectation"]
        ax.set_title(f"{MOD_CN[m]}：内置贡献 vs 积分梯度\ntop-k 重合率 {ov.get('mean', float('nan')):.2f}"
                     f"（随机期望 {rnd.get('mean', float('nan')):.2f}）", fontsize=10)
        ax.set_xlabel("逐样本 Spearman ρ（带符号）")
        ax.set_ylabel("样本数")
    fig.tight_layout()
    _save(fig, path)


def plot_error_covariates(ea: dict, path):
    plt = _plt()
    cov = ea.get("by_covariate", {})
    if not cov:
        return
    fig, axes = plt.subplots(1, len(cov), figsize=(5 * len(cov), 3.8), squeeze=False)
    for ax, (name, d) in zip(axes[0], cov.items()):
        keys = [k for k, v in d.items() if v.get("n", 0) > 0]
        mae = [d[k]["mae"] for k in keys]
        acc = [d[k]["acc3"] for k in keys]
        x = np.arange(len(keys))
        ax.bar(x - 0.2, mae, 0.4, color="#4C78A8", label="MAE")
        ax2 = ax.twinx()
        ax2.bar(x + 0.2, acc, 0.4, color="#F58518", label="Acc3")
        ax2.set_ylim(0, 1)
        for i, k in enumerate(keys):
            ax.text(i, mae[i], f"n={d[k]['n']}", ha="center", va="bottom", fontsize=7)
        ax.set_xticks(x, keys, rotation=15, fontsize=8)
        ax.set_ylabel("MAE")
        ax2.set_ylabel("Acc3")
        ax.set_title(f"按{name}分桶的误差")
        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(h1 + h2, l1 + l2, fontsize=7, loc="upper right")
    fig.tight_layout()
    _save(fig, path)


def plot_ablation(df, path):
    """df: ablation.csv（decision == projected），按变体画 验证/测试 的 Acc3、F1_w、MAE、Corr。"""
    plt = _plt()
    d = df[df["decision"] == "projected"]
    variants = list(dict.fromkeys(d["variant"]))
    metrics = [("acc3", "Acc3"), ("f1_weighted", "F1_w"), ("mae", "MAE"), ("corr", "Corr")]
    fig, axes = plt.subplots(1, 4, figsize=(17, 3.8))
    for ax, (k, name) in zip(axes, metrics):
        for j, split in enumerate(("valid", "test")):
            vals = [float(d[(d["variant"] == v) & (d["split"] == split)][k].iloc[0])
                    if ((d["variant"] == v) & (d["split"] == split)).any() else np.nan for v in variants]
            ax.bar(np.arange(len(variants)) + (j - 0.5) * 0.38, vals, 0.38, label=SPLIT_CN[split],
                   color=["#4C78A8", "#F58518"][j])
        ax.set_xticks(range(len(variants)), variants, rotation=25, fontsize=8)
        ax.set_title(name)
        ax.legend(fontsize=7)
    fig.suptitle("可解释性代价与消融（投影后指标）")
    fig.tight_layout()
    _save(fig, path)


def plot_att4_overview(records: list[dict], path):
    """附件4：每条样本的三模态 Shapley 份额（堆叠横条）+ 预测强度。"""
    plt = _plt()
    n = len(records)
    fig, axes = plt.subplots(1, 2, figsize=(14, 0.42 * n + 1.8), gridspec_kw={"width_ratios": [2, 1]})
    ax = axes[0]
    left = np.zeros(n)
    for m in MODALITIES:
        v = np.array([100 * r["share"][m] for r in records])
        ax.barh(range(n), v, left=left, color=MOD_COLORS[m], label=MOD_CN[m])
        for i in range(n):
            if v[i] > 8:
                ax.text(left[i] + v[i] / 2, i, f"{v[i]:.0f}", ha="center", va="center", fontsize=7, color="white")
        left += v
    ax.set_yticks(range(n), [f"{esc(r['id'])}（主：{MOD_CN[r['main_modality']]}）" for r in records], fontsize=7)
    ax.invert_yaxis()
    ax.set_xlabel("Shapley 作用份额 (%)")
    ax.set_title("附件4：逐样本模态作用份额")
    ax.legend(fontsize=8, loc="lower right")
    ax = axes[1]
    y = [r["pred_intensity"] for r in records]
    ax.barh(range(n), y, color=["#2166AC" if v > 0 else ("#B2182B" if v < 0 else "gray") for v in y])
    ax.axvline(0, color="black", lw=0.6)
    ax.set_yticks(range(n), [""] * n)
    ax.invert_yaxis()
    ax.set_xlim(-3, 3)
    ax.set_xlabel("预测情感强度")
    ax.set_title("附件4：预测强度（颜色 = 极性）")
    fig.tight_layout()
    _save(fig, path)
