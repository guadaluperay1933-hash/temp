"""问题2 附件3 推理：读入带随机局部缺失的无标签特征文件 → 极性 + 强度（自洽）→ CSV / 汇总 JSON / 图。

    python -m q2.predict --run outputs/q2/mrgnet_aligned --input data/附件3/aligned_50.pkl
    python -m q2.predict --model outputs/q2/mrgnet_aligned/model.pt --input <pkl> --split auto --out <dir>

* 附件3 的缺失本身就是"名义有效范围内的全零行"，读入后自动成为 avail=False 的位置，模型无需另给掩码。
* 标准化器、决策偏置、特征版本全部取自 model.pt；文件版本与模型版本不一致时报错（题面要求同一版本）。
* 输出 CSV 用 UTF-8-SIG 编码（Excel 直接打开中文不乱码）。
"""
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import CLASS_CN, CLASS_NAMES, MOD_CN, MOD_SHORT, MODALITIES, Normalizer, get_split, list_splits, \
    load_pkl, split_id
from common.decision import decide
from common.metrics import evaluate_all
from common.missing import missing_statistics
from common.utils import dump_json, get_device, get_logger

from . import plots
from .config import resolve_path
from .engine import load_checkpoint, make_batches, predict_batches

CSV_NAME = "附件3_预测结果.csv"
SUMMARY_NAME = "附件3_summary.json"
CSV_COLUMNS = ["index", "id", "video_id", "clip_id", "pred_polarity", "pred_polarity_cn", "pred_class",
               "pred_intensity", "raw_intensity", "prob_negative", "prob_neutral", "prob_positive",
               "miss_ratio_text", "miss_ratio_audio", "miss_ratio_vision", "missing_modalities"]


def pick_split(raw, split: str = "auto") -> str:
    splits = list_splits(raw)
    if split and split != "auto":
        if split not in splits:
            raise ValueError(f"文件中没有划分 {split}，可选 {splits}")
        return split
    if len(splits) == 1:
        return splits[0]
    if "test" in splits:
        return "test"
    raise ValueError(f"文件包含多个划分 {splits}，请用 --split 指定")


def _stats(x: np.ndarray) -> dict:
    x = np.asarray(x, np.float64)
    if x.size == 0:
        return {"n": 0}
    q = np.quantile(x, [0.05, 0.25, 0.5, 0.75, 0.95])
    return {"n": int(x.size), "mean": float(x.mean()), "std": float(x.std()), "min": float(x.min()),
            "q05": float(q[0]), "q25": float(q[1]), "median": float(q[2]), "q75": float(q[3]), "q95": float(q[4]),
            "max": float(x.max())}


def _hist(x: np.ndarray, edges) -> dict:
    c, e = np.histogram(np.asarray(x, np.float64), bins=edges)
    return {"edges": [float(v) for v in e], "counts": c.astype(int).tolist()}


def _position_label(center_rel: float) -> str:
    if not np.isfinite(center_rel):
        return "none"
    return "head" if center_rel < 1 / 3 else ("tail" if center_rel > 2 / 3 else "middle")


def missing_summary(rows: list[dict], n: int) -> dict:
    """由 common.missing.missing_statistics 的逐样本×模态行，汇总缺失率 / 段数 / 位置分布与组合频次。"""
    df = pd.DataFrame(rows)
    out = {"definition": "缺失 = 名义有效范围内特征整行为零的位置（附件3 人为置零 + 原始天然零行，二者无法区分）；"
                         "位置按缺失中心在有效范围内的相对位置 center_rel 三等分：<1/3 开头，>2/3 结尾，否则中间"}
    per = {}
    for m in MODALITIES:
        d = df[df["modality"] == m]
        has = d[d["n_missing"] > 0]
        pos = has["center_rel"].map(_position_label).value_counts().to_dict()
        per[m] = {"samples_with_missing": int(len(has)), "fraction_samples_with_missing": float(len(has) / max(n, 1)),
                  "ratio_all_samples": _stats(d["ratio"].to_numpy()),
                  "ratio_among_missing": _stats(has["ratio"].to_numpy()),
                  "ratio_hist_among_missing": _hist(has["ratio"].to_numpy(), np.linspace(0, 1, 11)),
                  "n_spans_counts": {str(k): int(v) for k, v in has["n_spans"].value_counts().sort_index().items()},
                  "position_counts": {k: int(pos.get(k, 0)) for k in ("head", "middle", "tail")},
                  "center_rel": _stats(has["center_rel"].to_numpy()),
                  "first_rel_start": _stats(has["first_rel_start"].to_numpy()),
                  "valid_len": _stats(d["valid_len"].to_numpy())}
    out["per_modality"] = per
    return out


def run_predict(model_path, input_path, split="auto", out_dir=None, device=None, batch_size=256, logger=None) -> dict:
    device = device or get_device()
    model, ck = load_checkpoint(model_path, device)
    raw = load_pkl(input_path)
    s = pick_split(raw, split)
    sd = get_split(raw, s, name=f"附件3:{s}")
    if sd.version != ck["version"]:
        raise ValueError(f"附件3 文件是 {sd.version} 版本，模型是 {ck['version']} 版本；题面要求同一特征版本")
    for m in MODALITIES:
        if sd.feats[m].shape[-1] != ck["feat_dims"][m]:
            raise ValueError(f"{m} 维度 {sd.feats[m].shape[-1]} 与模型 {ck['feat_dims'][m]} 不一致")
    norm = Normalizer.from_state_dict(ck["normalizer"])
    sdn = norm.transform(sd)
    pred = predict_batches(model, make_batches(sdn, None, batch_size), device)
    bias = np.asarray(ck["bias"], np.float32)
    c, y = decide(pred["logits"], pred["y_raw"], bias, project=True)
    z = pred["logits"] + bias[None, :]
    prob = np.exp(z - z.max(1, keepdims=True))
    prob /= prob.sum(1, keepdims=True)

    rows = missing_statistics(sd)  # 行顺序：样本优先、模态其次（与 sd.ids 一一对应，id 重复也不会错位）
    ratio = {m: np.zeros(len(sd)) for m in MODALITIES}
    for k, r in enumerate(rows):
        ratio[r["modality"]][k // len(MODALITIES)] = r["ratio"]
    combo = ["+".join(MOD_SHORT[m] for m in MODALITIES if ratio[m][i] > 0) or "none" for i in range(len(sd))]

    vc = [split_id(i) for i in sd.ids]
    df = pd.DataFrame({
        "index": np.arange(len(sd)), "id": sd.ids, "video_id": [v for v, _ in vc], "clip_id": [k for _, k in vc],
        "pred_polarity": [CLASS_NAMES[k] for k in c], "pred_polarity_cn": [CLASS_CN[CLASS_NAMES[k]] for k in c],
        "pred_class": c.astype(int), "pred_intensity": np.round(y, 4), "raw_intensity": np.round(np.clip(pred["y_raw"], -3, 3), 4),
        "prob_negative": np.round(prob[:, 0], 4), "prob_neutral": np.round(prob[:, 1], 4),
        "prob_positive": np.round(prob[:, 2], 4),
        **{f"miss_ratio_{m}": np.round(ratio[m], 4) for m in MODALITIES}, "missing_modalities": combo})[CSV_COLUMNS]

    out_dir = Path(out_dir or Path(model_path).parent)
    out_dir.mkdir(parents=True, exist_ok=True)
    df.to_csv(out_dir / CSV_NAME, index=False, encoding="utf-8-sig")
    pd.DataFrame(rows).to_csv(out_dir / "附件3_缺失检测明细.csv", index=False, encoding="utf-8-sig")

    # ---- 自洽性核查
    sign_ok = bool(np.all((c != 0) | (y < 0)) and np.all((c != 2) | (y > 0)) and np.all((c != 1) | (y == 0)))
    argmax_ok = bool(np.all(prob.argmax(1) == c))
    counts = {n: int((c == k).sum()) for k, n in enumerate(CLASS_NAMES)}
    by_combo = {}
    for cb in sorted(set(combo), key=lambda t: (t != "none", len(t), t)):
        sel = np.array([x == cb for x in combo])
        by_combo[cb] = {"n": int(sel.sum()), "class_counts": {n: int((c[sel] == k).sum()) for k, n in enumerate(CLASS_NAMES)},
                        "mean_intensity": float(y[sel].mean()), "mean_abs_intensity": float(np.abs(y[sel]).mean()),
                        "mean_max_prob": float(prob[sel].max(1).mean())}
    n_miss_mod = np.array([sum(ratio[m][i] > 0 for m in MODALITIES) for i in range(len(sd))])
    summ = {
        "input": str(input_path), "split": s, "n": int(len(sd)), "version": sd.version,
        "model": {"path": str(model_path), "variant": ck.get("variant"), "seed": ck.get("seed"), "bias": ck["bias"],
                  "best_epoch": ck.get("best_epoch")},
        "decision_rule": "pred_class = argmax(ℓ + (0,β,0))；prob_* = softmax(ℓ + (0,β,0))（与 pred_class 一致）；"
                         "pred_intensity = 一致性投影（Negative→min(ŷ,-0.01)，Neutral→0，Positive→max(ŷ,0.01)）",
        "class_counts": counts, "class_ratio": {k: v / max(len(sd), 1) for k, v in counts.items()},
        "pred_intensity": {**_stats(y), "hist": _hist(y, np.linspace(-3, 3, 13))},
        "raw_intensity": _stats(np.clip(pred["y_raw"], -3, 3)),
        "mean_prob": {n: float(prob[:, k].mean()) for k, n in enumerate(CLASS_NAMES)},
        "missing": missing_summary(rows, len(sd)),
        "combo_frequency": {cb: v["n"] for cb, v in by_combo.items()},
        "n_missing_modalities": {str(k): int((n_miss_mod == k).sum()) for k in range(4)},
        "by_combo": by_combo,
        "consistency": {"polarity_intensity_consistent": sign_ok, "prob_argmax_equals_pred_class": argmax_ok},
        "input_summary": sd.summary(),
    }
    if sd.has_labels:  # 附件3 应无标签；若有（例如自测），顺带给出指标
        summ["metrics_if_labeled"] = evaluate_all(sd.y_reg, y, c)
    dump_json(summ, out_dir / SUMMARY_NAME)
    make_figures(df, summ, out_dir / "figures")
    if logger:
        logger.info(f"附件3 预测 → {out_dir / CSV_NAME}（{len(df)} 行）；类别分布 {counts}；缺失组合 {summ['combo_frequency']}")
    return {"df": df, "summary": summ}


def make_figures(df: pd.DataFrame, summ: dict, fd: Path) -> None:
    plt = plots.plt
    fd.mkdir(parents=True, exist_ok=True)
    # 1) 各模态检测到的缺失率分布
    fig, axes = plt.subplots(1, 3, figsize=(15, 3.8), sharey=True)
    for ax, (k, m) in zip(axes, enumerate(MODALITIES)):
        r = df[f"miss_ratio_{m}"].to_numpy()
        has = r[r > 0]
        ax.hist(has, bins=np.linspace(0, 1, 21), color=plots.PALETTE[k], edgecolor="white", linewidth=1.0)
        ax.set_title(f"{MOD_CN[m]}：{len(has)}/{len(r)} 条样本有缺失，均值 {has.mean() if len(has) else 0:.2f}")
        ax.set_xlabel("缺失率（有效范围内全零行占比）")
    axes[0].set_ylabel("样本数")
    fig.suptitle("附件3：检测到的各模态缺失率分布（仅统计有缺失的样本）")
    plots._save(fig, fd / "att3_missing_ratio_hist.png")
    # 2) 预测分布
    fig, axes = plt.subplots(1, 2, figsize=(12, 4))
    names = list(CLASS_NAMES)
    cnt = [summ["class_counts"][n] for n in names]
    axes[0].bar(range(3), cnt, color=[plots.POL_COLORS[n] for n in names], width=0.6)
    for i, v in enumerate(cnt):
        axes[0].text(i, v, str(v), ha="center", va="bottom")
    axes[0].set_xticks(range(3), ["负向", "中性", "正向"])
    axes[0].set_title("预测极性分布")
    edges = np.linspace(-3.125, 3.125, 26)  # 以 0.25 的整数倍为中心，中性（=0）单独落在中心柱
    bottom = np.zeros(len(edges) - 1)
    for k, n in enumerate(names):
        h, _ = np.histogram(df.loc[df["pred_class"] == k, "pred_intensity"], bins=edges)
        axes[1].bar(edges[:-1], h, width=np.diff(edges) * 0.92, align="edge", bottom=bottom, color=plots.POL_COLORS[n],
                    label=["负向", "中性", "正向"][k])
        bottom += h
    axes[1].set_xlabel("预测强度（投影后）")
    axes[1].set_title("预测强度分布")
    axes[1].legend(fontsize=8)
    fig.suptitle(f"附件3 预测结果（n={len(df)}）")
    plots._save(fig, fd / "att3_pred_distribution.png")
    # 3) 按缺失组合的预测构成
    bc = summ["by_combo"]
    keys = list(bc)
    fig, ax = plt.subplots(figsize=(max(6, 1.0 * len(keys) + 2), 4))
    bottom = np.zeros(len(keys))
    for k, n in enumerate(names):
        v = np.array([bc[c]["class_counts"][n] for c in keys], float)
        ax.bar(range(len(keys)), v, bottom=bottom, color=plots.POL_COLORS[n], width=0.6, label=["负向", "中性", "正向"][k])
        bottom += v
    ax.set_xticks(range(len(keys)), keys)
    ax.set_xlabel("检测到的缺失模态组合")
    ax.set_ylabel("样本数")
    ax.legend(fontsize=8)
    ax.set_title("附件3：各缺失组合下的预测极性构成")
    plots._save(fig, fd / "att3_by_combo.png")


def main():
    ap = argparse.ArgumentParser(description="问题2：附件3 推理")
    ap.add_argument("--run", default=None, help="运行目录（用其中 model.pt，输出也写到这里）")
    ap.add_argument("--model", default=None, help="model.pt 路径（优先于 --run）")
    ap.add_argument("--input", required=True, help="附件3 特征文件 .pkl")
    ap.add_argument("--split", default="auto", help="auto / test / ... / __flat__")
    ap.add_argument("--out", default=None, help="输出目录（缺省 = 模型所在目录）")
    ap.add_argument("--device", default=None)
    ap.add_argument("--threads", type=int, default=None)
    a = ap.parse_args()
    if a.threads:
        import torch

        torch.set_num_threads(a.threads)
    mp = resolve_path(a.model) if a.model else resolve_path(a.run) / "model.pt"
    out = resolve_path(a.out) if a.out else mp.parent
    logger = get_logger("q2", out / "logs" / "q2.log")
    run_predict(mp, resolve_path(a.input), a.split, out, get_device(a.device), logger=logger)


if __name__ == "__main__":
    main()
