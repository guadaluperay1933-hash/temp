"""问题3 附件4 推理：完整三模态无标签特征 → 极性 + 强度（自洽）+ 可量化、可回看的解释。

    python -m q3.predict --run outputs/q3/amee_aligned --input data/附件4/<文件>.pkl --videos data/附件4/<视频目录>
    python -m q3.predict --model outputs/q3/amee_aligned/model.pt --input <pkl> --videos <dir> --out <dir>

产物（写到运行目录）：
    附件4_预测与解释结果.csv（UTF-8-SIG）、附件4_explanations.json（逐位置贡献与时间）、附件4_summary.json、
    timemap_report.json（帧率估计 / 词对齐来源 / 文本映射方式）、cards/<id>.png（每条样本一张解释卡）、figures/att4_*.png
* 标准化器、决策偏置、特征版本全部取自 model.pt；文件版本与模型版本不一致时报错（题面要求同一版本）。
"""
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import CLASS_NAMES, MOD_CN, MODALITIES, Normalizer, get_split, list_splits, load_pkl
from common.utils import dump_json, get_device, get_logger, load_json

from . import plots
from .cards import render_card, safe_name
from .config import load_config, resolve_path
from .engine import PooledSplit, check_version, load_checkpoint
from .explain import decomposition_checks, explanation_stats, extract_all_evidence, faithfulness, run_explain
from .records import build_record
from .timemap import build_timemaps, read_frame

CSV_NAME = "附件4_预测与解释结果.csv"
JSON_NAME = "附件4_explanations.json"
SUMMARY_NAME = "附件4_summary.json"
CSV_COLUMNS = ["index", "id", "video_id", "clip_id", "pred_polarity", "pred_polarity_cn", "pred_intensity",
               "raw_intensity", "prob_negative", "prob_neutral", "prob_positive", "main_modality",
               "share_text", "share_audio", "share_vision", "shapley_text", "shapley_audio", "shapley_vision",
               "gate_text", "gate_audio", "gate_vision", "gate_interaction",
               "text_evidence", "text_evidence_positions", "audio_evidence", "audio_evidence_positions",
               "vision_evidence", "vision_evidence_positions", "time_mapping", "deletion_drop",
               "deletion_drop_random", "shapley_base"]


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


def shares_to_percent(sh: np.ndarray) -> np.ndarray:
    """份额 → 百分数（两位小数），用最大余数法保证三者之和恰为 100.00。"""
    raw = np.asarray(sh, np.float64) * 10000.0
    base = np.floor(raw)
    rem = int(round(10000 - base.sum()))
    order = np.argsort(-(raw - base))
    for j in order[:max(rem, 0)]:
        base[j] += 1
    return base / 100.0


def _stats(x) -> dict:
    x = np.asarray(x, np.float64)
    if x.size == 0:
        return {"n": 0}
    return {"n": int(x.size), "mean": float(x.mean()), "std": float(x.std()), "min": float(x.min()),
            "median": float(np.median(x)), "max": float(x.max())}


def csv_row(rec: dict) -> dict:
    pct = shares_to_percent([rec["share"][m] for m in MODALITIES])
    g = rec.get("gate", {})
    return {"index": rec["index"], "id": rec["id"], "video_id": rec["video_id"], "clip_id": rec["clip_id"],
            "pred_polarity": rec["pred_polarity"], "pred_polarity_cn": rec["pred_polarity_cn"],
            "pred_intensity": round(rec["pred_intensity"], 4), "raw_intensity": round(rec["raw_intensity"], 4),
            "prob_negative": round(rec["probs"][0], 4), "prob_neutral": round(rec["probs"][1], 4),
            "prob_positive": round(rec["probs"][2], 4), "main_modality": MOD_CN[rec["main_modality"]],
            "share_text": pct[0], "share_audio": pct[1], "share_vision": pct[2],
            **{f"shapley_{m}": round(rec["shapley"][m], 5) for m in MODALITIES},
            **{f"gate_{e}": (round(g[e], 4) if e in g else np.nan) for e in ("text", "audio", "vision", "interaction")},
            "text_evidence": rec["text_evidence"], "text_evidence_positions": rec["text_evidence_positions"],
            "audio_evidence": rec["audio_evidence"], "audio_evidence_positions": rec["audio_evidence_positions"],
            "vision_evidence": rec["vision_evidence"], "vision_evidence_positions": rec["vision_evidence_positions"],
            "time_mapping": rec["time_mapping"],
            "deletion_drop": None if rec["deletion_drop"] is None else round(rec["deletion_drop"], 5),
            "deletion_drop_random": None if rec["deletion_drop_random"] is None else round(rec["deletion_drop_random"], 5),
            "shapley_base": round(rec["shapley_base"], 5)}


def run_predict(model_path, input_path, split: str, videos_dir, out_dir, device, cfg: dict | None, logger,
                batch_size: int = 256, make_cards: bool = True) -> dict:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    model, ck = load_checkpoint(model_path, device)
    cfg = cfg or ck["config"]
    ecfg = dict(cfg.get("explain", ck["config"]["explain"]))
    tcfg = dict(cfg.get("timemap", ck["config"].get("timemap", {})))
    bias = ck["bias"]
    raw = load_pkl(input_path)
    sp = pick_split(raw, split)
    sd_raw = get_split(raw, sp, name="附件4")
    check_version(sd_raw, ck["version"], f"附件4 {Path(input_path).name}")
    logger.info(f"[附件4] 读入 {input_path}（划分 {sp}）：{len(sd_raw)} 条，版本 {sd_raw.version}，"
                f"有标签={sd_raw.has_labels}；视频目录 {videos_dir}")
    sd = Normalizer.from_state_dict(ck["normalizer"]).transform(sd_raw)
    ps = PooledSplit(sd, model.strides)
    b = run_explain(model, ps, device, bias, batch_size)
    ev = extract_all_evidence(b, ecfg)
    fa = faithfulness(model, ps, b, ev, device, bias, ecfg, batch_size, with_curves=False, logger=logger)
    maps, trep = build_timemaps(sd_raw, videos_dir, tcfg, ck["version"], logger)
    for mp in maps:
        mp["rates"] = trep.get("rates", {})
    dump_json(trep, out_dir / "timemap_report.json")
    recs = [build_record(b, i, ev[i], sd_raw, word_idx=maps[i]["pos"]["text"]["word"], tmap=maps[i],
                         deletion_drop=fa["per_sample_deletion_drop"][i],
                         deletion_drop_random=fa["per_sample_deletion_drop_random"][i]) for i in range(b["n"])]
    df = pd.DataFrame([csv_row(r) for r in recs], columns=CSV_COLUMNS)
    df.to_csv(out_dir / CSV_NAME, index=False, encoding="utf-8-sig")
    dump_json({"model": str(model_path), "input": str(input_path), "split": sp, "version": ck["version"],
               "bias": bias, "notes": {
                   "positions": "pos = 模型使用的（下采样后）全局位置；orig_lo/hi = 原始序列位置（含两端）",
                   "contrib": "e_{m,t} = w_m·a_{m,t}·v_{m,t}，Σ_t e_{m,t} = contribution[m]，Σ_k contribution[k] = raw_intensity",
                   "shapley": "精确 3 人 Shapley，raw_intensity = shapley_base + Σ_m shapley[m]",
                   "time": "t_start/t_end 单位秒；时间映射方式见 time_mapping_method 与 timemap_report.json"},
               "samples": recs}, out_dir / JSON_NAME)
    # ---- 汇总
    st = explanation_stats(b)
    comp = fa["comprehensiveness"]
    summary = {"n": len(recs), "input": str(input_path), "split": sp, "version": ck["version"], "bias": bias,
               "pred_class_counts": {n: int((b["pred_cls"] == c).sum()) for c, n in enumerate(CLASS_NAMES)},
               "pred_intensity": _stats([r["pred_intensity"] for r in recs]),
               "raw_intensity": _stats(b["y_raw"]), "confidence": _stats(b["probs"].max(1)),
               "main_modality_counts": {MOD_CN[m]: v for m, v in st["main_modality_counts"].items()},
               "shapley_share_mean": {MOD_CN[m]: v for m, v in st["shapley_share_mean"].items()},
               "shapley_signed_mean": {MOD_CN[m]: v for m, v in st["shapley_signed_mean"].items()},
               "shapley_base_value": st["shapley_base_value"], "gate_mean": st.get("gate_mean"),
               "by_predicted_class": st["by_predicted_class"],
               "evidence": {m: {"mean_segments": float(np.mean([len(r["evidence"][m]) for r in recs])),
                                "mean_duration_s": _stats([s["t_end"] - s["t_start"] for r in recs for s in r["evidence"][m]
                                                           if "t_start" in s])} for m in MODALITIES},
               "faithfulness": {"deletion_all_evidence": comp["all_modalities"]["abs_dy"],
                                "deletion_main_modality": comp["main_modality"]["abs_dy"],
                                "sufficiency": fa["sufficiency"]["abs_dy"],
                                "note": "附件4 样本数少，删除检验只作逐样本参考（deletion_drop 列）"},
               "decomposition_checks": decomposition_checks(b), "timemap": trep,
               "files": {"csv": CSV_NAME, "explanations": JSON_NAME, "cards": "cards/"}}
    dump_json(summary, out_dir / SUMMARY_NAME)
    # ---- 图与解释卡
    fig_dir = out_dir / "figures"
    plots.plot_att4_overview(recs, fig_dir / "att4_overview.png")
    cache = out_dir / "explain_cache.npz"
    shown = {}
    if cache.exists():
        z = np.load(cache)
        for s in ("valid", "test"):
            if f"{s}_share" in z:
                shown[s] = {"share": z[f"{s}_share"], "main": z[f"{s}_main"]}
    shown["att4"] = {"share": b["share"], "main": b["main"]}
    plots.plot_share_distribution(shown, fig_dir / "modality_share_distribution.png")
    if make_cards:
        cdir = out_dir / "cards"
        index = []
        for r in recs:
            p = render_card(r, cdir / f"{safe_name(r['id'])}.png", int(tcfg.get("max_frames", 3)),
                            frame_reader=read_frame if videos_dir else None)
            index.append({"id": r["id"], "file": str(p.relative_to(out_dir))})
        dump_json(index, cdir / "att4_index.json")
    logger.info(f"[附件4] 完成：{len(recs)} 条 → {out_dir / CSV_NAME}；主要参考模态 {summary['main_modality_counts']}；"
                f"预测类别 {summary['pred_class_counts']}")
    return {"records": recs, "summary": summary, "bundle": b}


def main():
    ap = argparse.ArgumentParser(description="问题3：附件4 预测 + 解释")
    ap.add_argument("--run", default=None, help="运行目录（含 model.pt）；结果也写到这里")
    ap.add_argument("--model", default=None, help="直接给 model.pt（与 --out 搭配）")
    ap.add_argument("--input", required=True, help="附件4 特征文件")
    ap.add_argument("--split", default="auto")
    ap.add_argument("--videos", default=None, help="附件4 原始视频根目录 <video_id>/<clip_id>.mp4")
    ap.add_argument("--out", default=None)
    ap.add_argument("--config", default=None, help="取其中 explain / timemap 段（缺省用 model.pt 里保存的配置）")
    ap.add_argument("--device", default=None)
    ap.add_argument("--threads", type=int, default=None)
    ap.add_argument("--no-cards", action="store_true")
    a = ap.parse_args()
    if not a.run and not a.model:
        ap.error("需要 --run 或 --model")
    model_path = resolve_path(a.model) if a.model else resolve_path(a.run) / "model.pt"
    out = resolve_path(a.out) if a.out else (resolve_path(a.run) if a.run else model_path.parent)
    if a.threads:
        import torch

        torch.set_num_threads(a.threads)
    cfg = load_config(a.config) if a.config else None
    logger = get_logger("q3", out / "logs" / "q3.log")
    run_predict(model_path, resolve_path(a.input), a.split, resolve_path(a.videos) if a.videos else None, out,
                get_device(a.device), cfg, logger, make_cards=not a.no_cards)


if __name__ == "__main__":
    main()
