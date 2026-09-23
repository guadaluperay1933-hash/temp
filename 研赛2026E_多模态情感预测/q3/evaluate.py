"""问题3 评测与解释：验证/测试集解释统计、忠实性检验、错误归因、可解释性代价（消融）表、论文用图、典型样本解释卡。

    python -m q3.evaluate --run-name amee_aligned --data data/附件2/aligned_50.pkl
    python -m q3.evaluate --out outputs/q3/amee_aligned --tasks explain faith error ablation figures

产物（outputs/q3/<run>/）：
    explain_stats.json（并写入 metrics_valid/test.json 的 'explanation' 键）、test_explanations.csv / .jsonl
    faithfulness.json、error_analysis.json、ablation.csv、ablation_explain.json、cards/typical_*.png、figures/*.png
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import CLASS_CN, CLASS_NAMES, MOD_CN, MODALITIES
from common.decision import project_intensity
from common.error_analysis import error_attribution, plot_error_analysis
from common.metrics import EXTRA_KEYS, MAIN_KEYS
from common.utils import dump_json, get_logger, load_json

from . import plots
from .cards import render_card, safe_name
from .config import VARIANT_CN, dump_yaml, load_config, run_dir
from .engine import PooledSplit, load_checkpoint, prepare_data, resolve_device, setup_threads
from .explain import (decomposition_checks, explanation_stats, extract_all_evidence, faithfulness,
                      gate_shapley_agreement, ig_consistency, integrated_gradients, position_profile, run_explain)
from .records import build_record, text_word_index
from .timemap import TokenWordMapper
from .train import add_common_args, build_cfg

METRIC_KEYS = MAIN_KEYS + EXTRA_KEYS
TASKS = ("explain", "faith", "error", "ablation", "figures")


def _strip_private(d):
    if isinstance(d, dict):
        return {k: _strip_private(v) for k, v in d.items() if not str(k).startswith("_")}
    return d


def _merge_into(path: Path, key: str, value) -> None:
    if path.exists():
        d = load_json(path)
        d[key] = value
        dump_json(d, path)


def explanation_table(b: dict, ev: list, sd, mapper) -> tuple[pd.DataFrame, list[dict]]:
    """逐样本解释表（CSV）与完整记录（JSONL）。"""
    rows, recs = [], []
    for i in range(b["n"]):
        rec = build_record(b, i, ev[i], sd, word_idx=text_word_index(sd, i, mapper))
        recs.append(rec)
        row = {"index": i, "id": rec["id"], "y_true": rec.get("y_true"), "cls_true": rec.get("cls_true"),
               "pred_polarity": rec["pred_polarity"], "pred_intensity": rec["pred_intensity"],
               "raw_intensity": rec["raw_intensity"], "correct": rec.get("cls_true") == rec["pred_polarity"],
               "main_modality": MOD_CN[rec["main_modality"]],
               **{f"share_{m}": 100 * rec["share"][m] for m in MODALITIES},
               **{f"shapley_{m}": rec["shapley"][m] for m in MODALITIES}, "shapley_base": rec["shapley_base"]}
        for e, v in rec.get("gate", {}).items():
            row[f"gate_{e}"] = v
        for m in MODALITIES:
            row[f"{m}_evidence"] = rec[f"{m}_evidence"]
            row[f"{m}_evidence_positions"] = rec[f"{m}_evidence_positions"]
        rows.append(row)
    return pd.DataFrame(rows), recs


def pick_typical(b: dict, sd, n: int) -> list[tuple[int, str]]:
    """典型样本：每个主要参考模态各取一条"预测正确、非中性、该模态份额最高"的样本；再加中性正确样本、误差最大样本，
    不足 n 条时用置信度最高的正确样本补足。返回 [(下标, 标签)]。"""
    y_proj = project_intensity(b["pred_cls"], b["y_raw"])
    correct = b["pred_cls"] == sd.y_cls
    out = []
    for i_m, m in enumerate(MODALITIES):
        cand = np.where(correct & (b["main"] == i_m) & (sd.y_cls != 1))[0]
        if cand.size:
            out.append((int(cand[np.argmax(b["share"][cand, i_m])]), f"main_{m}"))
    cand = np.where(correct & (sd.y_cls == 1))[0]
    if cand.size:
        out.append((int(cand[np.argmax(b["probs"][cand, 1])]), "neutral"))
    err = np.abs(y_proj - sd.y_reg)
    for i in np.argsort(-err):
        if int(i) not in {j for j, _ in out}:
            out.append((int(i), "worst_error"))
            break
    for i in np.argsort(-b["probs"].max(1)):
        if len(out) >= n:
            break
        if correct[i] and int(i) not in {j for j, _ in out}:
            out.append((int(i), "confident"))
    return out[:n]


def ablation_table(cfg: dict, rdir: Path, data: dict, device, logger, test_ps_cache: dict) -> tuple[pd.DataFrame, dict]:
    """读 models/<variant>/metrics_*.json → 性能表；对每个变体再算解释统计（主模态分布、删除检验）→ 可解释性代价。"""
    rows, extra = [], {}
    ecfg = dict(cfg["explain"])
    for v in cfg["ablation"]["variants"]:
        vdir = rdir / "models" / v
        if not (vdir / "metrics_test.json").exists():
            continue
        for split in ("valid", "test"):
            mt = load_json(vdir / f"metrics_{split}.json")
            for stage, dec in (("after_projection", "projected"), ("before_projection", "raw")):
                r = {"variant": v, "description": VARIANT_CN.get(v, v), "split": split, "decision": dec,
                     "selected_seed": mt["selected_seed"], "n_params": mt["selected"].get("n_params")}
                r.update({k: mt["selected"][stage].get(k) for k in METRIC_KEYS})
                ms = mt["mean_std_over_seeds"][stage]
                r.update({f"{k}_std": ms[k]["std"] for k in ("acc3", "f1_weighted", "mae", "corr")})
                rows.append(r)
        sel = load_json(vdir / "selected.json")
        model, ck = load_checkpoint(vdir / sel["path"], device)
        key = tuple(sorted(model.strides.items()))
        if key not in test_ps_cache:
            test_ps_cache[key] = PooledSplit(data["test"], model.strides)
        ps = test_ps_cache[key]
        b = run_explain(model, ps, device, ck["bias"])
        is_amee = b.get("c") is not None
        ev = extract_all_evidence(b, ecfg, key="contrib" if is_amee else "attn")
        fa = faithfulness(model, ps, b, ev, device, ck["bias"], ecfg, with_curves=False)
        comp = fa["comprehensiveness"]["all_modalities"]["abs_dy"]
        st = explanation_stats(b, data["test"].y_cls)
        info = {"arch": ck["arch"], "evidence_source": "内置贡献 e_{m,t}" if is_amee else "注意力 a_{m,t}（无内置贡献）",
                "main_modality_rate": st["main_modality_rate"], "shapley_share_mean": st["shapley_share_mean"],
                "deletion_evidence_mean": comp["evidence_mean"], "deletion_random_mean": comp["random_mean"],
                "deletion_win_rate": comp["win_rate"], "sufficiency": fa["sufficiency"]["abs_dy"]}
        if is_amee:
            info["agreement"] = gate_shapley_agreement(b)
            info["decomposition_checks"] = decomposition_checks(b)
        extra[v] = info
        for r in rows:
            if r["variant"] == v and r["split"] == "test":
                r["test_deletion_ratio_evidence_vs_random"] = comp["evidence_mean"] / max(comp["random_mean"], 1e-12)
                r["test_deletion_win_rate"] = comp["win_rate"]
                r["test_main_text_rate"] = st["main_modality_rate"]["text"]
                if is_amee and "interaction_contribution_share" in info["agreement"]:
                    r["test_interaction_contrib_share"] = info["agreement"]["interaction_contribution_share"]["mean"]
        logger.info(f"[消融] {v}: 删除证据 |Δŷ| {comp['evidence_mean']:.4f} vs 随机 {comp['random_mean']:.4f}，"
                    f"主模态分布 {st['main_modality_rate']}")
    return pd.DataFrame(rows), extra


def run_evaluate(cfg: dict, rdir: Path, data: dict, device, logger, tasks=TASKS) -> dict:
    rdir = Path(rdir)
    fig_dir = rdir / "figures"
    fig_dir.mkdir(parents=True, exist_ok=True)
    model, ck = load_checkpoint(rdir / "model.pt", device)
    bias = ck["bias"]
    ecfg = dict(cfg["explain"])
    eb = int(cfg["train"].get("eval_batch_size", 256))
    splits = list(dict.fromkeys(list(ecfg.get("splits", ["valid", "test"])) + [ecfg.get("faith_split", "test")]))
    pooled = {s: PooledSplit(data[s], model.strides) for s in splits}
    bundles, evid, out = {}, {}, {}
    for s in splits:
        bundles[s] = run_explain(model, pooled[s], device, bias, eb)
        evid[s] = extract_all_evidence(bundles[s], ecfg)
    mapper = TokenWordMapper(cfg.get("timemap", {}).get("tokenizer"), logger)

    if "explain" in tasks:
        stats = {}
        for s in splits:
            b = bundles[s]
            st = explanation_stats(b, data[s].y_cls)
            st["agreement_gate_contrib_vs_shapley"] = gate_shapley_agreement(b)
            st["decomposition_checks"] = decomposition_checks(b)
            st["evidence_summary"] = {m: {"mean_positions": float(np.mean([len(e[m]["positions"]) for e in evid[s]])),
                                          "mean_segments": float(np.mean([len(e[m]["segments"]) for e in evid[s]]))}
                                      for m in MODALITIES}
            stats[s] = st
            _merge_into(rdir / f"metrics_{s}.json", "explanation", st)
            ag = st["agreement_gate_contrib_vs_shapley"]
            logger.info(f"[解释|{s}] 主要参考模态 {st['main_modality_rate']}；平均份额 {st['shapley_share_mean']}；"
                        f"门控主模态一致率 {ag.get('main_modality_agreement_gate_vs_shapley')}；"
                        f"核验 {st['decomposition_checks']}")
        stats["profile"] = {s: position_profile(bundles[s]) for s in splits}
        dump_json(stats, rdir / "explain_stats.json")
        df, recs = explanation_table(bundles["test"], evid["test"], data["test"], mapper)
        df.to_csv(rdir / "test_explanations.csv", index=False, encoding="utf-8-sig")
        with open(rdir / "test_explanations.jsonl", "w", encoding="utf-8") as f:
            for r in recs:
                f.write(json.dumps(r, ensure_ascii=False, default=float) + "\n")
        np.savez_compressed(rdir / "explain_cache.npz",
                            **{f"{s}_share": bundles[s]["share"] for s in splits},
                            **{f"{s}_main": bundles[s]["main"] for s in splits})
        # 典型样本解释卡
        cdir = rdir / "cards"
        typ = []
        for i, tag in pick_typical(bundles["test"], data["test"], int(ecfg.get("n_typical_cards", 6))):
            p = render_card(recs[i], cdir / f"typical_{tag}_{safe_name(recs[i]['id'])}.png")
            typ.append({"tag": tag, "id": recs[i]["id"], "file": str(p.relative_to(rdir))})
        dump_json(typ, cdir / "typical_index.json")
        out["explain"] = stats

    if "faith" in tasks:
        s = ecfg.get("faith_split", "test")
        b, ev, ps = bundles[s], evid[s], pooled[s]
        logger.info(f"[忠实性] 在 {s} 上计算积分梯度（{ecfg.get('ig_steps', 32)} 步）与删除/保留检验 …")
        igr = integrated_gradients(model, ps, device, int(ecfg.get("ig_steps", 32)), int(ecfg.get("ig_batch", 64)))
        igc = ig_consistency(b, igr, ev, ecfg)
        fa = faithfulness(model, ps, b, ev, device, bias, ecfg, eb, with_curves=True, logger=logger)
        res = {"split": s, "n": b["n"], "decomposition_checks": decomposition_checks(b),
               "integrated_gradients": {"steps": int(ecfg.get("ig_steps", 32)), "baseline": "全零特征（掩码不变）",
                                        **_strip_private(igc)},
               "agreement_gate_contrib_vs_shapley": gate_shapley_agreement(b),
               **{k: v for k, v in fa.items() if not k.startswith("per_sample")}}
        dump_json(res, rdir / "faithfulness.json")
        plots.plot_faithfulness(res, fig_dir / "faithfulness_curves.png")
        plots.plot_comprehensiveness(res, fig_dir / "faithfulness_comprehensiveness.png")
        plots.plot_ig_consistency(igc["_raw_spearman_signed"], igc, fig_dir / "ig_consistency.png")
        pl = igc["position_level"]
        logger.info("[忠实性] IG 一致性（逐样本 Spearman 中位数）："
                    + "，".join(f"{MOD_CN[m]} {pl[m]['spearman_signed'].get('median', float('nan')):.3f}" for m in MODALITIES))
        out["faith"] = res

    if "error" in tasks:
        ea_all = {}
        for s in ("test", "valid"):
            if s not in bundles:
                continue
            b, sd = bundles[s], data[s]
            y_proj = project_intensity(b["pred_cls"], b["y_raw"])
            n_words = np.array([len((t or "").split()) for t in (sd.raw_text or [""] * len(sd))])
            cov = {"主要参考模态": np.array([MOD_CN[MODALITIES[k]] for k in b["main"]]),
                   "最大作用份额": b["share"].max(1), "文本长度(词)": n_words}
            ea = error_attribution(sd.ids, sd.raw_text, sd.y_reg, y_proj, b["pred_cls"], covariates=cov)
            ea["split"] = s
            ea_all[s] = ea
            plot_error_analysis(ea, fig_dir / f"error_analysis_{s}.png", title=f"问题3 AMEE 错误归因（{plots.SPLIT_CN[s]}）")
            plots.plot_error_covariates(ea, fig_dir / f"error_covariates_{s}.png")
        ea = dict(ea_all.get("test", {}))
        ea["valid"] = ea_all.get("valid")
        dump_json(ea, rdir / "error_analysis.json")
        out["error"] = ea

    if "ablation" in tasks:
        df, extra = ablation_table(cfg, rdir, data, device, logger, {tuple(sorted(model.strides.items())): pooled["test"]}
                                   if "test" in pooled else {})
        if len(df):
            df.to_csv(rdir / "ablation.csv", index=False, encoding="utf-8-sig")
            dump_json(extra, rdir / "ablation_explain.json")
            plots.plot_ablation(df, fig_dir / "ablation.png")
        out["ablation"] = extra

    if "figures" in tasks:
        if (rdir / "train_log.csv").exists():
            plots.plot_training(pd.read_csv(rdir / "train_log.csv"), fig_dir / "training_curves.png", "AMEE 训练过程")
        shown = {s: bundles[s] for s in splits}
        plots.plot_share_distribution(shown, fig_dir / "modality_share_distribution.png")
        plots.plot_position_profile({s: position_profile(bundles[s]) for s in splits}, fig_dir / "position_profile.png")
        for s in splits:
            plots.plot_role_by_polarity(bundles[s], data[s].y_cls, fig_dir / f"modality_role_by_polarity_{s}.png",
                                        f"三模态作用差异（{plots.SPLIT_CN[s]}）")
        if bundles[splits[-1]].get("w") is not None:
            s = ecfg.get("faith_split", "test")
            plots.plot_gate_vs_shapley(bundles[s], gate_shapley_agreement(bundles[s]), fig_dir / "gate_vs_shapley.png")
    return out


def inherit_run_config(cfg: dict, rdir: Path, a) -> dict:
    """单独运行评测 / 推理时：命令行没给的数据路径与特征版本，沿用该运行目录训练时保存的 config.yaml。"""
    p = Path(rdir) / "config.yaml"
    if p.exists():
        saved = load_config(p)
        if not getattr(a, "version", None):
            cfg["version"] = saved.get("version", cfg["version"])
        if not getattr(a, "data", None):
            cfg["data"] = saved.get("data", cfg["data"])
        cfg["ablation"] = saved.get("ablation", cfg.get("ablation"))
    return cfg


def main():
    ap = argparse.ArgumentParser(description="问题3：解释统计 / 忠实性 / 错误归因 / 消融表 / 图")
    add_common_args(ap)
    ap.add_argument("--tasks", nargs="+", default=list(TASKS), choices=list(TASKS))
    a = ap.parse_args()
    cfg = build_cfg(a)
    rdir = run_dir(cfg, a.out)
    cfg = inherit_run_config(cfg, rdir, a)
    logger = get_logger("q3", rdir / "logs" / "q3.log")
    setup_threads(cfg)
    device = resolve_device(cfg)
    data, _ = prepare_data(cfg, logger)
    run_evaluate(cfg, rdir, data, device, logger, a.tasks)
    logger.info(f"评测完成：{rdir}")


if __name__ == "__main__":
    main()
