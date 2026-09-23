"""问题3 合成数据自检：核对一次 run_all 的产物是否齐全、自洽，并用合成真值检验"解释机制能否找回埋入的证据"。

    python -m q3.selftest --run outputs/q3/synth_aligned \\
        --truth outputs/synthetic/附件2/_truth_aligned.json --truth4 outputs/synthetic/附件4/_truth_aligned.json

合成数据（common.synth）的真值：每条样本的强度只由各模态少数"证据位置"决定（文本 1–2 个词元，语音/视觉一段约 15% 长度的
连续区间），各模态信息量 ∝ Dirichlet 权重 w_m，真值主导模态 = argmax_m w_m；y = 0 的样本没有埋入任何证据。
检查项（数值只说明代码跑通、解释机制有效，不能进论文）：
    1. 接口合同文件齐全、图非空；训练损失下降；三条分解恒等式的数值误差 < 1e-4
    2. 测试集：Shapley 主要参考模态 vs 真值主导模态的一致率（全部非中性样本 / 主导权重领先 ≥ 0.2 的"清晰"样本），
       与"总猜最常见模态"的基线比较；门控、内置贡献给出的主模态也一并报告
    3. 测试集：关键证据命中率 —— precision = |E∩T|/|E|、hit = 1[E∩T≠∅]、recall = |E∩T|/|T|，
       与随机取同样数量候选位置的期望 precision = |T∩C|/|C| 比较；同时给出按 |IG|、按注意力 a 排序的命中率作参照
    4. 删除检验：删除关键证据的 |Δŷ| 显著大于删除随机位置
    5. 附件4：CSV 列名/行数/UTF-8-SIG/份额和=100/极性与强度自洽；每条样本一张解释卡；
       非对齐版本：估计的特征帧率 vs 真值（语音 20 Hz、视觉 15 Hz）
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import MODALITIES
from common.utils import dump_json, get_device, load_json

from .config import load_config, resolve_path
from .engine import PooledSplit, load_checkpoint, prepare_data
from .explain import evidence_k, extract_all_evidence, extract_evidence, integrated_gradients, run_explain
from .predict import CSV_COLUMNS, CSV_NAME, JSON_NAME, SUMMARY_NAME

REQUIRED = ["model.pt", "config.yaml", "train_log.csv", "metrics_valid.json", "metrics_test.json", "faithfulness.json",
            "error_analysis.json", "ablation.csv", CSV_NAME, JSON_NAME, SUMMARY_NAME, "timemap_report.json"]


def _mean(x) -> float | None:
    x = [v for v in x if v is not None and np.isfinite(v)]
    return float(np.mean(x)) if x else None


def evidence_hits(b: dict, truth: dict, ids: list[str], ecfg: dict, ranking: dict | None = None,
                  min_weight: float = 0.0) -> dict:
    """ranking[m] (N,Lp)：用于排序的分数（缺省 = 内置贡献 |e|）。只统计 y≠0 且该模态真值权重 ≥ min_weight 的样本。"""
    out = {}
    for m in MODALITIES:
        st = int(b["strides"][m])
        prec, hit, rec, rnd = [], [], [], []
        for i, sid in enumerate(ids):
            t = truth.get(sid)
            if t is None or abs(t["y"]) < 1e-6 or t["weights"][m] < min_weight:
                continue
            cand = b["pos"][m]["cand"][i]
            T = {p // st for p in t["evidence"][m]} & set(np.where(cand)[0].tolist())
            if not T or cand.sum() <= len(T):
                continue
            score = b["pos"][m]["contrib"][i] if ranking is None else ranking[m][i]
            E = set(extract_evidence(np.asarray(score, np.float64), cand, m, ecfg)["positions"])
            if not E:
                continue
            inter = len(E & T)
            prec.append(inter / len(E))
            hit.append(float(inter > 0))
            rec.append(inter / len(T))
            rnd.append(len(T) / cand.sum())
        out[m] = {"n": len(prec), "precision": _mean(prec), "hit_rate": _mean(hit), "recall": _mean(rec),
                  "random_precision_expectation": _mean(rnd),
                  "precision_lift": (_mean(prec) / _mean(rnd)) if prec else None}
    return out


def main_modality_agreement(b: dict, truth: dict, ids: list[str], margin: float = 0.2) -> dict:
    tm, sel_all, sel_clear = [], [], []
    for i, sid in enumerate(ids):
        t = truth[sid]
        w = np.array([t["weights"][m] for m in MODALITIES])
        tm.append(int(np.argmax(w)))
        s = np.sort(w)[::-1]
        sel_all.append(abs(t["y"]) >= 1e-6)
        sel_clear.append(abs(t["y"]) >= 1e-6 and s[0] - s[1] >= margin)
    tm, sel_all, sel_clear = np.array(tm), np.array(sel_all), np.array(sel_clear)
    res = {}
    for name, sel in (("non_neutral", sel_all), (f"clear_margin_{margin}", sel_clear)):
        if not sel.any():
            continue
        r = {"n": int(sel.sum()), "shapley": float((b["main"][sel] == tm[sel]).mean()),
             "majority_baseline": float(np.bincount(tm[sel], minlength=3).max() / sel.sum())}
        if b.get("w") is not None:
            r["gate"] = float((np.argmax(b["w"][sel, :3], 1) == tm[sel]).mean())
            r["builtin_contribution"] = float((np.argmax(np.abs(b["c"][sel, :3]), 1) == tm[sel]).mean())
        res[name] = r
    return res


def check(run: Path, truth_path: Path | None, truth4_path: Path | None, device) -> dict:
    rep = {"run": str(run), "checks": {}}
    ok_all = True

    def put(name, ok, **info):
        nonlocal ok_all
        rep["checks"][name] = {"status": "PASS" if ok is True else ("WARN" if ok is None else "FAIL"), **info}
        if ok is False:
            ok_all = False
        print(f"[{rep['checks'][name]['status']}] {name}: {json.dumps(info, ensure_ascii=False, default=str)[:900]}")

    # ---- 1. 文件 / 图 / 损失 / 恒等式
    missing = [f for f in REQUIRED if not (run / f).exists()]
    put("required_files", not missing, missing=missing)
    figs = sorted((run / "figures").glob("*.png"))
    small = [f.name for f in figs if f.stat().st_size < 5000]
    put("figures", len(figs) >= 14 and not small, n_png=len(figs), too_small=small)
    dec = {}
    for log in sorted((run / "models").glob("*/train_log.csv")):
        df = pd.read_csv(log)
        for s, g in df.groupby("seed"):
            dec[f"{log.parent.name}|seed{s}"] = (round(float(g["train_loss"].iloc[0]), 4),
                                                 round(float(g["train_loss"].iloc[-1]), 4), int(len(g)))
    bad = {k: v for k, v in dec.items() if not v[1] < v[0]}
    put("loss_decreases", bool(dec) and not bad, first_last_epochs=dec, not_decreasing=bad)
    fa = load_json(run / "faithfulness.json")
    errs = fa["decomposition_checks"]
    put("decomposition_identities", all(v < 1e-4 for k, v in errs.items() if k.endswith("err")), **errs)
    comp = fa["comprehensiveness"]["all_modalities"]["abs_dy"]
    put("deletion_beats_random", comp["evidence_mean"] > comp["random_mean"] and comp["win_rate"] > 0.5,
        evidence=round(comp["evidence_mean"], 4), random=round(comp["random_mean"], 4), win_rate=comp["win_rate"],
        p=comp.get("wilcoxon_p_one_sided"))
    igl = fa["integrated_gradients"]["position_level"]
    put("ig_consistency", all((igl[m]["spearman_signed"].get("median") or 0) > 0 for m in MODALITIES),
        median_spearman={m: igl[m]["spearman_signed"].get("median") for m in MODALITIES},
        topk_overlap={m: igl[m]["topk_overlap"].get("mean") for m in MODALITIES},
        topk_overlap_random={m: igl[m]["topk_overlap_random_expectation"].get("mean") for m in MODALITIES})

    # ---- 2/3. 真值对照（测试集）
    cfg = load_config(run / "config.yaml")
    if truth_path is not None and truth_path.exists():
        truth = load_json(truth_path)
        data, _ = prepare_data(cfg)
        model, ck = load_checkpoint(run / "model.pt", device)
        ps = PooledSplit(data["test"], model.strides)
        b = run_explain(model, ps, device, ck["bias"])
        ecfg = cfg["explain"]
        ids = data["test"].ids
        mm = main_modality_agreement(b, truth, ids)
        r = mm.get("non_neutral", {})
        put("main_modality_vs_truth", r.get("shapley", 0) > r.get("majority_baseline", 1), **mm)
        eh = evidence_hits(b, truth, ids, ecfg)
        eh_strong = evidence_hits(b, truth, ids, ecfg, min_weight=0.2)
        igr = integrated_gradients(model, ps, device, int(ecfg.get("ig_steps", 32)), int(ecfg.get("ig_batch", 64)))
        eh_ig = evidence_hits(b, truth, ids, ecfg, ranking=igr["ig"])
        eh_att = evidence_hits(b, truth, ids, ecfg, ranking={m: b["pos"][m]["attn"] for m in MODALITIES})
        ok = all((eh[m]["precision"] or 0) > (eh[m]["random_precision_expectation"] or 1) for m in MODALITIES)
        put("evidence_hit_vs_truth", ok, builtin=eh, builtin_weight_ge_0_2=eh_strong, integrated_gradients=eh_ig,
            attention_only=eh_att)
        rep["truth_test"] = {"main_modality": mm, "evidence_builtin": eh, "evidence_builtin_weight_ge_0_2": eh_strong,
                             "evidence_ig": eh_ig, "evidence_attention": eh_att}

    # ---- 5. 附件4
    csv = run / CSV_NAME
    if csv.exists():
        raw = csv.read_bytes()
        df = pd.read_csv(csv, encoding="utf-8-sig")
        n = len(load_json(run / JSON_NAME)["samples"])
        sh = df[["share_text", "share_audio", "share_vision"]].sum(1)
        cons = (((df["pred_polarity"] == "Negative") & (df["pred_intensity"] < 0)) |
                ((df["pred_polarity"] == "Positive") & (df["pred_intensity"] > 0)) |
                ((df["pred_polarity"] == "Neutral") & (df["pred_intensity"] == 0)))
        probs = df[["prob_negative", "prob_neutral", "prob_positive"]].sum(1)
        put("att4_csv", list(df.columns) == CSV_COLUMNS and len(df) == n and raw[:3] == b"\xef\xbb\xbf"
            and bool(np.allclose(sh, 100.0, atol=0.011)) and bool(cons.all()) and bool(np.allclose(probs, 1, atol=2e-3)),
            n_rows=len(df), columns_ok=list(df.columns) == CSV_COLUMNS, bom=raw[:3] == b"\xef\xbb\xbf",
            share_sum_range=[float(sh.min()), float(sh.max())], polarity_intensity_consistent=bool(cons.all()),
            example=df.iloc[0][["text_evidence", "audio_evidence", "vision_evidence", "time_mapping"]].to_dict())
        cards = [p for p in (run / "cards").glob("*.png") if not p.name.startswith("typical_")]
        put("att4_cards", len(cards) == n, n_cards=len(cards), n_samples=n)
        trep = load_json(run / "timemap_report.json")
        if truth4_path is not None and truth4_path.exists():
            t4 = load_json(truth4_path)
            if trep.get("version") == "unaligned":
                est = {}
                ok = True
                for m, key in (("audio", "audio_rate"), ("vision", "vision_rate")):
                    tr = next(iter(t4.values()))[key]
                    r = trep["rates"][m]
                    rel = abs(r["rate_hz"] - tr) / tr if r.get("rate_hz") else None
                    est[m] = {"estimated_hz": r.get("rate_hz"), "truth_hz": tr, "rel_err": rel, "r2": r.get("r2"),
                              "method": r.get("method"), "n_used": r.get("n_used")}
                    ok = ok and rel is not None and rel < 0.02 and r.get("method") == "fixed-rate"
                put("att4_rate_estimation", ok, **est)
                rep["att4_rates"] = est
            recs = load_json(run / JSON_NAME)["samples"]
            agree = [r["main_modality"] == t4[r["id"]]["main_modality"] for r in recs if abs(t4[r["id"]]["y"]) > 1e-6]
            put("att4_main_modality_vs_truth", None, n_non_neutral=len(agree),
                agreement=float(np.mean(agree)) if agree else None,
                note="附件4 只有 10 条，仅作参考")
        put("att4_timemap", trep.get("videos_found") == trep.get("n"), videos_found=trep.get("videos_found"),
            n=trep.get("n"), text_mapping=trep.get("text_position_mapping"), word_source=trep.get("word_time_source"),
            tokenizer=trep.get("tokenizer_status"))
    rep["all_pass"] = ok_all
    dump_json(rep, run / "selftest_report.json")
    print(f"\n自检{'全部通过' if ok_all else '存在 FAIL'} → {run / 'selftest_report.json'}")
    return rep


def main():
    ap = argparse.ArgumentParser(description="问题3 合成数据自检")
    ap.add_argument("--run", required=True)
    ap.add_argument("--truth", default=None, help="附件2 合成真值 _truth_<version>.json")
    ap.add_argument("--truth4", default=None, help="附件4 合成真值 _truth_<version>.json")
    ap.add_argument("--threads", type=int, default=None)
    a = ap.parse_args()
    if a.threads:
        import torch

        torch.set_num_threads(a.threads)
    rep = check(resolve_path(a.run), resolve_path(a.truth) if a.truth else None,
                resolve_path(a.truth4) if a.truth4 else None, get_device("cpu"))
    raise SystemExit(0 if rep["all_pass"] else 1)


if __name__ == "__main__":
    main()
