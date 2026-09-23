"""问题2 合成数据自检：核对一次 run_all 的产物是否齐全、自洽，并与合成真值对照。

    python -m q2.selftest --run outputs/q2/synth_aligned --truth outputs/synthetic/附件3/_truth_aligned.json

检查项（合成数据上的数值只说明代码跑通，不能进论文）：
    1. 接口合同要求的文件全部存在；图全部生成且非空
    2. 训练损失下降（每个变体、每个种子：最后一轮 < 第一轮）
    3. 缺失条件下完整模型 vs 无增广模型（消融表混合协议 + 缺失率网格 r≥0.5 的平均）——报告数值
    4. 附件3 缺失检测与真值一致：文本/语音逐区间完全一致；视觉允许多出"天然零行"（检测 ⊇ 真值）
    5. 附件3 CSV：列名与顺序、行数、UTF-8-SIG、极性/强度/概率自洽
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import MODALITIES
from common.utils import dump_json

from .config import resolve_path
from .predict import CSV_COLUMNS, CSV_NAME, SUMMARY_NAME

REQUIRED = ["model.pt", "config.yaml", "train_log.csv", "metrics_valid.json", "metrics_test.json", "ablation.csv",
            "missing_type_ratio.csv", "missing_position.csv", "missing_fragmentation.csv", "degradation_fit.json",
            "error_analysis.json", CSV_NAME, SUMMARY_NAME]


def _spans(s) -> set:
    if not isinstance(s, str) or not s:
        return set()
    return {tuple(int(v) for v in p.split("-")) for p in s.split(";")}


def check(run: Path, truth_path: Path | None) -> dict:
    rep = {"run": str(run), "checks": {}}
    ok_all = True

    def put(name, ok, **info):
        nonlocal ok_all
        rep["checks"][name] = {"status": "PASS" if ok is True else ("WARN" if ok is None else "FAIL"), **info}
        if ok is False:
            ok_all = False
        print(f"[{rep['checks'][name]['status']}] {name}: {json.dumps(info, ensure_ascii=False, default=str)[:600]}")

    # 1) 文件
    missing = [f for f in REQUIRED if not (run / f).exists()]
    put("required_files", not missing, missing=missing)
    figs = sorted((run / "figures").glob("*.png"))
    empty = [f.name for f in figs if f.stat().st_size < 5000]
    put("figures", len(figs) >= 15 and not empty, n_png=len(figs), too_small=empty, names=[f.name for f in figs])

    # 2) 损失下降
    dec = {}
    for log in sorted((run / "models").glob("*/train_log.csv")):
        df = pd.read_csv(log)
        for s, g in df.groupby("seed"):
            dec[f"{log.parent.name}|seed{s}"] = (float(g["train_loss"].iloc[0]), float(g["train_loss"].iloc[-1]),
                                                 float(g["train_loss"].min()), int(len(g)))
    bad = {k: v for k, v in dec.items() if not v[1] < v[0]}
    put("loss_decreases", bool(dec) and not bad, first_last_min_epochs=dec, not_decreasing=bad)

    # 3) 完整模型 vs 无增广（缺失条件）
    abl = pd.read_csv(run / "ablation.csv")
    a = abl[(abl["decision"] == "projected")].set_index(["variant", "split", "protocol"])
    cmp = {}
    for split in ("valid", "test"):
        for proto in ("complete", "mixed"):
            for v in ("full", "no_aug", "naive_impute", "late_fusion"):
                if (v, split, proto) in a.index:
                    r = a.loc[(v, split, proto)]
                    cmp[f"{v}|{split}|{proto}"] = {k: round(float(r[k]), 4) for k in ("acc3", "f1_weighted", "mae", "corr")}
    tr = pd.read_csv(run / "missing_type_ratio.csv")
    hi = tr[(tr["ratio"] >= 0.5) & (tr["split"] == "test")]
    grid = {m: {k: round(float(hi[hi["model"] == m][k].mean()), 4) for k in ("acc3", "f1_weighted", "mae", "corr")}
            for m in sorted(set(hi["model"]))}
    wins = None
    if "full" in grid and "no_aug" in grid:
        f, n = grid["full"], grid["no_aug"]
        wins = {"mae": f["mae"] < n["mae"], "f1_weighted": f["f1_weighted"] > n["f1_weighted"],
                "corr": f["corr"] > n["corr"], "acc3": f["acc3"] > n["acc3"]}
    # 软检查：合成数据小、轮数少，赢不了只记 WARN（数值照实报告），不算代码错误
    put("full_vs_no_aug_under_missing", True if wins and sum(wins.values()) >= 3 else None,
        ablation_projected=cmp, grid_test_ratio_ge_0_5_mean=grid, full_better_on=wins)

    # 4) 附件3 缺失检测 vs 真值
    det = pd.read_csv(run / "附件3_缺失检测明细.csv")
    if truth_path and truth_path.exists():
        truth = json.load(open(truth_path, encoding="utf-8"))
        res = {}
        for m in MODALITIES:
            d = det[det["modality"] == m].set_index("id")
            exact, superset, n, err = 0, 0, 0, []
            for sid, t in truth.items():
                tsp = {tuple(x) for x in t.get("missing_spans", {}).get(m, [])}
                L = int(t["lengths"][m])
                tr_ratio = sum(e - s for s, e in tsp) / max(L, 1)
                got = _spans(d.loc[sid, "spans"]) if sid in d.index else set()
                cover_true = set()
                for s, e in tsp:
                    cover_true |= set(range(s, e))
                cover_got = set()
                for s, e in got:
                    cover_got |= set(range(s, e))
                n += 1
                exact += int(cover_got == cover_true)
                superset += int(cover_true <= cover_got)
                err.append(float(d.loc[sid, "ratio"]) - tr_ratio if sid in d.index else np.nan)
            err = np.asarray(err)
            res[m] = {"n": n, "exact_position_match": exact, "detected_superset_of_truth": superset,
                      "ratio_err_mean": float(np.nanmean(err)), "ratio_err_max_abs": float(np.nanmax(np.abs(err)))}
        ok = (res["text"]["exact_position_match"] == res["text"]["n"] and
              res["audio"]["exact_position_match"] == res["audio"]["n"] and
              res["vision"]["detected_superset_of_truth"] == res["vision"]["n"])
        put("att3_missing_detection_vs_truth", ok, per_modality=res,
            note="视觉合成数据含 ~5% 天然零行（未检出人脸），检测结果必然 ⊇ 真值")
    else:
        put("att3_missing_detection_vs_truth", None, note="未提供真值文件")

    # 5) CSV
    raw = (run / CSV_NAME).read_bytes()
    df = pd.read_csv(run / CSV_NAME, encoding="utf-8-sig")
    summ = json.load(open(run / SUMMARY_NAME, encoding="utf-8"))
    cols_ok = list(df.columns) == CSV_COLUMNS
    n_ok = len(df) == summ["n"] and (truth_path is None or not truth_path.exists() or len(df) == len(truth))
    probs = df[["prob_negative", "prob_neutral", "prob_positive"]].to_numpy()
    cons = (((df["pred_class"] == 0) & (df["pred_intensity"] < 0)) | ((df["pred_class"] == 1) & (df["pred_intensity"] == 0))
            | ((df["pred_class"] == 2) & (df["pred_intensity"] > 0))).all()
    name_ok = (df["pred_polarity"] == df["pred_class"].map({0: "Negative", 1: "Neutral", 2: "Positive"})).all() and \
              (df["pred_polarity_cn"] == df["pred_class"].map({0: "负向", 1: "中性", 2: "正向"})).all()
    argmax_ok = bool((probs.argmax(1) == df["pred_class"].to_numpy()).all())
    ids_ok = (df["id"] == df["video_id"] + "$_$" + df["clip_id"].astype(str)).all()
    put("att3_csv", bool(cols_ok and n_ok and cons and name_ok and argmax_ok and ids_ok and raw[:3] == b"\xef\xbb\xbf"),
        columns_ok=cols_ok, rows=len(df), rows_ok=n_ok, utf8_sig=raw[:3] == b"\xef\xbb\xbf", polarity_intensity_consistent=bool(cons),
        names_ok=bool(name_ok), prob_argmax_ok=argmax_ok, prob_sum_max_dev=float(np.abs(probs.sum(1) - 1).max()),
        id_ok=bool(ids_ok), class_counts=summ["class_counts"], combo_frequency=summ["combo_frequency"])
    rep["all_pass"] = ok_all
    dump_json(rep, run / "selftest_report.json")
    print("SELFTEST", "PASS" if ok_all else "FAIL", "→", run / "selftest_report.json")
    return rep


def main():
    ap = argparse.ArgumentParser(description="问题2 合成数据自检")
    ap.add_argument("--run", required=True)
    ap.add_argument("--truth", default=None, help="合成附件3 真值 _truth_<version>.json")
    a = ap.parse_args()
    rep = check(resolve_path(a.run), resolve_path(a.truth) if a.truth else None)
    sys.exit(0 if rep["all_pass"] else 1)


if __name__ == "__main__":
    main()
