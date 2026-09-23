"""问题2 评测：消融表、缺失规律（类型×缺失率 / 位置 / 碎片化 / 混合协议）、退化律拟合、错误归因、作图。

    python -m q2.evaluate --run outputs/q2/mrgnet_aligned --data data/附件2/aligned_50.pkl
    python -m q2.evaluate --run outputs/q2/mrgnet_aligned --tasks law        # 只重跑缺失规律

所有评测掩码都由 (划分的名义有效范围, 协议参数, 固定种子) 唯一确定，每个评测条件先生成一次掩码，
再依次喂给所有模型 —— 各模型比较的是完全相同的缺失输入。
模型取每个变体中验证综合分最高的种子（selected.json）；朴素插补基线 = 线性插值 + no_aug 模型，
其中性偏置在"插补后的验证集混合缺失协议"上单独调（它的输入管线不同）。
"""
from __future__ import annotations

import argparse
import time
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit

from common.data import MODALITIES, MOD_SHORT
from common.error_analysis import error_attribution, plot_error_analysis
from common.metrics import EXTRA_KEYS, MAIN_KEYS
from common.decision import decide
from common.missing import MODALITY_COMBOS
from common.utils import dump_json, get_logger, load_json

from . import plots
from .config import VARIANT_CN, finalize, load_config, parse_set_args, resolve_path
from .engine import (decision_metrics, drop_ratios, load_checkpoint, make_batches, mixed_drop_masks, predict_batches,
                     prepare_data, protocol_drop_masks, resolve_device, setup_threads)
from .train import PROTOCOLS, aggregate_seeds, evaluate_trained, protocol_batches

ROW_KEYS = ("acc3", "f1_weighted", "f1_macro", "mae", "corr", "acc2_has0", "acc2_non0", "acc7")


# ============================================================================ 模型库

class Entry:
    def __init__(self, name, model, bias, impute, seed):
        self.name, self.model, self.bias, self.impute, self.seed = name, model, np.asarray(bias, np.float32), impute, seed


def trained_variants(rdir: Path) -> list[str]:
    return sorted(p.parent.name for p in (rdir / "models").glob("*/selected.json"))


def load_selected(rdir: Path, variant: str, device):
    sel = load_json(rdir / "models" / variant / "selected.json")
    model, ck = load_checkpoint(rdir / "models" / variant / sel["path"], device)
    return model, ck


def naive_impute_eval(cfg, rdir, data, pb, device, logger) -> dict:
    """朴素插补基线：对 base 变体（默认 no_aug）的每个种子做"插补 + 重调偏置"评测，写 models/naive_impute/。"""
    base = cfg["ablation"].get("naive_impute_base", "no_aug")
    bdir = rdir / "models" / base
    if not (bdir / "selected.json").exists():
        logger.warning(f"朴素插补基线需要先训练 {base}，跳过")
        return {}
    per_seed = {}
    for ckp in sorted(bdir.glob("seed*/model.pt")):
        model, ck = load_checkpoint(ckp, device)
        ev, _ = evaluate_trained(model, ck["config"], data, pb, device, impute=True)
        per_seed[int(ck["seed"])] = {"val_score": ck["best_score"], "bias": ev["bias"], "valid": ev["valid"],
                                     "test": ev["test"], "bias_search": ev["bias_search"]}
    sel = load_json(bdir / "selected.json")["selected_seed"]
    ndir = rdir / "models" / "naive_impute"
    for split in ("valid", "test"):
        dump_json({"variant": "naive_impute", "description": VARIANT_CN["naive_impute"], "base_variant": base,
                   "split": split, "seeds": sorted(per_seed), "selected_seed": sel,
                   "selected": {"bias": per_seed[sel]["bias"], **per_seed[sel][split]},
                   "mean_std_over_seeds": aggregate_seeds(per_seed, split),
                   "per_seed": {str(s): {"bias": d["bias"], split: d[split]} for s, d in per_seed.items()}},
                  ndir / f"metrics_{split}.json")
    dump_json({"variant": "naive_impute", "base_variant": base, "selected_seed": sel,
               "bias": {str(s): d["bias"] for s, d in per_seed.items()}}, ndir / "selected.json")
    return per_seed


def load_bank(cfg, rdir, names, data, device, logger) -> dict[str, Entry]:
    bank = {}
    for name in names:
        if name == "naive_impute":
            base = cfg["ablation"].get("naive_impute_base", "no_aug")
            if not (rdir / "models" / base / "selected.json").exists():
                logger.warning("缺少 no_aug 模型，缺失规律评测里跳过 naive_impute")
                continue
            model, ck = load_selected(rdir, base, device)
            nsel = rdir / "models" / "naive_impute" / "selected.json"
            if nsel.exists():
                bias = load_json(nsel)["bias"][str(ck["seed"])]
            else:
                pb = protocol_batches(cfg, data)
                ev, _ = evaluate_trained(model, ck["config"], data, pb, device, impute=True)
                bias = ev["bias"]
            bank[name] = Entry(name, model, bias, True, ck["seed"])
        elif (rdir / "models" / name / "selected.json").exists():
            model, ck = load_selected(rdir, name, device)
            bank[name] = Entry(name, model, ck["bias"], False, ck["seed"])
        else:
            logger.warning(f"变体 {name} 尚未训练，缺失规律评测里跳过")
    return bank


# ============================================================================ 消融表

def build_ablation(cfg, rdir, data, device, logger) -> pd.DataFrame:
    pb = protocol_batches(cfg, data)
    naive_impute_eval(cfg, rdir, data, pb, device, logger)
    rows = []
    names = [v for v in cfg["ablation"]["variants"] if (rdir / "models" / v / "metrics_valid.json").exists()]
    if (rdir / "models" / "naive_impute" / "metrics_valid.json").exists():
        names.append("naive_impute")
    for v in names:
        for split in ("valid", "test"):
            mj = load_json(rdir / "models" / v / f"metrics_{split}.json")
            for proto in PROTOCOLS:
                for stage, tag in (("after_projection", "projected"), ("before_projection", "raw")):
                    ms = mj["mean_std_over_seeds"][proto][stage]
                    row = {"variant": v, "description": VARIANT_CN.get(v, v), "split": split, "protocol": proto,
                           "decision": tag, "n_seeds": len(mj["seeds"]), "selected_seed": mj["selected_seed"]}
                    for k in MAIN_KEYS + EXTRA_KEYS:
                        row[k] = ms[k]["mean"]
                        if k in MAIN_KEYS:
                            row[f"{k}_std"] = ms[k]["std"]
                    rows.append(row)
    df = pd.DataFrame(rows)
    df.to_csv(rdir / "ablation.csv", index=False, encoding="utf-8-sig")
    logger.info(f"消融表 → {rdir / 'ablation.csv'}（{len(df)} 行）")
    return df


# ============================================================================ 缺失规律

def _span_count(drop: np.ndarray) -> np.ndarray:
    """(N,L) 布尔 → 每行连续段数。"""
    d = drop.astype(np.int8)
    starts = (d[:, 1:] > d[:, :-1]).sum(1) + d[:, 0]
    return starts


def _eval_condition(bank, sd, drop, eb, device, base_info: dict, combo_mods=None) -> list[dict]:
    batches = make_batches(sd, drop, eb)
    ratios = drop_ratios(sd, drop)
    info = dict(base_info)
    for m in MODALITIES:
        info[f"miss_{MOD_SHORT[m]}"] = float(ratios[m].mean())
    if drop is not None and combo_mods:
        ns = [_span_count(drop[m])[drop[m].any(1)] for m in combo_mods]
        ns = np.concatenate(ns) if ns else np.array([])
        info["n_spans_actual"] = float(ns.mean()) if ns.size else 0.0
    rows = []
    for name, e in bank.items():
        p = predict_batches(e.model, batches, device, impute=e.impute)
        dm = decision_metrics(sd.y_reg, p, e.bias)
        a, b = dm["after_projection"], dm["before_projection"]
        row = {"model": name, **info, **{k: a[k] for k in ROW_KEYS}, "mae_raw": b["mae"], "corr_raw": b["corr"]}
        for i, m in enumerate(MODALITIES):
            mods = p["modalities"]
            row[f"alpha_{m}"] = float(p["alpha"][:, mods.index(m)].mean()) if ("alpha" in p and m in mods) else np.nan
            row[f"rho_{m}"] = float(p["rho"][:, mods.index(m)].mean()) if m in mods else np.nan
        rows.append(row)
    return rows


def _avg_over_seeds(df: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    num = [c for c in df.columns if c not in keys + ["mask_seed"] and pd.api.types.is_numeric_dtype(df[c])]
    g = df.groupby(keys, sort=False)
    out = g[num].mean()
    std = g[[k for k in ("acc3", "f1_weighted", "mae", "corr") if k in num]].std(ddof=1).fillna(0.0)
    std.columns = [f"{c}_std" for c in std.columns]
    out = out.join(std)  # 按分组索引对齐，不依赖行顺序
    out.insert(0, "n_mask_seeds", g.size())
    return out.reset_index()


def run_law(cfg, rdir, data, device, logger) -> dict:
    law = cfg["protocol"]["law"]
    eb = int(cfg["train"].get("eval_batch_size", 256))
    bank = load_bank(cfg, rdir, law["models"], data, device, logger)
    if not bank:
        raise RuntimeError("没有可评测的模型")
    logger.info(f"缺失规律评测模型：{ {k: f'seed{e.seed}' for k, e in bank.items()} }")
    tr, pos, frag, mix = [], [], [], []
    for split in law["splits"]:
        sd = data[split]
        t0 = time.time()
        # (1) 类型 × 缺失率：r=0 与组合、种子无关 → 只算一次再复制
        r0 = _eval_condition(bank, sd, None, eb, device, {"split": split})
        for combo in law["combos"]:
            for seed in law["mask_seeds"]:
                for r in law["ratios"]:
                    base = {"split": split, "combo": combo, "ratio": float(r), "mask_seed": int(seed)}
                    if r <= 0:
                        tr += [{**row, **base} for row in r0]
                        continue
                    drop = protocol_drop_masks(sd, combo, float(r), None, 1, int(seed))
                    tr += _eval_condition(bank, sd, drop, eb, device, base, MODALITY_COMBOS[combo])
            logger.info(f"[law|{split}] 类型×缺失率 {combo} 完成（累计 {time.time() - t0:.0f}s）")
        # (2) 位置：头/中/尾，单段，位置确定 → 与掩码种子无关
        for combo in law["position_combos"]:
            for r in law["position_ratios"]:
                for p in ("head", "middle", "tail"):
                    drop = protocol_drop_masks(sd, combo, float(r), p, 1, 0)
                    pos += _eval_condition(bank, sd, drop, eb, device,
                                           {"split": split, "combo": combo, "ratio": float(r), "position": p},
                                           MODALITY_COMBOS[combo])
        logger.info(f"[law|{split}] 位置完成（累计 {time.time() - t0:.0f}s）")
        # (3) 碎片化：总缺失率固定，段数变化
        for combo in law["frag_combos"]:
            for r in law["frag_ratios"]:
                for k in law["frag_n_spans"]:
                    for seed in law["mask_seeds"]:
                        drop = protocol_drop_masks(sd, combo, float(r), None, int(k), int(seed))
                        frag += _eval_condition(bank, sd, drop, eb, device,
                                                {"split": split, "combo": combo, "ratio": float(r), "n_spans": int(k),
                                                 "mask_seed": int(seed)}, MODALITY_COMBOS[combo])
        logger.info(f"[law|{split}] 碎片化完成（累计 {time.time() - t0:.0f}s）")
        # (4) 混合协议（仿附件3）
        for seed in law["mixed_seeds"]:
            drop = mixed_drop_masks(sd, cfg["protocol"]["mixed"], int(seed))
            mix += _eval_condition(bank, sd, drop, eb, device, {"split": split, "mask_seed": int(seed)}, MODALITIES)
        logger.info(f"[law|{split}] 混合协议完成（累计 {time.time() - t0:.0f}s）")

    tr_df = pd.DataFrame(tr)
    tr_df.to_csv(rdir / "missing_type_ratio_by_seed.csv", index=False, encoding="utf-8-sig")
    tr_avg = _avg_over_seeds(tr_df, ["model", "split", "combo", "ratio"])
    tr_avg.to_csv(rdir / "missing_type_ratio.csv", index=False, encoding="utf-8-sig")
    pos_df = pd.DataFrame(pos)
    pos_df.to_csv(rdir / "missing_position.csv", index=False, encoding="utf-8-sig")
    frag_df = pd.DataFrame(frag)
    frag_df.to_csv(rdir / "missing_fragmentation_by_seed.csv", index=False, encoding="utf-8-sig")
    frag_avg = _avg_over_seeds(frag_df, ["model", "split", "combo", "ratio", "n_spans"])
    frag_avg.to_csv(rdir / "missing_fragmentation.csv", index=False, encoding="utf-8-sig")
    mix_df = pd.DataFrame(mix)
    mix_avg = _avg_over_seeds(mix_df, ["model", "split"])
    mix_df.to_csv(rdir / "missing_mixed_by_seed.csv", index=False, encoding="utf-8-sig")
    mix_avg.to_csv(rdir / "missing_mixed.csv", index=False, encoding="utf-8-sig")
    fits = degradation_fit(tr_avg, pos_df, frag_avg, mix_avg)
    fits["models"] = {k: {"seed": e.seed, "impute": e.impute, "bias": e.bias.tolist()} for k, e in bank.items()}
    fits["protocol"] = law
    dump_json(fits, rdir / "degradation_fit.json")
    make_law_figures(rdir, tr_avg, pos_df, frag_avg, fits, law)
    return {"type_ratio": tr_avg, "position": pos_df, "fragmentation": frag_avg, "mixed": mix_avg, "fits": fits}


# ============================================================================ 退化律拟合

def _power(r, beta, gamma):
    return beta * np.power(r, gamma)


def fit_power_law(r: np.ndarray, delta: np.ndarray) -> dict:
    """Δ(r) = β·r^γ，只用 r>0 的点：先在对数空间线性回归得初值（仅 Δ>0 的点），再用 curve_fit 做非线性最小二乘。"""
    r, delta = np.asarray(r, float), np.asarray(delta, float)
    m = r > 0
    r, delta = r[m], delta[m]
    if r.size < 2:
        return {"beta": None, "gamma": None, "r2": None, "n": int(r.size)}
    pos = delta > 1e-9
    if pos.sum() >= 2:
        g0, lb0 = np.polyfit(np.log(r[pos]), np.log(delta[pos]), 1)
        p0 = [float(np.exp(lb0)), float(np.clip(g0, 0.05, 10))]
    else:
        p0 = [max(float(delta.max()), 1e-4), 1.0]
    try:
        (beta, gamma), _ = curve_fit(_power, r, delta, p0=p0, bounds=([0.0, 0.05], [np.inf, 10.0]), maxfev=20000)
    except Exception:  # noqa: BLE001 - 退化曲线平坦等情况，退回对数空间初值
        beta, gamma = p0
    pred = _power(r, beta, gamma)
    ss_res = float(((delta - pred) ** 2).sum())
    ss_tot = float(((delta - delta.mean()) ** 2).sum())
    return {"beta": float(beta), "gamma": float(gamma), "r2": (1 - ss_res / ss_tot) if ss_tot > 1e-12 else None,
            "n": int(r.size), "method": "scipy.optimize.curve_fit（初值来自 log Δ = log β + γ log r 的线性回归）"}


def linear_slope(r: np.ndarray, y: np.ndarray) -> dict:
    r, y = np.asarray(r, float), np.asarray(y, float)
    ok = np.isfinite(y)
    if ok.sum() < 2:
        return {"slope": None, "intercept": None, "r2": None}
    b, a = np.polyfit(r[ok], y[ok], 1)
    pred = a + b * r[ok]
    ss_tot = float(((y[ok] - y[ok].mean()) ** 2).sum())
    return {"slope": float(b), "intercept": float(a),
            "r2": (1 - float(((y[ok] - pred) ** 2).sum()) / ss_tot) if ss_tot > 1e-12 else None}


def degradation_fit(tr: pd.DataFrame, pos: pd.DataFrame, frag: pd.DataFrame, mix: pd.DataFrame) -> dict:
    out = {"definition": {
        "power_law": "ΔM(r) = M(r) − M(0) = β·r^γ（M = MAE 投影后 / mae_raw 投影前；r>0 的点，掩码种子平均）",
        "linear": "M(r) ≈ a + b·r 的最小二乘斜率 b（全部 r，含 r=0）",
        "relative": "(M(r) − M(0)) / |M(0)|",
        "auc": "∫_0^1 M(r) dr（梯形法，按实际网格）"}, "fits": {}}
    for (model, split, combo), g in tr.groupby(["model", "split", "combo"]):
        g = g.sort_values("ratio")
        r = g["ratio"].to_numpy(float)
        base = g[g["ratio"] == 0]
        res = {"power_law": {}, "linear": {}, "relative": {}, "auc": {}}
        for met in ("mae", "mae_raw"):
            if len(base):
                res["power_law"][met] = fit_power_law(r, g[met].to_numpy(float) - float(base[met].iloc[0]))
        for met in ("acc3", "f1_weighted", "f1_macro", "corr", "mae"):
            y = g[met].to_numpy(float)
            res["linear"][met] = linear_slope(r, y)
            res["auc"][met] = float(np.trapezoid(y, r)) if r.size > 1 else None
            if len(base):
                b0 = float(base[met].iloc[0])
                res["relative"][met] = {f"{rv:g}": (float(yv) - b0) / max(abs(b0), 1e-9) for rv, yv in zip(r, y)}
        out["fits"].setdefault(model, {}).setdefault(split, {})[combo] = res
    # 汇总结论所需的关键数
    summ = {}
    for (model, split), g in tr.groupby(["model", "split"]):
        full = g[g["ratio"] == g["ratio"].max()]
        base = g[g["ratio"] == 0]
        if base.empty:
            continue
        b0 = base.iloc[0]
        rank = sorted(((c, float(full[full["combo"] == c]["mae"].mean() - b0["mae"])) for c in set(full["combo"])),
                      key=lambda t: -t[1])
        summ.setdefault(model, {})[split] = {
            "baseline_r0": {k: float(b0[k]) for k in ("acc3", "f1_weighted", "mae", "corr")},
            "delta_mae_at_max_ratio_rank": rank,
            "delta_f1_at_max_ratio": {c: float(full[full["combo"] == c]["f1_weighted"].mean() - b0["f1_weighted"])
                                      for c in set(full["combo"])}}
    out["summary"] = summ
    if not pos.empty:
        pe = {}
        for (model, split, combo, r), g in pos.groupby(["model", "split", "combo", "ratio"]):
            d = g.set_index("position")
            pe.setdefault(model, {}).setdefault(split, {})[f"{combo}@{r:g}"] = {
                "mae": {p: float(d.loc[p, "mae"]) for p in d.index},
                "f1_weighted": {p: float(d.loc[p, "f1_weighted"]) for p in d.index},
                "worst_position_by_mae": str(d["mae"].idxmax()), "mae_range": float(d["mae"].max() - d["mae"].min())}
        out["position_effect"] = pe
    if not frag.empty:
        fe = {}
        for (model, split, combo, r), g in frag.groupby(["model", "split", "combo", "ratio"]):
            g = g.sort_values("n_spans")
            fe.setdefault(model, {}).setdefault(split, {})[f"{combo}@{r:g}"] = {
                "n_spans": g["n_spans"].tolist(), "n_spans_actual": g.get("n_spans_actual", pd.Series()).tolist(),
                "mae": g["mae"].tolist(), "f1_weighted": g["f1_weighted"].tolist(),
                "mae_slope_per_span": linear_slope(g["n_spans"].to_numpy(float), g["mae"].to_numpy(float))["slope"],
                "f1_slope_per_span": linear_slope(g["n_spans"].to_numpy(float), g["f1_weighted"].to_numpy(float))["slope"]}
        out["fragmentation_effect"] = fe
    if not mix.empty:
        out["mixed_protocol"] = {f"{r['model']}|{r['split']}": {k: float(r[k]) for k in ("acc3", "f1_weighted", "mae", "corr")}
                                 for _, r in mix.iterrows()}
    return out


def make_law_figures(rdir: Path, tr, pos, frag, fits, law) -> None:
    fd = rdir / "figures"
    for split in law["splits"]:
        plots.plot_type_ratio(tr, "full", split, fd / f"law_type_ratio_full_{split}.png")
        plots.plot_models_vs_ratio(tr, split, ["T", "TAV"], fd / f"law_models_vs_ratio_{split}.png")
        plots.plot_models_vs_ratio(tr, split, ["A", "V"], fd / f"law_models_vs_ratio_AV_{split}.png")
        plots.plot_position(pos, "full", split, fd / f"law_position_full_{split}.png")
        plots.plot_fragmentation(frag, "full", split, fd / f"law_fragmentation_full_{split}.png")
        plots.plot_gate_alpha(tr, split, fd / f"gate_alpha_vs_ratio_{split}.png")
        plots.plot_degradation_fit(tr, fits["fits"], "full", split, fd / f"degradation_fit_full_{split}.png")


# ============================================================================ 错误归因

def _ratio_bucket(r: np.ndarray) -> np.ndarray:
    return np.select([r <= 0, r <= 0.3, r <= 0.6], ["0", "(0,0.3]", "(0.3,0.6]"], ">0.6").astype(str)


def run_error_analysis(cfg, rdir, data, device, logger) -> dict:
    model, ck = load_checkpoint(rdir / "model.pt", device)
    bias = np.asarray(ck["bias"], np.float32)
    mixed = cfg["protocol"]["mixed"]
    eb = int(cfg["train"].get("eval_batch_size", 256))
    out = {"model": {"variant": ck.get("variant"), "seed": ck.get("seed"), "bias": ck["bias"]},
           "note": "预测 = 偏置后 argmax + 强度一致性投影；协变量为该条件下各模态在名义有效范围内的总缺失率"}
    fd = rdir / "figures"
    scatter = []
    for split, seed_key in (("valid", "valid_seed"), ("test", "test_seed")):
        sd = data[split]
        for proto in ("complete", "mixed"):
            drop = mixed_drop_masks(sd, mixed, int(mixed[seed_key])) if proto == "mixed" else None
            p = predict_batches(model, make_batches(sd, drop, eb), device)
            c, y = decide(p["logits"], p["y_raw"], bias, project=True)
            rat = drop_ratios(sd, drop)
            cov = {f"miss_ratio_{m}": _ratio_bucket(rat[m]) for m in MODALITIES}
            combo = np.array(["+".join(MOD_SHORT[m] for m in MODALITIES if rat[m][i] > 0) or "none"
                              for i in range(len(sd))])
            cov["missing_combo"] = combo
            cov["n_missing_modalities"] = np.array([str(sum(rat[m][i] > 0 for m in MODALITIES)) for i in range(len(sd))])
            ea = error_attribution(sd.ids, sd.raw_text, sd.y_reg, y, c, covariates=cov)
            key = f"{split}_{proto}"
            out[key] = ea
            plot_error_analysis(ea, fd / f"error_analysis_{key}.png",
                                title=f"错误归因：{split}（{'完整输入' if proto == 'complete' else '混合缺失协议'}）")
            if proto == "mixed":
                plots.plot_error_covariates(ea, fd / f"error_by_missing_{key}.png",
                                            title=f"按缺失程度分桶的误差（{split}，混合缺失协议）")
            if split == "valid":
                scatter.append((f"验证集 {'完整输入' if proto == 'complete' else '混合缺失'}", sd.y_reg, y, c))
    plots.plot_scatter(scatter, fd / "valid_scatter.png", title="验证集：预测强度 vs 真实强度（颜色 = 预测极性）")
    dump_json(out, rdir / "error_analysis.json")
    logger.info(f"错误归因 → {rdir / 'error_analysis.json'}")
    return out


def run_training_figures(rdir: Path) -> None:
    if (rdir / "train_log.csv").exists():
        plots.plot_train_curves(pd.read_csv(rdir / "train_log.csv"), rdir / "figures" / "train_curves_full.png",
                                title="MRG-Net 训练曲线")
    if (rdir / "ablation.csv").exists():
        abl = pd.read_csv(rdir / "ablation.csv")
        for split in ("valid", "test"):
            plots.plot_ablation(abl, rdir / "figures" / f"ablation_{split}.png", split=split)


def main():
    ap = argparse.ArgumentParser(description="问题2：消融表 / 缺失规律 / 错误归因 / 作图")
    ap.add_argument("--run", required=True, help="训练输出目录 outputs/q2/<run>")
    ap.add_argument("--data", default=None, help="附件2 路径（缺省用运行目录 config.yaml 里的）")
    ap.add_argument("--tasks", nargs="+", default=["ablation", "law", "error", "figures"],
                    choices=["ablation", "law", "error", "figures"])
    ap.add_argument("--threads", type=int, default=None)
    ap.add_argument("--set", nargs="*", default=None, metavar="KEY=VALUE", help="点路径覆盖（如 protocol.law.splits=[test]）")
    a = ap.parse_args()
    rdir = resolve_path(a.run)
    cfg = finalize(load_config(rdir / "config.yaml"), sets=parse_set_args(a.set), data=a.data, threads=a.threads)
    logger = get_logger("q2", rdir / "logs" / "q2.log")
    setup_threads(cfg)
    device = resolve_device(cfg)
    data, _ = prepare_data(cfg, logger)
    run_tasks(cfg, rdir, data, device, logger, a.tasks)


def run_tasks(cfg, rdir, data, device, logger, tasks) -> None:
    if "ablation" in tasks:
        build_ablation(cfg, rdir, data, device, logger)
    if "law" in tasks:
        run_law(cfg, rdir, data, device, logger)
    if "error" in tasks:
        run_error_analysis(cfg, rdir, data, device, logger)
    if "figures" in tasks:
        run_training_figures(rdir)


if __name__ == "__main__":
    main()
