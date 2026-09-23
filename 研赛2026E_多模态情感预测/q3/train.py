"""问题3 训练入口：训练主模型 AMEE（多种子）及消融 / 不可解释对照变体。

    python -m q3.train --config configs/q3.yaml --data data/附件2/aligned_50.pkl --variants full --seeds 1 2 3
    python -m q3.train --variants no_interaction no_uni no_sparse concat

每个 (变体, 种子)：训练 + 早停（验证综合分）→ 在验证集 logits 上调中性类偏置 → 验证/测试集投影前后两套指标。
产物：outputs/q3/<run>/models/<variant>/seed<k>/{model.pt, train_log.csv}、
      models/<variant>/{metrics_valid.json, metrics_test.json, selected.json, config.yaml}；
      变体 full 把选中种子的 model.pt / train_log.csv / metrics_*.json 复制到运行目录根部（接口合同）。
"""
from __future__ import annotations

import argparse
import shutil
import time
from pathlib import Path

import numpy as np
import pandas as pd

from common.metrics import EXTRA_KEYS, MAIN_KEYS
from common.utils import dump_json, get_logger

from .config import VARIANT_CN, dump_yaml, finalize, load_config, parse_set_args, run_dir, variant_config
from .engine import (PooledSplit, decision_metrics, predict_pooled, prepare_data, resolve_device, save_checkpoint,
                     seq_lens_of, setup_threads, train_one, tune_bias)
from .model import resolve_strides

METRIC_KEYS = MAIN_KEYS + EXTRA_KEYS


def mean_std(values) -> dict:
    a = np.asarray([v for v in values if v is not None and np.isfinite(v)], np.float64)
    if a.size == 0:
        return {"mean": None, "std": None, "n": 0}
    return {"mean": float(a.mean()), "std": float(a.std(ddof=1)) if a.size > 1 else 0.0, "n": int(a.size)}


def evaluate_trained(model, cfg: dict, data: dict, pooled: dict, device) -> dict:
    """完整输入：验证集 logits 上调中性类偏置 → 验证 / 测试投影前后指标。"""
    eb = int(cfg["train"].get("eval_batch_size", 256))
    preds = {s: predict_pooled(model, pooled[s], device, eb) for s in ("valid", "test")}
    bias, search = tune_bias(cfg, data["valid"].y_cls, preds["valid"]["logits"])
    res = {"bias": bias.tolist(), "bias_search": search}
    for s in ("valid", "test"):
        res[s] = decision_metrics(data[s].y_reg, preds[s], bias)
    return res


def pooled_splits(data: dict, strides: dict) -> dict:
    return {s: PooledSplit(data[s], strides) for s in ("valid", "test")}


def train_variant(cfg: dict, variant: str, seeds: list[int], rdir: Path, data: dict, normalizer, device, logger,
                  pooled: dict | None = None) -> dict:
    vcfg = variant_config(cfg, variant)
    vdir = rdir / "models" / variant
    vdir.mkdir(parents=True, exist_ok=True)
    dump_yaml(vcfg, vdir / "config.yaml")
    strides = vcfg["model"]["temporal_stride"]
    pooled = pooled if (pooled and pooled["valid"].strides == strides) else pooled_splits(data, strides)
    per_seed, logs = {}, []
    for s in seeds:
        t0 = time.time()
        res = train_one(vcfg, data, int(s), device, logger, pooled_valid=pooled["valid"])
        model = res["model"]
        ev = evaluate_trained(model, vcfg, data, pooled, device)
        sdir = vdir / f"seed{s}"
        save_checkpoint(sdir / "model.pt", model, vcfg, normalizer, ev["bias"],
                        {"seed": int(s), "best_epoch": res["best_epoch"], "best_score": res["best_score"],
                         "n_params": res["n_params"], "bias_search": ev["bias_search"]})
        log = pd.DataFrame(res["history"])
        log.insert(0, "seed", int(s))
        log.insert(0, "variant", variant)
        log.to_csv(sdir / "train_log.csv", index=False, encoding="utf-8-sig")
        logs.append(log)
        per_seed[int(s)] = {"best_epoch": res["best_epoch"], "val_score": res["best_score"], "n_params": res["n_params"],
                            "bias": ev["bias"], "bias_search": ev["bias_search"], "valid": ev["valid"],
                            "test": ev["test"], "train_time_s": time.time() - t0}
        a = ev["valid"]["after_projection"]
        logger.info(f"[{variant}|seed={s}] 最佳轮 {res['best_epoch']}，验证分 {res['best_score']:.4f}，中性偏置 "
                    f"{ev['bias'][1]:+.2f}；验证(投影后) Acc={a['acc3']:.3f} F1={a['f1_weighted']:.3f} "
                    f"MAE={a['mae']:.3f} r={a['corr']:.3f}")
    sel = max(per_seed, key=lambda k: per_seed[k]["val_score"])
    pd.concat(logs, ignore_index=True).to_csv(vdir / "train_log.csv", index=False, encoding="utf-8-sig")
    for split in ("valid", "test"):
        out = {"variant": variant, "description": VARIANT_CN.get(variant, variant), "version": cfg["version"],
               "split": split, "input": "完整三模态输入（附件2 原样）", "seeds": [int(s) for s in seeds],
               "selected_seed": sel, "selection_rule": "验证综合分 composite = F1_w + Acc3 + Corr − MAE 最高的种子",
               "decision_rule": "极性 = argmax(ℓ + (0,β,0))，β 在验证集上网格搜索（加权F1最大）；"
                                "强度投影 Negative→min(ŷ,-0.01)，Neutral→0，Positive→max(ŷ,0.01)",
               "selected": {"bias": per_seed[sel]["bias"], "best_epoch": per_seed[sel]["best_epoch"],
                            "n_params": per_seed[sel]["n_params"], **per_seed[sel][split]},
               "mean_std_over_seeds": {stage: {k: mean_std([per_seed[s][split][stage].get(k) for s in per_seed])
                                               for k in METRIC_KEYS}
                                       for stage in ("after_projection", "before_projection")},
               "per_seed": {str(s): {"best_epoch": d["best_epoch"], "val_score": d["val_score"], "bias": d["bias"],
                                     "train_time_s": d["train_time_s"], split: d[split]} for s, d in per_seed.items()}}
        if split == "valid":
            out["bias_search_selected"] = per_seed[sel]["bias_search"]
        dump_json(out, vdir / f"metrics_{split}.json")
    dump_json({"variant": variant, "selected_seed": sel, "path": f"seed{sel}/model.pt",
               "val_score": per_seed[sel]["val_score"]}, vdir / "selected.json")
    if variant == "full":
        shutil.copy2(vdir / f"seed{sel}" / "model.pt", rdir / "model.pt")
        shutil.copy2(vdir / "train_log.csv", rdir / "train_log.csv")
        for split in ("valid", "test"):
            shutil.copy2(vdir / f"metrics_{split}.json", rdir / f"metrics_{split}.json")
    return {"variant": variant, "selected_seed": sel, "per_seed": per_seed}


def seeds_for(cfg: dict, variant: str, override: list[int] | None = None) -> list[int]:
    if override:
        return list(override)
    return list(cfg["train"]["seeds"]) if variant == "full" else list(cfg["ablation"].get("seeds", [1]))


def add_common_args(ap: argparse.ArgumentParser) -> None:
    ap.add_argument("--config", default=None, help="YAML 配置（缺省 configs/q3.yaml）")
    ap.add_argument("--data", default=None, help="附件2 特征文件路径")
    ap.add_argument("--version", default=None, choices=["aligned", "unaligned"])
    ap.add_argument("--run-name", default=None)
    ap.add_argument("--out", default=None, help="运行目录（缺省 outputs/q3/<run_name>）")
    ap.add_argument("--epochs", type=int, default=None)
    ap.add_argument("--threads", type=int, default=None)
    ap.add_argument("--quick", action="store_true", help="叠加配置里的 quick 段（仅合成数据跑通用）")
    ap.add_argument("--set", nargs="*", default=None, metavar="KEY=VALUE", help="点路径覆盖，如 train.lr=1e-4")


def build_cfg(a, extra_seeds=None) -> dict:
    return finalize(load_config(a.config), quick=a.quick, sets=parse_set_args(a.set), version=a.version, data=a.data,
                    att4=getattr(a, "att4", None), videos=getattr(a, "videos", None), run_name=a.run_name,
                    epochs=a.epochs, seeds=extra_seeds, threads=a.threads)


def main():
    ap = argparse.ArgumentParser(description="问题3：训练 AMEE 及消融/对照变体")
    add_common_args(ap)
    ap.add_argument("--variants", nargs="+", default=["full"], help="要训练的变体（见配置 ablation.variants）")
    ap.add_argument("--seeds", nargs="+", type=int, default=None, help="覆盖所有变体的种子列表")
    a = ap.parse_args()
    cfg = build_cfg(a)
    rdir = run_dir(cfg, a.out)
    rdir.mkdir(parents=True, exist_ok=True)
    logger = get_logger("q3", rdir / "logs" / "q3.log")
    setup_threads(cfg)
    device = resolve_device(cfg)
    data, normalizer = prepare_data(cfg, logger)
    cfg = resolve_strides(cfg, seq_lens_of(data["train"]))
    if "full" in a.variants or not (rdir / "config.yaml").exists():
        dump_yaml(cfg, rdir / "config.yaml")
    pooled = pooled_splits(data, cfg["model"]["temporal_stride"])
    for v in a.variants:
        train_variant(cfg, v, seeds_for(cfg, v, a.seeds), rdir, data, normalizer, device, logger, pooled)
    logger.info(f"训练完成：{rdir}")


if __name__ == "__main__":
    main()
