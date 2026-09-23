"""问题2 训练入口：训练一个或多个变体（主模型 full / 消融 / 基线），每个变体可多种子。

    python -m q2.train --config configs/q2.yaml --data data/附件2/aligned_50.pkl --variants full --seeds 1 2 3
    python -m q2.train --config configs/q2.yaml --variants no_aug no_rec no_kd no_gate late_fusion

每个 (变体, 种子)：训练 + 早停 → 在验证集混合缺失协议的 logits 上调中性类偏置 →
在验证/测试集的"完整输入"和"固定混合缺失协议"下给出投影前后两套指标。
产物：outputs/q2/<run>/models/<variant>/seed<k>/{model.pt, train_log.csv}、
      models/<variant>/{metrics_valid.json, metrics_test.json, selected.json, config.yaml}；
      变体 full 额外把选中种子的 model.pt / train_log.csv / metrics_*.json 放到运行目录根部（按接口合同）。
"""
from __future__ import annotations

import argparse
import shutil
import time
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import MODALITIES
from common.metrics import EXTRA_KEYS, MAIN_KEYS
from common.utils import dump_json, get_logger

from .config import (VARIANT_CN, dump_yaml, finalize, load_config, parse_set_args, run_dir, variant_config)
from .engine import (decision_metrics, make_batches, mixed_drop_masks, predict_batches, prepare_data,
                     resolve_device, save_checkpoint, seq_lens_of, setup_threads, train_one, tune_bias)
from .model import resolve_strides

METRIC_KEYS = MAIN_KEYS + EXTRA_KEYS
PROTOCOLS = ("complete", "mixed")


def protocol_batches(cfg: dict, data: dict) -> dict:
    """验证/测试集 × {完整, 固定混合缺失} 的 batch 缓存（所有变体、所有种子共用同一套掩码）。"""
    mixed = cfg["protocol"]["mixed"]
    eb = int(cfg["train"].get("eval_batch_size", 256))
    out = {}
    for split, seed_key in (("valid", "valid_seed"), ("test", "test_seed")):
        sd = data[split]
        out[(split, "complete")] = make_batches(sd, None, eb)
        drop = mixed_drop_masks(sd, mixed, int(mixed[seed_key]))
        out[(split, "mixed")] = make_batches(sd, drop, eb)
        out[(split, "mixed_drop")] = drop
    return out


def mean_std(values: list[float]) -> dict:
    a = np.asarray([v for v in values if v is not None and np.isfinite(v)], np.float64)
    if a.size == 0:
        return {"mean": None, "std": None, "n": 0}
    return {"mean": float(a.mean()), "std": float(a.std(ddof=1)) if a.size > 1 else 0.0, "n": int(a.size)}


def aggregate_seeds(per_seed: dict, split: str) -> dict:
    out = {}
    for proto in PROTOCOLS:
        out[proto] = {}
        for stage in ("after_projection", "before_projection"):
            out[proto][stage] = {k: mean_std([per_seed[s][split][proto][stage].get(k) for s in per_seed])
                                 for k in METRIC_KEYS}
    return out


def evaluate_trained(model, cfg: dict, data: dict, pb: dict, device, impute: bool = False) -> dict:
    """调偏置（验证集混合协议）+ 四种评测条件下的投影前后指标。"""
    preds = {k: predict_batches(model, pb[k], device, impute=impute) for k in pb if k[1] != "mixed_drop"}
    bias, search = tune_bias(cfg, data["valid"].y_cls, preds[("valid", "mixed")]["logits"])
    res = {"bias": bias.tolist(), "bias_search": search}
    for split in ("valid", "test"):
        res[split] = {proto: decision_metrics(data[split].y_reg, preds[(split, proto)], bias) for proto in PROTOCOLS}
    return res, preds


def train_variant(cfg: dict, variant: str, seeds: list[int], rdir: Path, data: dict, normalizer, device, logger,
                  pb: dict | None = None) -> dict:
    vcfg = resolve_strides(variant_config(cfg, variant), seq_lens_of(data["train"]))
    vdir = rdir / "models" / variant
    vdir.mkdir(parents=True, exist_ok=True)
    dump_yaml(vcfg, vdir / "config.yaml")
    pb = pb or protocol_batches(cfg, data)
    per_seed, logs = {}, []
    for s in seeds:
        t0 = time.time()
        res = train_one(vcfg, data, int(s), device, logger)
        model = res["model"]
        ev, _ = evaluate_trained(model, vcfg, data, pb, device)
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
                            "bias": ev["bias"], "bias_search": ev["bias_search"], "valid": ev["valid"], "test": ev["test"],
                            "train_time_s": time.time() - t0}
        a = ev["valid"]["mixed"]["after_projection"]
        logger.info(f"[{variant}|seed={s}] 最佳轮 {res['best_epoch']}，验证分 {res['best_score']:.4f}，中性偏置 "
                    f"{ev['bias'][1]:+.2f}；验证混合缺失(投影后) Acc={a['acc3']:.3f} F1={a['f1_weighted']:.3f} "
                    f"MAE={a['mae']:.3f} r={a['corr']:.3f}")
    sel = max(per_seed, key=lambda k: per_seed[k]["val_score"])
    pd.concat(logs, ignore_index=True).to_csv(vdir / "train_log.csv", index=False, encoding="utf-8-sig")
    mixed = cfg["protocol"]["mixed"]
    for split in ("valid", "test"):
        out = {"variant": variant, "description": VARIANT_CN.get(variant, variant), "version": cfg["version"],
               "split": split, "seeds": [int(s) for s in seeds], "selected_seed": sel,
               "selection_rule": "验证综合分 ½[composite(完整)+composite(混合缺失)] 最高的种子",
               "protocols": {"complete": "完整输入（仅数据自带的不可用位置）",
                             "mixed": {**mixed, "seed_used": int(mixed["valid_seed" if split == "valid" else "test_seed"])}},
               "decision_rule": "极性 = argmax(ℓ + (0,β,0))，β 在验证集混合缺失协议上网格搜索（加权F1最大）；"
                                "强度投影 Negative→min(ŷ,-0.01)，Neutral→0，Positive→max(ŷ,0.01)",
               "selected": {"bias": per_seed[sel]["bias"], "best_epoch": per_seed[sel]["best_epoch"],
                            **per_seed[sel][split]},
               "mean_std_over_seeds": aggregate_seeds(per_seed, split),
               "per_seed": {str(s): {k: v for k, v in d.items() if k not in ("valid", "test", "bias_search")}
                            | {split: d[split]} for s, d in per_seed.items()}}
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
    ap.add_argument("--config", default=None, help="YAML 配置（缺省 configs/q2.yaml）")
    ap.add_argument("--data", default=None, help="附件2 特征文件路径")
    ap.add_argument("--version", default=None, choices=["aligned", "unaligned"])
    ap.add_argument("--run-name", default=None)
    ap.add_argument("--out", default=None, help="运行目录（缺省 outputs/q2/<run_name>）")
    ap.add_argument("--epochs", type=int, default=None)
    ap.add_argument("--threads", type=int, default=None)
    ap.add_argument("--quick", action="store_true", help="叠加配置里的 quick 段（仅合成数据跑通用）")
    ap.add_argument("--set", nargs="*", default=None, metavar="KEY=VALUE", help="点路径覆盖，如 train.lr=1e-4")


def build_cfg(a, extra_seeds=None) -> dict:
    return finalize(load_config(a.config), quick=a.quick, sets=parse_set_args(a.set), version=a.version, data=a.data,
                    att3=getattr(a, "att3", None), run_name=a.run_name, epochs=a.epochs, seeds=extra_seeds,
                    threads=a.threads)


def main():
    ap = argparse.ArgumentParser(description="问题2：训练 MRG-Net 及消融/基线变体")
    add_common_args(ap)
    ap.add_argument("--variants", nargs="+", default=["full"], help="要训练的变体（见配置 ablation.variants）")
    ap.add_argument("--seeds", nargs="+", type=int, default=None, help="覆盖所有变体的种子列表")
    a = ap.parse_args()
    cfg = build_cfg(a)
    rdir = run_dir(cfg, a.out)
    rdir.mkdir(parents=True, exist_ok=True)
    logger = get_logger("q2", rdir / "logs" / "q2.log")
    setup_threads(cfg)
    device = resolve_device(cfg)
    data, normalizer = prepare_data(cfg, logger)
    cfg = resolve_strides(cfg, seq_lens_of(data["train"]))
    if "full" in a.variants or not (rdir / "config.yaml").exists():
        dump_yaml(cfg, rdir / "config.yaml")
    pb = protocol_batches(cfg, data)
    for v in a.variants:
        train_variant(cfg, v, seeds_for(cfg, v, a.seeds), rdir, data, normalizer, device, logger, pb)
    logger.info(f"训练完成：{rdir}")


if __name__ == "__main__":
    main()
