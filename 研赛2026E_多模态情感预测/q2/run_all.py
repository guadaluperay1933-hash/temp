"""问题2 一键运行：训练主模型与全部消融/基线 → 消融表 / 缺失规律 / 退化律 / 错误归因 / 图 → 附件3 推理。

    python -m q2.run_all --config configs/q2.yaml --data data/附件2/aligned_50.pkl --att3 data/附件3/<文件>.pkl
    python -m q2.run_all --version unaligned --data data/附件2/unaligned_50.pkl --att3 data/附件3/<文件>.pkl
    # 合成数据跑通（数值无意义）：
    python -m q2.run_all --data outputs/synthetic/附件2/aligned_50.pkl --att3 outputs/synthetic/附件3/aligned_50.pkl \\
        --run-name synth_aligned --quick --threads 2

等价的分步命令：q2.train（--variants ...）→ q2.evaluate（--tasks ablation law error figures）→ q2.predict。
"""
from __future__ import annotations

import argparse
import sys
import time

from common.utils import dump_json, env_info, get_logger

from .config import dump_yaml, run_dir
from .engine import prepare_data, resolve_device, seq_lens_of, setup_threads
from .model import resolve_strides
from .evaluate import run_tasks
from .predict import run_predict
from .train import add_common_args, build_cfg, protocol_batches, seeds_for, train_variant


def main():
    ap = argparse.ArgumentParser(description="问题2：训练 + 消融 + 缺失规律 + 附件3 推理（一键）")
    add_common_args(ap)
    ap.add_argument("--att3", default=None, help="附件3 特征文件（缺省取配置 att3；为空则跳过推理）")
    ap.add_argument("--att3-split", default=None)
    ap.add_argument("--variants", nargs="+", default=None, help="只训练这些变体（缺省 = 配置里全部）")
    ap.add_argument("--seeds", nargs="+", type=int, default=None, help="主模型 full 的种子（覆盖 train.seeds）")
    ap.add_argument("--skip-train", action="store_true", help="已训练过：直接评测与推理")
    ap.add_argument("--tasks", nargs="+", default=["ablation", "law", "error", "figures"],
                    choices=["ablation", "law", "error", "figures"])
    a = ap.parse_args()
    cfg = build_cfg(a, extra_seeds=a.seeds)
    if a.att3_split:
        cfg["att3_split"] = a.att3_split
    rdir = run_dir(cfg, a.out)
    rdir.mkdir(parents=True, exist_ok=True)
    logger = get_logger("q2", rdir / "logs" / "q2.log")
    setup_threads(cfg)
    device = resolve_device(cfg)
    t_all = time.time()
    timings = {}
    logger.info(f"问题2 一键运行：run={rdir}  version={cfg['version']}  device={device}  quick={cfg.get('is_quick', False)}")
    data, normalizer = prepare_data(cfg, logger)
    cfg = resolve_strides(cfg, seq_lens_of(data["train"]))  # 时间下采样步长按数据解析后写进 config.yaml
    dump_yaml(cfg, rdir / "config.yaml")

    variants = a.variants or list(cfg["ablation"]["variants"])
    if "full" in variants:  # 主模型先训
        variants = ["full"] + [v for v in variants if v != "full"]
    if not a.skip_train:
        pb = protocol_batches(cfg, data)
        for v in variants:
            t0 = time.time()
            train_variant(cfg, v, seeds_for(cfg, v), rdir, data, normalizer, device, logger, pb)
            timings[f"train_{v}_s"] = time.time() - t0
    t0 = time.time()
    run_tasks(cfg, rdir, data, device, logger, a.tasks)
    timings["evaluate_s"] = time.time() - t0
    if cfg.get("att3"):
        from .config import resolve_path

        t0 = time.time()
        run_predict(rdir / "model.pt", resolve_path(cfg["att3"]), cfg.get("att3_split", "auto"), rdir, device,
                    int(cfg["train"].get("eval_batch_size", 256)), logger)
        timings["predict_att3_s"] = time.time() - t0
    else:
        logger.warning("未给出附件3 路径（--att3），跳过附件3 推理")
    timings["total_s"] = time.time() - t_all
    dump_json({"command": " ".join(sys.argv), "config": cfg, "variants_trained": variants if not a.skip_train else [],
               "timings": timings, "env": env_info()}, rdir / "run_info.json")
    logger.info(f"全部完成，用时 {timings['total_s'] / 60:.1f} 分钟 → {rdir}")


if __name__ == "__main__":
    main()
