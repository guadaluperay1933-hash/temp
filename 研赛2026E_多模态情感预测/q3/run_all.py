"""问题3 一键运行：训练 AMEE 主模型与消融/对照 → 解释统计 / 忠实性 / 错误归因 / 消融表 / 图 → 附件4 推理与解释卡。

    python -m q3.run_all --config configs/q3.yaml --data data/附件2/aligned_50.pkl \\
        --att4 data/附件4/<文件>.pkl --videos data/附件4/<视频目录>
    python -m q3.run_all --version unaligned --data data/附件2/unaligned_50.pkl --att4 ... --videos ...
    # 合成数据跑通（数值无意义）：
    python -m q3.run_all --data outputs/synthetic/附件2/aligned_50.pkl --att4 outputs/synthetic/附件4/aligned_50.pkl \\
        --videos outputs/synthetic/附件4/videos --run-name synth_aligned --quick --threads 2

等价的分步命令：q3.train（--variants ...）→ q3.evaluate（--tasks explain faith error ablation figures）→ q3.predict。
"""
from __future__ import annotations

import argparse
import sys
import time

from common.utils import dump_json, env_info, get_logger

from .config import dump_yaml, resolve_path, run_dir
from .engine import prepare_data, resolve_device, seq_lens_of, setup_threads
from .evaluate import TASKS, run_evaluate
from .model import resolve_strides
from .predict import run_predict
from .train import add_common_args, build_cfg, pooled_splits, seeds_for, train_variant


def main():
    ap = argparse.ArgumentParser(description="问题3：训练 + 解释 + 忠实性 + 消融 + 附件4 推理（一键）")
    add_common_args(ap)
    ap.add_argument("--att4", default=None, help="附件4 特征文件（缺省取配置 att4；为空则跳过推理）")
    ap.add_argument("--att4-split", default=None)
    ap.add_argument("--videos", default=None, help="附件4 原始视频根目录（<video_id>/<clip_id>.mp4）")
    ap.add_argument("--variants", nargs="+", default=None, help="只训练这些变体（缺省 = 配置里全部）")
    ap.add_argument("--seeds", nargs="+", type=int, default=None, help="主模型 full 的种子（覆盖 train.seeds）")
    ap.add_argument("--skip-train", action="store_true", help="已训练过：直接评测与推理")
    ap.add_argument("--tasks", nargs="+", default=list(TASKS), choices=list(TASKS))
    a = ap.parse_args()
    cfg = build_cfg(a, extra_seeds=a.seeds)
    if a.att4_split:
        cfg["att4_split"] = a.att4_split
    rdir = run_dir(cfg, a.out)
    rdir.mkdir(parents=True, exist_ok=True)
    logger = get_logger("q3", rdir / "logs" / "q3.log")
    setup_threads(cfg)
    device = resolve_device(cfg)
    t_all = time.time()
    timings = {}
    logger.info(f"问题3 一键运行：run={rdir}  version={cfg['version']}  device={device}  quick={cfg.get('is_quick', False)}")
    data, normalizer = prepare_data(cfg, logger)
    cfg = resolve_strides(cfg, seq_lens_of(data["train"]))
    dump_yaml(cfg, rdir / "config.yaml")

    variants = a.variants or list(cfg["ablation"]["variants"])
    if "full" in variants:
        variants = ["full"] + [v for v in variants if v != "full"]
    if not a.skip_train:
        pooled = pooled_splits(data, cfg["model"]["temporal_stride"])
        for v in variants:
            t0 = time.time()
            train_variant(cfg, v, seeds_for(cfg, v), rdir, data, normalizer, device, logger, pooled)
            timings[f"train_{v}_s"] = time.time() - t0
    t0 = time.time()
    run_evaluate(cfg, rdir, data, device, logger, a.tasks)
    timings["evaluate_s"] = time.time() - t0
    if cfg.get("att4"):
        t0 = time.time()
        run_predict(rdir / "model.pt", resolve_path(cfg["att4"]), cfg.get("att4_split", "auto"),
                    resolve_path(cfg["videos"]) if cfg.get("videos") else None, rdir, device, cfg, logger,
                    int(cfg["train"].get("eval_batch_size", 256)))
        timings["predict_att4_s"] = time.time() - t0
    else:
        logger.warning("未给出附件4 路径（--att4），跳过附件4 推理")
    timings["total_s"] = time.time() - t_all
    dump_json({"command": " ".join(sys.argv), "config": cfg, "variants_trained": variants if not a.skip_train else [],
               "timings": timings, "env": env_info()}, rdir / "run_info.json")
    logger.info(f"全部完成，用时 {timings['total_s'] / 60:.1f} 分钟 → {rdir}")


if __name__ == "__main__":
    main()
