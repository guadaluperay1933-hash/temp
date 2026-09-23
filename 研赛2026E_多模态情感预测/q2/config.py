"""问题2 配置：读 YAML、点路径覆盖（--set a.b=v / --quick）、解析数据路径与运行目录。

变体（消融/基线）= 主配置 + ablation.variants[name] 的点路径覆盖，保证除被消融的部分外其余超参数完全相同。
"""
from __future__ import annotations

import copy
from pathlib import Path

import yaml

from common.utils import OUTPUT_DIR, ROOT

DEFAULT_CONFIG = ROOT / "configs" / "q2.yaml"

VARIANT_CN = {
    "full": "MRG-Net（完整模型）",
    "no_aug": "w/o 缺失增广（只用完整数据训练）",
    "no_rec": "w/o 跨模态重构（去掉重构模块与 L_rec）",
    "no_kd": "w/o 一致性蒸馏（λ_kd=0）",
    "no_gate": "w/o 可靠性门控（α 均匀）",
    "late_fusion": "晚期融合基线（掩码均值池化 + MLP，无增广）",
    "naive_impute": "朴素插补基线（线性插值 + w/o 增广模型）",
    "uni_text": "单模态：文本",
    "uni_audio": "单模态：语音",
    "uni_vision": "单模态：视觉",
}


def load_config(path: str | Path | None = None) -> dict:
    path = Path(path) if path else DEFAULT_CONFIG
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def set_by_path(cfg: dict, key: str, value) -> None:
    cur = cfg
    parts = key.split(".")
    for p in parts[:-1]:
        if p not in cur or not isinstance(cur[p], dict):
            cur[p] = {}
        cur = cur[p]
    cur[parts[-1]] = value


def get_by_path(cfg: dict, key: str, default=None):
    cur = cfg
    for p in key.split("."):
        if not isinstance(cur, dict) or p not in cur:
            return default
        cur = cur[p]
    return cur


def apply_overrides(cfg: dict, overrides: dict | None) -> dict:
    out = copy.deepcopy(cfg)
    for k, v in (overrides or {}).items():
        set_by_path(out, k, copy.deepcopy(v))
    return out


def parse_set_args(items: list[str] | None) -> dict:
    """['a.b=1', 'c=[1,2]'] → {'a.b': 1, 'c': [1, 2]}（值按 YAML 解析）。"""
    out = {}
    for it in items or []:
        if "=" not in it:
            raise ValueError(f"--set 需要 key=value 形式：{it}")
        k, v = it.split("=", 1)
        out[k.strip()] = yaml.safe_load(v)
    return out


def finalize(cfg: dict, *, quick: bool = False, sets: dict | None = None, version: str | None = None,
             data: str | None = None, att3: str | None = None, run_name: str | None = None,
             epochs: int | None = None, seeds: list[int] | None = None, threads: int | None = None) -> dict:
    """把命令行参数合并进配置（优先级：--quick < 显式参数 < --set），并补全缺省路径。"""
    cfg = copy.deepcopy(cfg)
    if quick:
        cfg = apply_overrides(cfg, cfg.get("quick", {}))
        cfg["is_quick"] = True
    explicit = {}
    if version:
        explicit["version"] = version
    if data:
        explicit["data"] = data
    if att3:
        explicit["att3"] = att3
    if run_name:
        explicit["run_name"] = run_name
    if epochs is not None:
        explicit["train.epochs"] = int(epochs)
    if seeds:
        explicit["train.seeds"] = [int(s) for s in seeds]
    if threads:
        explicit["threads"] = int(threads)
    cfg = apply_overrides(cfg, explicit)
    cfg = apply_overrides(cfg, sets or {})
    if cfg["version"] not in ("aligned", "unaligned"):
        raise ValueError("version 只能是 aligned / unaligned")
    if not cfg.get("data"):
        cfg["data"] = str(Path("data") / "附件2" / f"{cfg['version']}_50.pkl")
    if not cfg.get("run_name"):
        cfg["run_name"] = f"mrgnet_{cfg['version']}"
    return cfg


def resolve_path(p: str | Path | None) -> Path | None:
    """相对路径按项目根目录解析（命令都在根目录运行，这里再兜底一次）。"""
    if p is None:
        return None
    p = Path(p)
    return p if p.is_absolute() else (ROOT / p)


def run_dir(cfg: dict, out: str | Path | None = None) -> Path:
    return resolve_path(out) if out else OUTPUT_DIR / "q2" / cfg["run_name"]


def variant_config(cfg: dict, variant: str) -> dict:
    """主配置 + 该变体的覆盖。naive_impute 不训练，复用 ablation.naive_impute_base 的模型。"""
    variants = cfg.get("ablation", {}).get("variants", {})
    if variant not in variants:
        raise KeyError(f"未知变体 {variant}，可选：{list(variants)}")
    out = apply_overrides(cfg, variants[variant])
    out["variant"] = variant
    return out


def dump_yaml(obj, path: str | Path) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(obj, f, allow_unicode=True, sort_keys=False)
