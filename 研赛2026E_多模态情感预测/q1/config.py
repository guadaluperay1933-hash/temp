"""问题1 配置：读取 configs/q1.yaml，支持命令行 `--set a.b=值` 覆盖（值按 YAML 语法解析），相对路径以项目根目录为基准。"""
from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path

import yaml

from common.utils import ROOT

DEFAULT_CONFIG = ROOT / "configs" / "q1.yaml"


def deep_update(base: dict, upd: dict) -> dict:
    for k, v in upd.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            deep_update(base[k], v)
        else:
            base[k] = v
    return base


def apply_overrides(cfg: dict, items: list[str] | None) -> dict:
    """items: ['text.model=outputs/q1_test/dummy_bert', 'vision.sample_fps=10', ...]"""
    for it in items or []:
        if "=" not in it:
            raise ValueError(f"--set 需要 key=value 形式：{it}")
        key, val = it.split("=", 1)
        node = cfg
        parts = key.strip().split(".")
        for p in parts[:-1]:
            node = node.setdefault(p, {})
        node[parts[-1]] = yaml.safe_load(val)
    return cfg


def load_config(path: str | Path | None = None, overrides: list[str] | None = None) -> dict:
    path = Path(path) if path else DEFAULT_CONFIG
    if not path.is_absolute():
        path = ROOT / path
    with open(path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    cfg = apply_overrides(copy.deepcopy(cfg), overrides)
    cfg["_config_file"] = str(path)
    return cfg


def resolve(p: str | Path | None) -> Path | None:
    """相对路径 → 以项目根目录为基准的绝对路径。"""
    if p is None:
        return None
    p = Path(p)
    return p if p.is_absolute() else ROOT / p


def model_ref(name: str | None) -> str | None:
    """模型名既可能是 HuggingFace 仓库名（原样返回），也可能是本地目录（相对项目根目录解析）。"""
    if not name:
        return name
    local = resolve(name)
    return str(local) if local.exists() else name


def config_hash(cfg: dict, keys=("text", "audio", "vision", "align", "sequence", "media")) -> str:
    """影响特征数值的配置段的哈希（断点续跑时判断缓存是否可复用）。"""
    sub = {k: cfg.get(k) for k in keys}
    return hashlib.md5(json.dumps(sub, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()[:12]
