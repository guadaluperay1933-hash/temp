"""模态局部缺失：区间构造（训练增广 / 评测协议）、施加、以及对附件3 的缺失检测统计。

题面对缺失的定义：一个或多个模态中存在"特征值全部为零的随机连续序列区间"。
因此这里所有缺失都以"把名义有效范围内一段连续位置整行置零"的方式模拟，与附件3 完全同构。

缺失率 r 的定义：该模态被置零的位置数 / 该模态名义有效长度。
缺失位置 position：
    'head'   区间从有效范围起点开始
    'middle' 区间居中
    'tail'   区间在有效范围末尾结束
    None     随机放置（多段时各段互不重叠）
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from .data import MODALITIES, SplitData

POSITIONS = ("head", "middle", "tail")
MODALITY_COMBOS = {  # 缺失模态类型（评测协议用）
    "T": ("text",), "A": ("audio",), "V": ("vision",),
    "TA": ("text", "audio"), "TV": ("text", "vision"), "AV": ("audio", "vision"),
    "TAV": ("text", "audio", "vision"),
}


def spans_from_mask(mask_1d) -> list[tuple[int, int]]:
    """布尔序列 → 连续 True 段列表 [(start, end_exclusive), ...]。"""
    m = np.asarray(mask_1d, dtype=bool).astype(np.int8)
    if m.size == 0:
        return []
    d = np.diff(np.concatenate([[0], m, [0]]))
    starts = np.where(d == 1)[0]
    ends = np.where(d == -1)[0]
    return list(zip(starts.tolist(), ends.tolist()))


def _composition(total: int, parts: int, rng: np.random.Generator, min_part: int = 0) -> np.ndarray:
    """把非负整数 total 随机拆成 parts 份（每份 >= min_part），均匀地在所有拆法中取一个。"""
    total_free = total - parts * min_part
    if total_free < 0:
        raise ValueError("total 太小")
    if parts == 1:
        return np.array([total])
    cuts = np.sort(rng.choice(total_free + parts - 1, size=parts - 1, replace=False))
    bounds = np.concatenate([[-1], cuts, [total_free + parts - 1]])
    return np.diff(bounds) - 1 + min_part


def sample_spans(valid_positions: np.ndarray, ratio: float, n_spans: int = 1, position: str | None = None,
                 rng: np.random.Generator | None = None) -> list[tuple[int, int]]:
    """在有效位置索引序列 valid_positions（升序，通常连续）上放置总长约 ratio*n 的缺失区间。

    返回的是原序列坐标下的 [(start, end_exclusive), ...]。
    """
    rng = rng or np.random.default_rng()
    vp = np.asarray(valid_positions)
    n = vp.size
    if n == 0 or ratio <= 0:
        return []
    total = int(round(ratio * n))
    total = min(max(total, 1), n)
    if total >= n:
        return [(int(vp[0]), int(vp[-1]) + 1)]
    n_spans = int(max(1, min(n_spans, total, n - total + 1)))
    if position in POSITIONS:
        n_spans = 1
        if position == "head":
            s = 0
        elif position == "tail":
            s = n - total
        else:
            s = (n - total) // 2
        local = [(s, s + total)]
    else:
        lens = _composition(total, n_spans, rng, min_part=1)
        # 剩余 n-total 个位置作为 n_spans+1 个间隙；内部间隙至少 1（保证各段不相邻、不重叠）
        free = n - total
        if n_spans == 1:
            gaps = _composition(free, 2, rng, min_part=0)
        else:
            gaps = _composition(free - (n_spans - 1), n_spans + 1, rng, min_part=0)
            gaps[1:-1] += 1
        local, cur = [], 0
        for k in range(n_spans):
            cur += int(gaps[k])
            local.append((cur, cur + int(lens[k])))
            cur += int(lens[k])
    # 映射回原坐标（valid_positions 若不连续，按局部索引取值）
    out = []
    for s, e in local:
        out.append((int(vp[s]), int(vp[e - 1]) + 1))
    return out


@dataclass
class MissingSpec:
    """对一个模态的缺失设定。ratio / n_spans 可以是定值，也可以是 (lo, hi) 区间（每条样本随机抽）。"""
    ratio: float | tuple[float, float] = 0.0
    n_spans: int | tuple[int, int] = 1
    position: str | None = None

    def draw(self, rng: np.random.Generator) -> tuple[float, int, str | None]:
        r = self.ratio if np.isscalar(self.ratio) else float(rng.uniform(*self.ratio))
        k = self.n_spans if np.isscalar(self.n_spans) else int(rng.integers(self.n_spans[0], self.n_spans[1] + 1))
        return float(r), int(k), self.position


def make_drop_masks(sd: SplitData, spec: dict[str, MissingSpec], seed: int = 0,
                    sample_prob: float = 1.0) -> dict[str, np.ndarray]:
    """按 spec 为每条样本生成"要置零的位置"掩码 drop[m] (N, L_m)。只在名义有效范围内放置区间。

    sample_prob：每条样本以该概率施加缺失（评测协议一般取 1.0）。
    固定 seed → 完全可复现（评测协议对所有模型使用相同的缺失掩码）。
    """
    rng = np.random.default_rng(seed)
    N = len(sd)
    drop = {m: np.zeros_like(sd.avail[m]) for m in MODALITIES}
    for i in range(N):
        if sample_prob < 1.0 and rng.random() >= sample_prob:
            continue
        for m, sp in spec.items():
            r, k, pos = sp.draw(rng)
            if r <= 0:
                continue
            vp = np.where(sd.valid[m][i])[0]
            for s, e in sample_spans(vp, r, k, pos, rng):
                drop[m][i, s:e] = True
    return drop


def apply_drop(sd: SplitData, drop: dict[str, np.ndarray]) -> SplitData:
    """返回一个新 SplitData：drop 位置整行置零、avail 同步更新（valid 不变 → 这些位置计为缺失）。"""
    out = sd.copy()
    for m in MODALITIES:
        if m in drop and drop[m].any():
            out.feats[m] = out.feats[m].copy()
            out.feats[m][drop[m]] = 0.0
            out.avail[m] = out.avail[m] & ~drop[m]
    return out


def simulate_missing(sd: SplitData, spec: dict[str, MissingSpec], seed: int = 0, sample_prob: float = 1.0) -> SplitData:
    return apply_drop(sd, make_drop_masks(sd, spec, seed=seed, sample_prob=sample_prob))


def protocol_spec(combo: str, ratio: float, position: str | None = None, n_spans: int = 1) -> dict[str, MissingSpec]:
    """评测协议：缺失模态类型 combo ∈ MODALITY_COMBOS，缺失率 ratio，缺失位置 position。"""
    return {m: MissingSpec(ratio=ratio, n_spans=n_spans, position=position) for m in MODALITY_COMBOS[combo]}


# ----------------------------------------------------------------------------- 训练期随机增广（numpy，逐样本）

@dataclass
class AugmentConfig:
    p_sample: float = 0.8            # 一条样本施加缺失的概率（其余保持完整，维持完整输入下的性能）
    p_modality: float = 0.5          # 施加时，每个模态被选中的概率（至少选一个）
    ratio_range: tuple[float, float] = (0.05, 0.9)
    max_spans: int = 3
    p_full_drop: float = 0.1         # 被选中的模态整段缺失（r=1）的概率
    keep_one_intact: bool = False    # True 时保证至少一个模态完整


def random_drop_for_sample(valid_rows: dict[str, np.ndarray], cfg: AugmentConfig,
                           rng: np.random.Generator) -> dict[str, np.ndarray]:
    """给一条样本随机生成缺失掩码。valid_rows[m] 为 (L_m,) 布尔名义有效范围。"""
    drop = {m: np.zeros_like(valid_rows[m]) for m in MODALITIES}
    if rng.random() >= cfg.p_sample:
        return drop
    chosen = [m for m in MODALITIES if rng.random() < cfg.p_modality]
    if not chosen:
        chosen = [MODALITIES[rng.integers(3)]]
    if cfg.keep_one_intact and len(chosen) == 3:
        chosen.pop(int(rng.integers(3)))
    for m in chosen:
        vp = np.where(valid_rows[m])[0]
        if vp.size == 0:
            continue
        if rng.random() < cfg.p_full_drop:
            r, k = 1.0, 1
        else:
            r = float(rng.uniform(*cfg.ratio_range))
            k = int(rng.integers(1, cfg.max_spans + 1))
        for s, e in sample_spans(vp, r, k, None, rng):
            drop[m][s:e] = True
    return drop


# ----------------------------------------------------------------------------- 附件3 缺失检测统计

def missing_statistics(sd: SplitData) -> list[dict]:
    """逐样本、逐模态统计名义有效范围内的全零区间（附件3 的缺失即此）。

    返回行字典列表，字段：id, modality, valid_len, n_missing, ratio, n_spans, spans(str),
    first_rel_start, center_rel（缺失中心在有效范围中的相对位置，0=开头 1=结尾）。
    """
    rows = []
    for i, sid in enumerate(sd.ids):
        for m in MODALITIES:
            v = sd.valid[m][i]
            miss = v & ~sd.avail[m][i]
            vp = np.where(v)[0]
            n_valid = int(vp.size)
            spans = spans_from_mask(miss)
            n_miss = int(miss.sum())
            if n_valid and n_miss:
                lo = vp[0]
                centers = np.where(miss)[0]
                center_rel = float((centers.mean() - lo) / max(n_valid - 1, 1))
                first_rel = float((spans[0][0] - lo) / max(n_valid - 1, 1))
            else:
                center_rel = first_rel = float("nan")
            rows.append({
                "id": sid, "modality": m, "valid_len": n_valid, "n_missing": n_miss,
                "ratio": n_miss / n_valid if n_valid else 0.0, "n_spans": len(spans),
                "spans": ";".join(f"{s}-{e}" for s, e in spans),
                "first_rel_start": first_rel, "center_rel": center_rel,
            })
    return rows
