"""附件2 / 附件3 / 附件4 特征文件（aligned_50.pkl / unaligned_50.pkl 格式）的统一读取接口。

约定（全项目通用，问题2、问题3 都只通过本模块读数据）：

* 模态顺序固定为 MODALITIES = ('text', 'audio', 'vision')。
* 一条样本第 m 个模态的第 t 个序列位置是否"可用"，由 ``avail[m][i, t]`` 给出：
  该行特征（非有限值先置 0 之后）不全为 0，且（文本）不在 BERT 注意力掩码之外。
  填充位置、赛题人为置零的缺失区间、原始数据中本就为 0 的位置（如视觉未检出人脸），
  在模型看来一律是"不可用"，统一被掩码屏蔽 —— 模型不需要区分这三者就能做预测。
* ``valid[m][i, t]`` 是"名义有效范围"：由 *_lengths 字段 / BERT 掩码 / 非零包络推出的有效长度，
  再结合自动检测到的填充方向（前填充 or 后填充）得到。
  缺失 = valid & ~avail，只在做缺失统计、构造缺失区间时使用。
* 情感类别按题面定义由连续强度 y 推出：y<0 → 0 Negative，y==0 → 1 Neutral，y>0 → 2 Positive。
  文件自带的 classification_labels / annotations 只用于一致性核对，不参与训练。
"""
from __future__ import annotations

import pickle
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

MODALITIES = ("text", "audio", "vision")
FEAT_DIMS = {"text": 768, "audio": 74, "vision": 35}
MOD_SHORT = {"text": "T", "audio": "A", "vision": "V"}
MOD_CN = {"text": "文本", "audio": "语音", "vision": "视觉"}
CLASS_NAMES = ("Negative", "Neutral", "Positive")
CLASS_CN = {"Negative": "负向", "Neutral": "中性", "Positive": "正向"}
SPLIT_KEYS = ("train", "valid", "test")
ID_SEP = "$_$"
ZERO_EPS = 1e-8          # 一行特征 max|x| <= ZERO_EPS 视为全零
NEUTRAL_EPS = 1e-6       # |y| < NEUTRAL_EPS 视为中性（标签是 1/3 的倍数，浮点误差很小）

_FEATURE_FIELDS = {"text", "audio", "vision"}


# ----------------------------------------------------------------------------- 标签

def reg_to_cls(y) -> np.ndarray:
    """连续强度 → 三分类编号（0 Negative / 1 Neutral / 2 Positive），严格按题面定义。"""
    y = np.asarray(y, dtype=np.float64).reshape(-1)
    out = np.ones(y.shape, dtype=np.int64)
    out[y <= -NEUTRAL_EPS] = 0
    out[y >= NEUTRAL_EPS] = 2
    return out


def cls_name(c) -> str:
    return CLASS_NAMES[int(c)]


def split_id(sample_id: str) -> tuple[str, str]:
    """'video_id$_$clip_id' → (video_id, clip_id)。无分隔符时 clip_id 为空串。"""
    s = str(sample_id)
    if ID_SEP in s:
        v, c = s.rsplit(ID_SEP, 1)
        return v, c
    return s, ""


# ----------------------------------------------------------------------------- 原始读取

def load_pkl(path: str | Path):
    """读取 .pkl；也支持 gzip 压缩的 .pkl.gz（问题1 自生成特征为控制附件体积可能压缩保存）。"""
    path = Path(path)
    if path.suffix == ".gz":
        import gzip

        with gzip.open(path, "rb") as f:
            return pickle.load(f)
    with open(path, "rb") as f:
        return pickle.load(f)


def _is_split_dict(d) -> bool:
    return isinstance(d, dict) and len(_FEATURE_FIELDS & set(d.keys())) >= 2


def list_splits(raw) -> list[str]:
    """返回文件里可用的划分名。

    * 附件2：{'train':{...}, 'valid':{...}, 'test':{...}}
    * 附件3/4 可能同样按划分组织，也可能直接是字段字典（记为 '__flat__'），
      也可能外面只包了一层任意名字的键 —— 三种都兼容。
    """
    if _is_split_dict(raw):
        return ["__flat__"]
    if not isinstance(raw, dict):
        raise TypeError(f"无法识别的特征文件结构：{type(raw)}")
    keys = [k for k in SPLIT_KEYS if k in raw and _is_split_dict(raw[k])]
    if keys:
        return keys
    keys = [k for k, v in raw.items() if _is_split_dict(v)]
    if keys:
        return keys
    raise ValueError(f"特征文件中找不到含 text/audio/vision 字段的划分，顶层键：{list(raw.keys())[:20]}")


def _to_str_list(x) -> list[str]:
    out = []
    for v in list(x):
        if isinstance(v, bytes):
            v = v.decode("utf-8", errors="replace")
        elif isinstance(v, np.ndarray):
            v = v.item() if v.size == 1 else " ".join(map(str, v.tolist()))
            if isinstance(v, bytes):
                v = v.decode("utf-8", errors="replace")
        out.append(str(v))
    return out


def _to_feat(x) -> np.ndarray:
    """转成 float32 (N, L, D)；非有限值（MOSEI 语音里有 -inf/NaN）置 0。"""
    arr = np.asarray(x)
    if arr.dtype == object:  # 变长列表 → 后填充到最长
        seqs = [np.asarray(s, dtype=np.float32) for s in x]
        L = max(s.shape[0] for s in seqs)
        D = seqs[0].shape[1]
        out = np.zeros((len(seqs), L, D), dtype=np.float32)
        for i, s in enumerate(seqs):
            out[i, : s.shape[0]] = s
        arr = out
    arr = arr.astype(np.float32, copy=True)
    if arr.ndim == 2:  # (N, D) → (N, 1, D)
        arr = arr[:, None, :]
    bad = ~np.isfinite(arr)
    if bad.any():
        arr[bad] = 0.0
    return arr


def _bert_mask(text_bert) -> np.ndarray | None:
    """从 text_bert (N,3,L) 中找出注意力掩码那一路（只含 0/1、且 1 最多的那一路）。"""
    if text_bert is None:
        return None
    tb = np.asarray(text_bert)
    if tb.ndim != 3:
        return None
    best, best_sum = None, -1
    for k in range(tb.shape[1]):
        row = tb[:, k, :]
        vals = np.unique(row)
        if np.all(np.isin(vals, (0, 1))):
            s = row.sum()
            if s > best_sum:
                best, best_sum = row.astype(bool), s
    return best


def _row_nonzero(x: np.ndarray) -> np.ndarray:
    return np.abs(x).max(axis=-1) > ZERO_EPS


def _hull_length(avail: np.ndarray) -> np.ndarray:
    """(N,L) 可用掩码 → 以"最后一个可用位置+1"计的长度（后填充假设下的有效长度下界）。"""
    N, L = avail.shape
    any_ = avail.any(axis=1)
    last = L - 1 - np.argmax(avail[:, ::-1], axis=1)
    return np.where(any_, last + 1, 0).astype(np.int64)


def _detect_padding_side(avail: np.ndarray, lengths: np.ndarray) -> str:
    """对比"可用位置落在前 len 个"与"落在后 len 个"的样本数，多数决定填充方向。"""
    N, L = avail.shape
    idx = np.arange(L)[None, :]
    post_region = idx < lengths[:, None]
    pre_region = idx >= (L - lengths)[:, None]
    post_hits = (avail & ~post_region).sum()
    pre_hits = (avail & ~pre_region).sum()
    return "pre" if pre_hits < post_hits else "post"


def _valid_region(lengths: np.ndarray, L: int, side: str) -> np.ndarray:
    idx = np.arange(L)[None, :]
    if side == "pre":
        return idx >= (L - lengths)[:, None]
    return idx < lengths[:, None]


# ----------------------------------------------------------------------------- 数据容器

@dataclass
class SplitData:
    name: str
    ids: list[str]
    feats: dict[str, np.ndarray]            # m -> float32 (N, L_m, D_m)，非有限值已置 0
    avail: dict[str, np.ndarray]            # m -> bool (N, L_m)   该位置可用
    valid: dict[str, np.ndarray]            # m -> bool (N, L_m)   名义有效范围
    lengths: dict[str, np.ndarray]          # m -> int  (N,)       名义有效长度
    padding_side: dict[str, str]            # m -> 'post' | 'pre'
    raw_text: list[str] | None = None
    y_reg: np.ndarray | None = None         # float32 (N,)
    y_cls: np.ndarray | None = None         # int64 (N,)  由 y_reg 按题面定义推出
    extra: dict = field(default_factory=dict)  # text_bert / annotations / classification_labels / 其它字段

    # --- 便捷属性
    def __len__(self) -> int:
        return len(self.ids)

    @property
    def has_labels(self) -> bool:
        return self.y_reg is not None

    @property
    def version(self) -> str:
        """'aligned'：三模态序列长度相同；否则 'unaligned'。"""
        Ls = {self.feats[m].shape[1] for m in MODALITIES}
        return "aligned" if len(Ls) == 1 else "unaligned"

    @property
    def missing(self) -> dict[str, np.ndarray]:
        """名义有效范围内、但特征全零的位置（附件3 的人为缺失 + 原始数据中的天然零行）。"""
        return {m: self.valid[m] & ~self.avail[m] for m in MODALITIES}

    def video_clip(self) -> list[tuple[str, str]]:
        return [split_id(s) for s in self.ids]

    def subset(self, index) -> "SplitData":
        index = np.asarray(index)
        pick = lambda a: a[index]  # noqa: E731
        return SplitData(
            name=self.name,
            ids=[self.ids[i] for i in index],
            feats={m: pick(v) for m, v in self.feats.items()},
            avail={m: pick(v) for m, v in self.avail.items()},
            valid={m: pick(v) for m, v in self.valid.items()},
            lengths={m: pick(v) for m, v in self.lengths.items()},
            padding_side=dict(self.padding_side),
            raw_text=[self.raw_text[i] for i in index] if self.raw_text is not None else None,
            y_reg=pick(self.y_reg) if self.y_reg is not None else None,
            y_cls=pick(self.y_cls) if self.y_cls is not None else None,
            extra={k: (pick(v) if isinstance(v, np.ndarray) and v.shape[:1] == (len(self.ids),) else v)
                   for k, v in self.extra.items()},
        )

    def copy(self) -> "SplitData":
        return self.subset(np.arange(len(self)))

    def summary(self) -> dict:
        out = {"name": self.name, "n": len(self), "version": self.version, "has_labels": self.has_labels}
        for m in MODALITIES:
            v = self.valid[m]
            n_valid = int(v.sum())
            out[m] = {
                "shape": list(self.feats[m].shape),
                "padding_side": self.padding_side[m],
                "mean_length": float(self.lengths[m].mean()) if len(self) else 0.0,
                "max_length": int(self.lengths[m].max()) if len(self) else 0,
                "missing_ratio_in_valid": float((v & ~self.avail[m]).sum() / max(n_valid, 1)),
                "samples_with_missing": int(((v & ~self.avail[m]).any(axis=1)).sum()),
            }
        if self.has_labels:
            out["label_mean"] = float(self.y_reg.mean())
            out["class_counts"] = {CLASS_NAMES[c]: int((self.y_cls == c).sum()) for c in range(3)}
        return out


def get_split(raw, split: str | None = None, name: str | None = None) -> SplitData:
    """从 load_pkl() 得到的原始字典中取出一个划分并标准化。"""
    splits = list_splits(raw)
    if split is None:
        if len(splits) != 1:
            raise ValueError(f"文件包含多个划分 {splits}，请指定 split")
        split = splits[0]
    d = raw if split == "__flat__" else raw[split]
    if not _is_split_dict(d):
        raise ValueError(f"划分 {split} 中没有特征字段")

    feats = {m: _to_feat(d[m]) for m in MODALITIES}
    N = feats["text"].shape[0]
    for m in MODALITIES:
        if feats[m].shape[0] != N:
            raise ValueError(f"{m} 的样本数 {feats[m].shape[0]} 与 text 的 {N} 不一致")

    ids = _to_str_list(d["id"]) if "id" in d else [f"sample{i:05d}{ID_SEP}0" for i in range(N)]
    raw_text = _to_str_list(d["raw_text"]) if "raw_text" in d else None

    bert_mask = _bert_mask(d.get("text_bert"))
    avail = {m: _row_nonzero(feats[m]) for m in MODALITIES}
    if bert_mask is not None and bert_mask.shape == avail["text"].shape:
        avail["text"] = avail["text"] & bert_mask

    # ---- 名义有效长度
    lengths: dict[str, np.ndarray] = {}
    aligned = len({feats[m].shape[1] for m in MODALITIES}) == 1
    if bert_mask is not None and bert_mask.shape == avail["text"].shape:
        lengths["text"] = bert_mask.sum(axis=1).astype(np.int64)
    else:
        lengths["text"] = _hull_length(avail["text"])
    for m in ("audio", "vision"):
        key = f"{m}_lengths"
        if key in d and d[key] is not None:
            lengths[m] = np.asarray(d[key]).reshape(-1).astype(np.int64)
        elif aligned:
            union = avail["text"] | avail["audio"] | avail["vision"]
            lengths[m] = np.maximum(lengths["text"], _hull_length(union))
        else:
            lengths[m] = _hull_length(avail[m])
    if aligned and bert_mask is None:
        # 对齐版本三模态共享位置：用三模态并集的包络作为共同长度（缺失区间恰好落在末尾时也不至于低估太多）
        union_len = _hull_length(avail["text"] | avail["audio"] | avail["vision"])
        for m in MODALITIES:
            lengths[m] = np.maximum(lengths[m], union_len)
    for m in MODALITIES:
        lengths[m] = np.clip(lengths[m], 0, feats[m].shape[1])

    padding_side = {m: _detect_padding_side(avail[m], lengths[m]) for m in MODALITIES}
    valid = {m: _valid_region(lengths[m], feats[m].shape[1], padding_side[m]) for m in MODALITIES}
    # 可用位置一定在有效范围内（极少数异常行：以数据为准扩大有效范围）
    for m in MODALITIES:
        valid[m] = valid[m] | avail[m]

    y_reg = y_cls = None
    if "regression_labels" in d and d["regression_labels"] is not None:
        y_reg = np.asarray(d["regression_labels"], dtype=np.float32).reshape(-1)
        if y_reg.shape[0] != N:
            raise ValueError("regression_labels 长度与样本数不一致")
        y_cls = reg_to_cls(y_reg)

    extra = {}
    for k, v in d.items():
        if k in _FEATURE_FIELDS or k in ("id", "raw_text", "regression_labels", "audio_lengths", "vision_lengths"):
            continue
        extra[k] = np.asarray(v) if not isinstance(v, (str, bytes)) else v

    return SplitData(name=name or split, ids=ids, feats=feats, avail=avail, valid=valid, lengths=lengths,
                     padding_side=padding_side, raw_text=raw_text, y_reg=y_reg, y_cls=y_cls, extra=extra)


def load_split(path: str | Path, split: str | None = None) -> SplitData:
    raw = load_pkl(path)
    return get_split(raw, split, name=f"{Path(path).stem}:{split or 'auto'}")


def load_all_splits(path: str | Path) -> dict[str, SplitData]:
    raw = load_pkl(path)
    return {s: get_split(raw, s, name=s) for s in list_splits(raw)}


def label_consistency_report(sd: SplitData) -> dict:
    """核对文件自带 classification_labels / annotations 与"由 y 按题面定义推出的类别"是否一致。"""
    rep: dict = {"n": len(sd)}
    if sd.y_cls is None:
        return rep
    if "annotations" in sd.extra:
        ann = _to_str_list(sd.extra["annotations"])
        name2c = {n.lower(): i for i, n in enumerate(CLASS_NAMES)}
        mapped = np.array([name2c.get(a.strip().lower(), -1) for a in ann])
        rep["annotations_agree"] = float((mapped == sd.y_cls).mean())
    if "classification_labels" in sd.extra:
        cl = np.asarray(sd.extra["classification_labels"]).reshape(-1)
        if cl.shape[0] == len(sd):
            # 可能的编码未知：枚举 0/1/2 的全部排列，报告最佳一致率及对应映射
            import itertools

            best = (-1.0, None)
            vals = np.unique(cl)
            if len(vals) <= 3:
                for perm in itertools.permutations(range(3), len(vals)):
                    mp = {v: p for v, p in zip(vals, perm)}
                    agree = float((np.vectorize(mp.get)(cl) == sd.y_cls).mean())
                    if agree > best[0]:
                        best = (agree, {str(k): CLASS_NAMES[v] for k, v in mp.items()})
            rep["classification_labels_best_agree"] = best[0]
            rep["classification_labels_mapping"] = best[1]
    return rep


# ----------------------------------------------------------------------------- 标准化

class Normalizer:
    """按维度 z-score 标准化：统计量只用训练集"可用"位置计算；变换后不可用位置仍保持为 0。"""

    def __init__(self, modalities=MODALITIES, clip: float = 10.0):
        self.modalities = tuple(modalities)
        self.clip = clip
        self.mean: dict[str, np.ndarray] = {}
        self.std: dict[str, np.ndarray] = {}

    def fit(self, sd: SplitData) -> "Normalizer":
        for m in self.modalities:
            rows = sd.feats[m][sd.avail[m]]
            if rows.shape[0] == 0:
                mu = np.zeros(sd.feats[m].shape[-1], np.float32)
                sd_ = np.ones_like(mu)
            else:
                mu = rows.mean(axis=0).astype(np.float32)
                sd_ = rows.std(axis=0).astype(np.float32)
                sd_[sd_ < 1e-6] = 1.0
            self.mean[m], self.std[m] = mu, sd_
        return self

    def transform(self, sd: SplitData) -> SplitData:
        out = sd.copy()
        for m in self.modalities:
            x = (out.feats[m] - self.mean[m]) / self.std[m]
            if self.clip:
                np.clip(x, -self.clip, self.clip, out=x)
            x[~out.avail[m]] = 0.0
            out.feats[m] = x.astype(np.float32)
        return out

    def state_dict(self) -> dict:
        return {"modalities": list(self.modalities), "clip": self.clip,
                "mean": {m: v.tolist() for m, v in self.mean.items()},
                "std": {m: v.tolist() for m, v in self.std.items()}}

    @classmethod
    def from_state_dict(cls, st: dict) -> "Normalizer":
        n = cls(st["modalities"], st["clip"])
        n.mean = {m: np.asarray(v, np.float32) for m, v in st["mean"].items()}
        n.std = {m: np.asarray(v, np.float32) for m, v in st["std"].items()}
        return n
