"""合成数据生成器 —— 只用于在拿不到赛题附件时把全部代码跑通、做回归测试。

生成的文件与附件2/3/4 的结构完全一致（字段名、张量形状、按字段集中存储、id 形如 video_id$_$clip_id），
并且内置"可核验的真值"：每条样本的情感强度只由各模态少数几个"证据位置"决定，
各模态信息量由随机权重控制（真值主导模态 = 权重最大者）。
这样可以检验：模型能不能学到、缺失时能不能稳住、问题3 的解释能不能找回真正的证据位置。

    python -m common.synth --out outputs/synthetic --version aligned
    python -m common.synth --out outputs/synthetic --version unaligned

注意：合成数据上的任何数值都不能写进论文，论文结果必须来自真实附件。
"""
from __future__ import annotations

import argparse
import pickle
import shutil
import subprocess
from pathlib import Path

import numpy as np

from .data import CLASS_NAMES, ID_SEP, MODALITIES, reg_to_cls
from .utils import dump_json

L_TEXT = 50
L_UNALIGNED = 500
DIMS = {"text": 768, "audio": 74, "vision": 35}
LATENT = 32
AUDIO_RATE = 20.0    # 合成"非对齐"语音特征帧率（Hz），用于检验位置→时间映射的帧率估计
VISION_RATE = 15.0   # 合成"非对齐"视觉特征帧率（Hz）
SEC_PER_TOKEN = 0.35

NEUTRAL_WORDS = ["i", "think", "the", "movie", "was", "really", "it", "and", "so", "but", "just", "like",
                 "you", "know", "that", "this", "story", "actor", "plot", "um", "well", "kind", "of"]
POS_WORDS = ["great", "awesome", "love", "amazing", "good", "fantastic", "enjoyed"]
NEG_WORDS = ["terrible", "hate", "awful", "bad", "boring", "worst", "disappointing"]


class _World:
    """固定的随机投影与"情感方向"，保证 train/valid/test/附件3/附件4 同分布。"""

    def __init__(self, seed: int):
        rng = np.random.default_rng(seed)
        self.proj = {m: rng.normal(0, 1 / np.sqrt(LATENT), size=(LATENT, DIMS[m])).astype(np.float32)
                     for m in MODALITIES}
        self.direction = {}
        for m in MODALITIES:
            d = rng.normal(size=DIMS[m]).astype(np.float32)
            self.direction[m] = d / np.linalg.norm(d) * np.sqrt(DIMS[m]) * 0.5
        self.offset = {m: rng.normal(0, 0.5, size=DIMS[m]).astype(np.float32) for m in MODALITIES}


def _sample_label(rng) -> float:
    if rng.random() < 0.2:
        return 0.0
    sign = 1.0 if rng.random() < 0.65 else -1.0
    steps = np.arange(1, 10)  # 1/3 ... 3
    p = np.exp(-0.25 * steps)
    p /= p.sum()
    return float(sign * rng.choice(steps, p=p) / 3.0)


def _make_video_id(rng) -> str:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-"
    return "".join(rng.choice(list(alphabet), size=11))


def _gen_sample(world: _World, rng, version: str):
    y = _sample_label(rng)
    w = rng.dirichlet([0.8, 0.8, 0.8]).astype(np.float32)
    main = MODALITIES[int(np.argmax(w))]
    Lt = int(np.clip(rng.normal(22, 9), 6, L_TEXT))
    n_words = Lt - 2
    if version == "aligned":
        lens = {"text": Lt, "audio": Lt, "vision": Lt}
        duration = Lt * SEC_PER_TOKEN
        Lmax = {"text": L_TEXT, "audio": L_TEXT, "vision": L_TEXT}
    else:
        duration = Lt * SEC_PER_TOKEN * float(rng.uniform(0.9, 1.1))
        lens = {"text": Lt,
                "audio": int(min(L_UNALIGNED, max(4, round(duration * AUDIO_RATE)))),
                "vision": int(min(L_UNALIGNED, max(4, round(duration * VISION_RATE))))}
        Lmax = {"text": L_TEXT, "audio": L_UNALIGNED, "vision": L_UNALIGNED}

    feats, evidence = {}, {}
    for mi, m in enumerate(MODALITIES):
        L = lens[m]
        z = rng.normal(size=(L, LATENT)).astype(np.float32)
        x = z @ world.proj[m] + world.offset[m] + rng.normal(0, 0.3, size=(L, DIMS[m])).astype(np.float32)
        strength = y * w[mi] * 3.0
        if m == "text":
            k = int(rng.integers(1, 3))
            pos = np.sort(rng.choice(np.arange(1, Lt - 1), size=min(k, n_words), replace=False)) if n_words > 0 else np.array([], int)
            ev = pos.tolist()
        else:
            span = max(1, int(round(0.15 * L)))
            s = int(rng.integers(0, max(1, L - span + 1)))
            ev = list(range(s, s + span))
        for p in ev:
            x[p] += strength * world.direction[m] / np.sqrt(len(ev)) * 1.5
        # 全段弱信号，让"平均池化"类基线也能学到一点
        x += 0.1 * strength * world.direction[m] / np.sqrt(L)
        full = np.zeros((Lmax[m], DIMS[m]), np.float32)
        full[:L] = x
        if m == "vision" and rng.random() < 0.1:  # 天然零行：部分帧未检出人脸
            nf = rng.random(L) < 0.05
            nf[ev] = False
            full[:L][nf] = 0.0
        feats[m] = full
        evidence[m] = ev
    if rng.random() < 0.02:  # MOSEI 语音里偶有 -inf / NaN
        feats["audio"][0, int(rng.integers(DIMS["audio"]))] = -np.inf if rng.random() < 0.5 else np.nan

    words = []
    ev_t = set(evidence["text"])
    for p in range(1, Lt - 1):
        if p in ev_t and y > 0:
            words.append(str(rng.choice(POS_WORDS)))
        elif p in ev_t and y < 0:
            words.append(str(rng.choice(NEG_WORDS)))
        else:
            words.append(str(rng.choice(NEUTRAL_WORDS)))
    raw_text = " ".join(words).upper()
    ids = np.zeros(L_TEXT, np.int64)
    ids[0], ids[Lt - 1] = 101, 102
    ids[1:Lt - 1] = rng.integers(1000, 30000, size=n_words)
    mask = np.zeros(L_TEXT, np.int64)
    mask[:Lt] = 1
    text_bert = np.stack([ids, mask, np.zeros(L_TEXT, np.int64)])
    truth = {"y": y, "weights": {m: float(w[i]) for i, m in enumerate(MODALITIES)}, "main_modality": main,
             "evidence": evidence, "lengths": lens, "duration": duration, "words": words}
    return feats, lens, raw_text, text_bert, y, truth


def _gen_split(world, rng, n, version, with_labels=True):
    cols = {m: [] for m in MODALITIES}
    out = {"id": [], "raw_text": [], "text_bert": [], "audio_lengths": [], "vision_lengths": [],
           "regression_labels": [], "annotations": [], "classification_labels": []}
    truths = {}
    for _ in range(n):
        vid = _make_video_id(rng)
        cid = str(int(rng.integers(0, 40)))
        sid = f"{vid}{ID_SEP}{cid}"
        feats, lens, raw_text, tb, y, truth = _gen_sample(world, rng, version)
        for m in MODALITIES:
            cols[m].append(feats[m])
        out["id"].append(sid)
        out["raw_text"].append(raw_text)
        out["text_bert"].append(tb)
        out["audio_lengths"].append(lens["audio"])
        out["vision_lengths"].append(lens["vision"])
        c = int(reg_to_cls([y])[0])
        out["regression_labels"].append(y)
        out["annotations"].append(CLASS_NAMES[c])
        out["classification_labels"].append(c)
        truths[sid] = truth
    d = {m: np.stack(cols[m]).astype(np.float32) for m in MODALITIES}
    d["id"] = np.array(out["id"], dtype=object)
    d["raw_text"] = np.array(out["raw_text"], dtype=object)
    d["text_bert"] = np.stack(out["text_bert"]).astype(np.int64)
    if version == "unaligned":
        d["audio_lengths"] = np.array(out["audio_lengths"], np.int64)
        d["vision_lengths"] = np.array(out["vision_lengths"], np.int64)
    if with_labels:
        d["regression_labels"] = np.array(out["regression_labels"], np.float32)
        d["annotations"] = np.array(out["annotations"], dtype=object)
        d["classification_labels"] = np.array(out["classification_labels"], np.float32)
    return d, truths


def _apply_random_missing(d: dict, rng, version: str) -> dict:
    """附件3 风格：随机 1~3 个模态，各放 1~2 段整行置零的连续区间。返回每条样本的真值区间。"""
    from .missing import sample_spans

    N = d["text"].shape[0]
    record = []
    for i in range(N):
        mods = [m for m in MODALITIES if rng.random() < 0.5] or [MODALITIES[int(rng.integers(3))]]
        rec = {}
        for m in mods:
            if m == "text":
                L = int(d["text_bert"][i, 1].sum())
            elif version == "unaligned":
                L = int(d[f"{m}_lengths"][i])
            else:
                L = int(d["text_bert"][i, 1].sum())
            r = float(rng.uniform(0.1, 0.7))
            spans = sample_spans(np.arange(L), r, int(rng.integers(1, 3)), None, rng)
            for s, e in spans:
                d[m][i, s:e] = 0.0
            rec[m] = spans
        record.append(rec)
    return record


def _make_video(path: Path, duration: float, ffmpeg: str, size="160x120", fps=25):
    path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [ffmpeg, "-y", "-loglevel", "error",
           "-f", "lavfi", "-i", f"testsrc=size={size}:rate={fps}:duration={duration:.3f}",
           "-f", "lavfi", "-i", f"sine=frequency=220:sample_rate=16000:duration={duration:.3f}",
           "-shortest", "-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "aac", str(path)]
    subprocess.run(cmd, check=True)


def make_synthetic(out_dir: str | Path, version: str = "aligned", n_train=600, n_valid=150, n_test=150,
                   n_att3=120, n_att4=10, seed=0, make_videos=True, att_layout: str = "split") -> dict:
    """生成合成版 附件2/3/4。att_layout='split' → 附件3/4 顶层为 {'test': {...}}；'flat' → 直接字段字典。"""
    assert version in ("aligned", "unaligned")
    out = Path(out_dir)
    world = _World(seed)
    rng = np.random.default_rng(seed + 1)
    fname = f"{version}_50.pkl"

    a2, t2 = {}, {}
    for split, n in (("train", n_train), ("valid", n_valid), ("test", n_test)):
        a2[split], t = _gen_split(world, rng, n, version)
        t2.update(t)
    (out / "附件2").mkdir(parents=True, exist_ok=True)
    with open(out / "附件2" / fname, "wb") as f:
        pickle.dump(a2, f, protocol=4)
    dump_json(t2, out / "附件2" / f"_truth_{version}.json")

    d3, t3 = _gen_split(world, rng, n_att3, version, with_labels=False)
    spans3 = _apply_random_missing(d3, rng, version)
    for sid, sp in zip(d3["id"], spans3):
        t3[sid]["missing_spans"] = sp
    (out / "附件3").mkdir(parents=True, exist_ok=True)
    with open(out / "附件3" / fname, "wb") as f:
        pickle.dump({"test": d3} if att_layout == "split" else d3, f, protocol=4)
    dump_json(t3, out / "附件3" / f"_truth_{version}.json")

    d4, t4 = _gen_split(world, rng, n_att4, version, with_labels=False)
    (out / "附件4").mkdir(parents=True, exist_ok=True)
    with open(out / "附件4" / fname, "wb") as f:
        pickle.dump({"test": d4} if att_layout == "split" else d4, f, protocol=4)
    for sid in t4:
        t4[sid]["audio_rate"] = AUDIO_RATE if version == "unaligned" else None
        t4[sid]["vision_rate"] = VISION_RATE if version == "unaligned" else None
    dump_json(t4, out / "附件4" / f"_truth_{version}.json")

    if make_videos:
        try:
            import imageio_ffmpeg

            ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
        except ImportError:
            ffmpeg = shutil.which("ffmpeg")
        if ffmpeg:
            for sid, tr in t4.items():
                vid, cid = sid.split(ID_SEP)
                _make_video(out / "附件4" / "videos" / vid / f"{cid}.mp4", max(tr["duration"], 1.0), ffmpeg)
    return {"附件2": str(out / "附件2" / fname), "附件3": str(out / "附件3" / fname),
            "附件4": str(out / "附件4" / fname), "附件4_videos": str(out / "附件4" / "videos")}


def main():
    ap = argparse.ArgumentParser(description="生成合成版附件2/3/4（仅用于跑通代码）")
    ap.add_argument("--out", default="outputs/synthetic")
    ap.add_argument("--version", default="aligned", choices=["aligned", "unaligned"])
    ap.add_argument("--n-train", type=int, default=600)
    ap.add_argument("--n-valid", type=int, default=150)
    ap.add_argument("--n-test", type=int, default=150)
    ap.add_argument("--n-att3", type=int, default=120)
    ap.add_argument("--n-att4", type=int, default=10)
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--no-videos", action="store_true")
    ap.add_argument("--layout", default="split", choices=["split", "flat"])
    a = ap.parse_args()
    paths = make_synthetic(a.out, a.version, a.n_train, a.n_valid, a.n_test, a.n_att3, a.n_att4, a.seed,
                           not a.no_videos, a.layout)
    for k, v in paths.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    main()
