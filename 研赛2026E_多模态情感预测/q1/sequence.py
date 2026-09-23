"""跨模态时序对齐：词级时间戳 + 帧级特征 → 附件2 同构的两种序列（对齐版 50 位置 / 非对齐版 80 ms 栅格）。

记号：D_a = 音频时长（样点数/sr），D_v = 视频时长（解码帧数/fps）；
      音频帧 k 的中心 τ_k = 0.01k（100 Hz），特征 a_k ∈ R^74；
      视频帧 n 的时间戳 t_n = n/fps，特征 v_n ∈ R^35，人脸标志 f_n ∈ {0,1}（未检出人脸 v_n = 0）。

一、对齐版本（aligned，三模态共用 50 个位置，位置 i ↔ 文本第 i 个词元）
  位置区间 I_i = [s_i, e_i)：
    * i = 0 的 [CLS] 与 i = L−1 的 [SEP]：整段 [0, D]（句级词元，取全片段均值；两者特征相同）；
    * 词元 i 属于第 j 个词（时间 [S_j, E_j)），是该词 P_j 个词元中的第 r 个：
          s_i = S_j + (E_j − S_j)·r/P_j，  e_i = S_j + (E_j − S_j)·(r+1)/P_j   （词区间在其词元间均分）；
    * i ≥ L（填充）：三模态整行为 0。
  语音：F_i = {k : τ_k ∈ [s_i, e_i)}，A_i = mean_{k∈F_i} a_k；F_i 为空（区间短于帧移）时取离区间中点最近的一帧。
  视觉：G_i = {n : t_n ∈ [s_i, e_i)}（为空时同样取最近一帧），V_i = mean_{n∈G_i, f_n=1} v_n；G_i 内没有人脸帧 → V_i = 0。
  文本长度 L = 词元数（含 [CLS]/[SEP]），audio_lengths = vision_lengths = L（位置共享）。
  说明：若区间终点恰为片段终点 D，则闭到 D（使最后一帧不被漏掉）。

二、非对齐版本（unaligned，语音/视觉各自保留时间轴，统一重采样到 Δ = 80 ms 栅格）
  位置 k ↔ [kΔ, (k+1)Δ)；K_a = ⌈D_a/Δ⌉，K_v = ⌈D_v/Δ⌉，上限 500（超出截断并记录）。
  帧 → 位置：p(τ) = min(⌊τ/Δ⌋, K−1)。语音位置 k 取落入该位置的 8 个 100 Hz 帧均值；
  视觉位置 k 取落入该位置的人脸帧均值（无帧 → 最近帧；有帧但都无人脸 → 0 行）。
  audio_lengths = K_a，vision_lengths = K_v；k ≥ K 为填充（0 行）。文本与对齐版本相同。
"""
from __future__ import annotations

import numpy as np

EPS_T = 1e-9


def token_intervals(tok: dict, words_t: list[dict], dur: float) -> list[tuple[float | None, float | None, str]]:
    """每个位置的时间区间与规则：(start, end, rule)。填充位置返回 (None, None, 'pad')。"""
    L = int(tok["text_len"])
    out = []
    for i in range(len(tok["tokens"])):
        if i >= L:
            out.append((None, None, "pad"))
            continue
        j = int(tok["word_index"][i])
        if j < 0:  # [CLS] / [SEP]
            out.append((0.0, float(dur), "special_whole_clip"))
            continue
        w = words_t[j]
        S, E = float(w["start"]), float(w["end"])
        P = max(int(tok["piece_count"][i]), 1)
        r = int(tok["piece_rank"][i])
        out.append((S + (E - S) * r / P, S + (E - S) * (r + 1) / P, "word_piece"))
    return out


def _frames_in(times: np.ndarray, a: float, b: float, dur: float) -> tuple[int, int]:
    """已排序时刻数组中落在 [a, b) 的下标范围 [lo, hi)；b ≥ dur 时闭到右端。"""
    lo = int(np.searchsorted(times, a - EPS_T, side="left"))
    if b >= dur - EPS_T:
        hi = int(np.searchsorted(times, max(b, dur) + EPS_T, side="right"))
    else:
        hi = int(np.searchsorted(times, b - EPS_T, side="left"))
    return lo, max(hi, lo)


def _nearest(times: np.ndarray, t: float) -> int:
    k = int(np.searchsorted(times, t))
    if k <= 0:
        return 0
    if k >= len(times):
        return len(times) - 1
    return k if abs(times[k] - t) < abs(times[k - 1] - t) else k - 1


def pool_interval(feats: np.ndarray, times: np.ndarray, a: float, b: float, dur: float,
                  usable: np.ndarray | None = None) -> tuple[np.ndarray, dict]:
    """区间均值池化。usable（如人脸标志）为 None 表示所有帧都可用。返回 (行, 记录)。"""
    D = feats.shape[1]
    rec = {"frames": None, "n_frames": 0, "n_used": 0, "rule": "no_frames"}
    if len(times) == 0:
        return np.zeros(D, np.float32), rec
    lo, hi = _frames_in(times, a, b, dur)
    rule = "mean"
    if hi <= lo:
        lo = _nearest(times, 0.5 * (a + b))
        hi = lo + 1
        rule = "nearest"
    idx = np.arange(lo, hi)
    if usable is not None:
        idx = idx[usable[lo:hi]]
    rec.update({"frames": [lo, hi], "n_frames": hi - lo, "n_used": int(len(idx))})
    if len(idx) == 0:
        rec["rule"] = rule + "_no_face"
        return np.zeros(D, np.float32), rec
    rec["rule"] = rule
    return feats[idx].mean(axis=0).astype(np.float32), rec


def build_aligned(tok: dict, words_t: list[dict], audio: dict | None, vision: dict | None, max_len: int = 50) -> dict:
    """audio: dict(feats (T,74), times (T,), dur)；vision: dict(feats (F,35), times (F,), frame_idx (F,), face (F,), dur)。

    返回 dict(audio (max_len,74), vision (max_len,35), positions [...], length L)。
    某模态为 None（提取失败/无音轨）→ 该模态整段为 0 行。
    """
    L = int(tok["text_len"])
    A = np.zeros((max_len, 74), np.float32)
    V = np.zeros((max_len, 35), np.float32)
    d_ref = audio["dur"] if audio is not None else (vision["dur"] if vision is not None else 0.0)
    ivs = token_intervals(tok, words_t, d_ref)
    positions = []
    for i, (s, e, rule) in enumerate(ivs):
        j = int(tok["word_index"][i])
        rec = {"pos": i, "token": tok["tokens"][i], "token_id": int(tok["input_ids"][i]), "word_index": j,
               "word": words_t[j]["word"] if j >= 0 else None, "rule": rule,
               "start_s": None if s is None else round(s, 4), "end_s": None if e is None else round(e, 4),
               "audio_frames": None, "audio_rule": None, "video_frames": None, "n_video_frames": 0,
               "n_face_frames": 0, "vision_rule": None}
        if rule != "pad":
            if audio is not None:
                # 特殊词元区间取该模态自己的整段 [0, D_m]
                a_e = audio["dur"] if rule == "special_whole_clip" else e
                row, ra = pool_interval(audio["feats"], audio["times"], s, a_e, audio["dur"])
                A[i] = row
                rec["audio_frames"], rec["audio_rule"] = ra["frames"], ra["rule"]
            if vision is not None:
                v_e = vision["dur"] if rule == "special_whole_clip" else e
                row, rv = pool_interval(vision["feats"], vision["times"], s, v_e, vision["dur"], usable=vision["face"])
                V[i] = row
                if rv["frames"] is not None:
                    lo, hi = rv["frames"]
                    rec["video_frames"] = [int(vision["frame_idx"][lo]), int(vision["frame_idx"][hi - 1]) + 1]
                rec["n_video_frames"], rec["n_face_frames"], rec["vision_rule"] = rv["n_frames"], rv["n_used"], rv["rule"]
        positions.append(rec)
    return {"audio": A, "vision": V, "positions": positions, "length": L,
            "audio_length": L if audio is not None else 0, "vision_length": L if vision is not None else 0}


def _grid_pool(feats: np.ndarray, times: np.ndarray, K: int, grid: float, usable: np.ndarray | None):
    """把帧均值池化到 K 个栅格位置。返回 (X (K,D), first (K,), last_excl (K,), n_frames (K,), n_used (K,), rule list)。"""
    D = feats.shape[1]
    X = np.zeros((K, D), np.float32)
    first = np.full(K, -1, np.int64)
    last = np.full(K, -1, np.int64)
    n_fr = np.zeros(K, np.int64)
    n_use = np.zeros(K, np.int64)
    rules = ["no_frames"] * K
    if K == 0 or len(times) == 0:
        return X, first, last, n_fr, n_use, rules
    pos = np.minimum(np.floor(times / grid + EPS_T).astype(np.int64), K - 1)
    ok = np.ones(len(times), bool) if usable is None else usable.astype(bool)
    for k in range(K):
        sel = np.where(pos == k)[0]
        rule = "mean"
        if sel.size == 0:
            sel = np.array([_nearest(times, (k + 0.5) * grid)])
            rule = "nearest"
        first[k], last[k], n_fr[k] = sel[0], sel[-1] + 1, sel.size
        use = sel[ok[sel]]
        n_use[k] = use.size
        if use.size:
            X[k] = feats[use].mean(axis=0)
            rules[k] = rule
        else:
            rules[k] = rule + "_no_face"
    return X, first, last, n_fr, n_use, rules


def build_unaligned(audio: dict | None, vision: dict | None, grid: float = 0.08, max_len: int = 500) -> dict:
    out = {"grid_s": grid}
    for m, src, D in (("audio", audio, 74), ("vision", vision, 35)):
        X = np.zeros((max_len, D), np.float32)
        rec = {"length": 0, "length_untruncated": 0, "truncated_s": 0.0}
        if src is not None and src["dur"] > 0:
            K_full = int(np.ceil(src["dur"] / grid - EPS_T))
            K = min(K_full, max_len)
            usable = src.get("face") if m == "vision" else None
            Xk, first, last, n_fr, n_use, rules = _grid_pool(src["feats"], src["times"], K, grid, usable)
            X[:K] = Xk
            fidx = src.get("frame_idx")
            if fidx is not None and K:  # 视觉：记录原视频帧号范围
                has = first >= 0
                f_first = np.where(has, fidx[np.maximum(first, 0)], -1)
                f_last = np.where(has, fidx[np.maximum(last - 1, 0)] + 1, -1)
            else:
                f_first, f_last = first, last
            rec = {"length": K, "length_untruncated": K_full, "truncated_s": max(0.0, src["dur"] - K * grid),
                   "start_s": [round(k * grid, 4) for k in range(K)],
                   "end_s": [round(min((k + 1) * grid, src["dur"]), 4) for k in range(K)],
                   "frame_start": f_first.tolist(), "frame_end": f_last.tolist(), "n_frames": n_fr.tolist(),
                   "n_used": n_use.tolist(), "rule": rules}
        out[m] = X
        out[f"{m}_record"] = rec
        out[f"{m}_length"] = rec["length"]
    return out
