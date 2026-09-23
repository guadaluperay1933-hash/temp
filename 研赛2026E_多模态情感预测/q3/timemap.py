"""序列位置 → 原始素材时间 的映射（附件4 回看原始视频用）。

* 视频信息 / 音频 / 取帧 / 词级对齐：优先调用问题1 的 q1.media.probe_video / load_audio / read_frame_at、
  q1.align.align_words（延迟导入）；导入失败时退回本模块自带的 OpenCV / imageio-ffmpeg 实现与"按字符数比例"的词时间。
* 文本位置 → 词：能加载 BERT 分词器（配置 timemap.tokenizer）时，逐词分出 wordpiece，
  有效范围首位置 = [CLS]、末位置 = [SEP]，中间依次为各词的 wordpiece（超过 50 截断）；
  加载失败或长度对不上时退回"位置 1..n 依次对应空格分词的第 1..n 个词"（长度不等时按比例对应），并记录回退原因。
  词 w 的第 q 个（共 Q 个）wordpiece 的时间 = 把词区间 [s_w, e_w] 等分为 Q 段中的第 q 段；
  [CLS] = [0, s_1]，[SEP] = [e_n, D]。
* 对齐版本：语音/视觉位置 i 与文本词元 i 共用同一时间区间（word-aligned）。
* 非对齐版本：在附件4 全部样本上估计每个模态的特征帧率 —— 有效长度 L 对视频时长 D 做过原点最小二乘
  L ≈ r·D（排除 L 达到序列上限的截断样本），r = Σ L_i D_i / Σ D_i²，
  R² = 1 − Σ(L_i − r D_i)² / Σ(L_i − L̄)²；R² ≥ 阈值（0.9）→ 固定帧率 t_k = [k/r, (k+1)/r]（fixed-rate），
  否则均匀铺满 t_k = [k·D/L, (k+1)·D/L]（uniform）。k 为有效范围内的序号。决策写入 timemap_report.json。
* 视觉关键帧：帧号 = round(t_center · fps)（截到 [0, n_frames−1]）。
"""
from __future__ import annotations

import logging
import re
import subprocess
from functools import lru_cache
from pathlib import Path

import numpy as np

from common.data import MODALITIES, split_id

_LOG = logging.getLogger("q3.timemap")
VIDEO_EXTS = (".mp4", ".avi", ".mkv", ".mov", ".webm", ".flv", ".m4v")


# ============================================================================ 媒体工具（q1 优先，失败回退）

@lru_cache(maxsize=None)
def _q1(module: str):
    try:
        if module == "media":
            from q1 import media
            return media
        from q1 import align
        return align
    except Exception as e:  # noqa: BLE001 - q1 由另一部分代码提供，缺失时回退
        _LOG.warning(f"无法导入 q1.{module}（{type(e).__name__}: {e}），改用问题3 自带的回退实现")
        return None


def _ffmpeg() -> str:
    try:
        import imageio_ffmpeg

        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:  # noqa: BLE001
        import shutil

        exe = shutil.which("ffmpeg")
        if not exe:
            raise FileNotFoundError("找不到 ffmpeg")
        return exe


def probe(path) -> dict:
    media = _q1("media")
    if media is not None:
        try:
            return media.probe_video(path)
        except Exception as e:  # noqa: BLE001
            _LOG.warning(f"q1.media.probe_video 失败（{e}），改用 OpenCV")
    import cv2

    cap = cv2.VideoCapture(str(path))
    fps = float(cap.get(cv2.CAP_PROP_FPS)) or None
    n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) or None
    w, h = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH)) or None, int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT)) or None
    cap.release()
    return {"duration": (n / fps) if (n and fps) else None, "fps": fps, "n_frames": n, "width": w, "height": h,
            "has_audio": None}


def load_audio(path, sr: int = 16000) -> np.ndarray:
    media = _q1("media")
    if media is not None:
        try:
            return media.load_audio(path, sr=sr)
        except Exception as e:  # noqa: BLE001
            _LOG.warning(f"q1.media.load_audio 失败（{e}），改用 ffmpeg 管道")
    cmd = [_ffmpeg(), "-hide_banner", "-nostdin", "-loglevel", "error", "-i", str(path), "-vn", "-ac", "1",
           "-ar", str(int(sr)), "-f", "f32le", "-acodec", "pcm_f32le", "pipe:1"]
    res = subprocess.run(cmd, capture_output=True, timeout=600)
    if res.returncode != 0:
        return np.zeros(0, np.float32)
    return np.frombuffer(res.stdout, dtype="<f4").astype(np.float32, copy=True)


def read_frame(path, t_sec: float) -> np.ndarray | None:
    media = _q1("media")
    if media is not None:
        try:
            return media.read_frame_at(path, t_sec)
        except Exception as e:  # noqa: BLE001
            _LOG.warning(f"q1.media.read_frame_at 失败（{e}），改用 OpenCV")
    import cv2

    cap = cv2.VideoCapture(str(path))
    try:
        fps = float(cap.get(cv2.CAP_PROP_FPS)) or 25.0
        n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) or 0
        idx = int(round(max(0.0, t_sec) * fps))
        if n:
            idx = min(idx, n - 1)
        cap.set(cv2.CAP_PROP_POS_FRAMES, idx)
        ok, bgr = cap.read()
        return cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB) if ok else None
    finally:
        cap.release()


def proportional_words(words: list[str], duration: float) -> list[dict]:
    """回退的词时间：按字符数比例把词铺满 [0, D]。"""
    if not words:
        return []
    w = np.array([max(1, len(re.sub(r"[^A-Za-z0-9]", "", x))) for x in words], np.float64)
    c = np.concatenate([[0.0], np.cumsum(w)]) / w.sum() * duration
    return [{"word": x, "start": float(c[j]), "end": float(c[j + 1]), "source": "proportional"}
            for j, x in enumerate(words)]


def word_times(audio: np.ndarray, sr: int, text: str, duration: float, method: str = "auto") -> tuple[list[dict], str]:
    """词级时间戳（与 text.split() 一一对应）。返回 (words, 来源)。"""
    words = str(text or "").split()
    if not words:
        return [], "none"
    align = _q1("align")
    if align is not None and audio is not None and len(audio) > 0:
        try:
            res = align.align_words(audio, sr, words, method=method)
            if len(res) == len(words):
                src = sorted({str(r.get("source")) for r in res})
                return [{"word": r["word"], "start": float(r["start"]), "end": float(r["end"]),
                         "source": r.get("source")} for r in res], "+".join(src)
        except Exception as e:  # noqa: BLE001
            _LOG.warning(f"q1.align.align_words 失败（{type(e).__name__}: {e}），按字符比例分配词时间")
    return proportional_words(words, duration), "proportional"


def find_video(videos_dir, sample_id: str) -> Path | None:
    """<dir>/<video_id>/<clip_id>.<ext>；也兼容 <dir>/<video_id>_<clip_id>.<ext> 与 <dir>/<id>.<ext>。"""
    if not videos_dir:
        return None
    root = Path(videos_dir)
    vid, cid = split_id(sample_id)
    stems = [root / vid / cid, root / f"{vid}_{cid}", root / f"{vid}$_${cid}", root / vid / f"{vid}_{cid}"]
    for s in stems:
        for ext in VIDEO_EXTS:
            p = s.with_name(s.name + ext)
            if p.exists():
                return p
    return None


# ============================================================================ 文本位置 → 词

class TokenWordMapper:
    """BERT 分词器可用时按 wordpiece 精确对应；否则按空格分词顺序对应（记录回退）。"""

    def __init__(self, tokenizer_name: str | None, logger=None):
        self.tok = None
        self.status = "disabled" if not tokenizer_name else None
        log = logger or _LOG
        if tokenizer_name:
            try:
                from transformers import AutoTokenizer

                try:
                    self.tok = AutoTokenizer.from_pretrained(tokenizer_name, local_files_only=True)
                except Exception:  # noqa: BLE001 - 本地没有缓存再尝试下载
                    self.tok = AutoTokenizer.from_pretrained(tokenizer_name)
                self.status = f"tokenizer:{tokenizer_name}"
            except Exception as e:  # noqa: BLE001
                self.status = f"fallback（分词器 {tokenizer_name} 加载失败：{type(e).__name__}）"
                log.warning(f"[时间映射] BERT 分词器 {tokenizer_name} 加载失败（{type(e).__name__}），"
                            f"文本位置按'位置 1..n ↔ 第 1..n 个空格分词'对应")

    def map(self, raw_text: str, n_valid: int, token_ids=None, max_len: int = 50) -> tuple[list, str]:
        """返回有效范围内每个序号 r（0..n_valid−1）对应的 (词序号 j 或 None, 第 q 片, 共 Q 片) 与所用方法。
        max_len：文本序列上限（附件2 为 50），有效长度达到上限时允许 wordpiece 被截断。"""
        words = str(raw_text or "").split()
        n_content = max(n_valid - 2, 0)
        special = [None] * n_valid
        if n_valid <= 0:
            return [], "empty"
        prefix = ""
        if self.tok is not None and words:
            pieces = []
            for j, w in enumerate(words):
                ps = self.tok.tokenize(w)
                for q, p in enumerate(ps):
                    pieces.append((j, q, len(ps), p))
            if len(pieces) == n_content or (len(pieces) > n_content and n_valid >= max_len):
                pieces = pieces[:n_content]
                ok_ids = None
                if token_ids is not None and len(token_ids) >= n_valid:
                    ids = self.tok.convert_tokens_to_ids([p[3] for p in pieces])
                    ok_ids = bool(np.array_equal(np.asarray(ids), np.asarray(token_ids[1:1 + len(pieces)])))
                out = [None] + [(j, q, Q) for j, q, Q, _ in pieces] + [None]
                return out[:n_valid], "tokenizer" + ("" if ok_ids is None else ("(ids一致)" if ok_ids else "(ids不一致)"))
            prefix = "tokenizer长度不符→"
        if not words:
            return special, "no_words"
        if n_content == len(words):
            return [None] + [(j, 0, 1) for j in range(len(words))] + [None], prefix + "whitespace-1to1"
        # 长度对不上：按比例对应（多个位置可对应同一词，按片等分词区间）
        idx = [min(len(words) - 1, int(r * len(words) / max(n_content, 1))) for r in range(n_content)]
        out = [None]
        for r, j in enumerate(idx):
            Q = idx.count(j)
            q = idx[:r].count(j)
            out.append((j, q, Q))
        out.append(None)
        return out[:n_valid], prefix + "whitespace-proportional"


# ============================================================================ 帧率估计

def estimate_rate(lengths, durations, max_len: int, r2_threshold: float = 0.9) -> dict:
    L = np.asarray(lengths, np.float64)
    D = np.asarray(durations, np.float64)
    ok = np.isfinite(D) & (D > 0) & (L > 0) & (L < max_len)
    rep = {"n_total": int(L.size), "n_used": int(ok.sum()), "n_saturated_excluded": int((L >= max_len).sum()),
           "max_len": int(max_len), "r2_threshold": float(r2_threshold)}
    if ok.sum() < 2:
        rep.update({"rate_hz": None, "r2": None, "method": "uniform", "reason": "可用样本不足 2 条，无法估计帧率"})
        return rep
    r = float((L[ok] * D[ok]).sum() / (D[ok] ** 2).sum())
    res = L[ok] - r * D[ok]
    sst = float(((L[ok] - L[ok].mean()) ** 2).sum())
    r2 = float(1.0 - (res ** 2).sum() / sst) if sst > 1e-12 else float("nan")
    r2_unc = float(1.0 - (res ** 2).sum() / max((L[ok] ** 2).sum(), 1e-12))
    method = "fixed-rate" if (np.isfinite(r2) and r2 >= r2_threshold) else "uniform"
    rep.update({"rate_hz": r, "r2": r2, "r2_uncentered": r2_unc, "method": method,
                "residual_mae_positions": float(np.abs(res).mean()),
                "reason": f"R²={r2:.4f} {'≥' if method == 'fixed-rate' else '<'} {r2_threshold}"})
    return rep


# ============================================================================ 构建时间映射

def _interval_of_piece(word: dict, q: int, Q: int) -> tuple[float, float]:
    s, e = float(word["start"]), float(word["end"])
    return s + (e - s) * q / Q, s + (e - s) * (q + 1) / Q


def build_timemaps(sd_raw, videos_dir, tcfg: dict, version: str, logger=None) -> tuple[list[dict], dict]:
    """为附件4 每条样本建立"原始位置 → 时间区间"表。sd_raw：未标准化的 SplitData（用它的 lengths/valid/raw_text）。

    返回 (maps, report)。maps[i] = {video, duration, fps, n_frames, word_source, words,
        text_map_method, pos: {m: {'start': (L_m,), 'end': (L_m,), 'word': (L_text,) 仅文本}}, method: {m: ...}}
    时间未知处为 NaN。
    """
    log = logger or _LOG
    N = len(sd_raw)
    mapper = TokenWordMapper(tcfg.get("tokenizer"), log)
    sr = int(tcfg.get("audio_sr", 16000))
    method = tcfg.get("align_method", "auto")
    maps = []
    token_ids = None
    tb = sd_raw.extra.get("text_bert")
    if tb is not None and np.asarray(tb).ndim == 3:
        token_ids = np.asarray(tb)[:, 0, :]
    n_found = 0
    for i, sid in enumerate(sd_raw.ids):
        vp = find_video(videos_dir, sid)
        info = {"video": str(vp) if vp else None, "duration": None, "fps": None, "n_frames": None}
        words, wsrc = [], "none"
        if vp is not None:
            n_found += 1
            try:
                pv = probe(vp)
                info.update({"duration": pv.get("duration"), "fps": pv.get("fps"), "n_frames": pv.get("n_frames")})
                audio = load_audio(vp, sr)
                dur = info["duration"] or (len(audio) / sr if len(audio) else None)
                info["duration"] = dur
                words, wsrc = word_times(audio, sr, sd_raw.raw_text[i] if sd_raw.raw_text else "", dur or 0.0, method)
            except Exception as e:  # noqa: BLE001
                log.warning(f"[时间映射] {sid} 读取视频失败：{type(e).__name__}: {e}")
        info["word_source"] = wsrc
        info["words"] = words
        maps.append(info)
    if videos_dir and n_found < N:
        log.warning(f"[时间映射] {N - n_found}/{N} 条样本找不到视频（{videos_dir}），这些样本只给位置、不给时间")

    # --- 非对齐：帧率估计
    rates = {}
    if version == "unaligned":
        durs = np.array([mp["duration"] if mp["duration"] else np.nan for mp in maps], np.float64)
        for m in ("audio", "vision"):
            rates[m] = estimate_rate(sd_raw.lengths[m], durs, sd_raw.feats[m].shape[1],
                                     float(tcfg.get("r2_threshold", 0.9)))
            log.info(f"[时间映射] {m} 帧率估计：r={rates[m].get('rate_hz')} Hz，{rates[m]['reason']} → {rates[m]['method']}")

    # --- 逐样本逐位置
    text_methods = {}
    for i, mp in enumerate(maps):
        D = mp["duration"]
        pos, meth = {}, {}
        Lt = sd_raw.feats["text"].shape[1]
        vt = np.where(sd_raw.valid["text"][i])[0]
        t_start, t_end = np.full(Lt, np.nan), np.full(Lt, np.nan)
        t_word = np.full(Lt, -1, np.int64)
        tmap, tmeth = mapper.map(sd_raw.raw_text[i] if sd_raw.raw_text else "", int(vt.size),
                                 token_ids[i, vt[0]:] if (token_ids is not None and vt.size) else None, max_len=Lt)
        text_methods[tmeth] = text_methods.get(tmeth, 0) + 1
        words = mp["words"]
        for r, p in enumerate(vt):
            ent = tmap[r] if r < len(tmap) else None
            if ent is not None:
                t_word[p] = ent[0]
            if D is None or not words:
                continue
            if ent is None:           # [CLS] / [SEP]
                if r == 0:
                    t_start[p], t_end[p] = 0.0, float(words[0]["start"])
                else:
                    t_start[p], t_end[p] = float(words[-1]["end"]), float(D)
            else:
                t_start[p], t_end[p] = _interval_of_piece(words[ent[0]], ent[1], ent[2])
        pos["text"] = {"start": t_start, "end": t_end, "word": t_word}
        meth["text"] = "word-aligned" if (D is not None and words) else "none"
        for m in ("audio", "vision"):
            Lm = sd_raw.feats[m].shape[1]
            s_, e_ = np.full(Lm, np.nan), np.full(Lm, np.nan)
            vm = np.where(sd_raw.valid[m][i])[0]
            if version == "aligned":
                if Lm == Lt:
                    s_, e_ = t_start.copy(), t_end.copy()
                meth[m] = meth["text"]
            elif D is not None and vm.size:
                rank = vm - vm[0]
                if rates[m]["method"] == "fixed-rate":
                    r_hz = rates[m]["rate_hz"]
                    s_[vm], e_[vm] = rank / r_hz, (rank + 1) / r_hz
                else:
                    n = float(sd_raw.lengths[m][i]) or float(vm.size)
                    s_[vm], e_[vm] = rank * D / n, (rank + 1) * D / n
                s_, e_ = np.minimum(s_, D), np.minimum(e_, D)
                meth[m] = rates[m]["method"]
            else:
                meth[m] = "none"
            pos[m] = {"start": s_, "end": e_}
        mp["pos"], mp["method"], mp["text_map_method"] = pos, meth, tmeth
    report = {"version": version, "n": N, "videos_dir": str(videos_dir) if videos_dir else None, "videos_found": n_found,
              "tokenizer_status": mapper.status, "text_position_mapping": text_methods,
              "word_time_source": _count([mp["word_source"] for mp in maps]),
              "rates": rates,
              "aligned_rule": "对齐版本：语音/视觉位置 i 与文本词元 i 共用时间区间" if version == "aligned" else None}
    return maps, report


def _count(xs) -> dict:
    out = {}
    for x in xs:
        out[str(x)] = out.get(str(x), 0) + 1
    return out


def pooled_interval(mp: dict, m: str, k: int, stride: int, valid_row: np.ndarray) -> tuple[float, float] | None:
    """下采样位置 k（原始 [k·s, (k+1)·s)）的时间区间 = 窗口内有效原始位置的 [min start, max end]。"""
    lo, hi = k * stride, min((k + 1) * stride, len(valid_row))
    idx = [p for p in range(lo, hi) if valid_row[p]]
    if not idx:
        return None
    s = mp["pos"][m]["start"][idx]
    e = mp["pos"][m]["end"][idx]
    if not np.isfinite(s).any():
        return None
    return float(np.nanmin(s)), float(np.nanmax(e))


def frame_index(t: float, fps: float | None, n_frames: int | None) -> int | None:
    if fps is None or t is None or not np.isfinite(t):
        return None
    k = int(round(t * fps))
    if n_frames:
        k = min(k, int(n_frames) - 1)
    return max(k, 0)
