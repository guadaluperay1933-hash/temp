"""词级时间戳：把"给定的转写文本"强制对齐到语音上（不做语音识别，不改动文本）。

接口（问题3 直接 import，签名不要改）：
    align_words(audio, sr, transcript, method='auto', device=None, **kw) -> list[dict(word, start, end, source, ...)]
    返回的列表与 transcript.split() 一一对应（同样的词、同样的顺序），start/end 单位秒，单调不减且落在 [0, 音频时长]。

两级方案：
  * 'ctc'：wav2vec2 CTC 声学模型（默认 facebook/wav2vec2-base-960h，字符级英文）输出逐帧对数后验
    E ∈ R^{T×C}（帧移 20 ms），在"只允许按给定文本顺序输出"的约束下用自写的 Viterbi 动态规划求最优 CTC 路径：
        扩展标签序列 z = (b, y_1, b, y_2, …, b, y_U, b)，长 S = 2U+1（b 为空白）；
        δ_0(0)=E[0,b]，δ_0(1)=E[0,y_1]，其余 −∞；
        δ_t(s) = E[t, z_s] + max{ δ_{t−1}(s), δ_{t−1}(s−1), δ_{t−1}(s−2)·1[z_s≠b ∧ z_s≠z_{s−2}] }；
        终点取 max{δ_{T−1}(S−1), δ_{T−1}(S−2)}，回溯得到每帧所在的 s。
    词 w 的区间 = [其第一个字符首次出现的帧, 其最后一个字符最后出现的帧 + 1) × 20 ms。
    文本规范化：大写，只保留 A–Z 与撇号 '，词间以 '|' 分隔；含数字的词（如 "1990"）和规范化后为空的词
    不进入 CTC 目标序列，时间由前后已对齐词之间的空隙均分插值得到（source='interp'）。
  * 'vad_proportional'（回退）：能量 VAD（librosa.effects.split，top_db=30）得到有声段，
    把所有词按字符数比例铺在拼接后的有声时间轴上，再映射回真实时间；跨越停顿的词只保留交集最长的那个有声段
    （source='vad_proportional'）。
  * 'auto'：CTC 模型能加载就用 'ctc'，否则回退并在日志里写明；单条样本 CTC 失败（如帧数不足）也逐条回退。

为什么不用 torchaudio.functional.forced_align：新版 torchaudio 已宣布移除该函数，自写 DP 只依赖 numpy，
行为完全可控、可单元测试（tests/test_q1.py 用手工构造的发射矩阵 + 穷举验证最优路径）。
"""
from __future__ import annotations

import logging
import re

import numpy as np

DEFAULT_CTC_MODEL = "facebook/wav2vec2-base-960h"
CTC_SR = 16000
_LOG = logging.getLogger("q1.align")
_MODEL_CACHE: dict = {}
_FAILED_MODELS: dict = {}


# ----------------------------------------------------------------------------- 文本规范化

def normalize_word(w: str) -> str:
    """大写 + 只保留 A–Z 和撇号；含数字的词返回空串（其读音与字母无关，交给插值）。"""
    if re.search(r"\d", w):
        return ""
    s = re.sub(r"[^A-Z']", "", w.upper())
    return s if re.search(r"[A-Z]", s) else ""


# ----------------------------------------------------------------------------- Viterbi 强制对齐

def ctc_viterbi(emission: np.ndarray, targets, blank: int = 0) -> tuple[np.ndarray, float]:
    """CTC 强制对齐（最优单路径）。

    emission: (T, C) 对数概率；targets: 长 U 的标签序列（不含空白）。
    返回 (path_states (T,) —— 每帧所处扩展序列下标 s ∈ [0, 2U]，偶数=空白、奇数 2u+1=第 u 个标签；最优路径对数分)。
    无可行路径（T 小于所需最少帧数）时抛 ValueError。
    """
    E = np.asarray(emission, dtype=np.float64)
    T = E.shape[0]
    y = np.asarray(list(targets), dtype=np.int64)
    U = len(y)
    if U == 0:
        return np.zeros(T, np.int64), float(E[:, blank].sum())
    n_rep = int((y[1:] == y[:-1]).sum())
    if T < U + n_rep:
        raise ValueError(f"帧数 T={T} 少于对齐 {U} 个标签所需的最少帧数 {U + n_rep}")
    S = 2 * U + 1
    z = np.full(S, blank, dtype=np.int64)
    z[1::2] = y
    # s-2 跳转只允许：z_s 不是空白，且与 z_{s-2} 不同（连续相同字符之间必须隔一个空白）
    skip = np.zeros(S, dtype=bool)
    skip[3::2] = y[1:] != y[:-1]
    NEG = -np.inf
    delta = np.full(S, NEG)
    delta[0] = E[0, z[0]]
    delta[1] = E[0, z[1]]
    back = np.zeros((T, S), dtype=np.int8)  # 0: 停留, 1: 从 s-1 来, 2: 从 s-2 来
    for t in range(1, T):
        stay = delta
        from1 = np.concatenate([[NEG], delta[:-1]])
        from2 = np.where(skip, np.concatenate([[NEG, NEG], delta[:-2]]), NEG)
        cand = np.stack([stay, from1, from2])
        arg = np.argmax(cand, axis=0)
        best = cand[arg, np.arange(S)]
        delta = best + E[t, z]
        back[t] = arg
    if delta[S - 1] >= delta[S - 2]:
        s, score = S - 1, delta[S - 1]
    else:
        s, score = S - 2, delta[S - 2]
    if not np.isfinite(score):
        raise ValueError("无可行 CTC 路径")
    path = np.zeros(T, np.int64)
    for t in range(T - 1, -1, -1):
        path[t] = s
        s -= int(back[t, s])
    if path[0] not in (0, 1):
        raise ValueError("回溯起点异常")
    return path, float(score)


def path_to_label_spans(path: np.ndarray, U: int) -> list[tuple[int, int]]:
    """由 Viterbi 路径求第 u 个标签占据的帧区间 [first, last]（闭区间；奇数状态 2u+1 的帧）。"""
    spans = []
    for u in range(U):
        fr = np.where(path == 2 * u + 1)[0]
        spans.append((int(fr[0]), int(fr[-1])) if fr.size else (-1, -1))
    return spans


# ----------------------------------------------------------------------------- CTC 模型

class CTCAligner:
    """wav2vec2 CTC 模型封装：输出 (T, C) 对数后验与帧移（秒）。"""

    def __init__(self, model_name: str = DEFAULT_CTC_MODEL, device=None):
        import torch
        from transformers import AutoFeatureExtractor, AutoTokenizer, Wav2Vec2ForCTC

        from common.utils import get_device

        self.model_name = model_name
        self.device = get_device(device)
        self.feature_extractor = AutoFeatureExtractor.from_pretrained(model_name)
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
        self.model = Wav2Vec2ForCTC.from_pretrained(model_name).eval().to(self.device)
        vocab = self.tokenizer.get_vocab()
        self.vocab = {k: int(v) for k, v in vocab.items()}
        self.blank = int(self.tokenizer.pad_token_id if self.tokenizer.pad_token_id is not None else 0)
        self.sep = self.vocab.get("|", self.vocab.get(" "))
        # 卷积特征提取器的总步长（wav2vec2-base 为 320 点 = 20 ms）
        stride = int(np.prod(self.model.config.conv_stride))
        self.frame_s = stride / CTC_SR
        self._torch = torch

    def info(self) -> dict:
        cfg = self.model.config
        return {"model": self.model_name, "name_or_path": getattr(cfg, "_name_or_path", None),
                "commit_hash": getattr(cfg, "_commit_hash", None), "vocab_size": len(self.vocab),
                "blank_id": self.blank, "frame_s": self.frame_s, "device": str(self.device)}

    def emission(self, audio16k: np.ndarray, chunk_s: float = 60.0) -> np.ndarray:
        """(T, C) 对数后验。长音频按 chunk_s 分块推理后按帧拼接（块边界处会损失少量上下文）。"""
        torch = self._torch
        x = np.asarray(audio16k, np.float32)
        step = int(chunk_s * CTC_SR)
        outs = []
        with torch.no_grad():
            for b in range(0, max(len(x), 1), step):
                seg = x[b:b + step]
                if len(seg) < 400:  # 不足一个卷积感受野：补零
                    seg = np.pad(seg, (0, 400 - len(seg)))
                inp = self.feature_extractor(seg, sampling_rate=CTC_SR, return_tensors="pt")
                logits = self.model(inp.input_values.to(self.device)).logits[0]
                outs.append(torch.log_softmax(logits.float(), dim=-1).cpu().numpy())
        return np.concatenate(outs, axis=0)

    def encode_words(self, norm_words: list[str]) -> tuple[list[int], list[tuple[int, int]]]:
        """规范化词 → 目标标签序列（词间插 '|'）+ 每个词在标签序列中的 [起, 止) 下标。词表外字符丢弃。"""
        labels, spans = [], []
        for k, w in enumerate(norm_words):
            if k > 0 and self.sep is not None:
                labels.append(self.sep)
            s = len(labels)
            labels.extend(self.vocab[ch] for ch in w if ch in self.vocab)
            spans.append((s, len(labels)))
        return labels, spans


def get_ctc_aligner(model_name: str = DEFAULT_CTC_MODEL, device=None) -> CTCAligner | None:
    """带缓存的加载；加载失败返回 None（并缓存失败原因，避免每条样本重复尝试联网）。"""
    key = (model_name, str(device))
    if key in _MODEL_CACHE:
        return _MODEL_CACHE[key]
    if key in _FAILED_MODELS:
        return None
    try:
        _MODEL_CACHE[key] = CTCAligner(model_name, device)
        return _MODEL_CACHE[key]
    except Exception as e:  # noqa: BLE001 - 网络/文件问题都走回退
        _FAILED_MODELS[key] = f"{type(e).__name__}: {str(e)[:300]}"
        _LOG.warning("CTC 对齐模型 %s 加载失败，改用 VAD 比例回退：%s", model_name, _FAILED_MODELS[key])
        return None


def ctc_failure_reason(model_name: str = DEFAULT_CTC_MODEL, device=None) -> str | None:
    return _FAILED_MODELS.get((model_name, str(device)))


# ----------------------------------------------------------------------------- 插值与回退

def _fill_unaligned(words: list[dict], duration: float) -> None:
    """start 为 None 的词：在前一个已定位词的 end 与后一个已定位词的 start 之间均分（连续多个时依次排开）。"""
    n = len(words)
    i = 0
    while i < n:
        if words[i]["start"] is not None:
            i += 1
            continue
        j = i
        while j < n and words[j]["start"] is None:
            j += 1
        left = words[i - 1]["end"] if i > 0 else 0.0
        right = words[j]["start"] if j < n else duration
        right = max(right, left)
        k = j - i
        for q in range(k):
            words[i + q]["start"] = left + (right - left) * q / k
            words[i + q]["end"] = left + (right - left) * (q + 1) / k
            words[i + q]["source"] = "interp"
        i = j


def _proportional(words: list[str], intervals: np.ndarray, duration: float) -> list[tuple[float, float]]:
    """把词按字符数权重铺在有声段上。intervals (K,2) 秒；为空时视整段为有声。

    记有声段总长 V、词 j 的权重 c_j = max(1, 字母数字个数)，累计位置 C_j = V·Σ_{i<j}c_i / Σc_i，
    词 j 占据拼接有声时间轴上的 [C_j, C_{j+1}]；由分段平移 g（第 k 段：g(c) = a_k + c − cum_k）映回真实时间。
    若 [C_j, C_{j+1}] 跨越了有声段边界（即跨过一个停顿），只保留它与各有声段交集中最长的那一段，
    —— 一个词不应覆盖静音停顿；这样得到的词区间仍然有序、互不重叠，且都落在有声段内。
    """
    if len(words) == 0:
        return []
    iv = np.asarray(intervals, np.float64).reshape(-1, 2)
    iv = iv[iv[:, 1] > iv[:, 0]]
    if iv.size == 0:
        iv = np.array([[0.0, duration]])
    seg_len = iv[:, 1] - iv[:, 0]
    cum = np.concatenate([[0.0], np.cumsum(seg_len)])
    V = cum[-1]
    w = np.array([max(1, len(re.sub(r"[^A-Za-z0-9]", "", x))) for x in words], np.float64)
    C = np.concatenate([[0.0], np.cumsum(w)]) / w.sum() * V

    K = len(seg_len)
    out = []
    for j in range(len(words)):
        a, b = min(max(C[j], 0.0), V), min(max(C[j + 1], 0.0), V)
        k0 = min(max(int(np.searchsorted(cum, a, side="right") - 1), 0), K - 1)  # a 落在段边界时归下一段
        k1 = min(max(int(np.searchsorted(cum, b, side="left") - 1), 0), K - 1)   # b 落在段边界时归上一段
        k1 = max(k1, k0)
        ks = np.arange(k0, k1 + 1)
        ov = np.minimum(b, cum[ks + 1]) - np.maximum(a, cum[ks])
        k = int(ks[int(np.argmax(ov))])
        lo, hi = max(a, cum[k]), min(b, cum[k + 1])
        s = float(iv[k, 0] + (lo - cum[k]))
        e = float(iv[k, 0] + (max(hi, lo) - cum[k]))
        out.append((s, e))
    return out


def _vad_align(audio: np.ndarray, sr: int, words: list[str], duration: float, top_db: float) -> list[dict]:
    from .audio_feats import voiced_intervals

    iv = voiced_intervals(audio, sr, top_db=top_db, hop=max(1, int(0.01 * sr)), win=max(2, int(0.025 * sr)))
    spans = _proportional(words, iv, duration)
    return [{"word": w, "start": s, "end": e, "source": "vad_proportional", "score": None}
            for w, (s, e) in zip(words, spans)]


def _ctc_align(aligner: CTCAligner, audio16k: np.ndarray, words: list[str], duration: float,
               chunk_s: float = 60.0) -> list[dict]:
    norm = [normalize_word(w) for w in words]
    idx = [i for i, w in enumerate(norm) if w]
    out = [{"word": w, "start": None, "end": None, "source": "ctc", "score": None} for w in words]
    if idx:
        labels, spans = aligner.encode_words([norm[i] for i in idx])
        E = aligner.emission(audio16k, chunk_s=chunk_s)
        path, _ = ctc_viterbi(E, labels, blank=aligner.blank)
        lab_spans = path_to_label_spans(path, len(labels))
        probs = np.exp(E)
        for k, i in enumerate(idx):
            a, b = spans[k]
            if b <= a:  # 词的所有字符都不在词表中
                continue
            f0 = lab_spans[a][0]
            f1 = lab_spans[b - 1][1]
            frames = np.where((path >= 2 * a + 1) & (path <= 2 * (b - 1) + 1) & (path % 2 == 1))[0]
            sc = float(np.mean(probs[frames, np.asarray(labels)[(path[frames] - 1) // 2]])) if frames.size else None
            out[i]["start"] = min(f0 * aligner.frame_s, duration)
            out[i]["end"] = min((f1 + 1) * aligner.frame_s, duration)
            out[i]["score"] = sc
    _fill_unaligned(out, duration)
    return out


def _resample(audio: np.ndarray, sr: int, target: int) -> np.ndarray:
    if sr == target:
        return np.asarray(audio, np.float32)
    from math import gcd

    from scipy.signal import resample_poly

    g = gcd(int(sr), int(target))
    return resample_poly(np.asarray(audio, np.float64), target // g, sr // g).astype(np.float32)


def _finalize(words: list[dict], duration: float) -> list[dict]:
    """保证单调、落在 [0, duration] 内，并补上序号。"""
    prev_end = 0.0
    for k, w in enumerate(words):
        s = float(min(max(w["start"], 0.0), duration))
        e = float(min(max(w["end"], 0.0), duration))
        s = max(s, prev_end)   # 各方法本身已保证词区间有序且不重叠，这里只防数值误差
        e = max(e, s)
        w["start"], w["end"], w["index"] = round(s, 4), round(e, 4), k
        prev_end = w["end"]
    return words


def align_words(audio, sr, transcript, method: str = "auto", device=None, model_name: str | None = None,
                vad_top_db: float = 30.0, chunk_s: float = 60.0, return_info: bool = False, **_):
    """给定转写文本的词级强制对齐。

    audio: 一维波形（任意采样率 sr）；transcript: 字符串（按空白切词）或词列表。
    method: 'auto' | 'ctc' | 'vad_proportional'。
    返回 list[dict(index, word, start, end, source, score)]；return_info=True 时返回 (words, info)。
    """
    words = transcript if isinstance(transcript, (list, tuple)) else (str(transcript).split() if transcript else [])
    audio = np.asarray(audio, np.float32).reshape(-1)
    duration = len(audio) / float(sr) if sr else 0.0
    info = {"method_requested": method, "method_used": None, "model": None, "fallback_reason": None,
            "duration_s": duration}
    if not words:
        info["method_used"] = "none"
        return ([], info) if return_info else []
    if duration <= 0:
        out = [{"word": w, "start": 0.0, "end": 0.0, "source": "no_audio", "score": None} for w in words]
        info["method_used"] = "no_audio"
        out = _finalize(out, 0.0)
        return (out, info) if return_info else out
    model_name = model_name or DEFAULT_CTC_MODEL
    res = None
    if method in ("auto", "ctc"):
        aligner = get_ctc_aligner(model_name, device)
        if aligner is None:
            reason = ctc_failure_reason(model_name, device)
            if method == "ctc":
                raise RuntimeError(f"CTC 对齐模型加载失败：{reason}")
            info["fallback_reason"] = f"model_load_failed: {reason}"
        else:
            try:
                a16 = _resample(audio, int(sr), CTC_SR)
                res = _ctc_align(aligner, a16, list(words), duration, chunk_s=chunk_s)
                info["method_used"], info["model"] = "ctc", model_name
            except Exception as e:  # noqa: BLE001 - 单条失败逐条回退
                if method == "ctc":
                    raise
                info["fallback_reason"] = f"ctc_failed: {type(e).__name__}: {e}"
                res = None
    elif method != "vad_proportional":
        raise ValueError(f"未知对齐方法：{method}")
    if res is None:
        res = _vad_align(audio, int(sr), list(words), duration, vad_top_db)
        info["method_used"] = "vad_proportional"
    res = _finalize(res, duration)
    return (res, info) if return_info else res
