"""语音模态：16 kHz 单声道波形 → 100 Hz 帧级 74 维声学特征（维度布局仿 COVAREP 的 74 维）。

分帧：窗长 25 ms（400 点）、帧移 10 ms（160 点）、Hann 窗、n_fft=512（补零）、center=True，
第 k 帧的中心时刻 τ_k = k·hop/sr = 0.01·k 秒，帧数 T = 1 + ⌊n/hop⌋（最后一帧中心 ≤ 音频时长）。

74 维 = 11 维韵律/音质 + 25 维 MFCC + 25 维 ΔMFCC + 13 维 Δ²MFCC：
     0  F0（Hz，pYIN；清音/静音帧为 0）            1  VUV 浊音标志（0/1）
     2  浊音概率（pYIN voiced_prob）                 3  对数能量 20·log10(max(RMS, 1e-5)) dB（下限 -100）
     4  过零率 ZCR                                   5  谱质心（Hz）
     6  谱带宽（Hz，p=2）                            7  85% 谱滚降频率（Hz）
     8  谱平坦度（0~1）                              9  H1−H2（dB；浊音帧，F0 与 2F0 处谐波幅度差；否则 0）
    10  谱倾斜（dB/kHz；对数幅度谱对频率的最小二乘斜率）
 11–35  MFCC 0–24（40 个 Mel 滤波器，0–sr/2 Hz）
 36–60  ΔMFCC 0–24（librosa.feature.delta，width=9）
 61–73  Δ²MFCC 0–12（order=2，width=9）

静音/清音帧仍保留其（非零的）谱特征，只有序列的填充位置才整行为 0 —— 这与 common.data 的
"整行为 0 = 不可用"语义一致：真实存在的音频帧永远可用。
F0 用更长的分析窗（默认 64 ms=1024 点）：pYIN 需要窗内至少包含两个最低基频周期（65 Hz → 15.4 ms/周期）。
"""
from __future__ import annotations

import numpy as np

AUDIO_DIM = 74
AUDIO_FEATURE_NAMES = (
    ["F0_Hz", "VUV", "voiced_prob", "logRMS_dB", "ZCR", "spec_centroid_Hz", "spec_bandwidth_Hz",
     "spec_rolloff85_Hz", "spec_flatness", "H1_H2_dB", "spec_tilt_dB_per_kHz"]
    + [f"MFCC{i}" for i in range(25)] + [f"dMFCC{i}" for i in range(25)] + [f"ddMFCC{i}" for i in range(13)]
)
assert len(AUDIO_FEATURE_NAMES) == AUDIO_DIM

DEFAULT_AUDIO_CFG = {
    "sr": 16000, "win_ms": 25.0, "hop_ms": 10.0, "n_fft": 512, "window": "hann", "n_mels": 40,
    "fmin_mel": 0.0, "fmax_mel": None, "n_mfcc": 25, "n_delta2": 13, "delta_width": 9,
    "f0_method": "pyin", "f0_min": 65.0, "f0_max": 500.0, "f0_frame_ms": 64.0,
    "rolloff_percent": 0.85, "energy_floor_db": -100.0, "h1h2_search_rel": 0.15, "vad_top_db": 30.0,
}


def _cfg(cfg: dict | None) -> dict:
    c = dict(DEFAULT_AUDIO_CFG)
    if cfg:
        c.update({k: v for k, v in cfg.items() if v is not None or k in ("fmax_mel",)})
    return c


def frame_params(cfg: dict | None = None) -> tuple[int, int, int]:
    c = _cfg(cfg)
    sr = int(c["sr"])
    win = int(round(c["win_ms"] * sr / 1000))
    hop = int(round(c["hop_ms"] * sr / 1000))
    return sr, win, hop


def n_audio_frames(n_samples: int, hop: int) -> int:
    return 0 if n_samples <= 0 else 1 + n_samples // hop


def _f0(y: np.ndarray, sr: int, hop: int, T: int, c: dict):
    """返回 (f0_hz, vuv, voiced_prob)，长度 T。清音帧 f0=0。"""
    import librosa

    frame_len = int(2 ** np.ceil(np.log2(c["f0_frame_ms"] * sr / 1000)))
    if len(y) < frame_len:  # 极短音频补零到一个 F0 分析窗（帧数仍由调用方的 T 决定）
        y = np.pad(y, (0, frame_len - len(y)))
    if c["f0_method"] == "pyin":
        f0, vflag, vprob = librosa.pyin(y, fmin=c["f0_min"], fmax=c["f0_max"], sr=sr, frame_length=frame_len,
                                        hop_length=hop, center=True)
        vflag = vflag.astype(np.float32)
        vprob = np.nan_to_num(vprob.astype(np.float32))
    else:  # 'yin'：更快，无浊音概率 → 用 RMS 门限给出 VUV，概率列取同值
        f0 = librosa.yin(y, fmin=c["f0_min"], fmax=c["f0_max"], sr=sr, frame_length=frame_len, hop_length=hop,
                         center=True)
        rms = librosa.feature.rms(y=y, frame_length=frame_len, hop_length=hop, center=True)[0]
        vflag = (20 * np.log10(np.maximum(rms, 1e-5)) > 20 * np.log10(max(rms.max(), 1e-5)) - c["vad_top_db"])
        vflag = vflag.astype(np.float32)
        vprob = vflag.copy()
    f0 = np.nan_to_num(np.asarray(f0, np.float32)) * (vflag > 0)
    return _fit_len(f0, T), _fit_len(vflag, T), _fit_len(vprob, T)


def _fit_len(x: np.ndarray, T: int) -> np.ndarray:
    """各 librosa 函数在 center=True 下帧数应一致；保险起见截断/末帧延拓到 T。"""
    x = np.asarray(x)
    if x.shape[-1] == T:
        return x
    if x.shape[-1] > T:
        return x[..., :T]
    pad = [(0, 0)] * (x.ndim - 1) + [(0, T - x.shape[-1])]
    return np.pad(x, pad, mode="edge") if x.shape[-1] > 0 else np.zeros(x.shape[:-1] + (T,), x.dtype)


def _harmonic_amp_db(mag: np.ndarray, freqs: np.ndarray, f0: np.ndarray, k: int, rel: float) -> np.ndarray:
    """第 k 次谐波幅度（dB）：在 k·F0 ± max(1 bin, rel·F0) 范围内取幅度谱最大值。"""
    df = freqs[1] - freqs[0]
    T = mag.shape[1]
    out = np.zeros(T, np.float32)
    for t in np.where(f0 > 0)[0]:
        fc = k * f0[t]
        half = max(df, rel * f0[t])
        lo = int(np.floor((fc - half) / df))
        hi = int(np.ceil((fc + half) / df))
        lo, hi = max(lo, 1), min(hi, len(freqs) - 1)
        if hi < lo:
            continue
        out[t] = 20 * np.log10(max(float(mag[lo:hi + 1, t].max()), 1e-10))
    return out


def extract_audio_features(y: np.ndarray, cfg: dict | None = None) -> tuple[np.ndarray, np.ndarray, dict]:
    """y: float32 单声道（采样率 = cfg['sr']）→ (feats (T,74) float32, times (T,) 帧中心秒, info)。"""
    import librosa

    c = _cfg(cfg)
    sr, win, hop = frame_params(c)
    y = np.asarray(y, np.float32)
    T = n_audio_frames(len(y), hop)
    info = {"sr": sr, "win": win, "hop": hop, "n_fft": int(c["n_fft"]), "n_frames": T,
            "duration_s": len(y) / sr if sr else 0.0}
    if T == 0:
        return np.zeros((0, AUDIO_DIM), np.float32), np.zeros(0), info
    n_fft = max(int(c["n_fft"]), win)
    if len(y) < n_fft:  # 极短音频：补零到一个分析窗，帧数仍按原长度计
        y_an = np.pad(y, (0, n_fft - len(y)))
    else:
        y_an = y

    S = librosa.stft(y_an, n_fft=n_fft, hop_length=hop, win_length=win, window=c["window"], center=True)
    mag = _fit_len(np.abs(S), T)
    power = mag ** 2
    freqs = librosa.fft_frequencies(sr=sr, n_fft=n_fft)

    f0, vuv, vprob = _f0(y_an, sr, hop, T, c)

    # 对数能量：时域 25 ms 矩形帧的 RMS（与 STFT 同一组帧中心），下限 energy_floor_db 保证静音帧也非零
    rms = _fit_len(librosa.feature.rms(y=y_an, frame_length=win, hop_length=hop, center=True)[0], T)
    log_e = 20 * np.log10(np.maximum(rms, 10 ** (c["energy_floor_db"] / 20)))
    zcr = _fit_len(librosa.feature.zero_crossing_rate(y_an, frame_length=win, hop_length=hop, center=True)[0], T)
    centroid = librosa.feature.spectral_centroid(S=mag, sr=sr, n_fft=n_fft, hop_length=hop)[0]
    bandwidth = librosa.feature.spectral_bandwidth(S=mag, sr=sr, n_fft=n_fft, hop_length=hop)[0]
    rolloff = librosa.feature.spectral_rolloff(S=mag, sr=sr, n_fft=n_fft, hop_length=hop,
                                               roll_percent=c["rolloff_percent"])[0]
    flatness = librosa.feature.spectral_flatness(S=mag, n_fft=n_fft, hop_length=hop)[0]

    h1 = _harmonic_amp_db(mag, freqs, f0, 1, c["h1h2_search_rel"])
    h2 = _harmonic_amp_db(mag, freqs, f0, 2, c["h1h2_search_rel"])
    h1h2 = np.where(f0 > 0, h1 - h2, 0.0).astype(np.float32)

    # 谱倾斜：对 1..N 号频点（去掉直流）拟合 dB = a + b·f(kHz)，取斜率 b
    fk = freqs[1:] / 1000.0
    ydb = 20 * np.log10(np.maximum(mag[1:], 1e-10))
    fc = fk - fk.mean()
    tilt = (fc[:, None] * (ydb - ydb.mean(axis=0, keepdims=True))).sum(axis=0) / (fc ** 2).sum()

    fmax = c["fmax_mel"] or sr / 2
    mel = librosa.feature.melspectrogram(S=power, sr=sr, n_fft=n_fft, n_mels=int(c["n_mels"]),
                                         fmin=float(c["fmin_mel"]), fmax=float(fmax))
    # power_to_db 取 top_db=None：默认的 80 dB 截断以"本片最大值"为参照，会让同一声音在不同片段里取值不同
    mfcc = librosa.feature.mfcc(S=librosa.power_to_db(mel, ref=1.0, amin=1e-10, top_db=None), n_mfcc=int(c["n_mfcc"]))
    width = int(c["delta_width"])
    if T >= 3:
        width = min(width, T if T % 2 == 1 else T - 1)
        d1 = librosa.feature.delta(mfcc, width=width, order=1)
        d2 = librosa.feature.delta(mfcc[: int(c["n_delta2"])], width=width, order=2)
    else:  # 帧数不足以做差分 → 差分置 0（行内其它维仍非零）
        d1 = np.zeros_like(mfcc)
        d2 = np.zeros_like(mfcc[: int(c["n_delta2"])])
    info["delta_width_used"] = width

    cols = [f0, vuv, vprob, log_e, zcr, _fit_len(centroid, T), _fit_len(bandwidth, T), _fit_len(rolloff, T),
            _fit_len(flatness, T), h1h2, _fit_len(tilt, T)]
    feats = np.concatenate([np.stack(cols, axis=1), _fit_len(mfcc, T).T, _fit_len(d1, T).T, _fit_len(d2, T).T],
                           axis=1).astype(np.float32)
    assert feats.shape == (T, AUDIO_DIM), feats.shape
    feats[~np.isfinite(feats)] = 0.0
    times = np.arange(T) * hop / sr
    info["voiced_ratio"] = float(vuv.mean()) if T else 0.0
    return feats, times, info


def voiced_intervals(y: np.ndarray, sr: int, top_db: float = 30.0, hop: int = 160, win: int = 400) -> np.ndarray:
    """能量 VAD（librosa.effects.split）：RMS 低于"全片最大 RMS − top_db"的帧视为静音。返回 (K,2) 秒。"""
    import librosa

    if len(y) == 0:
        return np.zeros((0, 2))
    iv = librosa.effects.split(np.asarray(y, np.float32), top_db=top_db, frame_length=win, hop_length=hop)
    return np.asarray(iv, dtype=np.float64).reshape(-1, 2) / sr
