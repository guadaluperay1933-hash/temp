"""问题1 单元测试（不依赖任何预训练模型）。

    python -m tests.test_q1          # 直接运行
    python -m pytest tests/test_q1.py  # 或用 pytest（若已安装）

覆盖：CTC Viterbi 强制对齐（手工构造的发射矩阵 + 穷举最优路径比对）、VAD 比例回退的单调性、
词元区间 → 帧池化规则、80 ms 栅格长度、blendshape → AU 布局、整词截断分词。
"""
from __future__ import annotations

import itertools

import numpy as np

from q1.align import _fill_unaligned, _proportional, align_words, ctc_viterbi, normalize_word, path_to_label_spans
from q1.sequence import build_aligned, build_unaligned, pool_interval, token_intervals
from q1.text_feats import tokenize_with_words
from q1.visual_feats import AU_INTENSITY, AU_PRESENCE, VISION_DIM, blendshapes_to_au


# ----------------------------------------------------------------------------- CTC Viterbi

def _collapse(path, blank=0):
    out, prev = [], None
    for p in path:
        if p != prev and p != blank:
            out.append(p)
        prev = p
    return out


def _brute_force(E, y, blank=0):
    """穷举所有 C^T 条帧级路径，取 collapse 后等于 y 的最大对数分。"""
    T, C = E.shape
    best, best_path = -np.inf, None
    for path in itertools.product(range(C), repeat=T):
        if _collapse(path, blank) == list(y):
            s = float(sum(E[t, c] for t, c in enumerate(path)))
            if s > best:
                best, best_path = s, path
    return best, best_path


def _logsoftmax(x):
    x = x - x.max(axis=1, keepdims=True)
    return x - np.log(np.exp(x).sum(axis=1, keepdims=True))


def test_viterbi_known_path():
    """手工构造：5 帧、目标 [1,2]，发射矩阵在路径 (b,1,1,b,2) 上概率 0.9 → DP 必须找回这条路径。"""
    target_frames = [0, 1, 1, 0, 2]
    E = np.full((5, 3), np.log(0.05))
    for t, c in enumerate(target_frames):
        E[t, c] = np.log(0.9)
    path, score = ctc_viterbi(E, [1, 2], blank=0)
    # 扩展序列 z = (b,1,b,2,b)：帧 → 状态 (0,1,1,2,3)
    assert path.tolist() == [0, 1, 1, 2, 3], path
    assert abs(score - (5 * np.log(0.9))) < 1e-9
    assert path_to_label_spans(path, 2) == [(1, 2), (4, 4)]


def test_viterbi_repeated_label_needs_blank():
    """连续相同标签 [1,1] 至少需要 3 帧（中间必须隔空白）；2 帧应判为不可行。"""
    E = _logsoftmax(np.random.default_rng(0).normal(size=(2, 3)))
    try:
        ctc_viterbi(E, [1, 1])
        raise AssertionError("应当抛出 ValueError")
    except ValueError:
        pass
    E = np.full((3, 3), np.log(0.01))
    E[0, 1] = E[1, 0] = E[2, 1] = np.log(0.98)
    path, _ = ctc_viterbi(E, [1, 1])
    assert path.tolist() == [1, 2, 3]


def test_viterbi_matches_brute_force():
    """随机发射矩阵（T≤6，C=4）上与穷举结果比对：最优分相同，且 DP 路径 collapse 后等于目标。"""
    rng = np.random.default_rng(123)
    n_cases = 0
    for T in range(2, 7):
        for y in ([1], [2, 3], [1, 1], [3, 1, 3], [2, 2, 1]):
            if T < len(y) + sum(a == b for a, b in zip(y[1:], y[:-1])):
                continue
            E = _logsoftmax(rng.normal(size=(T, 4)) * 2.0)
            bf, _ = _brute_force(E, y)
            path, score = ctc_viterbi(E, y)
            assert abs(bf - score) < 1e-9, (T, y, bf, score)
            labels = [0 if s % 2 == 0 else y[(s - 1) // 2] for s in path]
            assert _collapse(labels) == list(y), (labels, y)
            assert np.all(np.diff(path) >= 0) and np.all(np.diff(path) <= 2)
            n_cases += 1
    assert n_cases >= 20


def test_normalize_word():
    assert normalize_word("didn't,") == "DIDN'T"
    assert normalize_word("2019") == ""
    assert normalize_word("3rd") == ""
    assert normalize_word("--") == ""
    assert normalize_word("Well-known") == "WELLKNOWN"


# ----------------------------------------------------------------------------- 回退与插值

def test_proportional_monotonic_and_in_voiced():
    iv = np.array([[0.5, 1.5], [2.0, 3.0]])
    words = ["a", "bb", "ccc", "dddd"]
    spans = _proportional(words, iv, 4.0)
    starts = [s for s, _ in spans]
    ends = [e for _, e in spans]
    assert all(e >= s for s, e in spans)
    assert all(ends[k] <= starts[k + 1] + 1e-9 for k in range(len(spans) - 1))
    assert abs(starts[0] - 0.5) < 1e-9 and abs(ends[-1] - 3.0) < 1e-9
    for s, e in spans:  # 每个词整体落在某一个有声段内（不跨停顿）
        assert (0.5 - 1e-9 <= s <= e <= 1.5 + 1e-9) or (2.0 - 1e-9 <= s <= e <= 3.0 + 1e-9), (s, e)
    # 总有声时长 2 s，按 1:2:3:4 分配 → 累计 [0,.2,.6,1.2,2.0]；第一个词占 0.2 s
    assert abs((spans[0][1] - spans[0][0]) - 0.2) < 1e-9
    # 第 3 个词 [0.6,1.2] 跨越停顿：与第 1 段交 0.4、与第 2 段交 0.2 → 只保留第 1 段内的 [1.1, 1.5]
    assert np.allclose(spans[2], (1.1, 1.5)) and np.allclose(spans[3], (2.2, 3.0))


def test_fill_unaligned_interpolation():
    ws = [{"word": "a", "start": 0.1, "end": 0.3}, {"word": "1990", "start": None, "end": None},
          {"word": "b", "start": 0.5, "end": 0.7}, {"word": "7", "start": None, "end": None}]
    _fill_unaligned(ws, 1.0)
    assert (ws[1]["start"], ws[1]["end"]) == (0.3, 0.5) and ws[1]["source"] == "interp"
    assert (ws[3]["start"], ws[3]["end"]) == (0.7, 1.0)


def test_align_words_vad_fallback():
    sr = 16000
    t = np.arange(int(3.0 * sr)) / sr
    y = np.zeros_like(t)
    for a, b in ((0.3, 1.2), (1.6, 2.6)):
        m = (t >= a) & (t < b)
        y[m] = 0.3 * np.sin(2 * np.pi * 150 * t[m])
    words = "hello there 42 friend".split()
    out = align_words(y.astype(np.float32), sr, " ".join(words), method="vad_proportional")
    assert [w["word"] for w in out] == words
    assert all(0 <= w["start"] <= w["end"] <= 3.0 for w in out)
    assert all(out[k]["end"] <= out[k + 1]["start"] + 1e-6 for k in range(len(out) - 1))
    assert out[0]["start"] >= 0.25 and out[-1]["end"] <= 2.65


# ----------------------------------------------------------------------------- 对齐 → 序列

class _CharTok:
    """极简分词器：每个词切成 2 个字符一组的"词元"，用于测试整词截断与词映射。"""
    cls_token, sep_token, pad_token, pad_token_id = "[CLS]", "[SEP]", "[PAD]", 0

    def tokenize(self, w):
        w = w.lower()
        return [w[i:i + 2] if i == 0 else "##" + w[i:i + 2] for i in range(0, len(w), 2)]

    def convert_tokens_to_ids(self, toks):
        return [101 if t == "[CLS]" else 102 if t == "[SEP]" else 1000 + len(t) for t in toks]


def test_tokenize_whole_word_truncation():
    tok = tokenize_with_words("abcd ef ghijkl mn", _CharTok(), max_len=7)  # 词元数 2,1,3,1；预算 5
    assert tok["tokens"][:tok["text_len"]] == ["[CLS]", "ab", "##cd", "ef", "[SEP]"]
    assert tok["text_len"] == 5 and tok["n_words_kept"] == 2
    assert tok["truncated_words"] == 2 and tok["truncated_tokens"] == 4
    assert tok["word_index"].tolist() == [-1, 0, 0, 1, -1, -1, -1]
    assert tok["attention_mask"].tolist() == [1, 1, 1, 1, 1, 0, 0]


def _toy_streams(D=2.0, fps=10.0, face_off=(0.8, 1.2)):
    Ta = int(D * 100) + 1
    a_times = np.arange(Ta) * 0.01
    audio = {"feats": np.tile(a_times[:, None], (1, 74)).astype(np.float32) + 1.0, "times": a_times, "dur": D}
    n = int(D * fps)
    v_times = np.arange(n) / fps
    face = ~((v_times >= face_off[0]) & (v_times < face_off[1]))
    vf = np.tile(v_times[:, None], (1, 35)).astype(np.float32) + 1.0
    vf[~face] = 0
    vision = {"feats": vf, "times": v_times, "frame_idx": np.arange(n), "face": face, "dur": n / fps}
    return audio, vision


def test_pool_interval_rules():
    audio, vision = _toy_streams()
    row, rec = pool_interval(audio["feats"], audio["times"], 0.10, 0.20, audio["dur"])
    assert rec["frames"] == [10, 20] and rec["rule"] == "mean"
    assert abs(row[0] - (1.0 + np.mean(np.arange(10, 20) * 0.01))) < 1e-6
    row, rec = pool_interval(audio["feats"], audio["times"], 0.101, 0.104, audio["dur"])  # 短于帧移
    assert rec["rule"] == "nearest" and rec["frames"] == [10, 11]
    row, rec = pool_interval(vision["feats"], vision["times"], 0.85, 1.15, vision["dur"], usable=vision["face"])
    assert rec["rule"] == "mean_no_face" and not row.any()


def test_build_aligned_and_unaligned():
    audio, vision = _toy_streams()
    tok = tokenize_with_words("abcd ef", _CharTok(), max_len=8)
    words_t = [{"word": "abcd", "start": 0.2, "end": 0.6}, {"word": "ef", "start": 0.9, "end": 1.1}]
    ivs = token_intervals(tok, words_t, 2.0)
    assert ivs[0] == (0.0, 2.0, "special_whole_clip")
    assert np.allclose(ivs[1][:2], (0.2, 0.4)) and np.allclose(ivs[2][:2], (0.4, 0.6))
    al = build_aligned(tok, words_t, audio, vision, max_len=8)
    L = tok["text_len"]
    assert L == 5 and al["audio_length"] == 5
    assert np.abs(al["audio"][:L]).max(axis=1).min() > 0 and not al["audio"][L:].any()
    assert np.allclose(al["audio"][0], al["audio"][L - 1])  # [CLS] 与 [SEP] 同为整段均值
    assert not al["vision"][3].any()                          # 'ef' 落在无人脸区间 → 0 行
    assert al["positions"][3]["vision_rule"] == "mean_no_face"
    ua = build_unaligned(audio, vision, grid=0.08, max_len=500)
    assert ua["audio_length"] == 25 and ua["vision_length"] == 25  # ceil(2.0/0.08)
    assert not ua["audio"][25:].any() and np.abs(ua["audio"][:25]).max(axis=1).min() > 0
    rec = ua["audio_record"]
    assert rec["frame_start"][0] == 0 and rec["frame_end"][0] == 8 and rec["n_frames"][1] == 8
    # 视觉 10 fps 时每个 80 ms 位置不一定有帧 → 取最近帧
    assert "nearest" in ua["vision_record"]["rule"]
    ua2 = build_unaligned(audio, vision, grid=0.08, max_len=10)
    assert ua2["audio_length"] == 10 and ua2["audio_record"]["truncated_s"] > 1.1


# ----------------------------------------------------------------------------- 视觉布局

def test_blendshape_to_au_layout():
    bs = {"mouthSmileLeft": 0.9, "mouthSmileRight": 0.7, "jawOpen": 0.3, "mouthClose": 0.1,
          "mouthRollLower": 0.5, "mouthRollUpper": 0.3, "eyeBlinkLeft": 0.1, "eyeBlinkRight": 0.1}
    row = blendshapes_to_au(bs, presence_threshold=1.0, au25_gain=2.0)
    assert row.shape == (VISION_DIM,) == (35,)
    r = dict(zip(AU_INTENSITY, row[:17]))
    c = dict(zip(AU_PRESENCE, row[17:]))
    assert abs(r["AU12"] - 4.0) < 1e-6 and c["AU12"] == 1.0
    assert abs(r["AU26"] - 1.5) < 1e-6
    assert abs(r["AU25"] - 5 * min(1.0, 2.0 * 0.2)) < 1e-6   # g·(jawOpen − mouthClose) = 0.4 → 2.0
    assert c["AU28"] == 1.0 and "AU28" not in AU_INTENSITY    # 0.4·5 = 2.0 ≥ 1.0
    assert c["AU45"] == 0.0                                  # 0.1·5 = 0.5 < 1.0
    assert not blendshapes_to_au({}).any()                   # 全部分数为 0 → 整行 0（与"不可用"同义）


if __name__ == "__main__":
    n = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            n += 1
            print(f"PASS {name}")
    print(f"全部 {n} 项通过")
