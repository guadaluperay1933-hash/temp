"""构造一个"附件1 形态"的小测试集，只用于在拿不到赛题附件、不能联网下载模型时把问题1 的代码跑通。

    python -m q1.make_test_set --out outputs/q1_test

生成：
  outputs/q1_test/附件1/<video_id>/<clip_id>.mp4   4 段视频、3 个文件夹（时长 2.7 s ~ 9 s，帧率 25/30/29.97）
      画面：skimage 自带的真实人脸照片 astronaut()，逐帧轻微缩放/平移/旋转（让 blendshape 有变化），
            其中一段含 0.6 s 的"无人脸"画面（检验未检出人脸 → 0 行）；
      声音：类语音的谐波复合音（F0 110–220 Hz 起伏、按"词"做幅度包络、句读处停顿），
            采样率/声道各不相同（16 k 单声道、44.1 k 立体声…），检验重采样与混缩。
  outputs/q1_test/附件1/label-100.xlsx               video_id / clip_id / text / label / annotation
      （video_id 有纯数字的、clip_id 以数值存储，检验 Excel 类型容错；一条文本超长，检验截断；含数字与撇号）
  outputs/q1_test/附件1/_truth_word_times.json       合成音频里每个"词"的真实起止时间（只用于粗查对齐回退的合理性）
  outputs/q1_test/dummy_bert/        tokenizers 训练的小 WordPiece 词表 + 随机初始化 BertModel（hidden=768, 2 层, 4 头）
  outputs/q1_test/dummy_wav2vec2/    随机初始化的极小 Wav2Vec2ForCTC（标准 32 字符词表），只为走通 CTC 代码路径
  outputs/q1_test/附件1_bad/          （--bad-copy）健壮性测试副本：Zq-_x9/12.mp4 换成随机字节（无法解码），
                                     标注表末尾追加一行文件不存在的样本 vidA_test/7 —— 提取不应中断，核验应报出这两处

dummy 模型只因为本环境无法访问 HuggingFace 才使用；正式运行用 bert-base-uncased 与 facebook/wav2vec2-base-960h。
这里的任何数值都不能写进论文。
"""
from __future__ import annotations

import argparse
import json
import subprocess
import tempfile
from pathlib import Path

import numpy as np

from common.utils import ROOT, dump_json, set_seed

from .media import find_ffmpeg

SAMPLES = [
    # video_id, clip_id(数值), 时长 s, fps, (W,H), 音频 sr, 声道, 无人脸区间, 文本, label, annotation
    dict(video_id="vidA_test", clip_id=0, dur=2.7, fps=25.0, size=(480, 360), sr=16000, ch=1, noface=None,
         text="I really loved this movie.", label=2.0, annotation="Positive"),
    dict(video_id="vidA_test", clip_id=3, dur=5.2, fps=30.0, size=(640, 360), sr=44100, ch=2, noface=(2.0, 2.6),
         text="Well, the plot was boring and I didn't like the ending at all.", label=-1.3333333333333333,
         annotation="Negative"),
    dict(video_id=1234567, clip_id=1, dur=9.0, fps=30.0, size=(640, 480), sr=22050, ch=1, noface=None,
         text=("So in 2019 I watched it 3 times with my family, and honestly the story is about a teacher who moves to "
               "a small town, meets the neighbors, starts a garden, argues with the mayor about the old library, "
               "and eventually everyone learns something about patience, kindness, community, friendship and "
               "forgiveness, which is fine I suppose."),
         label=0.0, annotation="Neutral"),
    dict(video_id="Zq-_x9", clip_id=12, dur=4.0, fps=30000 / 1001, size=(480, 270), sr=48000, ch=2, noface=None,
         text="It's okay, I guess. Not great, not terrible.", label=0.3333333333333333, annotation="Positive"),
]

W2V_VOCAB = ["<pad>", "<s>", "</s>", "<unk>", "|", "E", "T", "A", "O", "N", "I", "H", "S", "R", "D", "L", "U", "M",
             "W", "C", "F", "G", "Y", "P", "B", "V", "K", "'", "X", "J", "Q", "Z"]


# ----------------------------------------------------------------------------- 音频

def synth_speech(text: str, dur: float, sr: int, rng: np.random.Generator) -> tuple[np.ndarray, list[dict]]:
    """类语音信号：每个词是一段谐波复合音（汉宁包络），词间 50 ms 间隙，逗号/句号后 0.35 s 停顿。

    词长按字符数比例分配，首尾各留 0.25 s 静音。返回 (波形, 词真值时间)。
    """
    words = text.split()
    n = int(round(dur * sr))
    y = np.zeros(n, np.float64)
    lead, tail, gap, pause = 0.25, 0.25, 0.05, 0.35
    n_pause = sum(1 for w in words[:-1] if w[-1] in ",.;!?")
    speech = dur - lead - tail - gap * (len(words) - 1) - pause * n_pause
    if speech < 0.15 * len(words):  # 词太多：缩短停顿
        pause, gap = 0.1, 0.02
        speech = dur - lead - tail - gap * (len(words) - 1) - pause * n_pause
    w_len = np.array([max(2, len(w)) for w in words], float)
    w_dur = w_len / w_len.sum() * speech
    t = lead
    truth = []
    for k, (w, d) in enumerate(zip(words, w_dur)):
        s0, s1 = int(t * sr), int((t + d) * sr)
        tt = np.arange(s1 - s0) / sr
        f0 = 110 + 110 * (0.5 + 0.5 * np.sin(2 * np.pi * (0.7 + 0.3 * rng.random()) * (t + tt)))
        phase = 2 * np.pi * np.cumsum(f0) / sr
        sig = sum((1.0 / h) * np.sin(h * phase) for h in range(1, 16))
        env = np.hanning(len(tt)) * (0.6 + 0.4 * np.sin(2 * np.pi * 4.0 * tt) ** 2)  # 约 4 Hz 音节起伏
        y[s0:s1] += 0.25 * sig * env
        truth.append({"word": w, "start": round(t, 4), "end": round(t + d, 4)})
        t += d + gap + (pause if (w[-1] in ",.;!?" and k < len(words) - 1) else 0.0)
    y += 0.002 * rng.standard_normal(n)
    return y.astype(np.float32), truth


# ----------------------------------------------------------------------------- 画面

def _face_image():
    from skimage import data

    return np.ascontiguousarray(data.astronaut())  # (512,512,3) RGB，NASA 公有领域照片


def render_frames(face: np.ndarray, dur: float, fps: float, size: tuple[int, int], noface, rng):
    """逐帧：人脸照片缩放到画面高度，按时间做 ±5% 缩放、±10px 平移、±3° 旋转；noface 区间内画纯灰底（无人脸）。"""
    import cv2

    W, H = size
    n = int(round(dur * fps))
    side = int(H * 0.95)
    base = cv2.resize(face, (side, side), interpolation=cv2.INTER_AREA)
    bg = np.full((H, W, 3), 90, np.uint8)
    bg[:, :, 2] = 110
    ph = rng.random(3) * 2 * np.pi
    for i in range(n):
        t = i / fps
        if noface and noface[0] <= t < noface[1]:
            fr = bg.copy()
            cv2.putText(fr, "no face", (W // 3, H // 2), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (200, 200, 200), 2)
            yield fr
            continue
        z = 1.0 + 0.05 * np.sin(2 * np.pi * 0.5 * t + ph[0])
        ang = 3.0 * np.sin(2 * np.pi * 0.35 * t + ph[1])
        dx = 10 * np.sin(2 * np.pi * 0.3 * t + ph[2])
        dy = 6 * np.cos(2 * np.pi * 0.4 * t)
        M = cv2.getRotationMatrix2D((side / 2, side / 2), ang, z)
        M[:, 2] += (dx, dy)
        warped = cv2.warpAffine(base, M, (side, side), borderMode=cv2.BORDER_REFLECT)
        fr = bg.copy()
        x0, y0 = (W - side) // 2, (H - side) // 2
        fr[y0:y0 + side, x0:x0 + side] = warped
        yield fr


def write_video(path: Path, frames, fps: float, size, audio: np.ndarray, sr: int, ch: int) -> None:
    """原始 RGB 帧经管道送入 ffmpeg（libx264 / yuv420p），与 wav 复用为 mp4（AAC 音轨）。"""
    import soundfile as sf

    path.parent.mkdir(parents=True, exist_ok=True)
    W, H = size
    with tempfile.TemporaryDirectory() as td:
        wav = Path(td) / "a.wav"
        a = audio if ch == 1 else np.stack([audio, 0.9 * audio], axis=1)
        sf.write(wav, a, sr, subtype="PCM_16")
        fps_arg = "30000/1001" if abs(fps - 30000 / 1001) < 1e-6 else f"{fps:g}"
        cmd = [find_ffmpeg(), "-y", "-hide_banner", "-loglevel", "error",
               "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}", "-r", fps_arg, "-i", "pipe:0",
               "-i", str(wav), "-c:v", "libx264", "-pix_fmt", "yuv420p", "-preset", "fast", "-crf", "20",
               "-c:a", "aac", "-b:a", "96k", "-shortest", str(path)]
        proc = subprocess.Popen(cmd, stdin=subprocess.PIPE)
        for fr in frames:
            proc.stdin.write(np.ascontiguousarray(fr).tobytes())
        proc.stdin.close()
        if proc.wait() != 0:
            raise RuntimeError(f"ffmpeg 写视频失败：{path}")


def write_label_xlsx(path: Path, rows: list[dict]) -> None:
    """用 openpyxl 直接写，保留数值类型（纯数字 video_id、数值 clip_id），模拟真实表格的类型混杂。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["video_id", "clip_id", "text", "label", "annotation"])
    for r in rows:
        ws.append([r["video_id"], r["clip_id"], r["text"], r["label"], r["annotation"]])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ----------------------------------------------------------------------------- dummy 模型

def make_dummy_bert(out: Path, texts: list[str], seed: int = 0) -> None:
    import torch
    from tokenizers import BertWordPieceTokenizer
    from transformers import BertConfig, BertModel, BertTokenizerFast

    out.mkdir(parents=True, exist_ok=True)
    wp = BertWordPieceTokenizer(lowercase=True)
    wp.train_from_iterator(texts, vocab_size=160, min_frequency=1,
                           special_tokens=["[PAD]", "[UNK]", "[CLS]", "[SEP]", "[MASK]"])
    wp.save_model(str(out))
    vocab_lines = (out / "vocab.txt").read_text(encoding="utf-8").splitlines()
    vocab = {t: i for i, t in enumerate(vocab_lines)}
    try:  # transformers ≥ 5：BertTokenizer(vocab=dict)
        tok = BertTokenizerFast(vocab=vocab, do_lower_case=True)
    except TypeError:  # transformers 4.x
        tok = BertTokenizerFast(vocab_file=str(out / "vocab.txt"), do_lower_case=True)
    if len(tok) != len(vocab):
        raise RuntimeError(f"dummy 分词器词表大小 {len(tok)} ≠ {len(vocab)}")
    tok.save_pretrained(str(out))
    torch.manual_seed(seed)
    cfg = BertConfig(vocab_size=len(tok), hidden_size=768, num_hidden_layers=2, num_attention_heads=4,
                     intermediate_size=1024, max_position_embeddings=128)
    BertModel(cfg).eval().save_pretrained(str(out))


def make_dummy_wav2vec2(out: Path, seed: int = 0) -> None:
    import torch
    from transformers import (Wav2Vec2Config, Wav2Vec2CTCTokenizer, Wav2Vec2FeatureExtractor, Wav2Vec2ForCTC,
                              Wav2Vec2Processor)

    out.mkdir(parents=True, exist_ok=True)
    vocab_file = out / "vocab.json"
    vocab_file.write_text(json.dumps({t: i for i, t in enumerate(W2V_VOCAB)}), encoding="utf-8")
    tok = Wav2Vec2CTCTokenizer(str(vocab_file), unk_token="<unk>", pad_token="<pad>", word_delimiter_token="|")
    fe = Wav2Vec2FeatureExtractor(feature_size=1, sampling_rate=16000, padding_value=0.0, do_normalize=True,
                                  return_attention_mask=False)
    Wav2Vec2Processor(feature_extractor=fe, tokenizer=tok).save_pretrained(str(out))
    torch.manual_seed(seed)
    cfg = Wav2Vec2Config(vocab_size=len(W2V_VOCAB), hidden_size=32, num_hidden_layers=2, num_attention_heads=2,
                         intermediate_size=64, conv_dim=(16,) * 7, num_conv_pos_embeddings=16,
                         num_conv_pos_embedding_groups=2, pad_token_id=0, feat_extract_norm="group",
                         do_stable_layer_norm=False)
    Wav2Vec2ForCTC(cfg).eval().save_pretrained(str(out))


# ----------------------------------------------------------------------------- 主程序

def main(argv=None):
    ap = argparse.ArgumentParser(description="构造问题1 小测试集（合成，仅供跑通代码）")
    ap.add_argument("--out", default="outputs/q1_test")
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--no-models", action="store_true", help="不生成 dummy BERT / wav2vec2")
    ap.add_argument("--bad-copy", action="store_true", help="另生成含损坏文件/缺失文件的健壮性测试副本 附件1_bad")
    a = ap.parse_args(argv)
    out = Path(a.out)
    if not out.is_absolute():
        out = ROOT / out
    set_seed(a.seed)
    rng = np.random.default_rng(a.seed)
    face = _face_image()
    att = out / "附件1"
    truth = {}
    for s in SAMPLES:
        audio, wt = synth_speech(s["text"], s["dur"], s["sr"], rng)
        path = att / str(s["video_id"]) / f"{s['clip_id']}.mp4"
        write_video(path, render_frames(face, s["dur"], s["fps"], s["size"], s["noface"], rng), s["fps"], s["size"],
                    audio, s["sr"], s["ch"])
        truth[f"{s['video_id']}$_${s['clip_id']}"] = {"words": wt, "noface": s["noface"], "dur": s["dur"],
                                                      "fps": s["fps"]}
        print(f"写出 {path.relative_to(out)}  {s['dur']}s  {s['fps']:.3f}fps  {s['size']}  sr={s['sr']} ch={s['ch']}")
    write_label_xlsx(att / "label-100.xlsx", SAMPLES)
    dump_json(truth, att / "_truth_word_times.json")
    if not a.no_models:
        make_dummy_bert(out / "dummy_bert", [s["text"] for s in SAMPLES], a.seed)
        make_dummy_wav2vec2(out / "dummy_wav2vec2", a.seed)
        print(f"dummy 模型：{out / 'dummy_bert'}，{out / 'dummy_wav2vec2'}")
    if a.bad_copy:
        make_bad_copy(att, out / "附件1_bad", a.seed)
        print(f"健壮性测试副本：{out / '附件1_bad'}")
    print(f"完成：{att}")


def make_bad_copy(src: Path, dst: Path, seed: int = 0) -> None:
    import shutil

    from openpyxl import load_workbook

    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)
    (dst / "Zq-_x9" / "12.mp4").write_bytes(np.random.default_rng(seed).bytes(20000))  # 无法解码
    wb = load_workbook(dst / "label-100.xlsx")
    wb.active.append(["vidA_test", 7, "this clip file does not exist", -2.0, "Negative"])  # 文件缺失
    wb.save(dst / "label-100.xlsx")


if __name__ == "__main__":
    main()
