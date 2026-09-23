"""问题1 主流程：附件1（原始视频 + label-100.xlsx）→ 与附件2 同构的对齐 / 非对齐特征文件 + 逐样本对齐记录 + 清单 + 日志。

    python -m q1.extract --config configs/q1.yaml [--run-name default] [--limit N] [--resume] [--set key=value ...]

处理流程（每条样本，按标注表行序）：
    1 定位文件 <video_root>/<video_id>/<clip_id>.mp4，计算 MD5，ffmpeg/OpenCV 探测时长、帧率、分辨率、音轨
    2 音频：ffmpeg 解码为 16 kHz 单声道 → 100 Hz × 74 维声学特征（audio_feats）+ 能量 VAD
    3 视觉：逐帧 MediaPipe FaceLandmarker → 35 维 AU 特征（visual_feats），未检出人脸帧为 0 行
    4 文本：转写文本按空白切词 → BERT 词元（≤50，整词截断）→ 最后一层 768 维（text_feats）
    5 词级时间戳：给定文本对音频做 CTC 强制对齐（align，模型不可用时回退为 VAD 比例分配）
    6 对齐：词元区间 → 对齐版 50 位置；80 ms 栅格 → 非对齐版（sequence）
    7 写逐位置对齐记录 alignment/<id>.json
全部样本处理完后写 features/*.pkl(.gz)、manifest.csv/.xlsx、manifest_long.csv、run_info.json。
任何一条样本的任一模态失败都不会中断运行：该模态记 0 行、长度 0，status=error，保证"标注行 ↔ 视频 ↔ 特征"一一对应。
"""
from __future__ import annotations

import argparse
import gzip
import pickle
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import CLASS_NAMES, ID_SEP, MOD_CN, reg_to_cls
from common.utils import ROOT, dump_json, env_info, get_logger, set_seed

from . import media
from .align import _proportional, align_words, ctc_failure_reason, get_ctc_aligner
from .audio_feats import AUDIO_DIM, AUDIO_FEATURE_NAMES, extract_audio_features, frame_params, voiced_intervals
from .config import config_hash, load_config, model_ref, resolve
from .sequence import build_aligned, build_unaligned
from .text_feats import TextEncoder, text_bert_array
from .visual_feats import VISION_DIM, VISION_FEATURE_NAMES, FaceAUExtractor

TEXT_DIM = 768
VIDEO_EXTS = (".mp4", ".MP4", ".mkv", ".avi", ".mov", ".webm")


# ----------------------------------------------------------------------------- 标注表与文件定位

def norm_id(v) -> str:
    """Excel 里的 video_id / clip_id 可能被读成 int、float（3.0）或字符串 → 统一成字符串。"""
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return ""
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        return str(int(v)) if float(v).is_integer() else repr(float(v))
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
        s = s[:-2]
    return s


def _loose(s: str) -> str:
    """宽松比较键：小写；纯数字去前导零。"""
    s = s.strip().lower()
    return s.lstrip("0") or "0" if s.isdigit() else s


def read_label_table(path: Path, sheet=0) -> pd.DataFrame:
    """读取 label-100.xlsx（列名大小写/空格不敏感）。只读不改：标签、文本原样保留。"""
    if path.suffix.lower() in (".csv", ".txt"):
        df = pd.read_csv(path, dtype=object)
    else:
        df = pd.read_excel(path, sheet_name=sheet, dtype=object)
    cols = {c: str(c).strip().lower() for c in df.columns}
    df = df.rename(columns=cols)
    need = ["video_id", "clip_id"]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"标注表缺少列 {miss}，实际列：{list(df.columns)}")
    for c in ("text", "label", "annotation"):
        if c not in df.columns:
            df[c] = None
    df = df.reset_index(drop=True)
    df["row"] = np.arange(len(df)) + 2  # Excel 行号（第 1 行为表头）
    return df


class FileResolver:
    """<video_root>/<video_id>/<clip_id>.mp4 定位；精确匹配失败时做宽松匹配（大小写、数字前导零、扩展名）。"""

    def __init__(self, root: Path):
        self.root = root
        self.dirs = {d.name: d for d in sorted(root.iterdir()) if d.is_dir()} if root.exists() else {}
        self.loose_dirs = {}
        for name, d in self.dirs.items():
            self.loose_dirs.setdefault(_loose(name), d)

    def all_videos(self) -> list[Path]:
        out = []
        for d in self.dirs.values():
            out.extend(sorted(p for p in d.iterdir() if p.is_file() and p.suffix in VIDEO_EXTS))
        return out

    def find(self, vid: str, cid: str) -> tuple[Path | None, str]:
        d = self.dirs.get(vid) or self.loose_dirs.get(_loose(vid))
        if d is None:
            return None, f"找不到文件夹 {vid}"
        for ext in VIDEO_EXTS:
            p = d / f"{cid}{ext}"
            if p.exists():
                return p, "exact" if d.name == vid and ext == ".mp4" else "loose"
        for p in sorted(d.iterdir()):
            if p.is_file() and p.suffix in VIDEO_EXTS and _loose(p.stem) == _loose(cid):
                return p, "loose"
        return None, f"文件夹 {d.name} 中找不到 {cid}.mp4"


# ----------------------------------------------------------------------------- 单条样本

def _speech_span(words_t: list[dict]) -> float:
    if not words_t:
        return 0.0
    return float(max(0.0, words_t[-1]["end"] - words_t[0]["start"]))


def process_sample(row: dict, path: Path | None, enc: TextEncoder | None, face: FaceAUExtractor | None,
                   cfg: dict, logger) -> dict:
    """返回一条样本的全部结果（特征、记录、清单行）。各模态独立 try/except。"""
    t0 = time.time()
    seq = cfg["sequence"]
    Lt, Lav = int(seq["max_len_text"]), int(seq["max_len_av"])
    sid = row["sample_id"]
    errors, warns = [], []
    res = {
        "text": np.zeros((Lt, TEXT_DIM), np.float32), "text_bert": np.zeros((3, Lt), np.int64),
        "a_audio": np.zeros((Lt, AUDIO_DIM), np.float32), "a_vision": np.zeros((Lt, VISION_DIM), np.float32),
        "u_audio": np.zeros((Lav, AUDIO_DIM), np.float32), "u_vision": np.zeros((Lav, VISION_DIM), np.float32),
        "a_audio_len": 0, "a_vision_len": 0, "u_audio_len": 0, "u_vision_len": 0, "text_len": 0,
    }
    man = {"sample_id": sid, "video_id": row["video_id"], "clip_id": row["clip_id"],
           "file": str(path.relative_to(resolve(cfg["paths"]["video_root"]))) if path else None}
    record = {"id": sid, "video_id": row["video_id"], "clip_id": row["clip_id"], "file": man["file"],
              "excel_row": row["row"]}

    if path is None:
        errors.append(row.get("resolve_msg", "文件缺失"))
        return _finish(res, man, record, row, errors, warns, t0, None, None, None, None, None, cfg)

    # ---- 1 探测
    probe = None
    try:
        probe = media.probe_video(path)
        man["md5"] = media.file_md5(path)
    except Exception as e:  # noqa: BLE001
        errors.append(f"probe: {e}")

    # ---- 2 音频
    audio = None
    a_info = {}
    voiced = np.zeros((0, 2))
    wave = np.zeros(0, np.float32)
    try:
        sr = int(cfg["audio"]["sr"])
        wave = media.load_audio(path, sr)
        if len(wave) == 0:
            warns.append("无音轨：语音模态整段为 0 行")
        else:
            feats, times, a_info = extract_audio_features(wave, cfg["audio"])
            audio = {"feats": feats, "times": times, "dur": len(wave) / sr}
            _, win, hop = frame_params(cfg["audio"])
            voiced = voiced_intervals(wave, sr, top_db=float(cfg["audio"]["vad_top_db"]), hop=hop, win=win)
    except Exception as e:  # noqa: BLE001
        errors.append(f"audio: {type(e).__name__}: {e}")
        logger.debug(traceback.format_exc())
        audio = None

    # ---- 3 视觉
    vision = None
    v_raw = None
    if face is not None:
        try:
            v_raw = face.process_video(path, backend=cfg["media"].get("video_backend", "cv2"))
            fps = v_raw["fps"]
            if cfg["vision"].get("sample_fps"):
                n_dec = max(v_raw["n_decoded"], (probe or {}).get("n_frames") or 0)
            else:
                n_dec = v_raw["n_decoded"]
            vdur = n_dec / fps if fps else 0.0
            vision = {"feats": v_raw["feats"], "times": v_raw["times"], "frame_idx": v_raw["frame_idx"],
                      "face": v_raw["face"], "dur": vdur}
            if v_raw["n_processed"] == 0:
                errors.append("vision: 没有解码出任何帧")
                vision = None
        except Exception as e:  # noqa: BLE001
            errors.append(f"vision: {type(e).__name__}: {e}")
            logger.debug(traceback.format_exc())
    else:
        errors.append("vision: 视觉提取器不可用")

    # ---- 4 文本
    tok = None
    raw_text = row.get("text")
    raw_text = "" if raw_text is None or (isinstance(raw_text, float) and not np.isfinite(raw_text)) else str(raw_text)
    try:
        tok = enc.tokenize(raw_text)
        res["text"] = enc.encode([tok], batch_size=1)[0]
        res["text_bert"] = text_bert_array(tok)
        res["text_len"] = int(tok["text_len"])
        if tok["truncated_words"] > 0:
            warns.append(f"文本超长：截断 {tok['truncated_words']} 个词 / {tok['truncated_tokens']} 个词元")
        if tok["n_words_total"] == 0:
            warns.append("转写文本为空：只有 [CLS][SEP]")
    except Exception as e:  # noqa: BLE001
        errors.append(f"text: {type(e).__name__}: {e}")
        logger.debug(traceback.format_exc())
        tok = None

    # ---- 5 词级时间戳
    words_t, al_info = [], {"method_used": None}
    words = tok["words"] if tok is not None else raw_text.split()
    try:
        if audio is not None:
            words_t, al_info = align_words(wave, int(cfg["audio"]["sr"]), words, method=cfg["align"]["method"],
                                           device=cfg["align"].get("device"),
                                           model_name=model_ref(cfg["align"]["ctc_model"]),
                                           vad_top_db=float(cfg["align"]["vad_top_db"]),
                                           chunk_s=float(cfg["align"].get("chunk_s", 60)), return_info=True)
        elif vision is not None and words:  # 无音轨：按字符数比例铺满整段视频
            spans = _proportional(words, np.zeros((0, 2)), vision["dur"])
            words_t = [{"index": k, "word": w, "start": round(s, 4), "end": round(e, 4), "source": "no_audio_uniform",
                        "score": None} for k, (w, (s, e)) in enumerate(zip(words, spans))]
            al_info = {"method_used": "no_audio_uniform"}
        if al_info.get("fallback_reason") and cfg["align"]["method"] == "auto" and \
                not al_info["fallback_reason"].startswith("model_load_failed"):
            warns.append(f"CTC 对齐失败已回退：{al_info['fallback_reason'][:120]}")
    except Exception as e:  # noqa: BLE001
        errors.append(f"align: {type(e).__name__}: {e}")
        logger.debug(traceback.format_exc())
        words_t = []

    # ---- 6 对齐与序列化
    try:
        if tok is not None and (words_t or not words):
            al = build_aligned(tok, words_t, audio, vision, Lt)
            res["a_audio"], res["a_vision"] = al["audio"], al["vision"]
            res["a_audio_len"], res["a_vision_len"] = al["audio_length"], al["vision_length"]
            record["aligned"] = {"length": al["length"], "audio_length": al["audio_length"],
                                 "vision_length": al["vision_length"], "positions": al["positions"],
                                 "pad_rule": f"位置 ≥ {al['length']} 为填充：文本/语音/视觉整行为 0，text_bert 掩码为 0"}
        elif tok is not None:
            errors.append("aligned: 词级时间戳缺失，对齐版语音/视觉置 0")
        ua = build_unaligned(audio, vision, float(seq["grid_s"]), Lav)
        res["u_audio"], res["u_vision"] = ua["audio"], ua["vision"]
        res["u_audio_len"], res["u_vision_len"] = ua["audio_length"], ua["vision_length"]
        for m in ("audio", "vision"):
            if ua[f"{m}_record"].get("truncated_s", 0) > 0:
                warns.append(f"非对齐 {m} 超过 {Lav} 个位置，截断 {ua[f'{m}_record']['truncated_s']:.2f}s")
        record["unaligned"] = {"grid_s": float(seq["grid_s"]), "text_length": res["text_len"],
                               "audio": ua["audio_record"], "vision": ua["vision_record"],
                               "pad_rule": "语音/视觉位置 ≥ 各自 length 为填充（整行 0）；文本同对齐版本"}
    except Exception as e:  # noqa: BLE001
        errors.append(f"sequence: {type(e).__name__}: {e}")
        logger.debug(traceback.format_exc())

    return _finish(res, man, record, row, errors, warns, t0, probe, audio, vision, v_raw, (tok, words_t, al_info,
                                                                                          voiced, a_info), cfg)


def _finish(res, man, record, row, errors, warns, t0, probe, audio, vision, v_raw, extra, cfg):
    tok, words_t, al_info, voiced, a_info = extra if extra else (None, [], {}, np.zeros((0, 2)), {})
    probe = probe or {}
    label = row.get("label")
    try:
        label_f = float(label)
    except (TypeError, ValueError):
        label_f = float("nan")
    ann = row.get("annotation")
    ann = "" if ann is None or (isinstance(ann, float) and not np.isfinite(ann)) else str(ann)
    cls_from_label = CLASS_NAMES[int(reg_to_cls([label_f])[0])] if np.isfinite(label_f) else ""
    ann_ok = (ann.strip().lower() == cls_from_label.lower()) if ann and cls_from_label else None
    if ann_ok is False:
        warns.append(f"annotation={ann} 与 label={label_f} 按题面规则推出的 {cls_from_label} 不一致（原样保留）")
    face_rate = v_raw["detect_rate"] if v_raw else None
    if face_rate is not None and face_rate < 0.5:
        warns.append(f"人脸检出率偏低 {face_rate:.2f}")
    if v_raw and v_raw.get("pass") == "center_square":
        warns.append(f"整帧人脸检出率 {v_raw['detect_rate_full']:.2f} 偏低，已改用居中正方形裁剪（检出率 "
                     f"{v_raw['detect_rate']:.2f}）")
    a_dur = audio["dur"] if audio else 0.0
    v_dur = vision["dur"] if vision else 0.0
    if audio and vision and abs(a_dur - v_dur) > 0.2:
        warns.append(f"音视频时长相差 {abs(a_dur - v_dur):.2f}s")
    voiced_s = float((voiced[:, 1] - voiced[:, 0]).sum()) if len(voiced) else 0.0
    face_frames = int(v_raw["face"].sum()) if v_raw else 0
    proc_fps = (v_raw["n_processed"] / v_dur) if (v_raw and v_dur > 0) else None
    face_visible_s = face_frames / proc_fps if proc_fps else 0.0
    n_interp = sum(1 for w in words_t if w.get("source") == "interp")
    man.update({
        "md5": man.get("md5"), "duration_s": probe.get("duration"), "audio_duration_s": round(a_dur, 4),
        "video_duration_s": round(v_dur, 4), "fps": probe.get("fps"), "n_frames": probe.get("n_frames"),
        "n_frames_decoded": v_raw["n_decoded"] if v_raw else None,
        "n_frames_processed": v_raw["n_processed"] if v_raw else None,
        "resolution": f"{probe.get('width')}x{probe.get('height')}" if probe.get("width") else None,
        "audio_sr": probe.get("audio_sr"), "audio_channels": probe.get("audio_channels"),
        "n_words": tok["n_words_total"] if tok else None,
        "n_tokens_total": (int(sum(tok["pieces_per_word"])) + 2) if tok else None,
        "text_len": res["text_len"], "truncated_words": tok["truncated_words"] if tok else None,
        "truncated_tokens": tok["truncated_tokens"] if tok else None,
        "speech_span_s": round(_speech_span(words_t), 4), "voiced_duration_s": round(voiced_s, 4),
        "audio_frames_100hz": int(audio["feats"].shape[0]) if audio else 0,
        "audio_len_aligned": res["a_audio_len"], "vision_len_aligned": res["a_vision_len"],
        "audio_len_unaligned": res["u_audio_len"], "vision_len_unaligned": res["u_vision_len"],
        "face_detect_rate": round(face_rate, 4) if face_rate is not None else None,
        "face_detect_rate_full": v_raw.get("detect_rate_full") if v_raw else None,
        "vision_pass": v_raw.get("pass") if v_raw else None,
        "face_visible_s": round(face_visible_s, 4),
        "aligner_used": al_info.get("method_used"), "n_words_interp": n_interp,
        "label": label_f, "annotation": ann, "cls_from_label": cls_from_label, "annotation_consistent": ann_ok,
        "status": "error" if errors else ("warning" if warns else "ok"),
        "error": " | ".join(errors), "warnings": " | ".join(warns),
        "proc_time_s": round(time.time() - t0, 2),
    })
    record.update({
        "durations_s": {"container": probe.get("duration"), "audio": a_dur, "video": v_dur},
        "fps": probe.get("fps"), "n_video_frames_decoded": man["n_frames_decoded"],
        "audio_frame": {"hop_s": frame_params(cfg["audio"])[2] / frame_params(cfg["audio"])[0],
                        "win_s": frame_params(cfg["audio"])[1] / frame_params(cfg["audio"])[0],
                        "time_of_frame_k": "k * hop_s（帧中心）", "n_frames": man["audio_frames_100hz"]},
        "video_frame": {"time_of_frame_n": "n / fps", "n_processed": man["n_frames_processed"],
                        "n_face": face_frames, "detect_rate": man["face_detect_rate"],
                        "pass": man["vision_pass"], "detect_rate_full": man["face_detect_rate_full"],
                        "detect_rate_crop": v_raw.get("detect_rate_crop") if v_raw else None},
        "voiced_intervals_s": [[round(a, 3), round(b, 3)] for a, b in voiced.tolist()],
        "aligner": al_info, "words": words_t,
        "text": {"raw_text": row.get("text") if isinstance(row.get("text"), str) else "",
                 "text_len": res["text_len"], "n_words_total": man["n_words"],
                 "n_words_kept": tok["n_words_kept"] if tok else None,
                 "truncated_words": man["truncated_words"], "truncated_tokens": man["truncated_tokens"],
                 "tokens": tok["tokens"][: res["text_len"]] if tok else [],
                 "pieces_per_word": tok["pieces_per_word"] if tok else []},
        "lengths": {"text": res["text_len"], "aligned_audio": res["a_audio_len"],
                    "aligned_vision": res["a_vision_len"], "unaligned_audio": res["u_audio_len"],
                    "unaligned_vision": res["u_vision_len"]},
        "label": label_f, "annotation": ann, "status": man["status"], "errors": errors, "warnings": warns,
    })
    return {"res": res, "manifest": man, "record": record, "raw_text": row.get("text")}


# ----------------------------------------------------------------------------- 汇总写出

def _long_rows(man: dict, cfg: dict) -> list[dict]:
    grid_ms = int(round(float(cfg["sequence"]["grid_s"]) * 1000))
    fps = man.get("fps")
    rows = []
    base = {"sample_id": man["sample_id"], "video_id": man["video_id"], "clip_id": man["clip_id"],
            "status": man["status"]}
    rows.append({**base, "modality": MOD_CN["text"], "raw_source": "label-100.xlsx 的 text 列（英文转写）",
                 "raw_duration_s": man.get("audio_duration_s"),
                 "original_valid_duration_s": man.get("speech_span_s"),
                 "valid_duration_definition": "首词起点→末词终点（词级强制对齐）",
                 "feature_dim": TEXT_DIM, "feature_tool": "BERT 最后一层",
                 "aligned_len": man.get("text_len"), "unaligned_len": man.get("text_len"),
                 "alignment_granularity_aligned": "WordPiece 词元（区间=所属词的时间区间均分）",
                 "alignment_granularity_unaligned": "WordPiece 词元（同对齐版）",
                 "raw_rate": "词序列", "n_raw_units": man.get("n_words")})
    rows.append({**base, "modality": MOD_CN["audio"], "raw_source": "mp4 音轨 → 16 kHz 单声道",
                 "raw_duration_s": man.get("audio_duration_s"),
                 "original_valid_duration_s": man.get("voiced_duration_s"),
                 "valid_duration_definition": f"能量 VAD 有声段总长（top_db={cfg['audio']['vad_top_db']}）",
                 "feature_dim": AUDIO_DIM, "feature_tool": "librosa 帧级声学特征",
                 "aligned_len": man.get("audio_len_aligned"), "unaligned_len": man.get("audio_len_unaligned"),
                 "alignment_granularity_aligned": "词元（词级时间区间内帧均值）",
                 "alignment_granularity_unaligned": f"{grid_ms} ms 栅格",
                 "raw_rate": f"{1000 / float(cfg['audio']['hop_ms']):.0f} Hz（{cfg['audio']['hop_ms']} ms 帧移）",
                 "n_raw_units": man.get("audio_frames_100hz")})
    rows.append({**base, "modality": MOD_CN["vision"], "raw_source": "mp4 视频流逐帧",
                 "raw_duration_s": man.get("video_duration_s"),
                 "original_valid_duration_s": man.get("face_visible_s"),
                 "valid_duration_definition": "检出人脸的帧数 / 处理帧率",
                 "feature_dim": VISION_DIM, "feature_tool": "MediaPipe FaceLandmarker blendshape→AU",
                 "aligned_len": man.get("vision_len_aligned"), "unaligned_len": man.get("vision_len_unaligned"),
                 "alignment_granularity_aligned": "词元（词级时间区间内人脸帧均值）",
                 "alignment_granularity_unaligned": f"{grid_ms} ms 栅格",
                 "raw_rate": f"{fps:.3f} fps" if fps else None, "n_raw_units": man.get("n_frames_processed")})
    return rows


def _save_pkl(obj, path: Path, compress: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    if compress == "gzip":
        path = path.with_suffix(path.suffix + ".gz")
        # mtime=0：gzip 头里不写入当前时间 → 相同输入两次运行得到逐字节相同的文件（MD5 可复核）
        with open(path, "wb") as raw, gzip.GzipFile(filename="", mode="wb", fileobj=raw, compresslevel=9,
                                                    mtime=0) as f:
            pickle.dump(obj, f, protocol=4)
    else:
        with open(path, "wb") as f:
            pickle.dump(obj, f, protocol=4)
    return path


def assemble(results: list[dict], cfg: dict) -> tuple[dict, dict]:
    dt = np.float16 if str(cfg["output"]["dtype"]) == "float16" else np.float32
    ids = np.array([r["manifest"]["sample_id"] for r in results], dtype=object)
    raw_text = np.array(["" if not isinstance(r["raw_text"], str) else r["raw_text"] for r in results], dtype=object)
    y = np.array([r["manifest"]["label"] for r in results], dtype=np.float32)
    ann = np.array([r["manifest"]["annotation"] for r in results], dtype=object)
    cls = reg_to_cls(np.nan_to_num(y)).astype(np.float32)
    text = np.stack([r["res"]["text"] for r in results]).astype(dt)
    tb = np.stack([r["res"]["text_bert"] for r in results]).astype(np.int64)
    common_fields = {"id": ids, "raw_text": raw_text, "text": text, "text_bert": tb, "regression_labels": y,
                     "annotations": ann, "classification_labels": cls}
    aligned = {**common_fields,
               "audio": np.stack([r["res"]["a_audio"] for r in results]).astype(dt),
               "vision": np.stack([r["res"]["a_vision"] for r in results]).astype(dt),
               "audio_lengths": np.array([r["res"]["a_audio_len"] for r in results], np.int64),
               "vision_lengths": np.array([r["res"]["a_vision_len"] for r in results], np.int64)}
    unaligned = {**common_fields,
                 "audio": np.stack([r["res"]["u_audio"] for r in results]).astype(dt),
                 "vision": np.stack([r["res"]["u_vision"] for r in results]).astype(dt),
                 "audio_lengths": np.array([r["res"]["u_audio_len"] for r in results], np.int64),
                 "vision_lengths": np.array([r["res"]["u_vision_len"] for r in results], np.int64)}
    return {"all": aligned}, {"all": unaligned}


def feature_spec(cfg: dict, n: int) -> dict:
    Lt, Lav = int(cfg["sequence"]["max_len_text"]), int(cfg["sequence"]["max_len_av"])
    return {
        "top_level_key": "all",
        "read_example": "from common.data import load_all_splits; sd = load_all_splits('aligned_100.pkl.gz')['all']",
        "dtype": str(cfg["output"]["dtype"]), "compression": cfg["output"]["compress"], "n_samples": n,
        "fields": {
            "id": "object (N,)  'video_id$_$clip_id'，顺序 = label-100.xlsx 行序",
            "raw_text": "object (N,)  原始英文转写（未改动）",
            "text": f"(N, {Lt}, 768)  BERT 最后一层；位置 i ↔ 词元 i；填充位置为 0",
            "text_bert": f"int64 (N, 3, {Lt})  [input_ids; attention_mask; token_type_ids]",
            "audio": f"aligned (N, {Lt}, 74) / unaligned (N, {Lav}, 74)",
            "vision": f"aligned (N, {Lt}, 35) / unaligned (N, {Lav}, 35)",
            "audio_lengths": "int64 (N,)  aligned: = 文本长度；unaligned: 80 ms 栅格位置数",
            "vision_lengths": "int64 (N,)  同上",
            "regression_labels": "float32 (N,)  label 列原值",
            "annotations": "object (N,)  annotation 列原值",
            "classification_labels": "float32 (N,)  由 label 按题面规则推出：0 Negative / 1 Neutral / 2 Positive",
        },
        "audio_feature_names": AUDIO_FEATURE_NAMES, "vision_feature_names": VISION_FEATURE_NAMES,
        "padding": "后填充（有效位置在前），填充行为 0",
        "aligned_position_rule": "i=0 [CLS] 与 i=L-1 [SEP] ↔ 整段 [0,D]；词元 ↔ 所属词时间区间按词元数均分；"
                                 "语音=区间内 100 Hz 帧均值，视觉=区间内人脸帧均值（无人脸→0 行）",
        "unaligned_position_rule": f"位置 k ↔ [k·{cfg['sequence']['grid_s']}, (k+1)·{cfg['sequence']['grid_s']}) 秒",
    }


def _extra_versions() -> dict:
    """common.utils.env_info 之外、问题1 用到的库版本。"""
    out = {}
    for mod in ("tokenizers", "imageio_ffmpeg", "soundfile", "yaml", "openpyxl"):
        try:
            out[mod] = getattr(__import__(mod), "__version__", "?")
        except Exception:  # noqa: BLE001
            out[mod] = None
    return out


# ----------------------------------------------------------------------------- 主程序

def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="问题1：多模态特征提取与时序对齐")
    ap.add_argument("--config", default="configs/q1.yaml")
    ap.add_argument("--run-name", default=None)
    ap.add_argument("--video-root", default=None, help="覆盖 paths.video_root")
    ap.add_argument("--label-xlsx", default=None, help="覆盖 paths.label_xlsx")
    ap.add_argument("--output-root", default=None, help="覆盖 paths.output_root")
    ap.add_argument("--limit", type=int, default=None, help="只处理前 N 行（调试用）")
    ap.add_argument("--resume", action="store_true", help="复用 cache/ 中配置一致的逐样本结果")
    ap.add_argument("--set", nargs="*", default=[], metavar="KEY=VALUE", help="覆盖任意配置项，如 text.model=xxx")
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    cfg = load_config(args.config, args.set)
    if args.run_name:
        cfg["run_name"] = args.run_name
    for k, v in (("video_root", args.video_root), ("label_xlsx", args.label_xlsx), ("output_root", args.output_root)):
        if v:
            cfg["paths"][k] = v
    run_dir = resolve(cfg["paths"]["output_root"]) / str(cfg["run_name"])
    run_dir.mkdir(parents=True, exist_ok=True)
    logger = get_logger("q1.extract", run_dir / "logs" / "extract.log")
    for name in ("q1.align",):  # 让 align 模块的回退警告也进日志文件
        lg = get_logger(name, run_dir / "logs" / "extract.log")
        lg.setLevel("INFO")
    t_start = time.time()
    seed = int(cfg.get("seed", 42))
    set_seed(seed)
    if cfg.get("num_threads"):
        import torch

        torch.set_num_threads(int(cfg["num_threads"]))
    media.set_ffmpeg(cfg["media"].get("ffmpeg"))
    logger.info("===== 问题1 特征提取开始：run=%s  配置=%s", cfg["run_name"], cfg["_config_file"])
    logger.info("命令行：%s", " ".join(sys.argv))

    video_root = resolve(cfg["paths"]["video_root"])
    label_path = resolve(cfg["paths"]["label_xlsx"])
    df = read_label_table(label_path, cfg["paths"].get("label_sheet", 0))
    if args.limit:
        df = df.head(args.limit)
    resolver = FileResolver(video_root)
    logger.info("标注表 %s：%d 行；视频根目录 %s：%d 个文件夹、%d 个视频文件", label_path, len(df), video_root,
                len(resolver.dirs), len(resolver.all_videos()))

    rows = []
    for _, r in df.iterrows():
        vid, cid = norm_id(r["video_id"]), norm_id(r["clip_id"])
        path, msg = resolver.find(vid, cid)
        if path is not None and msg == "loose":
            logger.warning("行 %d：%s/%s 宽松匹配到 %s（样本编号以实际文件夹/文件名为准）", r["row"], vid, cid, path)
            vid, cid = path.parent.name, path.stem
        rows.append({"video_id": vid, "clip_id": cid, "sample_id": f"{vid}{ID_SEP}{cid}", "text": r["text"],
                     "label": r["label"], "annotation": r["annotation"], "row": int(r["row"]), "path": path,
                     "resolve_msg": msg})
    dup = pd.Series([r["sample_id"] for r in rows]).duplicated()
    if dup.any():
        logger.warning("标注表中存在重复样本编号：%s", [rows[i]["sample_id"] for i in np.where(dup)[0]])

    # ---- 模型
    text_model = model_ref(cfg["text"]["model"])
    enc = TextEncoder(text_model, cfg["text"].get("device"), int(cfg["sequence"]["max_len_text"]))
    logger.info("文本模型：%s", enc.info())
    try:
        face = FaceAUExtractor(cfg["vision"], root=ROOT)
        logger.info("视觉工具：%s", face.info())
    except Exception as e:  # noqa: BLE001
        logger.error("视觉提取器初始化失败（所有样本视觉模态将记为 error）：%s", e)
        face = None
    ctc_info = None
    if cfg["align"]["method"] in ("auto", "ctc"):
        al = get_ctc_aligner(model_ref(cfg["align"]["ctc_model"]), cfg["align"].get("device"))
        if al is not None:
            ctc_info = al.info()
            logger.info("词级对齐：CTC 强制对齐，模型 %s", ctc_info)
        elif cfg["align"]["method"] == "ctc":
            raise RuntimeError(f"align.method=ctc 但模型加载失败：{ctc_failure_reason(model_ref(cfg['align']['ctc_model']), cfg['align'].get('device'))}")
        else:
            logger.warning("词级对齐：CTC 模型 %s 不可用 → 全部样本使用 vad_proportional 回退。原因：%s",
                           cfg["align"]["ctc_model"],
                           ctc_failure_reason(model_ref(cfg["align"]["ctc_model"]), cfg["align"].get("device")))
    else:
        logger.info("词级对齐：%s", cfg["align"]["method"])

    # ---- 逐样本
    cache_dir = run_dir / "cache"
    chash = config_hash(cfg)
    results = []
    for n, row in enumerate(rows):
        sid = row["sample_id"]
        cpath = cache_dir / f"{sid}.pkl"
        if args.resume and cpath.exists():
            try:
                with open(cpath, "rb") as f:
                    c = pickle.load(f)
                if c.get("config_hash") == chash:
                    results.append(c["result"])
                    logger.info("[%d/%d] %s 复用缓存", n + 1, len(rows), sid)
                    continue
            except Exception:  # noqa: BLE001
                pass
        r = process_sample(row, row["path"], enc, face, cfg, logger)
        results.append(r)
        m = r["manifest"]
        logger.info("[%d/%d] %s  dur=%.2fs  words=%s tok=%s  A=%d/%d  V=%d/%d  face=%s(%s)  aligner=%s  %s  %.1fs",
                    n + 1, len(rows), sid, m["audio_duration_s"] or 0, m["n_words"], m["text_len"],
                    m["audio_len_aligned"], m["audio_len_unaligned"], m["vision_len_aligned"],
                    m["vision_len_unaligned"], m["face_detect_rate"], m["vision_pass"], m["aligner_used"],
                    m["status"], m["proc_time_s"])
        for w in r["record"]["warnings"]:
            logger.warning("  %s: %s", sid, w)
        for e in r["record"]["errors"]:
            logger.error("  %s: %s", sid, e)
        if cfg["output"].get("save_cache", True):
            cache_dir.mkdir(parents=True, exist_ok=True)
            with open(cpath, "wb") as f:
                pickle.dump({"config_hash": chash, "result": r}, f, protocol=4)

    # ---- 写出
    feat_dir = run_dir / "features"
    aligned, unaligned = assemble(results, cfg)
    p_al = _save_pkl(aligned, feat_dir / f"aligned_{len(results)}.pkl", cfg["output"]["compress"])
    p_un = _save_pkl(unaligned, feat_dir / f"unaligned_{len(results)}.pkl", cfg["output"]["compress"])
    dump_json(feature_spec(cfg, len(results)), feat_dir / "feature_spec.json")
    al_dir = run_dir / "alignment"
    al_dir.mkdir(parents=True, exist_ok=True)
    for r in results:
        dump_json(r["record"], al_dir / f"{r['manifest']['sample_id']}.json", indent=1)
    man = pd.DataFrame([r["manifest"] for r in results])
    man.to_csv(run_dir / "manifest.csv", index=False, encoding="utf-8-sig")
    long = pd.DataFrame([x for r in results for x in _long_rows(r["manifest"], cfg)])
    long.to_csv(run_dir / "manifest_long.csv", index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(run_dir / "manifest.xlsx") as xw:
        man.to_excel(xw, sheet_name="样本清单", index=False)
        long.to_excel(xw, sheet_name="样本x模态", index=False)
    sizes = {p.name: p.stat().st_size for p in (p_al, p_un)}
    total_mb = sum(sizes.values()) / 2 ** 20
    status_counts = man["status"].value_counts().to_dict()
    face_rates = man["face_detect_rate"].dropna()
    run_info = {
        "task": "问题1 多模态情感特征提取与时序对齐", "run_name": cfg["run_name"],
        "started": datetime.fromtimestamp(t_start).isoformat(timespec="seconds"),
        "finished": datetime.now().isoformat(timespec="seconds"), "elapsed_s": round(time.time() - t_start, 1),
        "command": " ".join(sys.argv), "python": sys.executable, "seed": seed, "config_hash": chash,
        "env": {**env_info(), **_extra_versions()},
        "ffmpeg": {"path": media.find_ffmpeg(), "version": media.ffmpeg_version()},
        "models": {"text": enc.info(), "vision": face.info() if face else None,
                   "align": {"method": cfg["align"]["method"], "ctc": ctc_info,
                             "ctc_failure": ctc_failure_reason(model_ref(cfg["align"]["ctc_model"]),
                                                               cfg["align"].get("device"))}},
        "inputs": {"video_root": str(video_root), "label_xlsx": str(label_path), "n_rows": len(rows),
                   "label_xlsx_md5": media.file_md5(label_path)},
        "outputs": {"aligned": str(p_al.relative_to(run_dir)), "unaligned": str(p_un.relative_to(run_dir)),
                    "sizes_bytes": sizes, "total_mb": round(total_mb, 2)},
        "summary": {"status_counts": status_counts,
                    "aligner_counts": man["aligner_used"].fillna("none").value_counts().to_dict(),
                    "face_detect_rate_mean": float(face_rates.mean()) if len(face_rates) else None,
                    "face_detect_rate_min": float(face_rates.min()) if len(face_rates) else None,
                    "truncated_samples": int((man["truncated_words"].fillna(0) > 0).sum())},
        "config": {k: v for k, v in cfg.items() if not k.startswith("_")},
        "config_file": cfg["_config_file"],
    }
    dump_json(run_info, run_dir / "run_info.json")
    logger.info("特征文件：%s（%.2f MB）、%s（%.2f MB），合计 %.2f MB（dtype=%s, 压缩=%s）", p_al.name,
                sizes[p_al.name] / 2 ** 20, p_un.name, sizes[p_un.name] / 2 ** 20, total_mb, cfg["output"]["dtype"],
                cfg["output"]["compress"])
    logger.info("状态统计：%s；对齐方式：%s；平均人脸检出率：%s", status_counts, run_info["summary"]["aligner_counts"],
                run_info["summary"]["face_detect_rate_mean"])
    logger.info("===== 完成，用时 %.1fs，输出目录 %s", time.time() - t_start, run_dir)
    return run_dir


if __name__ == "__main__":
    main()
