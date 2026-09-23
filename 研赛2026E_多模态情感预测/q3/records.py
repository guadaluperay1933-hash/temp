"""逐样本解释记录：把 bundle（explain.run_explain 的数组）+ 证据片段 + 时间映射 整理成一条 JSON 友好的字典，
供解释卡、附件4 CSV / explanations.json 共用，保证三处给出的数字完全一致。

位置坐标约定：
    'pos'        下采样后的全局位置 k（模型实际使用的分辨率）
    'orig_lo/hi' 原始序列位置（含两端）；对齐版本 / 未下采样时与 k 相同
    词序号 'word_lo/hi' 从 1 开始计（与论文、CSV 中"词7-8"的写法一致）
"""
from __future__ import annotations

import numpy as np

from common.data import CLASS_CN, CLASS_NAMES, MOD_CN, MODALITIES, split_id
from common.decision import project_intensity

from .timemap import frame_index, pooled_interval

EXPERT_CN = {"text": "文本", "audio": "语音", "vision": "视觉", "interaction": "跨模态交互"}


def _fmt_t(x: float) -> str:
    return f"{x:.2f}"


def text_word_index(sd, i: int, mapper) -> np.ndarray:
    """文本每个原始位置对应的词序号（0 起；[CLS]/[SEP]/无效位置为 −1）。"""
    Lt = sd.feats["text"].shape[1]
    out = np.full(Lt, -1, np.int64)
    vt = np.where(sd.valid["text"][i])[0]
    if not vt.size:
        return out
    tb = sd.extra.get("text_bert")
    ids = np.asarray(tb)[i, 0, vt[0]:] if (tb is not None and np.asarray(tb).ndim == 3) else None
    tmap, _ = mapper.map(sd.raw_text[i] if sd.raw_text else "", int(vt.size), ids, max_len=Lt)
    for r, p in enumerate(vt):
        if r < len(tmap) and tmap[r] is not None:
            out[p] = tmap[r][0]
    return out


def build_record(b: dict, i: int, evidence_i: dict, sd, word_idx: np.ndarray | None = None, tmap: dict | None = None,
                 deletion_drop: float | None = None, deletion_drop_random: float | None = None) -> dict:
    """一条样本的完整解释记录。sd：该划分的 SplitData（取 id / raw_text / 标签 / valid）。"""
    sid = b["ids"][i]
    vid, cid = split_id(sid)
    c = int(b["pred_cls"][i])
    y_raw = float(b["y_raw"][i])
    words = (sd.raw_text[i] if sd.raw_text else "").split()
    rec = {"index": int(i), "id": sid, "video_id": vid, "clip_id": cid, "raw_text": sd.raw_text[i] if sd.raw_text else "",
           "pred_class": c, "pred_polarity": CLASS_NAMES[c], "pred_polarity_cn": CLASS_CN[CLASS_NAMES[c]],
           "pred_intensity": float(project_intensity([c], [y_raw])[0]), "raw_intensity": y_raw,
           "probs": [float(p) for p in b["probs"][i]], "confidence": float(b["probs"][i].max()),
           "shapley": {m: float(b["phi"][i, k]) for k, m in enumerate(MODALITIES)},
           "shapley_prob": {m: float(b["phi_prob"][i, k]) for k, m in enumerate(MODALITIES)},
           "shapley_base": float(b["base"][i]),
           "coalition_values": {n: float(v) for n, v in zip(_coal_names(), b["v_coal"][i])},
           "share": {m: float(b["share"][i, k]) for k, m in enumerate(MODALITIES)},
           "main_modality": MODALITIES[int(b["main"][i])]}
    if sd.y_reg is not None:
        rec["y_true"] = float(sd.y_reg[i])
        rec["cls_true"] = CLASS_NAMES[int(sd.y_cls[i])]
    if b.get("w") is not None:
        rec["gate"] = {e: float(b["w"][i, k]) for k, e in enumerate(b["experts"])}
        rec["contribution"] = {e: float(b["c"][i, k]) for k, e in enumerate(b["experts"])}
        rec["expert_output"] = {e: float(b["s"][i, k]) for k, e in enumerate(b["experts"])}
    if tmap is not None:
        rec["video"] = tmap.get("video")
        rec["duration"] = tmap.get("duration")
        rec["fps"] = tmap.get("fps")
        rec["n_frames"] = tmap.get("n_frames")
        rec["word_time_source"] = tmap.get("word_source")
        rec["text_map_method"] = tmap.get("text_map_method")
        rec["time_mapping_method"] = dict(tmap.get("method", {}))
        rec["word_times"] = [{"word": w["word"], "start": round(float(w["start"]), 3), "end": round(float(w["end"]), 3)}
                             for w in tmap.get("words", [])]
    # ---- 逐位置
    pos = {}
    for k_m, m in enumerate(MODALITIES):
        st = int(b["strides"][m])
        valid = b["pos"][m]["valid"][i]
        ks = np.where(valid)[0]
        vrow = sd.valid[m][i]
        d = {"pos": ks.tolist(),
             "orig_lo": [int(k * st) for k in ks],
             "orig_hi": [int(min((k + 1) * st, len(vrow)) - 1) for k in ks],
             "cand": b["pos"][m]["cand"][i, ks].astype(bool).tolist(),
             "attn": np.round(b["pos"][m]["attn"][i, ks], 6).tolist(),
             "value": np.round(b["pos"][m]["value"][i, ks], 6).tolist(),
             "contrib": np.round(b["pos"][m]["contrib"][i, ks], 6).tolist(),
             "contrib_cls": np.round(b["pos"][m]["contrib_cls"][i, ks], 6).tolist()}
        if m == "text" and word_idx is not None:
            d["word"] = [int(word_idx[k]) if st == 1 else -1 for k in ks]
        if tmap is not None:
            iv = [pooled_interval(tmap, m, int(k), st, vrow) for k in ks]
            d["t_start"] = [round(v[0], 4) if v else None for v in iv]
            d["t_end"] = [round(v[1], 4) if v else None for v in iv]
        pos[m] = d
    rec["positions"] = pos
    # ---- 证据片段
    ev_out = {}
    for m in MODALITIES:
        segs = []
        st = int(b["strides"][m])
        vrow = sd.valid[m][i]
        for s in evidence_i[m]["segments"]:
            seg = dict(s)
            seg["orig_lo"] = int(s["lo"] * st)
            seg["orig_hi"] = int(min((s["hi"] + 1) * st, len(vrow)) - 1)
            if m == "text" and word_idx is not None:
                ws = sorted({int(word_idx[p]) for p in s["positions"] if word_idx[p] >= 0})
                if ws:
                    seg["word_lo"], seg["word_hi"] = ws[0] + 1, ws[-1] + 1
                    seg["words"] = " ".join(words[ws[0]:ws[-1] + 1]) if words else ""
            if tmap is not None:
                a = pooled_interval(tmap, m, s["lo"], st, vrow)
                z = pooled_interval(tmap, m, s["hi"], st, vrow)
                pk = pooled_interval(tmap, m, s["peak"], st, vrow)
                if a and z:
                    seg["t_start"], seg["t_end"] = round(a[0], 3), round(z[1], 3)
                if pk:
                    seg["peak_t"] = round(0.5 * (pk[0] + pk[1]), 3)
                    seg["peak_frame"] = frame_index(seg["peak_t"], tmap.get("fps"), tmap.get("n_frames"))
            segs.append(seg)
        ev_out[m] = segs
    rec["evidence"] = ev_out
    rec["text_evidence"] = format_text_evidence(ev_out["text"])
    rec["audio_evidence"] = format_av_evidence(ev_out["audio"], "audio")
    rec["vision_evidence"] = format_av_evidence(ev_out["vision"], "vision")
    for m in MODALITIES:
        rec[f"{m}_evidence_positions"] = ";".join(
            (f"{s['orig_lo']}-{s['orig_hi']}" if s["orig_hi"] > s["orig_lo"] else f"{s['orig_lo']}")
            for s in ev_out[m])
    if tmap is not None:
        meth = rec["time_mapping_method"]
        rates = tmap.get("rates", {})
        parts = []
        for m in MODALITIES:
            t = meth.get(m, "none")
            if t == "fixed-rate" and rates.get(m, {}).get("rate_hz"):
                t += f"({rates[m]['rate_hz']:.2f}Hz)"
            parts.append(f"{MOD_CN[m]}:{t}")
        rec["time_mapping"] = "; ".join(parts)
    else:
        rec["time_mapping"] = "none"
    rec["deletion_drop"] = None if deletion_drop is None else float(deletion_drop)
    rec["deletion_drop_random"] = None if deletion_drop_random is None else float(deletion_drop_random)
    return rec


def _coal_names():
    from .explain import COALITION_NAMES

    return COALITION_NAMES


def format_text_evidence(segs: list[dict]) -> str:
    """'blown away [词7-8, 3.45–4.10 s] (+0.52)'，多段按 |贡献| 降序以 '; ' 连接。"""
    out = []
    for s in sorted(segs, key=lambda s: -s["abs"]):
        frag = s.get("words") or f"位置{s['orig_lo']}" + (f"-{s['orig_hi']}" if s["orig_hi"] > s["orig_lo"] else "")
        loc = []
        if "word_lo" in s:
            loc.append(f"词{s['word_lo']}" + (f"-{s['word_hi']}" if s["word_hi"] > s["word_lo"] else ""))
        if "t_start" in s:
            loc.append(f"{_fmt_t(s['t_start'])}–{_fmt_t(s['t_end'])} s")
        out.append(f"{frag}" + (f" [{', '.join(loc)}]" if loc else "") + f" ({s['sum']:+.3f})")
    return "; ".join(out)


def format_av_evidence(segs: list[dict], m: str) -> str:
    """语音：'1.20–1.85 s (位置15–23)'；视觉：'1.93–2.60 s，关键帧 2.28 s (第68帧) (位置58–78)'。"""
    out = []
    for s in sorted(segs, key=lambda s: -s["abs"]):
        p = f"位置{s['orig_lo']}" + (f"–{s['orig_hi']}" if s["orig_hi"] > s["orig_lo"] else "")
        t = f"{_fmt_t(s['t_start'])}–{_fmt_t(s['t_end'])} s" if "t_start" in s else ""
        if m == "vision" and s.get("peak_t") is not None:
            kf = f"关键帧 {_fmt_t(s['peak_t'])} s" + (f" (第{s['peak_frame']}帧)" if s.get("peak_frame") is not None else "")
            t = f"{t}，{kf}" if t else kf
        out.append((f"{t} ({p})" if t else p) + f" ({s['sum']:+.3f})")
    return "; ".join(out)
