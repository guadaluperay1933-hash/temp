"""典型样本对齐可视化：文本片段 ↔ 语音时段 ↔ 视频帧 ↔ 三类特征。

    python -m q1.visualize_sample --run-dir outputs/q1/<run> [--id "video_id$_$clip_id" ...] [--note 文字]

每个样本输出两张图到 <run_dir>/figures/：
  typical_<id>.png            对齐版（主图）：
      ① 6 个关键帧（带时间戳、对应的词、是否检出人脸）
      ② 波形 + RMS 包络，词按强制对齐得到的时间区间标在时间轴上，VAD 有声段浅色底
      ③ 对应带：每个词元位置 i 从其时间区间 [s_i, e_i) 连到序列位置 i（直观展示"时间 → 位置"的映射）
      ④ 文本特征：各位置 768 维向量在本样本内做 PCA 的前 8 个主成分（逐成分标准化）
      ⑤ 语音特征：F0、对数能量、H1−H2、MFCC1–12（逐行标准化）
      ⑥ 视觉特征：17 个 AU 强度（0–5；灰色 = 该位置无人脸帧，特征为 0 行）
  typical_<id>_unaligned.png  非对齐版：同一时间轴上 80 ms 栅格的语音 / 视觉序列
不指定 --id 时自动挑一个"典型"样本：处理状态正常、人脸检出率高、有效长度接近中位数。
"""
from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.patches import Polygon  # noqa: E402

from common.data import load_pkl  # noqa: E402
from common.utils import load_json  # noqa: E402

from .audio_feats import AUDIO_FEATURE_NAMES  # noqa: E402
from .config import resolve  # noqa: E402
from .media import load_audio, read_frame_at  # noqa: E402
from .visual_feats import AU_INTENSITY  # noqa: E402

INK, INK2, MUTED = "#0b0b0b", "#52514e", "#9a9994"
BLUE, AQUA, ORANGE = "#2a78d6", "#1baf7a", "#eb6834"
plt.rcParams.update({
    "font.sans-serif": ["Noto Sans CJK SC", "WenQuanYi Zen Hei", "DejaVu Sans"], "axes.unicode_minus": False,
    "axes.edgecolor": MUTED, "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2,
    "axes.titlesize": 10, "axes.titlecolor": INK, "font.size": 8.5,
    "text.parse_math": False,  # 样本编号含 "$_$"，不能当作数学公式解析
})


def _zrows(X: np.ndarray) -> np.ndarray:
    """逐行（每个特征）在位置维上标准化，仅用于显示。"""
    mu = X.mean(axis=1, keepdims=True)
    sd = X.std(axis=1, keepdims=True)
    return (X - mu) / np.where(sd < 1e-8, 1.0, sd)


def _pca_rows(X: np.ndarray, k: int = 8) -> np.ndarray:
    Xc = X - X.mean(axis=0, keepdims=True)
    if Xc.shape[0] < 2:
        return np.zeros((k, X.shape[0]))
    U, S, _ = np.linalg.svd(Xc, full_matrices=False)
    Z = (U[:, :k] * S[:k]).T  # (k, L)
    if Z.shape[0] < k:
        Z = np.vstack([Z, np.zeros((k - Z.shape[0], Z.shape[1]))])
    return _zrows(Z)


def pick_typical(man: pd.DataFrame) -> str:
    m = man[(man["status"] != "error")].copy()
    if m.empty:
        return str(man["sample_id"].iloc[0])
    m["score"] = -(m["text_len"] - m["text_len"].median()).abs() + 20 * m["face_detect_rate"].fillna(0) \
        - 10 * (m["truncated_words"].fillna(0) > 0)
    return str(m.sort_values("score", ascending=False)["sample_id"].iloc[0])


def _load_sample(run_dir: Path, sid: str):
    info = load_json(run_dir / "run_info.json")
    rec = load_json(run_dir / "alignment" / f"{sid}.json")
    feats = {}
    for v in ("aligned", "unaligned"):
        d = load_pkl(run_dir / info["outputs"][v])["all"]
        ids = [str(x) for x in d["id"]]
        n = ids.index(sid)
        feats[v] = {k: (np.asarray(d[k][n], np.float32) if k in ("text", "audio", "vision") else d[k][n])
                    for k in ("text", "audio", "vision", "text_bert", "audio_lengths", "vision_lengths",
                              "regression_labels", "annotations", "raw_text")}
    video = Path(info["inputs"]["video_root"]) / rec["file"]
    return info, rec, feats, video


def _time_axis_words(ax, words, D, ymax):
    """词按时间区间标注：交替三个高度，浅色区间底。"""
    for k, w in enumerate(words):
        c = BLUE if k % 2 == 0 else AQUA
        ax.axvspan(w["start"], w["end"], ymin=0.0, ymax=0.08, color=c, alpha=0.9, lw=0)
        lvl = [0.93, 0.80, 0.67][k % 3]
        ax.text(0.5 * (w["start"] + w["end"]), lvl * ymax, w["word"], ha="center", va="center", fontsize=7.5,
                color=INK, clip_on=True)


def draw_aligned(run_dir: Path, sid: str, note: str | None, out: Path) -> Path:
    info, rec, feats, video = _load_sample(run_dir, sid)
    sr = 16000
    y = load_audio(video, sr)
    D = rec["durations_s"]["audio"] or rec["durations_s"]["video"] or len(y) / sr
    words = rec["words"]
    pos = [p for p in rec["aligned"]["positions"] if p["rule"] != "pad"]
    L = len(pos)
    fa = feats["aligned"]

    fig = plt.figure(figsize=(15, 13.5), dpi=150)
    gs = fig.add_gridspec(6, 2, height_ratios=[2.1, 1.6, 0.8, 1.35, 2.0, 2.4], width_ratios=[60, 1],
                          hspace=0.42, wspace=0.02, left=0.07, right=0.95, top=0.92, bottom=0.08)
    # ① 关键帧
    gsf = gs[0, 0].subgridspec(1, 6, wspace=0.05)
    if len(words) >= 6:
        pick = np.linspace(0, len(words) - 1, 6).round().astype(int)
        key = [(0.5 * (words[k]["start"] + words[k]["end"]), words[k]["word"]) for k in pick]
    else:
        key = [(t, None) for t in np.linspace(0.05 * D, 0.95 * D, 6)]
    vf = rec["video_frame"]
    fps = rec["fps"] or 25.0
    ua_v = rec["unaligned"]["vision"]
    if ua_v["length"] and min(ua_v["n_used"]) == 0:
        # 存在无人脸时段：把离它最近的关键帧换成最长无人脸段的中点，直观展示"无人脸 → 0 行"
        from common.missing import spans_from_mask

        runs = spans_from_mask(np.array(ua_v["n_used"]) == 0)
        a, b = max(runs, key=lambda r: r[1] - r[0])
        t_nf = 0.5 * (a + b) * rec["unaligned"]["grid_s"]
        j = int(np.argmin([abs(t - t_nf) for t, _ in key]))
        w_nf = next((w["word"] for w in words if w["start"] <= t_nf < w["end"]), None)
        key[j] = (t_nf, w_nf)
        key.sort(key=lambda x: x[0])
    for j, (t, w) in enumerate(key):
        ax = fig.add_subplot(gsf[0, j])
        try:
            ax.imshow(read_frame_at(video, t))
        except Exception:  # noqa: BLE001
            ax.text(0.5, 0.5, "读帧失败", ha="center", transform=ax.transAxes)
        k = min(int(t / rec["unaligned"]["grid_s"]), max(ua_v["length"] - 1, 0))
        has_face = ua_v["length"] > 0 and ua_v["n_used"][k] > 0
        ax.set_title(f"t={t:.2f}s  帧#{int(round(t * fps))}" + (f"\n“{w}”" if w else "") +
                     ("" if has_face else "\n(未检出人脸)"), fontsize=8, color=INK if has_face else ORANGE)
        ax.set_xticks([])
        ax.set_yticks([])
    # ② 波形 + RMS + 词
    ax_w = fig.add_subplot(gs[1, 0])
    t_ax = np.arange(len(y)) / sr
    step = max(1, len(y) // 20000)
    ax_w.plot(t_ax[::step], y[::step], color=MUTED, lw=0.4)
    hop = int(0.01 * sr)
    if len(y) > 400:
        import librosa

        rms = librosa.feature.rms(y=y, frame_length=400, hop_length=hop, center=True)[0]
        k_rms = 0.6 * max(float(np.abs(y).max()), 1e-3) / max(float(rms.max()), 1e-8)
        ax_w.plot(np.arange(len(rms)) * hop / sr, rms * k_rms, color=BLUE, lw=1.4, label="RMS 包络（缩放显示）")
    for a, b in rec.get("voiced_intervals_s", []):
        ax_w.axvspan(a, b, color=BLUE, alpha=0.06, lw=0)
    ymax = max(float(np.abs(y).max()) if len(y) else 1.0, 1e-3) * 1.15
    ax_w.set_ylim(-ymax, ymax)
    _time_axis_words(ax_w, words, D, ymax)
    for t, _ in key:
        ax_w.axvline(t, color=ORANGE, lw=0.8, ls="--")
    ax_w.set_xlim(0, D)
    ax_w.set_xlabel("时间 / s")
    ax_w.set_ylabel("波形")
    ax_w.set_title(f"② 语音波形与词级时间戳（对齐方式：{rec['aligner'].get('method_used')}；浅蓝底 = VAD 有声段；"
                   f"橙色虚线 = 上方关键帧时刻）", loc="left")
    ax_w.legend(loc="lower right", fontsize=7, frameon=False)
    for s in ("top", "right"):
        ax_w.spines[s].set_visible(False)
    # ③ 对应带
    ax_b = fig.add_subplot(gs[2, 0])
    ax_b.set_xlim(0, 1)
    ax_b.set_ylim(0, 1)
    ax_b.axis("off")
    for p in pos:
        i = p["pos"]
        if p["rule"] != "word_piece":
            continue
        c = BLUE if p["word_index"] % 2 == 0 else AQUA
        poly = Polygon([(p["start_s"] / D, 1), (p["end_s"] / D, 1), ((i + 0.9) / L, 0), ((i + 0.1) / L, 0)],
                       closed=True, color=c, alpha=0.35, lw=0)
        ax_b.add_patch(poly)
    ax_b.text(0.0, 0.5, "③ 时间区间 → 序列位置", fontsize=9, color=INK, va="center", ha="left",
              transform=ax_b.transAxes, bbox=dict(facecolor="white", edgecolor="none", alpha=0.8))
    ax_b.text(1.0, 0.5, "[CLS]/[SEP] ↔ 整段 [0, D]", fontsize=8, color=INK2, va="center", ha="right",
              transform=ax_b.transAxes, bbox=dict(facecolor="white", edgecolor="none", alpha=0.8))
    # ④⑤⑥ 位置轴热图
    tokens = [p["token"] for p in pos]
    Xt = fa["text"][:L]
    Xa = fa["audio"][:L]
    Xv = fa["vision"][:L]
    rows = [(3, "④ 文本：BERT 768 维 → 本样本 PCA 前 8 主成分（逐成分标准化）", _pca_rows(Xt, 8),
             [f"PC{k + 1}" for k in range(8)], "RdBu_r", (-2.5, 2.5), "z 分数"),
            (4, "⑤ 语音：词元区间内 100 Hz 帧均值（逐行标准化）",
             _zrows(Xa[:, [0, 3, 9] + list(range(12, 24))].T),
             ["F0", "logE", "H1−H2"] + [AUDIO_FEATURE_NAMES[k] for k in range(12, 24)], "RdBu_r", (-2.5, 2.5), "z 分数"),
            (5, "⑥ 视觉：词元区间内人脸帧的 AU 强度（灰 = 区间内无人脸帧 → 0 行）", Xv[:, :17].T, AU_INTENSITY,
             "Blues", (0, 5), "强度 0–5")]
    no_face = ~(np.abs(Xv).max(axis=1) > 1e-8)
    for r, title, M, ylabels, cmap, (vmin, vmax), clabel in rows:
        ax = fig.add_subplot(gs[r, 0])
        Mm = np.ma.array(M, mask=np.zeros_like(M, bool))
        if r == 5:
            Mm.mask[:, no_face] = True
        cm = plt.get_cmap(cmap).copy()
        cm.set_bad("#d9d8d4")
        im = ax.imshow(Mm, aspect="auto", cmap=cm, vmin=vmin, vmax=vmax, interpolation="nearest",
                       extent=(-0.5, L - 0.5, len(ylabels) - 0.5, -0.5))
        ax.set_yticks(range(len(ylabels)))
        ax.set_yticklabels(ylabels, fontsize=6.5)
        ax.set_xlim(-0.5, L - 0.5)
        ax.set_title(title, loc="left")
        if r == 5:
            ax.set_xticks(range(L))
            ax.set_xticklabels([f"{i}:{t}" for i, t in enumerate(tokens)], rotation=90, fontsize=6.5)
            ax.set_xlabel(f"序列位置（词元）；有效长度 L={L}，位置 {L}–49 为填充（三模态 0 行）")
        else:
            ax.set_xticks(range(L))
            ax.set_xticklabels([])
        cax = fig.add_subplot(gs[r, 1])
        cb = fig.colorbar(im, cax=cax)
        cb.set_label(clabel, fontsize=7)
        cb.ax.tick_params(labelsize=6.5)
    lab = float(fa["regression_labels"])
    ann = str(fa["annotations"])
    txt = str(fa["raw_text"])
    txt = txt if len(txt) <= 150 else txt[:147] + "…"
    fig.suptitle(f"样本 {sid}   情感强度 label = {lab:+.3f}（{ann}）   时长 {D:.2f}s   "
                 f"人脸检出率 {vf.get('detect_rate')}", fontsize=12, color=INK, x=0.07, ha="left", y=0.985)
    fig.text(0.07, 0.955, f"转写：{txt}" + (f"\n{note}" if note else ""), fontsize=9, color=INK2, va="top")
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out)
    plt.close(fig)
    return out


def draw_unaligned(run_dir: Path, sid: str, note: str | None, out: Path) -> Path:
    info, rec, feats, video = _load_sample(run_dir, sid)
    fu = feats["unaligned"]
    grid = rec["unaligned"]["grid_s"]
    Ka, Kv = int(fu["audio_lengths"]), int(fu["vision_lengths"])
    D = max(Ka, Kv) * grid
    fig = plt.figure(figsize=(15, 7.2), dpi=150)
    gs = fig.add_gridspec(3, 2, height_ratios=[1.3, 2.2, 2.4], width_ratios=[60, 1], hspace=0.38, wspace=0.02,
                          left=0.07, right=0.95, top=0.88, bottom=0.09)
    ax = fig.add_subplot(gs[0, 0])
    Xa = fu["audio"][:Ka]
    tk = (np.arange(Ka) + 0.5) * grid
    ax.plot(tk, Xa[:, 3], color=BLUE, lw=1.3)
    ax.set_ylabel("logE / dB")
    ymax = float(Xa[:, 3].max()) if Ka else 0.0
    ymin = float(Xa[:, 3].min()) if Ka else -100.0
    ax.set_ylim(ymin - 2, ymax + 0.35 * (ymax - ymin + 1))
    for k, w in enumerate(rec["words"]):
        ax.axvspan(w["start"], w["end"], ymin=0.9, ymax=1.0, color=BLUE if k % 2 == 0 else AQUA, alpha=0.8, lw=0)
    ax.set_xlim(0, D)
    ax.set_title(f"非对齐版：语音对数能量（{int(grid * 1000)} ms 栅格；顶部色条 = 词级时间区间）", loc="left")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for r, (title, M, ylabels, cmap, lim, K) in enumerate([
            (f"语音 {Ka} 个位置（F0、logE、MFCC1–12，逐行标准化）", _zrows(Xa[:, [0, 3] + list(range(12, 24))].T),
             ["F0", "logE"] + [AUDIO_FEATURE_NAMES[k] for k in range(12, 24)], "RdBu_r", (-2.5, 2.5), Ka),
            (f"视觉 {Kv} 个位置（17 个 AU 强度；灰 = 栅格内无人脸帧）", fu["vision"][:Kv, :17].T, AU_INTENSITY, "Blues",
             (0, 5), Kv)], start=1):
        ax = fig.add_subplot(gs[r, 0])
        Mm = np.ma.array(M, mask=np.zeros_like(M, bool))
        if r == 2:
            Mm.mask[:, ~(np.abs(fu["vision"][:Kv]).max(axis=1) > 1e-8)] = True
        cm = plt.get_cmap(cmap).copy()
        cm.set_bad("#d9d8d4")
        im = ax.imshow(Mm, aspect="auto", cmap=cm, vmin=lim[0], vmax=lim[1], interpolation="nearest",
                       extent=(0, K * grid, len(ylabels) - 0.5, -0.5))
        ax.set_xlim(0, D)
        ax.set_yticks(range(len(ylabels)))
        ax.set_yticklabels(ylabels, fontsize=6.5)
        ax.set_title(title, loc="left")
        if r == 2:
            ax.set_xlabel(f"时间 / s（位置 k ↔ [{grid}k, {grid}(k+1)) s；k ≥ 长度为填充 0 行）")
        cax = fig.add_subplot(gs[r, 1])
        cb = fig.colorbar(im, cax=cax)
        cb.ax.tick_params(labelsize=6.5)
    fig.suptitle(f"样本 {sid}   audio_lengths={Ka}  vision_lengths={Kv}" + (f"   {note}" if note else ""),
                 fontsize=12, color=INK, x=0.07, ha="left")
    fig.savefig(out)
    plt.close(fig)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="问题1 典型样本对齐可视化")
    ap.add_argument("--run-dir", required=True)
    ap.add_argument("--id", action="append", default=None, help="样本编号 video_id$_$clip_id，可重复；缺省自动挑选")
    ap.add_argument("--note", default=None, help="图中附加说明（如'合成测试数据，仅验证流程'）")
    a = ap.parse_args(argv)
    run_dir = resolve(a.run_dir)
    man = pd.read_csv(run_dir / "manifest.csv", dtype={"video_id": str, "clip_id": str})
    ids = a.id or [pick_typical(man)]
    for sid in ids:
        safe = sid.replace("/", "_")
        p1 = draw_aligned(run_dir, sid, a.note, run_dir / "figures" / f"typical_{safe}.png")
        p2 = draw_unaligned(run_dir, sid, a.note, run_dir / "figures" / f"typical_{safe}_unaligned.png")
        print(f"{sid} → {p1}\n{' ' * len(sid)} → {p2}")


if __name__ == "__main__":
    main()
