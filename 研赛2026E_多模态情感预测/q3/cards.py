"""解释卡：一条样本一张 PNG。

版面（自上而下）：
    1. 抬头：样本编号、预测极性 / 强度 / 置信度、主要参考模态、Shapley 分解式、真实标签（有则给）、时间映射方式
    2. 左：三模态 Shapley 作用份额 与 门控权重（含跨模态交互）对比柱状图，柱上标注带符号的 φ_m
       右：文本逐词贡献（蓝 = 推向正向，红 = 推向负向，颜色深浅 ∝ |贡献|），关键证据词加黑框
    3. 语音、视觉：逐位置贡献 e_{m,t} 随时间（秒；无视频时为原始位置序号）的柱状曲线，关键片段加阴影
    4. （有视频时）在关键片段峰值时刻抽取的 1–3 张关键帧，标注时间与帧号
"""
from __future__ import annotations

from pathlib import Path

import numpy as np

from common.data import MOD_CN, MODALITIES

from .records import EXPERT_CN

POS_COLOR = "#2166AC"   # 正向贡献（蓝）
NEG_COLOR = "#B2182B"   # 负向贡献（红）
SHADE = "#FDB863"
YH = r"$\hat{y}$"      # CJK 字体里没有 ŷ，用 mathtext 画


def esc(s) -> str:
    """matplotlib 会把成对的 $ 当作数学公式（样本编号里的 $_$ 就会触发），画字符串前统一转义。"""
    return str(s).replace("$", r"\$")


def setup_font(plt) -> None:
    try:
        from common.error_analysis import _setup_cjk_font

        _setup_cjk_font(plt)
    except Exception:  # noqa: BLE001 - 字体只影响显示
        pass


def _word_contrib(rec: dict) -> tuple[list[str], np.ndarray, set]:
    """词级贡献 = 该词各 wordpiece 位置贡献之和；证据词 = 文本证据片段覆盖的词。"""
    words = rec["raw_text"].split()
    p = rec["positions"]["text"]
    wc = np.zeros(len(words))
    if "word" in p:
        for w, e in zip(p["word"], p["contrib"]):
            if 0 <= w < len(words):
                wc[w] += e
    ev = set()
    for s in rec["evidence"]["text"]:
        if "word_lo" in s:
            ev.update(range(s["word_lo"] - 1, s["word_hi"]))
    return words, wc, ev


def _draw_words(ax, fig, words, wc, ev, fontsize=10.5):
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    if not words:
        ax.text(0.02, 0.5, "（无转写文本）", fontsize=fontsize, va="center")
        return
    import matplotlib.pyplot as plt

    vmax = max(np.abs(wc).max(), 1e-9)
    renderer = fig.canvas.get_renderer()
    inv = ax.transAxes.inverted()
    x, y, line_h = 0.01, 0.90, 0.13
    for j, w in enumerate(words):
        v = wc[j] / vmax
        col = plt.cm.RdBu(0.5 + 0.5 * v)
        t = ax.text(x, y, esc(w), fontsize=fontsize, va="top", ha="left",
                    bbox=dict(boxstyle="round,pad=0.25", fc=col, ec="black" if j in ev else "none",
                              lw=1.8 if j in ev else 0))
        bb = t.get_window_extent(renderer=renderer)
        (x0, _), (x1, _) = inv.transform([[bb.x0, bb.y0], [bb.x1, bb.y1]])
        width = x1 - x0 + 0.012
        if x + width > 0.99 and x > 0.02:
            t.remove()
            x, y = 0.01, y - line_h
            t = ax.text(x, y, esc(w), fontsize=fontsize, va="top", ha="left",
                        bbox=dict(boxstyle="round,pad=0.25", fc=col, ec="black" if j in ev else "none",
                                  lw=1.8 if j in ev else 0))
        x += width
        if y < 0.05:
            ax.text(x, y, "…", fontsize=fontsize, va="top")
            break


def _curve(ax, rec: dict, m: str, use_time: bool):
    p = rec["positions"][m]
    e = np.asarray(p["contrib"], float)
    if use_time and p.get("t_start") and all(v is not None for v in p["t_start"]):
        x0 = np.asarray(p["t_start"], float)
        x1 = np.asarray(p["t_end"], float)
        xl = "时间 (s)"
    else:
        x0 = np.asarray(p["orig_lo"], float)
        x1 = np.asarray(p["orig_hi"], float) + 1
        xl = "原始序列位置"
    if e.size == 0:
        ax.text(0.5, 0.5, "该模态无有效位置", transform=ax.transAxes, ha="center")
        return
    wdt = np.maximum(x1 - x0, 1e-6)
    ax.bar(x0, e, width=wdt * 0.92, align="edge", color=np.where(e >= 0, POS_COLOR, NEG_COLOR), lw=0)
    ax.axhline(0, color="black", lw=0.6)
    pos_of = {k: j for j, k in enumerate(p["pos"])}
    for s in rec["evidence"][m]:
        a, b = pos_of.get(s["lo"]), pos_of.get(s["hi"])
        if a is None or b is None:
            continue
        ax.axvspan(x0[a], x1[b], color=SHADE, alpha=0.35, lw=0)
        lab = (f"{s['t_start']:.2f}–{s['t_end']:.2f}s" if (use_time and "t_start" in s)
               else f"位置{s['orig_lo']}–{s['orig_hi']}")
        ax.text(0.5 * (x0[a] + x1[b]), 1.0, lab, transform=ax.get_xaxis_transform(), ha="center", va="bottom",
                fontsize=8, color="#8c4b00")
    ax.set_xlim(x0.min(), x1.max())
    ax.set_xlabel(xl, fontsize=9)
    ax.set_ylabel("贡献 e", fontsize=9)
    ax.tick_params(labelsize=8)
    c = rec.get("contribution", {}).get(m)
    extra = f"，Σe = c = {c:+.3f}" if c is not None else ""
    ax.set_title(f"{MOD_CN[m]}：逐位置贡献（阴影 = 关键片段）{extra}；Shapley φ = {rec['shapley'][m]:+.3f}，"
                 f"份额 {100 * rec['share'][m]:.1f}%", fontsize=10, loc="left", pad=12)


def _frames(rec: dict, max_frames: int) -> list[tuple[float, int | None, str]]:
    """关键帧时刻：先取视觉片段峰值（≤2），不足再补语音片段峰值。"""
    out = []
    for m in ("vision", "audio"):
        for s in sorted(rec["evidence"][m], key=lambda s: -s["abs"]):
            if s.get("peak_t") is not None and len(out) < max_frames:
                fr = s.get("peak_frame")
                if fr is None and rec.get("fps"):
                    fr = int(round(s["peak_t"] * rec["fps"]))
                out.append((float(s["peak_t"]), fr, MOD_CN[m]))
    return out


def render_card(rec: dict, path, max_frames: int = 3, frame_reader=None) -> Path:
    """画一张解释卡。frame_reader(video_path, t) → RGB 图（None 表示不抽帧）。"""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.gridspec import GridSpec

    setup_font(plt)
    frames = []
    if frame_reader is not None and rec.get("video"):
        for t, fr, src in _frames(rec, max_frames):
            try:
                img = frame_reader(rec["video"], t)
            except Exception:  # noqa: BLE001
                img = None
            if img is not None:
                frames.append((t, fr, src, img))
    use_time = rec.get("time_mapping", "none") != "none"
    n_rows = 5 if frames else 4
    heights = [0.95, 2.1, 1.45, 1.45] + ([1.9] if frames else [])
    fig = plt.figure(figsize=(13, sum(heights) + 0.9))
    gs = GridSpec(n_rows, 6, height_ratios=heights, hspace=0.75, wspace=0.6, figure=fig,
                  left=0.06, right=0.98, top=0.97, bottom=0.05)
    # ---- 1. 抬头
    ax = fig.add_subplot(gs[0, :])
    ax.set_axis_off()
    main = rec["main_modality"]
    col = POS_COLOR if rec["pred_class"] == 2 else (NEG_COLOR if rec["pred_class"] == 0 else "#555555")
    ax.text(0, 1.0, f"样本 {esc(rec['id'])}", fontsize=13, weight="bold", va="top")
    ax.text(0, 0.68, f"预测：{rec['pred_polarity_cn']}（{rec['pred_polarity']}）  强度 {rec['pred_intensity']:+.2f}"
                     f"（原始 {rec['raw_intensity']:+.3f}）  置信度 {rec['confidence']:.2f}", fontsize=11.5,
            color=col, weight="bold", va="top")
    ax.text(0.62, 0.68, f"主要参考模态：{MOD_CN[main]}（Shapley 份额 {100 * rec['share'][main]:.1f}%）", fontsize=11.5,
            weight="bold", va="top")
    phis = "  ".join(f"φ_{m[0].upper()}={rec['shapley'][m]:+.3f}" for m in MODALITIES)
    line3 = (f"概率 负/中/正 = {rec['probs'][0]:.2f}/{rec['probs'][1]:.2f}/{rec['probs'][2]:.2f}；  "
             f"{YH} = {YH}(∅) + Σφ = {rec['shapley_base']:+.3f} + ({phis})")
    ax.text(0, 0.36, line3, fontsize=9.5, va="top")
    extra = []
    if "y_true" in rec:
        extra.append(f"真实标签：{rec['cls_true']}（{rec['y_true']:+.2f}）")
    extra.append(f"时间映射：{rec.get('time_mapping', 'none')}")
    if rec.get("deletion_drop") is not None:
        extra.append(f"删除全部证据后 |Δ{YH}| = {rec['deletion_drop']:.3f}")
    ax.text(0, 0.08, "；  ".join(extra), fontsize=9, color="#444444", va="top")
    # ---- 2a. 份额 vs 门控
    ax = fig.add_subplot(gs[1, :2])
    names = ["text", "audio", "vision", "interaction"]
    xs = np.arange(4)
    sh = [100 * rec["share"][m] for m in MODALITIES] + [0.0]
    gate = [100 * rec.get("gate", {}).get(e, np.nan) for e in names]
    ax.bar(xs - 0.2, sh, 0.4, color="#4C78A8", label="Shapley 份额 |φ|/Σ|φ|")
    ax.bar(xs + 0.2, gate, 0.4, color="#F58518", label="门控权重 w")
    for k, m in enumerate(MODALITIES):
        ax.text(k - 0.2, sh[k] + 2, f"{rec['shapley'][m]:+.2f}", ha="center", fontsize=8)
    ax.text(3 - 0.2, 2, "—", ha="center", fontsize=9, color="#888888")
    ax.set_xticks(xs, [EXPERT_CN[e] for e in names], fontsize=9)
    for lab, e in zip(ax.get_xticklabels(), names):
        if e == main:
            lab.set_fontweight("bold")
            lab.set_color("#C0392B")
    ax.set_ylim(0, 115)
    ax.set_ylabel("%", fontsize=9)
    ax.legend(fontsize=7.5, loc="upper right", frameon=False)
    ax.set_title("模态作用程度（柱上数字 = 带符号 φ）", fontsize=10, loc="left")
    ax.tick_params(labelsize=8)
    # ---- 2b. 文本
    ax = fig.add_subplot(gs[1, 2:])
    words, wc, ev = _word_contrib(rec)
    fig.canvas.draw()
    _draw_words(ax, fig, words, wc, ev)
    ct = rec.get("contribution", {}).get("text")
    ax.set_title("文本：逐词贡献（蓝 = 推向正向，红 = 推向负向，黑框 = 关键证据）"
                 + (f"，Σ = {ct:+.3f}" if ct is not None else ""), fontsize=10, loc="left")
    if rec["text_evidence"]:
        ax.text(0.01, -0.04, "证据：" + esc(rec["text_evidence"]), fontsize=8.5, transform=ax.transAxes, va="top",
                color="#333333", wrap=True)
    # ---- 3. 语音 / 视觉
    for r, m in ((2, "audio"), (3, "vision")):
        _curve(fig.add_subplot(gs[r, :]), rec, m, use_time)
    # ---- 4. 关键帧
    if frames:
        n = len(frames)
        for j, (t, fr, src, img) in enumerate(frames):
            ax = fig.add_subplot(gs[4, 2 * j:2 * j + 2])
            ax.imshow(img)
            ax.set_axis_off()
            ax.set_title(f"{src}证据关键帧 {t:.2f} s" + (f"（第{fr}帧）" if fr is not None else ""), fontsize=9.5)
        for j in range(n, 3):
            fig.add_subplot(gs[4, 2 * j:2 * j + 2]).set_axis_off()
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=130)
    plt.close(fig)
    return path


def safe_name(sid: str) -> str:
    return "".join(ch if (ch.isalnum() or ch in "-_.") else "_" for ch in sid.replace("$_$", "__"))
