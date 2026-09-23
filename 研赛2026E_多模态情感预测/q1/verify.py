"""问题1 产物核验：覆盖完整性（标注行 ↔ 视频文件 ↔ 特征条目 ↔ 对齐记录 一一对应）、形状、数值、长度与填充规则、
时间单调性、标签原样性、文件校验和、common.data 可读性。

    python -m q1.verify --run-dir outputs/q1/<run_name>

输出 <run_dir>/verification_report.json 与 verification_report.md；有任何一项不通过时退出码为 1。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

from common.data import ZERO_EPS, load_all_splits, load_pkl, reg_to_cls
from common.utils import dump_json, load_json

from .config import resolve
from .extract import FileResolver, norm_id, read_label_table
from .media import file_md5

TOL_T = 1e-3  # 时间比较容差（秒；记录中时间保留 4 位小数）


class Report:
    def __init__(self):
        self.checks = []

    def add(self, name: str, ok: bool, detail: str = "", **extra):
        self.checks.append({"check": name, "passed": bool(ok), "detail": detail, **extra})
        return ok

    @property
    def passed(self) -> bool:
        return all(c["passed"] for c in self.checks)


def _nonzero(x: np.ndarray) -> np.ndarray:
    return np.abs(x.astype(np.float32)).max(axis=-1) > ZERO_EPS


def _find_feature(feat_dir: Path, prefix: str) -> Path | None:
    for suffix in (".pkl.gz", ".pkl"):
        c = sorted(feat_dir.glob(f"{prefix}_*{suffix}"))
        if c:
            return c[0]
    return None


def verify(run_dir: Path, video_root: Path | None = None, label_xlsx: Path | None = None) -> Report:
    rep = Report()
    info = load_json(run_dir / "run_info.json")
    cfg = info["config"]
    video_root = video_root or Path(info["inputs"]["video_root"])
    label_xlsx = label_xlsx or Path(info["inputs"]["label_xlsx"])
    Lt, Lav = int(cfg["sequence"]["max_len_text"]), int(cfg["sequence"]["max_len_av"])
    grid = float(cfg["sequence"]["grid_s"])

    # ---- 1 标注表 ↔ 视频文件
    df = read_label_table(label_xlsx, cfg["paths"].get("label_sheet", 0))
    res = FileResolver(video_root)
    exp_ids, exp_paths, missing = [], [], []
    for _, r in df.iterrows():
        vid, cid = norm_id(r["video_id"]), norm_id(r["clip_id"])
        p, msg = res.find(vid, cid)
        if p is None:
            missing.append(f"行{r['row']} {vid}/{cid}: {msg}")
            exp_ids.append(f"{vid}$_${cid}")
        else:
            if msg == "loose":
                vid, cid = p.parent.name, p.stem
            exp_ids.append(f"{vid}$_${cid}")
        exp_paths.append(p)
    all_videos = set(p.resolve() for p in res.all_videos())
    used = set(p.resolve() for p in exp_paths if p is not None)
    extra = sorted(str(p.relative_to(video_root.resolve())) for p in all_videos - used)
    rep.add("标注行→视频文件", not missing, f"{len(df)} 行，{len(df) - len(missing)} 行找到文件" +
            (f"；缺失：{missing[:10]}" if missing else ""))
    rep.add("视频文件→标注行", not extra, f"视频根目录 {len(res.dirs)} 个文件夹、{len(all_videos)} 个视频文件" +
            (f"；未出现在标注表中的文件：{extra[:10]}" if extra else "；全部出现在标注表中"))
    dup = pd.Series(exp_ids)[pd.Series(exp_ids).duplicated()].tolist()
    rep.add("样本编号唯一", not dup, f"重复：{dup}" if dup else f"{len(exp_ids)} 个编号互不相同")

    # ---- 2 特征文件
    feat_dir = run_dir / "features"
    paths = {"aligned": _find_feature(feat_dir, "aligned"), "unaligned": _find_feature(feat_dir, "unaligned")}
    raws = {}
    for v, p in paths.items():
        if rep.add(f"{v} 特征文件存在", p is not None, str(p.name) if p else "未找到"):
            raw = load_pkl(p)
            ok = isinstance(raw, dict) and list(raw.keys()) == ["all"]
            rep.add(f"{v} 顶层结构", ok, f"顶层键 {list(raw.keys()) if isinstance(raw, dict) else type(raw)}")
            raws[v] = raw["all"]
    if len(raws) < 2:
        return rep
    man = pd.read_csv(run_dir / "manifest.csv", dtype={"video_id": str, "clip_id": str})
    N = len(df)
    for v, d in raws.items():
        ids = [str(x) for x in d["id"]]
        rep.add(f"{v} 编号与标注行一一对应（含顺序）", ids == exp_ids,
                f"特征 {len(ids)} 条 / 标注 {N} 行" + ("" if ids == exp_ids else
                                                  f"；首个不一致位置 {next((i for i, (a, b) in enumerate(zip(ids, exp_ids)) if a != b), min(len(ids), N))}"))
    rep.add("manifest 行与标注行一一对应", man["sample_id"].astype(str).tolist() == exp_ids, f"manifest {len(man)} 行")
    al_dir = run_dir / "alignment"
    al_files = {p.stem for p in al_dir.glob("*.json")}
    miss_al = [i for i in exp_ids if i not in al_files]
    extra_al = sorted(al_files - set(exp_ids))
    rep.add("对齐记录与样本一一对应", not miss_al and not extra_al,
            f"{len(al_files)} 个记录文件" + (f"；缺 {miss_al[:5]}" if miss_al else "") + (f"；多 {extra_al[:5]}" if extra_al else ""))

    # ---- 3 形状 / dtype / 有限值
    exp_shape = {"aligned": {"text": (N, Lt, 768), "audio": (N, Lt, 74), "vision": (N, Lt, 35)},
                 "unaligned": {"text": (N, Lt, 768), "audio": (N, Lav, 74), "vision": (N, Lav, 35)}}
    for v, d in raws.items():
        bad = []
        for m, shp in exp_shape[v].items():
            if tuple(d[m].shape) != shp:
                bad.append(f"{m} {d[m].shape}≠{shp}")
        if tuple(d["text_bert"].shape) != (N, 3, Lt) or not np.issubdtype(d["text_bert"].dtype, np.integer):
            bad.append(f"text_bert {d['text_bert'].shape} {d['text_bert'].dtype}")
        for k in ("audio_lengths", "vision_lengths", "regression_labels", "annotations", "classification_labels",
                  "raw_text", "id"):
            if k not in d or len(d[k]) != N:
                bad.append(f"{k} 缺失或长度不符")
        rep.add(f"{v} 字段与形状", not bad, "；".join(bad) if bad else
                f"text{exp_shape[v]['text']} audio{exp_shape[v]['audio']} vision{exp_shape[v]['vision']} "
                f"text_bert{(N, 3, Lt)}，特征 dtype={d['text'].dtype}")
        nonfinite = {m: int((~np.isfinite(d[m].astype(np.float32))).sum()) for m in ("text", "audio", "vision")}
        rep.add(f"{v} 无 NaN/Inf", sum(nonfinite.values()) == 0, str(nonfinite))

    # ---- 4 文本长度 / 掩码 / 填充
    recs = {i: load_json(al_dir / f"{i}.json") for i in exp_ids if i in al_files}
    d = raws["aligned"]
    probs = []
    for n, sid in enumerate(exp_ids):
        tb = d["text_bert"][n]
        mask = tb[1]
        L = int(mask.sum())
        if not (mask[:L] == 1).all() or (mask[L:] != 0).any():
            probs.append(f"{sid}: 掩码不是前缀连续 1")
        if (tb[2] != 0).any():
            probs.append(f"{sid}: token_type_ids 非全 0")
        rec = recs.get(sid)
        if rec is not None and rec["lengths"]["text"] != L:
            probs.append(f"{sid}: 掩码长度 {L} ≠ 记录 text_len {rec['lengths']['text']}")
        if rec is not None and L > 0:
            toks = rec["aligned"]["positions"] if "aligned" in rec else []
            ids_rec = [p["token_id"] for p in toks[:L]]
            if toks and ids_rec != tb[0][:L].tolist():
                probs.append(f"{sid}: input_ids 与对齐记录不一致")
        for v, dd in raws.items():
            nz = _nonzero(dd["text"][n])
            if not (nz[:L].all() and not nz[L:].any()):
                probs.append(f"{sid}/{v}: 文本非零行 ≠ 掩码位置（非零 {int(nz.sum())}，掩码 {L}）")
            if not np.array_equal(dd["text_bert"][n], tb):
                probs.append(f"{sid}: 两个版本的 text_bert 不一致")
    rep.add("文本：掩码长度 = 有效长度 = 非零行（填充行为 0）", not probs, "；".join(probs[:8]) if probs else
            f"{N} 条全部一致；text_len 范围 {int(d['text_bert'][:, 1].sum(1).min())}–{int(d['text_bert'][:, 1].sum(1).max())}")

    # ---- 5 对齐版本长度规则
    probs = []
    for n, sid in enumerate(exp_ids):
        L = int(d["text_bert"][n, 1].sum())
        la, lv = int(d["audio_lengths"][n]), int(d["vision_lengths"][n])
        nza, nzv = _nonzero(d["audio"][n]), _nonzero(d["vision"][n])
        if la not in (0, L) or lv not in (0, L):
            probs.append(f"{sid}: audio/vision_lengths ({la},{lv}) ≠ text_len {L}")
        if nza[la:].any() or nzv[lv:].any():
            probs.append(f"{sid}: 有效长度之外存在非零行")
        if la and not nza[:la].all():
            probs.append(f"{sid}: 语音有效范围内出现 0 行（语音帧不应为 0）")
        rec = recs.get(sid)
        if rec is not None and "aligned" in rec and lv:
            exp_face = [p["pos"] for p in rec["aligned"]["positions"] if p["n_face_frames"] > 0]
            if sorted(np.where(nzv)[0].tolist()) != exp_face:
                probs.append(f"{sid}: 视觉非零位置与记录中含人脸帧的位置不一致")
    rep.add("对齐版：audio/vision 长度 = text_len，语音全有效，视觉 0 行 ⇔ 区间内无人脸帧", not probs,
            "；".join(probs[:8]) if probs else f"{N} 条全部一致")

    # ---- 6 非对齐版本长度规则
    u = raws["unaligned"]
    probs = []
    for n, sid in enumerate(exp_ids):
        la, lv = int(u["audio_lengths"][n]), int(u["vision_lengths"][n])
        nza, nzv = _nonzero(u["audio"][n]), _nonzero(u["vision"][n])
        if la > Lav or lv > Lav:
            probs.append(f"{sid}: 长度超过 {Lav}")
        if nza[la:].any() or nzv[lv:].any():
            probs.append(f"{sid}: 有效长度之外存在非零行")
        if la and not nza[:la].all():
            probs.append(f"{sid}: 语音有效范围内出现 0 行")
        rec = recs.get(sid)
        if rec is not None and "unaligned" in rec:
            ra, rv = rec["unaligned"]["audio"], rec["unaligned"]["vision"]
            if ra["length"] != la or rv["length"] != lv:
                probs.append(f"{sid}: 长度与记录不一致")
            da, dv = rec["durations_s"]["audio"], rec["durations_s"]["video"]
            if la and min(Lav, int(np.ceil(da / grid - 1e-9))) != la:
                probs.append(f"{sid}: audio_lengths {la} ≠ ⌈{da:.3f}/{grid}⌉")
            if lv and min(Lav, int(np.ceil(dv / grid - 1e-9))) != lv:
                probs.append(f"{sid}: vision_lengths {lv} ≠ ⌈{dv:.3f}/{grid}⌉")
            if lv:
                exp_face = [k for k, c in enumerate(rv["n_used"]) if c > 0]
                if np.where(nzv)[0].tolist() != exp_face:
                    probs.append(f"{sid}: 视觉非零位置与记录中含人脸帧的位置不一致")
    rep.add(f"非对齐版：长度 = ⌈时长/{grid}s⌉ ≤ {Lav}，有效范围外全 0，视觉 0 行 ⇔ 栅格内无人脸帧", not probs,
            "；".join(probs[:8]) if probs else
            f"{N} 条全部一致；audio_lengths {int(u['audio_lengths'].min())}–{int(u['audio_lengths'].max())}，"
            f"vision_lengths {int(u['vision_lengths'].min())}–{int(u['vision_lengths'].max())}")

    # ---- 7 时间单调性与范围
    probs = []
    n_words = 0
    for sid, rec in recs.items():
        D = max(rec["durations_s"]["audio"] or 0.0, rec["durations_s"]["video"] or 0.0)
        ws = rec["words"]
        n_words += len(ws)
        if ws and rec["text"]["n_words_total"] not in (None, len(ws)):
            probs.append(f"{sid}: 词数 {len(ws)} ≠ 转写词数 {rec['text']['n_words_total']}")
        for k, w in enumerate(ws):
            if not (-TOL_T <= w["start"] <= w["end"] + TOL_T <= D + 2 * TOL_T):
                probs.append(f"{sid}: 词 {k} 时间 [{w['start']}, {w['end']}] 越界（D={D:.3f}）")
            if k and ws[k - 1]["end"] > w["start"] + TOL_T:
                probs.append(f"{sid}: 词 {k - 1}/{k} 时间不单调")
        if "aligned" in rec:
            prev = -1.0
            for p in rec["aligned"]["positions"]:
                if p["rule"] == "word_piece":
                    if p["start_s"] < prev - TOL_T or p["end_s"] < p["start_s"] - TOL_T or p["end_s"] > D + TOL_T:
                        probs.append(f"{sid}: 位置 {p['pos']} 区间 [{p['start_s']},{p['end_s']}] 不单调或越界")
                    prev = p["end_s"]
    rep.add("词级时间戳与位置区间单调且落在 [0, 时长] 内", not probs,
            "；".join(probs[:8]) if probs else f"{len(recs)} 条记录、{n_words} 个词全部满足")

    # ---- 8 标签 / 文本原样
    probs = []
    y_x = pd.to_numeric(df["label"], errors="coerce").to_numpy(np.float64)
    for v, dd in raws.items():
        y = dd["regression_labels"].astype(np.float64)
        if not np.allclose(y, y_x.astype(np.float32), equal_nan=True, atol=1e-6):
            probs.append(f"{v}: regression_labels 与 label 列不一致")
        ann_x = ["" if not isinstance(a, str) else a for a in df["annotation"]]
        if [str(a) for a in dd["annotations"]] != ann_x:
            probs.append(f"{v}: annotations 与 annotation 列不一致")
        if not np.array_equal(dd["classification_labels"].astype(int), reg_to_cls(np.nan_to_num(y_x))):
            probs.append(f"{v}: classification_labels ≠ reg_to_cls(label)")
        txt_x = ["" if not isinstance(t, str) else t for t in df["text"]]
        if [str(t) for t in dd["raw_text"]] != txt_x:
            probs.append(f"{v}: raw_text 与 text 列不一致")
    ann_bad = int((man["annotation_consistent"].astype(str) == "False").sum())
    rep.add("标签与文本原样保留（未修改）", not probs, "；".join(probs) if probs else
            f"label/annotation/text 与标注表逐条相等；annotation 与 label 推出类别不一致的样本 {ann_bad} 条（原样保留，仅记录）")

    # ---- 9 MD5
    probs = []
    for p, (_, row) in zip(exp_paths, man.iterrows()):
        if p is None:
            continue
        if file_md5(p) != row["md5"]:
            probs.append(f"{row['sample_id']}: MD5 与提取时不同")
    rep.add("视频文件 MD5 与提取时一致", not probs, "；".join(probs[:5]) if probs else f"{sum(p is not None for p in exp_paths)} 个文件")

    # ---- 10 处理状态
    err = man[man["status"] == "error"]
    rep.add("无处理失败样本", len(err) == 0,
            f"error {len(err)} 条：{err['sample_id'].tolist()[:10]}" if len(err) else
            f"状态统计 {man['status'].value_counts().to_dict()}")

    # ---- 11 公共读取接口
    for v, p in paths.items():
        try:
            sds = load_all_splits(p)
            sd = sds["all"]
            ok = len(sd) == N and sd.version == v
            if v == "unaligned":
                ok &= bool(np.array_equal(sd.lengths["audio"], u["audio_lengths"]))
            rep.add(f"common.data.load_all_splits 读取 {v}", ok,
                    f"splits={list(sds)}，n={len(sd)}，version={sd.version}，"
                    f"lengths(mean) T/A/V = {sd.lengths['text'].mean():.1f}/{sd.lengths['audio'].mean():.1f}/"
                    f"{sd.lengths['vision'].mean():.1f}，有效范围内视觉缺失比例 "
                    f"{sd.summary()['vision']['missing_ratio_in_valid']:.3f}")
        except Exception as e:  # noqa: BLE001
            rep.add(f"common.data.load_all_splits 读取 {v}", False, f"{type(e).__name__}: {e}")
    return rep


def write_report(rep: Report, run_dir: Path) -> None:
    out = {"run_dir": str(run_dir), "passed": rep.passed, "n_checks": len(rep.checks),
           "n_failed": sum(not c["passed"] for c in rep.checks), "checks": rep.checks}
    dump_json(out, run_dir / "verification_report.json")
    lines = [f"# 问题1 特征文件核验报告", "", f"- 运行目录：`{run_dir}`",
             f"- 结论：**{'全部通过' if rep.passed else '存在未通过项'}**（{len(rep.checks)} 项，未通过 {out['n_failed']} 项）", "",
             "| # | 核验项 | 结果 | 说明 |", "|---|---|---|---|"]
    for i, c in enumerate(rep.checks, 1):
        detail = str(c["detail"]).replace("|", "/").replace("\n", " ")
        lines.append(f"| {i} | {c['check']} | {'通过' if c['passed'] else '未通过'} | {detail} |")
    (run_dir / "verification_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv=None):
    ap = argparse.ArgumentParser(description="问题1 产物核验")
    ap.add_argument("--run-dir", required=True, help="outputs/q1/<run_name>")
    ap.add_argument("--video-root", default=None, help="默认取 run_info.json 中的路径")
    ap.add_argument("--label-xlsx", default=None)
    a = ap.parse_args(argv)
    run_dir = resolve(a.run_dir)
    rep = verify(run_dir, resolve(a.video_root) if a.video_root else None,
                 resolve(a.label_xlsx) if a.label_xlsx else None)
    write_report(rep, run_dir)
    for c in rep.checks:
        print(f"[{'通过' if c['passed'] else '未通过'}] {c['check']}：{c['detail']}")
    print(f"\n结论：{'全部通过' if rep.passed else '存在未通过项'} → {run_dir / 'verification_report.md'}")
    sys.exit(0 if rep.passed else 1)


if __name__ == "__main__":
    main()
