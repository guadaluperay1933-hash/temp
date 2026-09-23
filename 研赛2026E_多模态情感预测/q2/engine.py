"""问题2 训练 / 推理引擎：数据准备、固定缺失协议、双视图训练、预测、决策层、检查点读写。

要点：
* 标准化器只在训练集可用位置上拟合（common.data.Normalizer），所有划分、附件3 共用。
* 评测时的缺失只以"drop 掩码"的形式给出，模型前向时施加（与训练增广同一代码路径），
  因此同一组掩码可以原封不动地喂给所有模型 —— "对所有模型使用相同掩码"由此保证。
* 朴素插补基线：对缺失视图做逐模态沿时间的线性插值（端点用最近可用行填充），插补后的位置视为可用。
"""
from __future__ import annotations

import copy
import math
import time
from pathlib import Path

import numpy as np
import torch
import torch.nn.functional as F

from common.data import MODALITIES, Normalizer, load_all_splits
from common.decision import evaluate_decisions, tune_neutral_bias
from common.metrics import composite_score, evaluate_all
from common.missing import AugmentConfig, make_drop_masks, protocol_spec, random_drop_for_sample
from common.torch_data import MSADataset, apply_drop_tensors, collate, make_loader, to_device
from common.utils import env_info, get_device, set_seed

from .config import resolve_path
from .model import build_model, count_params

CKPT_FORMAT = "q2-mrgnet-v1"


# ============================================================================ 数据

def prepare_data(cfg: dict, logger=None) -> tuple[dict, Normalizer]:
    """读附件2 → 核对特征版本 → 训练集拟合标准化器 → 变换 train/valid/test。"""
    path = resolve_path(cfg["data"])
    if not path.exists():
        raise FileNotFoundError(f"找不到附件2 特征文件：{path}")
    splits = load_all_splits(path)
    for k in ("train", "valid", "test"):
        if k not in splits:
            raise ValueError(f"{path} 中缺少划分 {k}（现有 {list(splits)}）")
        if not splits[k].has_labels:
            raise ValueError(f"划分 {k} 没有 regression_labels")
    ver = splits["train"].version
    if ver != cfg["version"]:
        raise ValueError(f"配置 version={cfg['version']}，但 {path.name} 是 {ver} 版本；训练/验证/测试必须同一版本")
    norm = Normalizer().fit(splits["train"])
    out = {k: norm.transform(splits[k]) for k in ("train", "valid", "test")}
    if logger:
        for k, sd in out.items():
            s = sd.summary()
            logger.info(f"[数据] {k}: n={s['n']} version={s['version']} 类别={s.get('class_counts')} "
                        + " ".join(f"{m}:L={s[m]['shape'][1]},均长={s[m]['mean_length']:.1f},"
                                   f"天然缺失率={s[m]['missing_ratio_in_valid']:.3f}" for m in MODALITIES))
    return out, norm


def feat_dims_of(sd) -> dict:
    return {m: int(sd.feats[m].shape[-1]) for m in MODALITIES}


def seq_lens_of(sd) -> dict:
    return {m: int(sd.feats[m].shape[1]) for m in MODALITIES}


# ============================================================================ 缺失协议（固定种子 → 所有模型同一掩码）

def mixed_drop_masks(sd, mixed_cfg: dict, seed: int) -> dict:
    """混合缺失协议（仿附件3）：每条样本各模态以 p_modality 概率被选中（至少一个），
    被选模态放 1~max_spans 段、总缺失率 ~ U(ratio_range) 的连续置零区间。复用 common.missing.random_drop_for_sample。"""
    acfg = AugmentConfig(p_sample=1.0, p_modality=float(mixed_cfg.get("p_modality", 0.5)),
                         ratio_range=tuple(mixed_cfg.get("ratio_range", (0.1, 0.7))),
                         max_spans=int(mixed_cfg.get("max_spans", 2)),
                         p_full_drop=float(mixed_cfg.get("p_full_drop", 0.0)), keep_one_intact=False)
    rng = np.random.default_rng(seed)
    drop = {m: np.zeros_like(sd.valid[m]) for m in MODALITIES}
    for i in range(len(sd)):
        d = random_drop_for_sample({m: sd.valid[m][i] for m in MODALITIES}, acfg, rng)
        for m in MODALITIES:
            drop[m][i] = d[m]
    return drop


def protocol_drop_masks(sd, combo: str, ratio: float, position=None, n_spans: int = 1, seed: int = 0) -> dict:
    return make_drop_masks(sd, protocol_spec(combo, ratio, position, n_spans), seed=seed)


def drop_ratios(sd, drop: dict | None) -> dict:
    """每条样本、每个模态在名义有效范围内的"总缺失率"（原本不可用 + 协议置零）。"""
    out = {}
    for m in MODALITIES:
        v = sd.valid[m]
        miss = v & ~sd.avail[m]
        if drop is not None and m in drop:
            miss = miss | (drop[m] & v)
        out[m] = miss.sum(1) / np.maximum(v.sum(1), 1)
    return out


# ============================================================================ Dataset / batch

class FixedDropDataset(MSADataset):
    """drop 掩码由调用方固定给出（评测协议），不做随机增广。"""

    def __init__(self, sd, drop: dict | None):
        super().__init__(sd, augment=None)
        self.fixed = drop

    def __getitem__(self, i: int) -> dict:
        item = super().__getitem__(i)
        if self.fixed is not None:
            for m in MODALITIES:
                if m in self.fixed:
                    item[f"drop_{m}"] = self.fixed[m][i]
        return item


def make_batches(sd, drop: dict | None = None, batch_size: int = 256) -> list[dict]:
    """预先 collate 成 batch 列表（CPU 张量），同一评测条件下多个模型可以复用。"""
    ds = FixedDropDataset(sd, drop)
    return [collate([ds[i] for i in range(s, min(s + batch_size, len(ds)))]) for s in range(0, len(ds), batch_size)]


def complete_view(batch: dict) -> dict:
    return {f"{p}_{m}": batch[f"{p}_{m}"] for p in ("x", "avail", "valid") for m in MODALITIES}


def missing_view(batch: dict) -> dict:
    return apply_drop_tensors(batch)


def interpolate_view(view: dict, modalities=MODALITIES) -> dict:
    """朴素插补：缺失位置（valid & ~avail）用同模态前后最近可用行的线性插值填充，只有一侧可用时取最近行。
    插补后的位置标为可用；某模态有效范围内完全没有可用行时保持缺失。"""
    out = dict(view)
    for m in modalities:
        x, va = view[f"x_{m}"], view[f"valid_{m}"]
        av = view[f"avail_{m}"] & va
        B, L, D = x.shape
        idx = torch.arange(L, device=x.device).unsqueeze(0).expand(B, L)
        left = torch.where(av, idx, torch.full_like(idx, -1)).cummax(dim=1).values
        right = torch.where(av, idx, torch.full_like(idx, L)).flip(1).cummin(dim=1).values.flip(1)
        has_l, has_r = left >= 0, right < L
        l_i, r_i = left.clamp(0, L - 1), right.clamp(0, L - 1)
        xl = x.gather(1, l_i.unsqueeze(-1).expand(B, L, D))
        xr = x.gather(1, r_i.unsqueeze(-1).expand(B, L, D))
        w = (idx - left).float() / (right - left).clamp(min=1).float()
        w = torch.where(has_l & has_r, w, torch.where(has_l, torch.zeros_like(w), torch.ones_like(w)))
        xi = (1 - w).unsqueeze(-1) * xl + w.unsqueeze(-1) * xr
        fill = va & ~av & (has_l | has_r)
        out[f"x_{m}"] = torch.where(fill.unsqueeze(-1), xi, x)
        out[f"avail_{m}"] = av | fill
    return out


@torch.no_grad()
def predict_batches(model, batches: list[dict], device, impute: bool = False) -> dict:
    """在（已施加 drop 的）缺失视图上推理。返回 numpy：logits (N,3)、y_raw (N,)、alpha/rho (N,M)。"""
    model.eval()
    outs = {"logits": [], "y_raw": [], "alpha": [], "rho": [], "index": []}
    for b in batches:
        b = to_device(b, device)
        view = missing_view(b)
        if impute:
            view = interpolate_view(view, model.modalities)
        o = model(view)
        outs["logits"].append(o["logits"].float().cpu())
        outs["y_raw"].append(o["y"].float().cpu())
        outs["rho"].append(o["rho"].float().cpu())
        if o["alpha"] is not None:
            outs["alpha"].append(o["alpha"].float().cpu())
        outs["index"].append(b["index"].cpu())
    res = {k: torch.cat(v).numpy() for k, v in outs.items() if v}
    order = np.argsort(res["index"])
    res = {k: v[order] for k, v in res.items()}
    res["modalities"] = list(model.modalities)
    return res


def predict(model, sd, drop=None, device=None, batch_size: int = 256, impute: bool = False) -> dict:
    device = device or next(model.parameters()).device
    return predict_batches(model, make_batches(sd, drop, batch_size), device, impute=impute)


def raw_metrics(y_true, pred: dict) -> dict:
    """训练期验证用：分类 = logits 的 argmax（不加偏置），强度 = 原始回归值。"""
    return evaluate_all(y_true, np.clip(pred["y_raw"], -3, 3), pred["logits"].argmax(1))


def bias_grid(cfg: dict) -> np.ndarray:
    lo, hi, n = cfg.get("decision", {}).get("bias_grid", [-2.0, 2.0, 41])
    return np.round(np.linspace(float(lo), float(hi), int(n)), 3)


def tune_bias(cfg: dict, y_cls, logits) -> tuple[np.ndarray, dict]:
    return tune_neutral_bias(logits, y_cls, grid=bias_grid(cfg), metric=cfg.get("decision", {}).get("bias_metric",
                                                                                                     "f1_weighted"))


def decision_metrics(y_true, pred: dict, bias) -> dict:
    return evaluate_decisions(y_true, pred["logits"], pred["y_raw"], bias)


# ============================================================================ 训练

def class_weights(y_cls: np.ndarray, mode: str = "inv_sqrt") -> torch.Tensor:
    """w_c ∝ n_c^{-1/2}，再归一化使 Σ_c (n_c/N)·w_c = 1（样本平均权重为 1，不改变损失量级）。"""
    counts = np.bincount(np.asarray(y_cls), minlength=3).astype(np.float64)
    if mode in (None, "none"):
        w = np.ones(3)
    else:
        w = 1.0 / np.sqrt(np.maximum(counts, 1.0))
        w = w / ((counts / counts.sum()) * w).sum()
    return torch.tensor(w, dtype=torch.float32)


def augment_config(cfg: dict) -> AugmentConfig | None:
    a = cfg.get("augment", {})
    if not a.get("enabled", True):
        return None
    return AugmentConfig(p_sample=float(a.get("p_sample", 0.8)), p_modality=float(a.get("p_modality", 0.5)),
                         ratio_range=tuple(a.get("ratio_range", (0.05, 0.9))), max_spans=int(a.get("max_spans", 3)),
                         p_full_drop=float(a.get("p_full_drop", 0.1)),
                         keep_one_intact=bool(a.get("keep_one_intact", False)))


def supervised_loss(out: dict, y_reg, y_cls, cw, lambda_cls: float) -> tuple[torch.Tensor, float, float]:
    l_reg = F.l1_loss(out["y"], y_reg)
    l_cls = F.cross_entropy(out["logits"], y_cls, weight=cw)
    return l_reg + lambda_cls * l_cls, float(l_reg.detach()), float(l_cls.detach())


def reconstruction_loss(out_m: dict, out_c: dict, modalities) -> torch.Tensor | None:
    """只在"完整视图可用、缺失视图不可用"（即原本可用、被增广置零）的位置上算 MSE，
    目标 = 完整视图的标准化特征；逐维平均后对模态取平均。位置与目标都取模型实际使用的分辨率
    （out['view']：长序列经时间下采样后的视图；步长为 1 时就是原始输入）。"""
    rec = out_m["rec"]
    if rec is None:
        return None
    vc, vm = out_c["view"], out_m["view"]
    terms = []
    for m in modalities:
        pos = vc[f"avail_{m}"] & ~vm[f"avail_{m}"] & vc[f"valid_{m}"]
        if pos.any():
            terms.append(((rec[m] - vc[f"x_{m}"].detach()) ** 2)[pos].mean())
    return torch.stack(terms).mean() if terms else None


def kd_loss(student: dict, teacher: dict, sel: torch.Tensor, T: float) -> torch.Tensor | None:
    """一致性蒸馏：教师 = 完整视图（stop-gradient），学生 = 缺失视图；只对确有"原本可用行被置零"的样本计算。"""
    if not sel.any():
        return None
    hs, ht = student["h"][sel], teacher["h"][sel].detach()
    l_h = ((hs - ht) ** 2).mean()
    pt = torch.softmax(teacher["logits"][sel].detach() / T, dim=-1)
    log_ps = torch.log_softmax(student["logits"][sel] / T, dim=-1)
    l_kl = (pt * (torch.log(pt.clamp_min(1e-8)) - log_ps)).sum(-1).mean() * (T ** 2)
    return l_h + l_kl


def _lr_lambda(total_steps: int, warmup_steps: int):
    def f(step):
        if warmup_steps > 0 and step < warmup_steps:
            return (step + 1) / warmup_steps
        p = (step - warmup_steps) / max(1, total_steps - warmup_steps)
        return 0.5 * (1.0 + math.cos(math.pi * min(1.0, p)))
    return f


def train_one(cfg: dict, data: dict, seed: int, device, logger=None) -> dict:
    """训练一个模型（一个变体、一个种子），早停并恢复验证综合分最佳的参数。

    验证综合分 = ½[composite(完整验证集) + composite(固定混合缺失协议下的验证集)]，
    composite = F1_w + Acc3 + Corr − MAE（common.metrics.composite_score）。
    """
    set_seed(seed)
    tr, va = data["train"], data["valid"]
    model = build_model(cfg, feat_dims_of(tr)).to(device)
    mods = model.modalities
    tcfg, lcfg = cfg["train"], cfg["loss"]
    aug = augment_config(cfg)
    loader = make_loader(tr, batch_size=int(tcfg["batch_size"]), shuffle=True, augment=aug, seed=seed)
    opt = torch.optim.AdamW(model.parameters(), lr=float(tcfg["lr"]), weight_decay=float(tcfg["weight_decay"]))
    epochs = int(tcfg["epochs"])
    steps = epochs * len(loader)
    sched_name = tcfg.get("scheduler", "cosine")
    if sched_name == "cosine":
        sched = torch.optim.lr_scheduler.LambdaLR(opt, _lr_lambda(steps, int(tcfg.get("warmup_epochs", 1)) * len(loader)))
    elif sched_name == "plateau":
        sched = torch.optim.lr_scheduler.ReduceLROnPlateau(opt, mode="max", factor=float(tcfg.get("plateau_factor", 0.5)),
                                                           patience=int(tcfg.get("plateau_patience", 3)))
    else:
        sched = None
    cw = class_weights(tr.y_cls, lcfg.get("class_weight", "inv_sqrt")).to(device)
    lam_cls, lam_rec, lam_kd = float(lcfg["lambda_cls"]), float(lcfg["lambda_rec"]), float(lcfg["lambda_kd"])
    T = float(lcfg.get("kd_temperature", 2.0))
    mixed_cfg = cfg["protocol"]["mixed"]
    va_mixed = mixed_drop_masks(va, mixed_cfg, int(mixed_cfg["valid_seed"]))
    eb = int(tcfg.get("eval_batch_size", 256))
    va_batches_c = make_batches(va, None, eb)
    va_batches_m = make_batches(va, va_mixed, eb)

    best = {"score": -np.inf, "epoch": -1, "state": None}
    history, bad = [], 0
    n_params = count_params(model)
    if logger:
        logger.info(f"[{cfg.get('variant', '?')}|seed={seed}] 参数量 {n_params:,}，增广={'开' if aug else '关'}，"
                    f"重构={'开' if getattr(model, 'use_rec', False) else '关'}，门控="
                    f"{'开' if getattr(model, 'use_gate', False) else '关'}，λ=(cls {lam_cls}, rec {lam_rec}, kd {lam_kd})")
    for ep in range(epochs):
        t0 = time.time()
        loader.dataset.set_epoch(ep)
        model.train()
        acc = {"loss": 0.0, "reg_c": 0.0, "cls_c": 0.0, "reg_m": 0.0, "cls_m": 0.0, "rec": 0.0, "kd": 0.0}
        nb = 0
        for batch in loader:
            batch = to_device(batch, device)
            y_reg, y_cls = batch["y_reg"], batch["y_cls"]
            out_c = model(complete_view(batch))
            loss, lr_, lc_ = supervised_loss(out_c, y_reg, y_cls, cw, lam_cls)
            acc["reg_c"] += lr_
            acc["cls_c"] += lc_
            if aug is not None:
                out_m = model(missing_view(batch))
                l_m, lr_, lc_ = supervised_loss(out_m, y_reg, y_cls, cw, lam_cls)
                loss = loss + l_m
                acc["reg_m"] += lr_
                acc["cls_m"] += lc_
                if lam_rec > 0:
                    l_rec = reconstruction_loss(out_m, out_c, mods)
                    if l_rec is not None:
                        loss = loss + lam_rec * l_rec
                        acc["rec"] += float(l_rec.detach())
                if lam_kd > 0:
                    sel = torch.zeros_like(y_cls, dtype=torch.bool)
                    for m in mods:
                        sel = sel | (batch[f"drop_{m}"] & batch[f"avail_{m}"] & batch[f"valid_{m}"]).any(1)
                    l_kd = kd_loss(out_m, out_c, sel, T)
                    if l_kd is not None:
                        loss = loss + lam_kd * l_kd
                        acc["kd"] += float(l_kd.detach())
            opt.zero_grad(set_to_none=True)
            loss.backward()
            if tcfg.get("grad_clip"):
                torch.nn.utils.clip_grad_norm_(model.parameters(), float(tcfg["grad_clip"]))
            opt.step()
            if sched is not None and sched_name == "cosine":
                sched.step()
            acc["loss"] += float(loss.detach())
            nb += 1
        pc = predict_batches(model, va_batches_c, device)
        pm = predict_batches(model, va_batches_m, device)
        mc, mm = raw_metrics(va.y_reg, pc), raw_metrics(va.y_reg, pm)
        sc, sm = composite_score(mc), composite_score(mm)
        score = 0.5 * (sc + sm)
        if sched is not None and sched_name == "plateau":
            sched.step(score)
        row = {"epoch": ep + 1, "lr": opt.param_groups[0]["lr"], **{f"train_{k}": v / max(nb, 1) for k, v in acc.items()},
               "val_score": score, "val_score_complete": sc, "val_score_mixed": sm,
               **{f"val_complete_{k}": mc[k] for k in ("acc3", "f1_weighted", "mae", "corr")},
               **{f"val_mixed_{k}": mm[k] for k in ("acc3", "f1_weighted", "mae", "corr")},
               "time_s": time.time() - t0}
        history.append(row)
        improved = score > best["score"] + 1e-6
        if improved:
            best = {"score": score, "epoch": ep + 1, "state": copy.deepcopy({k: v.detach().cpu()
                                                                              for k, v in model.state_dict().items()})}
            bad = 0
        else:
            bad += 1
        if logger:
            logger.info(f"[{cfg.get('variant', '?')}|seed={seed}] ep{ep + 1:02d} loss={row['train_loss']:.4f} "
                        f"rec={row['train_rec']:.4f} kd={row['train_kd']:.4f} | val 完整 F1={mc['f1_weighted']:.3f} "
                        f"MAE={mc['mae']:.3f} r={mc['corr']:.3f} | 混合 F1={mm['f1_weighted']:.3f} MAE={mm['mae']:.3f} "
                        f"r={mm['corr']:.3f} | score={score:.4f}{' *' if improved else ''} ({row['time_s']:.1f}s)")
        if bad >= int(tcfg.get("patience", 8)):
            if logger:
                logger.info(f"[{cfg.get('variant', '?')}|seed={seed}] 早停于第 {ep + 1} 轮（最佳第 {best['epoch']} 轮）")
            break
    model.load_state_dict(best["state"])
    return {"model": model, "history": history, "best_epoch": best["epoch"], "best_score": float(best["score"]),
            "n_params": n_params, "va_mixed_drop": va_mixed}


# ============================================================================ 检查点

def save_checkpoint(path: str | Path, model, cfg: dict, normalizer: Normalizer, bias, meta: dict) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    torch.save({"format": CKPT_FORMAT, "state_dict": {k: v.detach().cpu() for k, v in model.state_dict().items()},
                "config": cfg, "normalizer": normalizer.state_dict(), "bias": [float(b) for b in bias],
                "version": cfg["version"], "variant": cfg.get("variant"), "modalities": list(model.modalities),
                "feat_dims": {m: int(len(v)) for m, v in normalizer.mean.items()}, "env": env_info(), **meta}, path)


def load_checkpoint(path: str | Path, device=None):
    ck = torch.load(Path(path), map_location="cpu", weights_only=False)
    if ck.get("format") != CKPT_FORMAT:
        raise ValueError(f"{path} 不是问题2 的模型文件")
    model = build_model(ck["config"], ck.get("feat_dims"))
    model.load_state_dict(ck["state_dict"])
    model.to(device or torch.device("cpu")).eval()
    return model, ck


def resolve_device(cfg: dict):
    dev = cfg.get("device", "auto")
    return get_device(None if dev in (None, "auto") else dev)


def setup_threads(cfg: dict) -> None:
    if cfg.get("threads"):
        torch.set_num_threads(int(cfg["threads"]))
