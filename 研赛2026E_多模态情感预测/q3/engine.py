"""问题3 训练 / 推理引擎：数据准备、预下采样的整份划分、模态随机丢弃、损失、训练循环、预测、检查点读写。

要点：
* 标准化器只在训练集可用位置上拟合（common.data.Normalizer），所有划分、附件4 共用。
* 解释、删除检验、积分梯度都在"下采样后的整份划分"（PooledSplit）上做：位置坐标是全局的
  （下采样位置 k ↔ 原始位置 [k·s, (k+1)·s)），不随 batch 的裁剪起点变化，结果与 batch 组成无关。
* 训练时的"模态随机丢弃"：每条样本每个模态以 p 的概率把全部位置置为不可用（不会三个同时丢），
  使 Shapley 计算里"某些模态缺席"的联盟输入都落在训练分布内。
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
from common.missing import AugmentConfig
from common.torch_data import apply_drop_tensors, make_loader, to_device
from common.utils import env_info, get_device, set_seed

from .config import resolve_path
from .model import build_model, count_params, pool_time

CKPT_FORMAT = "q3-amee-v1"


# ============================================================================ 数据

def check_version(sd, version: str, what: str) -> None:
    if sd.version != version:
        raise ValueError(f"配置 version={version}，但 {what} 是 {sd.version} 版本；训练/验证/附件4 必须同一版本")


def prepare_data(cfg: dict, logger=None) -> tuple[dict, Normalizer]:
    """读附件2 → 核对特征版本 → 训练集拟合标准化器 → 变换 train/valid/test（原始未标准化的一份留在 raw 里）。"""
    path = resolve_path(cfg["data"])
    if not path.exists():
        raise FileNotFoundError(f"找不到附件2 特征文件：{path}")
    splits = load_all_splits(path)
    for k in ("train", "valid", "test"):
        if k not in splits:
            raise ValueError(f"{path} 中缺少划分 {k}（现有 {list(splits)}）")
        if not splits[k].has_labels:
            raise ValueError(f"划分 {k} 没有 regression_labels")
    check_version(splits["train"], cfg["version"], path.name)
    norm = Normalizer().fit(splits["train"])
    out = {k: norm.transform(splits[k]) for k in ("train", "valid", "test")}
    if logger:
        for k, sd in out.items():
            s = sd.summary()
            logger.info(f"[数据] {k}: n={s['n']} version={s['version']} 类别={s.get('class_counts')} "
                        + " ".join(f"{m}:L={s[m]['shape'][1]},均长={s[m]['mean_length']:.1f},"
                                   f"天然不可用率={s[m]['missing_ratio_in_valid']:.3f}" for m in MODALITIES))
    return out, norm


def feat_dims_of(sd) -> dict:
    return {m: int(sd.feats[m].shape[-1]) for m in MODALITIES}


def seq_lens_of(sd) -> dict:
    return {m: int(sd.feats[m].shape[1]) for m in MODALITIES}


class PooledSplit:
    """把一整份 SplitData 按模型的步长下采样好（全局坐标，不裁剪），供推理 / 解释反复使用。

    x[m] (N, Lp_m, D_m) float、avail[m] / valid[m] (N, Lp_m) bool，均为 CPU 张量；
    orig_range(m, k) = 原始位置 [k·s_m, min((k+1)·s_m, L_m))。
    """

    def __init__(self, sd, strides: dict, chunk: int = 512):
        self.sd = sd
        self.n = len(sd)
        self.strides = {m: int(strides.get(m, 1)) for m in MODALITIES}
        self.orig_len = {m: int(sd.feats[m].shape[1]) for m in MODALITIES}
        self.x, self.avail, self.valid = {}, {}, {}
        for m in MODALITIES:
            xs, av, va = [], [], []
            for s in range(0, self.n, chunk):
                x = torch.from_numpy(sd.feats[m][s:s + chunk]).float()
                a = torch.from_numpy(sd.avail[m][s:s + chunk])
                v = torch.from_numpy(sd.valid[m][s:s + chunk])
                x, a, v = pool_time(x, a & v, v, self.strides[m])
                xs.append(x)
                av.append(a & v)
                va.append(v)
            self.x[m], self.avail[m], self.valid[m] = torch.cat(xs), torch.cat(av), torch.cat(va)

    def __len__(self) -> int:
        return self.n

    def view(self, idx, device=None, drop: dict | None = None) -> dict:
        """取 idx 这些样本的（已下采样）视图；drop[m] (len(idx), Lp) 为真的位置置为不可用、特征置零。"""
        idx = torch.as_tensor(np.asarray(idx), dtype=torch.long)
        out = {}
        for m in MODALITIES:
            x, a = self.x[m][idx], self.avail[m][idx]
            if drop is not None and m in drop and drop[m] is not None:
                d = torch.as_tensor(drop[m], dtype=torch.bool)
                a = a & ~d
                x = x.masked_fill(d.unsqueeze(-1), 0.0)
            out[f"x_{m}"], out[f"avail_{m}"], out[f"valid_{m}"] = x, a, self.valid[m][idx]
        if device is not None:
            out = {k: v.to(device) for k, v in out.items()}
        return out

    def orig_range(self, m: str, k: int) -> tuple[int, int]:
        s = self.strides[m]
        return k * s, min((k + 1) * s, self.orig_len[m])

    def chunks(self, batch_size: int):
        for s in range(0, self.n, batch_size):
            yield np.arange(s, min(s + batch_size, self.n))


def model_strides(model) -> dict:
    return dict(getattr(model, "strides", {m: 1 for m in MODALITIES}))


# ============================================================================ 训练期视图

def align_batch_offsets(batch: dict, strides: dict) -> dict:
    """common.torch_data.collate 会把序列裁到 batch 实际用到的范围 [lo, hi)；lo 不是步长整数倍时在前面补空位置，
    使下采样窗口与推理时（全局坐标、从 0 开始分窗）一致。后填充的数据 lo 恒为 0，不受影响。"""
    for m in MODALITIES:
        s = int(strides.get(m, 1))
        lo = int(batch.get(f"offset_{m}", 0))
        r = lo % s if s > 1 else 0
        if r:
            batch[f"x_{m}"] = F.pad(batch[f"x_{m}"], (0, 0, r, 0))
            for k in ("avail", "valid", "drop"):
                batch[f"{k}_{m}"] = F.pad(batch[f"{k}_{m}"], (r, 0))
            batch[f"offset_{m}"] = lo - r
    return batch


def modality_dropout(view: dict, p: float, gen: torch.Generator) -> tuple[dict, torch.Tensor]:
    """每条样本每个模态以概率 p 整段置为不可用；本来就有可用位置的模态至少留一个（不会三个同时丢）。"""
    has = torch.stack([view[f"avail_{m}"].any(1) for m in MODALITIES], dim=1).cpu()
    B = has.shape[0]
    if p <= 0:
        return view, torch.zeros(B, len(MODALITIES), dtype=torch.bool)
    drop = (torch.rand(B, len(MODALITIES), generator=gen) < p) & has
    bad = has.any(1) & ~(has & ~drop).any(1)
    if bad.any():
        k = (torch.rand(B, len(MODALITIES), generator=gen) * has.float()).argmax(1)
        drop[bad, k[bad]] = False
    out = dict(view)
    for i, m in enumerate(MODALITIES):
        d = drop[:, i].to(view[f"avail_{m}"].device)
        if d.any():
            out[f"avail_{m}"] = view[f"avail_{m}"] & ~d.unsqueeze(1)
            out[f"x_{m}"] = view[f"x_{m}"].masked_fill(d.view(-1, 1, 1), 0.0)
    return out, drop


# ============================================================================ 损失

def class_weights(y_cls: np.ndarray, mode: str = "inv_sqrt") -> torch.Tensor:
    """w_c ∝ n_c^{-1/2}，再归一化使 Σ_c (n_c/N)·w_c = 1（样本平均权重为 1，不改变损失量级）。"""
    counts = np.bincount(np.asarray(y_cls), minlength=3).astype(np.float64)
    if mode in (None, "none"):
        w = np.ones(3)
    else:
        w = 1.0 / np.sqrt(np.maximum(counts, 1.0))
        w = w / ((counts / counts.sum()) * w).sum()
    return torch.tensor(w, dtype=torch.float32)


def attention_entropy(out: dict) -> torch.Tensor | None:
    """mean_m 归一化熵 H(a_m)/log|C_m|，只对候选位置 ≥ 2 的（样本, 模态）计算，先对样本平均再对模态平均。"""
    terms = []
    for m in MODALITIES:
        a, cand = out["attn"][m], out["cand"][m]
        n = cand.sum(1)
        sel = n >= 2
        if sel.any():
            h = -(a * torch.log(a.clamp_min(1e-12))).sum(1)
            terms.append((h[sel] / torch.log(n[sel].float())).mean())
    return torch.stack(terms).mean() if terms else None


def attention_tv(out: dict, modalities=("audio", "vision")) -> torch.Tensor | None:
    """语音/视觉：½ Σ_t |a_{t+1} − a_t|（只计相邻两位置都是候选的差分），对在场样本平均、再对模态平均。∈ [0,1]。"""
    terms = []
    for m in modalities:
        a, cand = out["attn"][m], out["cand"][m]
        pres = cand.any(1)
        if not pres.any() or a.shape[1] < 2:
            continue
        pair = (cand[:, 1:] & cand[:, :-1]).float()
        tv = 0.5 * ((a[:, 1:] - a[:, :-1]).abs() * pair).sum(1)
        terms.append(tv[pres].mean())
    return torch.stack(terms).mean() if terms else None


def compute_loss(model, out: dict, y_reg, y_cls, cw, lcfg: dict) -> tuple[torch.Tensor, dict]:
    """L = L1(ŷ,y) + λ_cls·CE_w(ℓ,c) + λ_uni·Σ_m L1(s_m,y) + λ_ent·Ent + λ_tv·TV + λ_null·[L1(ŷ(∅),y) + λ_cls·CE_w(ℓ(∅),c)]。
    concat 对照模型只有前两项。"""
    lam_cls = float(lcfg.get("lambda_cls", 0.5))
    l_reg = F.l1_loss(out["y"], y_reg)
    l_cls = F.cross_entropy(out["logits"], y_cls, weight=cw)
    loss = l_reg + lam_cls * l_cls
    parts = {"reg": float(l_reg.detach()), "cls": float(l_cls.detach())}
    if out.get("s") is None:  # concat
        return loss, parts
    lam_uni, lam_ent = float(lcfg.get("lambda_uni", 0.0)), float(lcfg.get("lambda_ent", 0.0))
    lam_tv, lam_null = float(lcfg.get("lambda_tv", 0.0)), float(lcfg.get("lambda_null", 0.0))
    if lam_uni > 0:
        uni = []
        for i, _m in enumerate(MODALITIES):
            pres = out["present"][:, i]
            if pres.any():
                uni.append(F.l1_loss(out["s"][pres, i], y_reg[pres]))
        if uni:
            l_uni = torch.stack(uni).sum()
            loss = loss + lam_uni * l_uni
            parts["uni"] = float(l_uni.detach())
    if lam_ent > 0:
        l_ent = attention_entropy(out)
        if l_ent is not None:
            loss = loss + lam_ent * l_ent
            parts["ent"] = float(l_ent.detach())
    if lam_tv > 0:
        l_tv = attention_tv(out)
        if l_tv is not None:
            loss = loss + lam_tv * l_tv
            parts["tv"] = float(l_tv.detach())
    if lam_null > 0 and getattr(model, "use_interaction", False):
        s0, l0 = model.null_output(y_reg.device)
        B = y_reg.shape[0]
        l_null = F.l1_loss(s0.expand(B), y_reg) + lam_cls * F.cross_entropy(l0.expand(B, 3), y_cls, weight=cw)
        loss = loss + lam_null * l_null
        parts["null"] = float(l_null.detach())
    return loss, parts


# ============================================================================ 预测

@torch.no_grad()
def predict_pooled(model, ps: PooledSplit, device, batch_size: int = 256, drop: dict | None = None) -> dict:
    """在（已下采样的）整份划分上推理。返回 numpy：y_raw (N,)、logits (N,3)、w (N,K) 或 None。"""
    model.eval()
    ys, ls, ws = [], [], []
    for idx in ps.chunks(batch_size):
        d = {m: drop[m][idx] for m in drop} if drop is not None else None
        o = model(ps.view(idx, device, d), prepared=True)
        ys.append(o["y"].float().cpu())
        ls.append(o["logits"].float().cpu())
        if o.get("w") is not None:
            ws.append(o["w"].float().cpu())
    return {"y_raw": torch.cat(ys).numpy(), "logits": torch.cat(ls).numpy(),
            "w": torch.cat(ws).numpy() if ws else None}


def raw_metrics(y_true, pred: dict) -> dict:
    """训练期验证用：分类 = logits 的 argmax（不加偏置），强度 = 原始回归值（裁剪到 [-3,3]）。"""
    return evaluate_all(y_true, np.clip(pred["y_raw"], -3, 3), pred["logits"].argmax(1))


def bias_grid(cfg: dict) -> np.ndarray:
    lo, hi, n = cfg.get("decision", {}).get("bias_grid", [-2.0, 2.0, 41])
    return np.round(np.linspace(float(lo), float(hi), int(n)), 3)


def tune_bias(cfg: dict, y_cls, logits) -> tuple[np.ndarray, dict]:
    return tune_neutral_bias(logits, y_cls, grid=bias_grid(cfg),
                             metric=cfg.get("decision", {}).get("bias_metric", "f1_weighted"))


def decision_metrics(y_true, pred: dict, bias) -> dict:
    return evaluate_decisions(y_true, pred["logits"], pred["y_raw"], bias)


def softmax_np(logits: np.ndarray, bias=None) -> np.ndarray:
    z = np.asarray(logits, np.float64) + (np.asarray(bias)[None, :] if bias is not None else 0.0)
    z = z - z.max(1, keepdims=True)
    e = np.exp(z)
    return e / e.sum(1, keepdims=True)


# ============================================================================ 训练

def span_drop_config(cfg: dict) -> AugmentConfig | None:
    a = cfg["train"].get("span_drop", {}) or {}
    if not a.get("enabled", False):
        return None
    return AugmentConfig(p_sample=float(a.get("p_sample", 0.2)), p_modality=float(a.get("p_modality", 0.5)),
                         ratio_range=tuple(a.get("ratio_range", (0.05, 0.3))), max_spans=int(a.get("max_spans", 3)),
                         p_full_drop=float(a.get("p_full_drop", 0.0)), keep_one_intact=True)


def _lr_lambda(total_steps: int, warmup_steps: int):
    def f(step):
        if warmup_steps > 0 and step < warmup_steps:
            return (step + 1) / warmup_steps
        p = (step - warmup_steps) / max(1, total_steps - warmup_steps)
        return 0.5 * (1.0 + math.cos(math.pi * min(1.0, p)))
    return f


def train_one(cfg: dict, data: dict, seed: int, device, logger=None, pooled_valid: PooledSplit | None = None) -> dict:
    """训练一个模型（一个变体、一个种子），按验证集综合分 composite = F1_w + Acc3 + Corr − MAE 早停。"""
    set_seed(seed)
    tr, va = data["train"], data["valid"]
    model = build_model(cfg, feat_dims_of(tr)).to(device)
    strides = model_strides(model)
    pv = pooled_valid if (pooled_valid is not None and pooled_valid.strides == strides) else PooledSplit(va, strides)
    tcfg, lcfg = cfg["train"], cfg["loss"]
    aug = span_drop_config(cfg)
    loader = make_loader(tr, batch_size=int(tcfg["batch_size"]), shuffle=True, augment=aug, seed=seed)
    opt = torch.optim.AdamW(model.parameters(), lr=float(tcfg["lr"]), weight_decay=float(tcfg["weight_decay"]))
    epochs = int(tcfg["epochs"])
    sched = torch.optim.lr_scheduler.LambdaLR(opt, _lr_lambda(epochs * len(loader),
                                                              int(tcfg.get("warmup_epochs", 1)) * len(loader)))
    cw = class_weights(tr.y_cls, lcfg.get("class_weight", "inv_sqrt")).to(device)
    p_md = float(tcfg.get("modality_dropout", 0.15))
    gen = torch.Generator().manual_seed(int(seed) * 7919 + 17)
    eb = int(tcfg.get("eval_batch_size", 256))
    n_params = count_params(model)
    if logger:
        logger.info(f"[{cfg.get('variant', '?')}|seed={seed}] 结构={cfg['model'].get('arch', 'amee')} 参数量 {n_params:,}，"
                    f"交互专家={'开' if getattr(model, 'use_interaction', False) else '关'}，模态丢弃 p={p_md}，"
                    f"区间置零增广={'开' if aug else '关'}，步长={strides}，λ={ {k: v for k, v in lcfg.items()} }")
    best = {"score": -np.inf, "epoch": -1, "state": None}
    history, bad = [], 0
    for ep in range(epochs):
        t0 = time.time()
        loader.dataset.set_epoch(ep)
        model.train()
        acc: dict = {}
        nb = 0
        for batch in loader:
            batch = align_batch_offsets(batch, strides)
            batch = to_device(batch, device)
            view = apply_drop_tensors(batch)
            view, _ = modality_dropout(view, p_md, gen)
            out = model(view)
            loss, parts = compute_loss(model, out, batch["y_reg"], batch["y_cls"], cw, lcfg)
            opt.zero_grad(set_to_none=True)
            loss.backward()
            if tcfg.get("grad_clip"):
                torch.nn.utils.clip_grad_norm_(model.parameters(), float(tcfg["grad_clip"]))
            opt.step()
            sched.step()
            acc["loss"] = acc.get("loss", 0.0) + float(loss.detach())
            for k, v in parts.items():
                acc[k] = acc.get(k, 0.0) + v
            nb += 1
        pred = predict_pooled(model, pv, device, eb)
        mv = raw_metrics(va.y_reg, pred)
        score = composite_score(mv)
        row = {"epoch": ep + 1, "lr": opt.param_groups[0]["lr"],
               **{f"train_{k}": v / max(nb, 1) for k, v in acc.items()},
               "val_score": score, **{f"val_{k}": mv[k] for k in ("acc3", "f1_weighted", "f1_macro", "mae", "corr")},
               "time_s": time.time() - t0}
        if pred["w"] is not None:
            for k, name in enumerate(model.experts):
                row[f"val_gate_{name}"] = float(pred["w"][:, k].mean())
        history.append(row)
        improved = score > best["score"] + 1e-6
        if improved:
            best = {"score": score, "epoch": ep + 1,
                    "state": copy.deepcopy({k: v.detach().cpu() for k, v in model.state_dict().items()})}
            bad = 0
        else:
            bad += 1
        if logger:
            gate = " ".join(f"{k[9:]}={v:.2f}" for k, v in row.items() if k.startswith("val_gate_"))
            logger.info(f"[{cfg.get('variant', '?')}|seed={seed}] ep{ep + 1:02d} loss={row['train_loss']:.4f} "
                        f"| val Acc={mv['acc3']:.3f} F1={mv['f1_weighted']:.3f} MAE={mv['mae']:.3f} r={mv['corr']:.3f} "
                        f"score={score:.4f}{' *' if improved else ''} | 门控均值 {gate} ({row['time_s']:.1f}s)")
        if bad >= int(tcfg.get("patience", 8)):
            if logger:
                logger.info(f"[{cfg.get('variant', '?')}|seed={seed}] 早停于第 {ep + 1} 轮（最佳第 {best['epoch']} 轮）")
            break
    model.load_state_dict(best["state"])
    model.eval()
    return {"model": model, "history": history, "best_epoch": best["epoch"], "best_score": float(best["score"]),
            "n_params": n_params}


# ============================================================================ 检查点

def save_checkpoint(path: str | Path, model, cfg: dict, normalizer: Normalizer, bias, meta: dict) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    torch.save({"format": CKPT_FORMAT, "state_dict": {k: v.detach().cpu() for k, v in model.state_dict().items()},
                "config": cfg, "normalizer": normalizer.state_dict(), "bias": [float(b) for b in bias],
                "version": cfg["version"], "variant": cfg.get("variant"), "arch": cfg["model"].get("arch", "amee"),
                "experts": list(getattr(model, "experts", ())), "strides": model_strides(model),
                "feat_dims": {m: int(len(v)) for m, v in normalizer.mean.items()}, "env": env_info(), **meta}, path)


def load_checkpoint(path: str | Path, device=None):
    ck = torch.load(Path(path), map_location="cpu", weights_only=False)
    if ck.get("format") != CKPT_FORMAT:
        raise ValueError(f"{path} 不是问题3 的模型文件")
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
