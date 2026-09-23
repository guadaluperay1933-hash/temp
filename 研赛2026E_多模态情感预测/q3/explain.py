"""问题3 解释计算：模态级精确 Shapley、位置级内置贡献、关键证据片段、积分梯度复核、删除/保留忠实性检验。

全部在"下采样后的整份划分"（engine.PooledSplit，全局位置坐标）上进行。记号：
    f(x) = ŷ（强度输出，模型原始值；AMEE 中 ŷ = Σ_k w_k s_k 天然落在 [−3,3]）
    缺席模态 = 该模态全部位置置为不可用（特征置零）—— 与训练时"模态随机丢弃"完全一致的输入形式。

1. 模态作用程度（主指标）：三人合作博弈的精确 Shapley 值，玩家 N = {T, A, V}，价值函数 v(S) = ŷ(只保留 S 中的模态)
       φ_m = Σ_{S⊆N\\{m}} |S|!(3−|S|−1)!/3! · [v(S∪{m}) − v(S)]      （权重：|S|=0 → 1/3，|S|=1 → 1/6，|S|=2 → 1/3）
   8 个联盟各一次前向即可精确计算；有效性 Σ_m φ_m = v(N) − v(∅) 严格成立。
   作用份额 share_m = |φ_m| / Σ_m' |φ_m'|（Σ|φ|≈0 时取 1/3），主要参考模态 = argmax_m share_m。
   同时对"预测类别的概率" p_ĉ 也算一套 Shapley（φ^prob），供极性判断的解释参考。
2. 内置贡献（精确分解）：c_k = w_k s_k（Σ_k c_k = ŷ），e_{m,t} = w_m a_{m,t} v_{m,t}（Σ_t e_{m,t} = c_m）。
3. 关键证据：按 |e_{m,t}| 排序取候选（文本 top-k 个词元；语音/视觉 k = clip(round(ρ·n), k_min, k_max)），
   把间隔 ≤ merge_gap 的候选合并为连续片段；语音/视觉按片段内 Σ|e| 保留至多 max_segments 段。
4. 复核：积分梯度（基线 = 全零特征，中点黎曼和），逐位置对特征维求和，与 e_{m,t} 算 Spearman；
   删除检验（comprehensiveness）、保留检验（sufficiency）、逐步删除 / 插入曲线与 AOPC，全部与随机位置对照。
"""
from __future__ import annotations

import itertools
import math

import numpy as np
import torch
from scipy import stats

from common.data import MODALITIES

from .engine import PooledSplit, softmax_np

# 8 个联盟：按 (T, A, V) 在场与否编码；名称用于记录
COALITIONS = [tuple(bool(b) for b in bits) for bits in itertools.product([0, 1], repeat=3)]
COALITION_NAMES = ["".join(s for s, k in zip("TAV", c) if k) or "∅" for c in COALITIONS]


def _shap_weight(size: int, n: int = 3) -> float:
    return math.factorial(size) * math.factorial(n - size - 1) / math.factorial(n)


def shapley_from_values(vals: np.ndarray) -> np.ndarray:
    """vals (N, 8)（列顺序同 COALITIONS）→ φ (N, 3)。"""
    col = {c: j for j, c in enumerate(COALITIONS)}
    N = vals.shape[0]
    phi = np.zeros((N, 3), np.float64)
    for m in range(3):
        for c in COALITIONS:
            if c[m]:
                continue
            with_m = tuple(True if j == m else c[j] for j in range(3))
            phi[:, m] += _shap_weight(sum(c)) * (vals[:, col[with_m]] - vals[:, col[c]])
    return phi


def shares_from_phi(phi: np.ndarray) -> np.ndarray:
    a = np.abs(phi)
    tot = a.sum(1, keepdims=True)
    return np.where(tot > 1e-12, a / np.maximum(tot, 1e-12), 1.0 / 3.0)


def keep_modalities(view: dict, keep) -> dict:
    """只保留 keep[m] 为真的模态，其余模态全部位置置为不可用、特征置零。"""
    out = dict(view)
    for m, k in zip(MODALITIES, keep):
        if not k:
            out[f"avail_{m}"] = torch.zeros_like(view[f"avail_{m}"])
            out[f"x_{m}"] = torch.zeros_like(view[f"x_{m}"])
    return out


# ============================================================================ 主解释计算

@torch.no_grad()
def run_explain(model, ps: PooledSplit, device, bias, batch_size: int = 256) -> dict:
    """对整份划分：预测 + 门控/专家输出 + 位置贡献 + 8 个联盟的输出 → Shapley。返回 numpy 组成的 bundle。"""
    model.eval()
    bias = np.asarray(bias, np.float64)
    N = len(ps)
    Lp = {m: int(ps.x[m].shape[1]) for m in MODALITIES}
    has_w = hasattr(model, "experts") and len(getattr(model, "experts", ())) > 0
    K = len(model.experts) if has_w else 0
    b = {"n": N, "ids": list(ps.sd.ids), "experts": list(model.experts) if has_w else [], "strides": dict(ps.strides),
         "orig_len": dict(ps.orig_len), "y_raw": np.zeros(N), "logits": np.zeros((N, 3)),
         "w": np.zeros((N, K)) if has_w else None, "s": np.zeros((N, K)) if has_w else None,
         "present": np.zeros((N, 3), bool), "v_coal": np.zeros((N, 8)), "p_coal": np.zeros((N, 8)),
         "pos": {m: {k: np.zeros((N, Lp[m]), np.float32) for k in ("attn", "value", "contrib", "contrib_cls")}
                 for m in MODALITIES}}
    for m in MODALITIES:
        b["pos"][m]["cand"] = np.zeros((N, Lp[m]), bool)
        b["pos"][m]["valid"] = ps.valid[m].numpy().copy()
        b["pos"][m]["avail"] = ps.avail[m].numpy().copy()
    for idx in ps.chunks(batch_size):
        view = ps.view(idx, device)
        o = model(view, prepared=True)
        logits = o["logits"].double().cpu().numpy()
        c_hat = np.argmax(logits + bias[None, :], axis=1)
        b["y_raw"][idx] = o["y"].double().cpu().numpy()
        b["logits"][idx] = logits
        b["present"][idx] = o["present"].cpu().numpy()
        if has_w:
            w = o["w"]
            b["w"][idx] = w.double().cpu().numpy()
            b["s"][idx] = o["s"].double().cpu().numpy()
        for i, m in enumerate(MODALITIES):
            a = o["attn"][m]
            b["pos"][m]["attn"][idx] = a.float().cpu().numpy()
            b["pos"][m]["cand"][idx] = o["cand"][m].cpu().numpy()
            if has_w:
                v = o["value"][m]
                b["pos"][m]["value"][idx] = v.float().cpu().numpy()
                b["pos"][m]["contrib"][idx] = (w[:, i:i + 1] * a * v).float().cpu().numpy()
                lt = o["pos_logits"][m]                                                     # (B, L, 3)
                ci = torch.as_tensor(c_hat, device=lt.device).view(-1, 1, 1).expand(-1, lt.shape[1], 1)
                b["pos"][m]["contrib_cls"][idx] = (w[:, i:i + 1] * a * lt.gather(2, ci).squeeze(-1)).float().cpu().numpy()
        for j, keep in enumerate(COALITIONS):
            oc = model(keep_modalities(view, keep), prepared=True)
            b["v_coal"][idx, j] = oc["y"].double().cpu().numpy()
            p = softmax_np(oc["logits"].double().cpu().numpy(), bias)
            b["p_coal"][idx, j] = p[np.arange(len(idx)), c_hat]
    b["probs"] = softmax_np(b["logits"], bias)
    b["pred_cls"] = np.argmax(b["logits"] + bias[None, :], axis=1)
    b["phi"] = shapley_from_values(b["v_coal"])
    b["phi_prob"] = shapley_from_values(b["p_coal"])
    b["base"] = b["v_coal"][:, COALITIONS.index((False, False, False))]
    b["share"] = shares_from_phi(b["phi"])
    b["main"] = np.argmax(b["share"], axis=1)
    if has_w:
        b["c"] = b["w"] * b["s"]
    return b


def decomposition_checks(b: dict) -> dict:
    """数值核验三条恒等式（应只有浮点误差）：Σφ = v(N) − v(∅)、Σ_k c_k = ŷ、Σ_t e_{m,t} = c_m。"""
    full = b["v_coal"][:, COALITIONS.index((True, True, True))]
    out = {"shapley_efficiency_max_abs_err": float(np.abs(b["phi"].sum(1) - (full - b["base"])).max()),
           "full_coalition_equals_prediction_max_abs_err": float(np.abs(full - b["y_raw"]).max())}
    if b.get("c") is not None:
        out["expert_sum_equals_prediction_max_abs_err"] = float(np.abs(b["c"].sum(1) - b["y_raw"]).max())
        errs = [np.abs(b["pos"][m]["contrib"].sum(1) - b["c"][:, i]).max() for i, m in enumerate(MODALITIES)]
        out["position_sum_equals_modality_contribution_max_abs_err"] = float(max(errs))
    out["prediction_out_of_[-3,3]_rate"] = float((np.abs(b["y_raw"]) > 3.0 + 1e-9).mean())
    return out


# ============================================================================ 关键证据

def evidence_k(m: str, n: int, ecfg: dict) -> int:
    if n <= 0:
        return 0
    if m == "text":
        return int(min(int(ecfg.get("topk_text", 3)), n))
    k = int(round(float(ecfg.get("av_frac", 0.1)) * n))
    k = max(int(ecfg.get("av_kmin", 2)), min(int(ecfg.get("av_kmax", 12)), k))
    return int(min(k, n))


def extract_evidence(contrib: np.ndarray, cand: np.ndarray, m: str, ecfg: dict) -> dict:
    """单条样本、单个模态：top-k 候选 → 合并为片段。返回 {'positions', 'top', 'segments'}（下采样后的全局位置）。

    片段 seg = {lo, hi（含）, positions（片段内的候选位置）, sum（Σe）, abs（Σ|e|）, peak（|e| 最大的位置）, sign}。
    文本：top-k 词元中相邻者（间隔 0）合并，全部保留；语音/视觉：间隔 ≤ merge_gap 者合并，按 Σ|e| 保留 max_segments 段。
    """
    idx = np.where(cand)[0]
    k = evidence_k(m, idx.size, ecfg)
    if k == 0:
        return {"positions": [], "top": [], "segments": []}
    order = idx[np.argsort(-np.abs(contrib[idx]), kind="stable")]
    top = np.sort(order[:k])
    gap = 0 if m == "text" else int(ecfg.get("merge_gap", 1))
    groups, cur = [], [int(top[0])]
    for p in top[1:]:
        if p - cur[-1] - 1 <= gap:
            cur.append(int(p))
        else:
            groups.append(cur)
            cur = [int(p)]
    groups.append(cur)
    segs = []
    for g in groups:
        lo, hi = g[0], g[-1]
        pos = [int(p) for p in range(lo, hi + 1) if cand[p]] if m != "text" else g
        e = contrib[pos]
        segs.append({"lo": int(lo), "hi": int(hi), "positions": pos, "sum": float(e.sum()),
                     "abs": float(np.abs(e).sum()), "peak": int(pos[int(np.argmax(np.abs(e)))]),
                     "sign": int(np.sign(e.sum()))})
    if m != "text":
        segs = sorted(segs, key=lambda s: -s["abs"])[: int(ecfg.get("max_segments", 2))]
        segs = sorted(segs, key=lambda s: s["lo"])
    positions = sorted({p for s in segs for p in s["positions"]})
    return {"positions": positions, "top": [int(p) for p in top], "segments": segs}


def extract_all_evidence(b: dict, ecfg: dict, key: str = "contrib") -> list[dict]:
    """全部样本 × 三模态的证据。key='contrib'（内置贡献）；concat 等无内置贡献的模型可用 'attn'。"""
    out = []
    for i in range(b["n"]):
        out.append({m: extract_evidence(b["pos"][m][key][i], b["pos"][m]["cand"][i], m, ecfg) for m in MODALITIES})
    return out


def evidence_masks(b: dict, evidence: list[dict]) -> dict:
    masks = {m: np.zeros_like(b["pos"][m]["cand"]) for m in MODALITIES}
    for i, ev in enumerate(evidence):
        for m in MODALITIES:
            if ev[m]["positions"]:
                masks[m][i, ev[m]["positions"]] = True
    return masks


# ============================================================================ 积分梯度

def integrated_gradients(model, ps: PooledSplit, device, steps: int = 32, batch_size: int = 64) -> dict:
    """IG_{m,t} = Σ_d x_{m,t,d} · (1/K) Σ_{k=0}^{K−1} ∂f(α_k x)/∂x_{m,t,d}，α_k = (k+½)/K，基线 x' = 0（掩码不变）。
    返回 {'ig': {m: (N,Lp)}, 'f_x', 'f_0', 'completeness_abs_err'}。"""
    model.eval()
    N = len(ps)
    ig = {m: np.zeros((N, int(ps.x[m].shape[1])), np.float32) for m in MODALITIES}
    f_x, f_0 = np.zeros(N), np.zeros(N)
    for idx in ps.chunks(batch_size):
        view = ps.view(idx, device)
        xs = {m: view[f"x_{m}"] for m in MODALITIES}
        grads = {m: torch.zeros_like(xs[m]) for m in MODALITIES}
        for k in range(steps):
            alpha = (k + 0.5) / steps
            v2 = dict(view)
            inp = []
            for m in MODALITIES:
                t = (alpha * xs[m]).detach().requires_grad_(True)
                v2[f"x_{m}"] = t
                inp.append(t)
            with torch.enable_grad():
                y = model(v2, prepared=True)["y"]
                g = torch.autograd.grad(y.sum(), inp)
            for m, gm in zip(MODALITIES, g):
                grads[m] += gm.detach()
        for m in MODALITIES:
            ig[m][idx] = (xs[m] * grads[m] / steps).sum(-1).float().cpu().numpy()
        with torch.no_grad():
            f_x[idx] = model(view, prepared=True)["y"].double().cpu().numpy()
            v0 = dict(view)
            for m in MODALITIES:
                v0[f"x_{m}"] = torch.zeros_like(xs[m])
            f_0[idx] = model(v0, prepared=True)["y"].double().cpu().numpy()
    tot = sum(ig[m].sum(1) for m in MODALITIES)
    return {"ig": ig, "f_x": f_x, "f_0": f_0, "completeness_abs_err": np.abs(tot - (f_x - f_0))}


def _spearman(a, b) -> float:
    a, b = np.asarray(a, np.float64), np.asarray(b, np.float64)
    if a.size < 3 or np.ptp(a) < 1e-12 or np.ptp(b) < 1e-12:
        return float("nan")
    return float(stats.spearmanr(a, b)[0])


def _summ(x) -> dict:
    x = np.asarray([v for v in x if v is not None and np.isfinite(v)], np.float64)
    if x.size == 0:
        return {"n": 0}
    return {"n": int(x.size), "mean": float(x.mean()), "median": float(np.median(x)), "std": float(x.std()),
            "frac_positive": float((x > 0).mean())}


def ig_consistency(b: dict, igr: dict, evidence: list[dict], ecfg: dict) -> dict:
    """内置贡献 e_{m,t} 与积分梯度 IG_{m,t} 的一致性（逐样本逐模态 Spearman、top-k 重合率），
    以及模态级 IG 总和与 Shapley φ_m / 内置 c_m 的一致性。"""
    out = {"completeness": _summ(igr["completeness_abs_err"]) | {
        "note": "|Σ_{m,t} IG − (f(x) − f(0))|，积分离散化误差；越小说明步数足够"}}
    per_mod, raw = {}, {}
    for i_m, m in enumerate(MODALITIES):
        sp_signed, sp_abs, overlap, overlap_rand = [], [], [], []
        for i in range(b["n"]):
            cand = b["pos"][m]["cand"][i]
            n = int(cand.sum())
            if n < 3:
                continue
            e = b["pos"][m]["contrib"][i][cand]
            g = igr["ig"][m][i][cand]
            sp_signed.append(_spearman(e, g))
            sp_abs.append(_spearman(np.abs(e), np.abs(g)))
            k = evidence_k(m, n, ecfg)
            if 0 < k < n:
                te = set(np.argsort(-np.abs(e), kind="stable")[:k].tolist())
                tg = set(np.argsort(-np.abs(g), kind="stable")[:k].tolist())
                overlap.append(len(te & tg) / k)
                overlap_rand.append(k / n)
        per_mod[m] = {"spearman_signed": _summ(sp_signed), "spearman_abs": _summ(sp_abs),
                      "topk_overlap": _summ(overlap), "topk_overlap_random_expectation": _summ(overlap_rand)}
        raw[m] = sp_signed
    out["position_level"] = per_mod
    out["_raw_spearman_signed"] = raw          # 只供画图，不写入 JSON
    ig_tot = np.stack([igr["ig"][m].sum(1) for m in MODALITIES], axis=1)
    phi = b["phi"]
    out["modality_level"] = {
        "spearman_IGsum_vs_shapley_pooled": _spearman(ig_tot.ravel(), phi.ravel()),
        "sign_agreement_IGsum_vs_shapley": float((np.sign(ig_tot) == np.sign(phi)).mean()),
        "main_modality_agreement_IG_vs_shapley": float((np.argmax(np.abs(ig_tot), 1) == b["main"]).mean()),
    }
    if b.get("c") is not None:
        c = b["c"][:, :3]
        out["modality_level"]["spearman_IGsum_vs_builtin_c_pooled"] = _spearman(ig_tot.ravel(), c.ravel())
    return out


# ============================================================================ 门控 / 内置贡献 与 Shapley 的一致性

def gate_shapley_agreement(b: dict) -> dict:
    if b.get("w") is None:
        return {}
    w = b["w"][:, :3]
    c = np.abs(b["c"][:, :3])
    c_share = c / np.maximum(c.sum(1, keepdims=True), 1e-12)
    sh = b["share"]
    out = {"main_modality_agreement_gate_vs_shapley": float((np.argmax(w, 1) == b["main"]).mean()),
           "main_modality_agreement_contrib_vs_shapley": float((np.argmax(c, 1) == b["main"]).mean()),
           "spearman_gate_vs_shapley_share": {m: _spearman(w[:, i], sh[:, i]) for i, m in enumerate(MODALITIES)},
           "spearman_contrib_share_vs_shapley_share": {m: _spearman(c_share[:, i], sh[:, i])
                                                       for i, m in enumerate(MODALITIES)},
           "spearman_gate_vs_shapley_share_pooled": _spearman(w.ravel(), sh.ravel()),
           "spearman_contrib_share_vs_shapley_share_pooled": _spearman(c_share.ravel(), sh.ravel()),
           "pearson_contrib_vs_shapley_signed_pooled": float(np.corrcoef(b["c"][:, :3].ravel(), b["phi"].ravel())[0, 1])
           if b["n"] > 1 else float("nan")}
    per = [_spearman(c_share[i], sh[i]) if np.ptp(sh[i]) > 1e-9 else np.nan for i in range(b["n"])]
    out["per_sample_spearman_contrib_vs_shapley"] = _summ(per)
    if "interaction" in b["experts"]:
        k = b["experts"].index("interaction")
        tot = np.abs(b["c"]).sum(1)
        out["interaction_contribution_share"] = _summ(np.abs(b["c"][:, k]) / np.maximum(tot, 1e-12))
        out["interaction_gate_weight"] = _summ(b["w"][:, k])
    return out


# ============================================================================ 删除 / 保留检验

@torch.no_grad()
def forward_masked(model, ps: PooledSplit, drop: dict, device, bias, cls_idx: np.ndarray,
                   batch_size: int = 256) -> tuple[np.ndarray, np.ndarray]:
    """drop[m] (N,Lp) 为真的位置置为不可用后的 ŷ 与 p_{cls_idx}。"""
    model.eval()
    ys, ps_ = np.zeros(len(ps)), np.zeros(len(ps))
    for idx in ps.chunks(batch_size):
        d = {m: (drop[m][idx] if m in drop else None) for m in MODALITIES}
        o = model(ps.view(idx, device, d), prepared=True)
        ys[idx] = o["y"].double().cpu().numpy()
        p = softmax_np(o["logits"].double().cpu().numpy(), bias)
        ps_[idx] = p[np.arange(len(idx)), cls_idx[idx]]
    return ys, ps_


def _random_like(cand: np.ndarray, counts: np.ndarray, rng) -> np.ndarray:
    """每行从候选位置里无放回随机取 counts[i] 个。"""
    out = np.zeros_like(cand)
    for i in range(cand.shape[0]):
        idx = np.where(cand[i])[0]
        k = int(min(counts[i], idx.size))
        if k > 0:
            out[i, rng.choice(idx, size=k, replace=False)] = True
    return out


def _ranked_prefix(score: np.ndarray, cand: np.ndarray, frac: float, rows=None) -> np.ndarray:
    """每行候选位置按 score 降序取前 ceil(frac·n) 个。rows：只处理这些行（其余全 False）。"""
    out = np.zeros_like(cand)
    rows = range(cand.shape[0]) if rows is None else rows
    for i in rows:
        idx = np.where(cand[i])[0]
        k = int(math.ceil(frac * idx.size - 1e-9))
        if k > 0:
            out[i, idx[np.argsort(-score[i, idx], kind="stable")[:k]]] = True
    return out


def _paired(ev: np.ndarray, rnd: np.ndarray, sel: np.ndarray, larger_better: bool = True) -> dict:
    ev, rnd = ev[sel], rnd[sel]
    if ev.size == 0:
        return {"n": 0}
    diff = ev - rnd if larger_better else rnd - ev
    res = {"n": int(ev.size), "evidence_mean": float(ev.mean()), "random_mean": float(rnd.mean()),
           "evidence_median": float(np.median(ev)), "random_median": float(np.median(rnd)),
           "win_rate": float((diff > 0).mean()), "ratio_mean": float(ev.mean() / max(rnd.mean(), 1e-12))}
    try:
        res["wilcoxon_p_one_sided"] = float(stats.wilcoxon(diff, alternative="greater").pvalue) \
            if np.any(np.abs(diff) > 0) else 1.0
    except ValueError:
        res["wilcoxon_p_one_sided"] = float("nan")
    return res


def faithfulness(model, ps: PooledSplit, b: dict, evidence: list[dict], device, bias, ecfg: dict,
                 batch_size: int = 256, with_curves: bool = True, logger=None) -> dict:
    """删除 / 保留检验。|Δŷ| = |ŷ(x) − ŷ(x 删除后)|；Δp = p_ĉ(x) − p_ĉ(x 删除后)（ĉ 为原预测类别）。

    comprehensiveness：删掉证据位置 vs 删掉同样数量的随机候选位置（越大越好）；
    sufficiency：每个模态只保留证据位置 vs 只保留同样数量的随机位置（|Δŷ| 越小越好）；
    曲线：按 |e| 降序逐步删除（deletion）/ 只保留前若干（insertion）比例 f 的候选位置，与随机顺序对照；
          AOPC = 各比例点 |Δŷ| 的平均（删除曲线越大越好，插入曲线越小越好）。
    """
    rng = np.random.default_rng(int(ecfg.get("seed", 2026)))
    R = int(ecfg.get("n_random", 5))
    cls_idx = b["pred_cls"]
    y0 = b["y_raw"]
    p0 = b["probs"][np.arange(b["n"]), cls_idx]
    cand = {m: b["pos"][m]["cand"] for m in MODALITIES}
    ev = evidence_masks(b, evidence)
    counts = {m: ev[m].sum(1) for m in MODALITIES}
    res = {"n": b["n"], "n_random": R, "evidence_rule": {k: ecfg.get(k) for k in
                                                          ("topk_text", "av_frac", "av_kmin", "av_kmax", "merge_gap",
                                                           "max_segments")}}

    def run(drop):
        return forward_masked(model, ps, drop, device, bias, cls_idx, batch_size)

    def rand_runs(make_drop):
        ys, pps = [], []
        for _ in range(R):
            y, p = run(make_drop())
            ys.append(np.abs(y0 - y))
            pps.append(p0 - p)
        return np.mean(ys, 0), np.mean(pps, 0)

    # --- 1. 单模态 comprehensiveness
    comp = {}
    for m in MODALITIES:
        sel = (counts[m] > 0) & (cand[m].sum(1) > counts[m])
        y, p = run({m: ev[m]})
        d_ev, dp_ev = np.abs(y0 - y), p0 - p
        d_r, dp_r = rand_runs(lambda m=m: {m: _random_like(cand[m], counts[m], rng)})
        comp[m] = {"abs_dy": _paired(d_ev, d_r, sel), "dprob": _paired(dp_ev, dp_r, sel)}
    # 主要参考模态
    main_drop = {m: np.where((b["main"] == i)[:, None], ev[m], False) for i, m in enumerate(MODALITIES)}
    y, p = run(main_drop)
    d_ev, dp_ev = np.abs(y0 - y), p0 - p
    d_r, dp_r = rand_runs(lambda: {m: np.where((b["main"] == i)[:, None], _random_like(cand[m], counts[m], rng), False)
                                   for i, m in enumerate(MODALITIES)})
    sel = np.array([counts[MODALITIES[k]][i] > 0 and cand[MODALITIES[k]][i].sum() > counts[MODALITIES[k]][i]
                    for i, k in enumerate(b["main"])], bool)
    comp["main_modality"] = {"abs_dy": _paired(d_ev, d_r, sel), "dprob": _paired(dp_ev, dp_r, sel)}
    # 三模态证据一起删
    y, p = run(ev)
    d_all, dp_all = np.abs(y0 - y), p0 - p
    d_r, dp_r = rand_runs(lambda: {m: _random_like(cand[m], counts[m], rng) for m in MODALITIES})
    sel = np.ones(b["n"], bool)
    comp["all_modalities"] = {"abs_dy": _paired(d_all, d_r, sel), "dprob": _paired(dp_all, dp_r, sel)}
    res["comprehensiveness"] = comp
    res["per_sample_deletion_drop"] = d_all
    res["per_sample_deletion_drop_random"] = d_r

    # --- 2. sufficiency：每个模态只保留证据
    keep_drop = {m: cand[m] & ~ev[m] for m in MODALITIES}
    y, p = run(keep_drop)
    s_ev, sp_ev = np.abs(y0 - y), p0 - p
    ys, pps = [], []
    for _ in range(R):
        dr = {m: cand[m] & ~_random_like(cand[m], counts[m], rng) for m in MODALITIES}
        yy, pp = run(dr)
        ys.append(np.abs(y0 - yy))
        pps.append(p0 - pp)
    res["sufficiency"] = {"abs_dy": _paired(s_ev, np.mean(ys, 0), sel, larger_better=False),
                          "dprob": _paired(sp_ev, np.mean(pps, 0), sel, larger_better=False),
                          "note": "只保留证据位置（其余候选位置删除）后的变化，越小越充分；与只保留同样数量随机位置对照"}
    if logger:
        c = comp["all_modalities"]["abs_dy"]
        logger.info(f"[忠实性] 删除全部证据 |Δŷ| 均值 {c['evidence_mean']:.4f} vs 随机 {c['random_mean']:.4f}"
                    f"（胜率 {c['win_rate']:.2f}）；只保留证据 |Δŷ| {res['sufficiency']['abs_dy']['evidence_mean']:.4f}"
                    f" vs 随机 {res['sufficiency']['abs_dy']['random_mean']:.4f}")

    # --- 3. 逐步删除 / 插入曲线
    if with_curves:
        fracs = [float(f) for f in ecfg.get("curve_fracs", [0, .1, .2, .3, .4, .5, .6, .7, .8, .9, 1.0])]
        Rc = int(ecfg.get("curve_random", 3))
        score = {m: np.abs(b["pos"][m]["contrib"]) for m in MODALITIES}
        curves = {}
        for target in list(MODALITIES) + ["main"]:
            rows = {m: (np.where(b["main"] == i)[0] if target == "main" else
                        (np.arange(b["n"]) if m == target else np.array([], int)))
                    for i, m in enumerate(MODALITIES)}
            rscores = [{m: rng.random(score[m].shape) for m in MODALITIES} for _ in range(Rc)]
            cur = {"fracs": fracs, "deletion_evidence": [], "deletion_random": [], "insertion_evidence": [],
                   "insertion_random": [], "deletion_evidence_dprob": [], "deletion_random_dprob": []}
            for f in fracs:
                de = {m: _ranked_prefix(score[m], cand[m], f, rows[m]) for m in MODALITIES}
                y, p = run(de)
                cur["deletion_evidence"].append(float(np.abs(y0 - y).mean()))
                cur["deletion_evidence_dprob"].append(float((p0 - p).mean()))
                ie = {m: np.where(np.isin(np.arange(b["n"]), rows[m])[:, None],
                                  cand[m] & ~_ranked_prefix(score[m], cand[m], f, rows[m]), False) for m in MODALITIES}
                y, _ = run(ie)
                cur["insertion_evidence"].append(float(np.abs(y0 - y).mean()))
                dr, dpr, ir = [], [], []
                for rs in rscores:
                    d = {m: _ranked_prefix(rs[m], cand[m], f, rows[m]) for m in MODALITIES}
                    y, p = run(d)
                    dr.append(np.abs(y0 - y).mean())
                    dpr.append((p0 - p).mean())
                    ins = {m: np.where(np.isin(np.arange(b["n"]), rows[m])[:, None],
                                       cand[m] & ~_ranked_prefix(rs[m], cand[m], f, rows[m]), False) for m in MODALITIES}
                    y, _ = run(ins)
                    ir.append(np.abs(y0 - y).mean())
                cur["deletion_random"].append(float(np.mean(dr)))
                cur["deletion_random_dprob"].append(float(np.mean(dpr)))
                cur["insertion_random"].append(float(np.mean(ir)))
            for k in ("deletion_evidence", "deletion_random", "insertion_evidence", "insertion_random"):
                cur[f"aopc_{k}"] = float(np.mean(cur[k]))
            curves[target] = cur
        res["curves"] = curves
        res["curves_note"] = ("deletion：按 |e| 降序删除该模态前 f 比例的候选位置后的平均 |Δŷ|（越大越忠实）；"
                              "insertion：该模态只保留前 f 比例候选位置时与原预测的平均 |Δŷ|（越小越忠实）；"
                              "random 为随机顺序对照；'main' 表示每条样本只在其主要参考模态上操作")
    return res


# ============================================================================ 汇总统计

def explanation_stats(b: dict, y_cls_true=None) -> dict:
    """模态作用份额分布、主要参考模态分布、按极性分组的作用差异、门控均值。"""
    from common.data import CLASS_NAMES

    sh, main = b["share"], b["main"]
    out = {"n": b["n"],
           "main_modality_counts": {m: int((main == i).sum()) for i, m in enumerate(MODALITIES)},
           "main_modality_rate": {m: float((main == i).mean()) for i, m in enumerate(MODALITIES)},
           "shapley_share_mean": {m: float(sh[:, i].mean()) for i, m in enumerate(MODALITIES)},
           "shapley_share_median": {m: float(np.median(sh[:, i])) for i, m in enumerate(MODALITIES)},
           "shapley_abs_mean": {m: float(np.abs(b["phi"][:, i]).mean()) for i, m in enumerate(MODALITIES)},
           "shapley_signed_mean": {m: float(b["phi"][:, i].mean()) for i, m in enumerate(MODALITIES)},
           "shapley_base_value": float(np.mean(b["base"])),
           "max_share": {"mean": float(sh.max(1).mean()), "median": float(np.median(sh.max(1)))}}
    if b.get("w") is not None:
        out["gate_mean"] = {e: float(b["w"][:, k].mean()) for k, e in enumerate(b["experts"])}
        out["contribution_abs_mean"] = {e: float(np.abs(b["c"][:, k]).mean()) for k, e in enumerate(b["experts"])}
    by_pred = {}
    for c, name in enumerate(CLASS_NAMES):
        s = b["pred_cls"] == c
        if s.any():
            by_pred[name] = {"n": int(s.sum()), "share_mean": {m: float(sh[s, i].mean()) for i, m in enumerate(MODALITIES)},
                             "main_rate": {m: float((main[s] == i).mean()) for i, m in enumerate(MODALITIES)},
                             "phi_mean": {m: float(b["phi"][s, i].mean()) for i, m in enumerate(MODALITIES)}}
    out["by_predicted_class"] = by_pred
    if y_cls_true is not None:
        by_true = {}
        for c, name in enumerate(CLASS_NAMES):
            s = np.asarray(y_cls_true) == c
            if s.any():
                by_true[name] = {"n": int(s.sum()),
                                 "share_mean": {m: float(sh[s, i].mean()) for i, m in enumerate(MODALITIES)},
                                 "main_rate": {m: float((main[s] == i).mean()) for i, m in enumerate(MODALITIES)},
                                 "phi_mean": {m: float(b["phi"][s, i].mean()) for i, m in enumerate(MODALITIES)}}
        out["by_true_class"] = by_true
    return out


def position_profile(b: dict, n_bins: int = 20, key: str = "contrib") -> dict:
    """各模态"平均局部重要性 vs 相对时间"：每条样本把 |e_{m,t}| 在候选位置上归一化为和 1，
    按相对时间 τ = (t − t_first)/(t_last − t_first) 分到 n_bins 个等宽桶再求和，最后对样本平均（按主要参考模态也分组）。"""
    out = {"n_bins": n_bins, "bin_centers": ((np.arange(n_bins) + 0.5) / n_bins).tolist(), "all": {}, "by_main": {}}
    for i_m, m in enumerate(MODALITIES):
        prof = np.zeros((b["n"], n_bins))
        ok = np.zeros(b["n"], bool)
        for i in range(b["n"]):
            valid = b["pos"][m]["valid"][i]
            cand = b["pos"][m]["cand"][i]
            vi = np.where(valid)[0]
            if vi.size < 2 or not cand.any():
                continue
            e = np.abs(b["pos"][m][key][i]) * cand
            if e.sum() <= 0:
                continue
            tau = (np.arange(len(valid)) - vi[0]) / max(vi[-1] - vi[0], 1)
            bins = np.clip((tau * n_bins).astype(int), 0, n_bins - 1)
            prof[i] = np.bincount(bins[cand], weights=e[cand] / e.sum(), minlength=n_bins)
            ok[i] = True
        out["all"][m] = prof[ok].mean(0).tolist() if ok.any() else [0.0] * n_bins
        out["by_main"][m] = {}
        for j, mm in enumerate(MODALITIES):
            s = ok & (b["main"] == j)
            out["by_main"][m][mm] = prof[s].mean(0).tolist() if s.any() else None
    return out
