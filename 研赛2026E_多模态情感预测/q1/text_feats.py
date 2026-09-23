"""文本模态：给定转写文本 → BERT 词元序列（50 个位置）+ 每个位置到原文"词"的映射 + 768 维上下文特征。

与附件2 的约定保持一致：
    text_bert (3, 50) = [input_ids; attention_mask; token_type_ids]，后填充；
    text (50, 768)    = BERT 最后一层隐状态，[CLS]…[SEP] 位置有值，填充位置整行为 0。

"词"的定义：转写文本按空白切分得到的片段（与 align.align_words 的词表完全相同，下标一一对应）。
标点粘在词上（如 "movie,"），BERT 会把它切成独立词元，但它仍映射回原来那个词 → 共用该词的时间区间。

截断规则：只保留完整的词。设词 w_j 切成 P_j 个词元，保留最大的 k 使 2 + Σ_{j<k} P_j ≤ max_len，
序列为 [CLS] w_0…w_{k-1} 的全部词元 [SEP]，其余词整体丢弃并记录丢弃的词数/词元数
（不把一个词切一半，是为了保证每个保留词元都能对应到完整的词级时间区间）。

下载说明：默认模型 bert-base-uncased（HuggingFace）。国内网络可先设置镜像
    export HF_ENDPOINT=https://hf-mirror.com    （Windows PowerShell: $env:HF_ENDPOINT="https://hf-mirror.com"）
或把模型目录下载到本地后在 configs/q1.yaml 的 text.model 填本地路径。
"""
from __future__ import annotations

import numpy as np

DEFAULT_TEXT_MODEL = "bert-base-uncased"


def split_words(transcript) -> list[str]:
    """转写文本 → 词列表（按空白切分）。NaN / None → 空列表。"""
    if transcript is None:
        return []
    if isinstance(transcript, float) and not np.isfinite(transcript):
        return []
    return str(transcript).split()


def tokenize_with_words(transcript, tokenizer, max_len: int = 50) -> dict:
    """带词映射的 BERT 分词。

    返回 dict：
        input_ids / attention_mask / token_type_ids : int64 (max_len,)
        word_index  : int64 (max_len,)  该位置所属词的下标；[CLS]/[SEP]/填充 为 -1
        piece_rank  : int64 (max_len,)  该词元在所属词内的序号（0 起）；特殊位置 -1
        piece_count : int64 (max_len,)  所属词的词元总数；特殊位置 0
        tokens      : list[str] 长 max_len（填充位置为 '[PAD]'）
        words       : 全部词（未截断）；pieces_per_word : 每个词的词元数（含被截断的词）
        text_len    : 有效长度（含 [CLS]、[SEP]）
        n_words_total / n_words_kept / truncated_words / truncated_tokens
    """
    words = split_words(transcript)
    pieces = [tokenizer.tokenize(w) for w in words]
    budget = max_len - 2
    kept, used = 0, 0
    for p in pieces:
        if used + len(p) > budget:
            break
        used += len(p)
        kept += 1
    cls_tok = tokenizer.cls_token or "[CLS]"
    sep_tok = tokenizer.sep_token or "[SEP]"
    pad_tok = tokenizer.pad_token or "[PAD]"
    tokens = [cls_tok]
    word_index, piece_rank, piece_count = [-1], [-1], [0]
    for j in range(kept):
        for r, tok in enumerate(pieces[j]):
            tokens.append(tok)
            word_index.append(j)
            piece_rank.append(r)
            piece_count.append(len(pieces[j]))
    tokens.append(sep_tok)
    word_index.append(-1)
    piece_rank.append(-1)
    piece_count.append(0)
    L = len(tokens)
    n_pad = max_len - L
    ids = tokenizer.convert_tokens_to_ids(tokens)
    pad_id = tokenizer.pad_token_id if tokenizer.pad_token_id is not None else 0
    out = {
        "input_ids": np.array(ids + [pad_id] * n_pad, dtype=np.int64),
        "attention_mask": np.array([1] * L + [0] * n_pad, dtype=np.int64),
        "token_type_ids": np.zeros(max_len, dtype=np.int64),
        "word_index": np.array(word_index + [-1] * n_pad, dtype=np.int64),
        "piece_rank": np.array(piece_rank + [-1] * n_pad, dtype=np.int64),
        "piece_count": np.array(piece_count + [0] * n_pad, dtype=np.int64),
        "tokens": tokens + [pad_tok] * n_pad,
        "words": words,
        "pieces_per_word": [len(p) for p in pieces],
        "text_len": L,
        "n_words_total": len(words),
        "n_words_kept": kept,
        "truncated_words": len(words) - kept,
        "truncated_tokens": int(sum(len(p) for p in pieces[kept:])),
    }
    return out


def text_bert_array(tok: dict) -> np.ndarray:
    """(3, max_len) int64：input_ids / attention_mask / token_type_ids（附件2 的 text_bert 格式）。"""
    return np.stack([tok["input_ids"], tok["attention_mask"], tok["token_type_ids"]]).astype(np.int64)


class TextEncoder:
    """BERT 编码器：懒加载，推理模式（eval + no_grad），逐批前向，填充位置置零。"""

    def __init__(self, model_name: str = DEFAULT_TEXT_MODEL, device=None, max_len: int = 50):
        from transformers import AutoModel, AutoTokenizer

        from common.utils import get_device

        self.model_name = model_name
        self.max_len = max_len
        self.device = get_device(device)
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
        self.model = AutoModel.from_pretrained(model_name)
        self.model.eval().to(self.device)
        self.hidden_size = int(self.model.config.hidden_size)

    def info(self) -> dict:
        cfg = self.model.config
        return {"model": self.model_name, "name_or_path": getattr(cfg, "_name_or_path", None),
                "commit_hash": getattr(cfg, "_commit_hash", None), "hidden_size": self.hidden_size,
                "num_hidden_layers": getattr(cfg, "num_hidden_layers", None),
                "vocab_size": getattr(cfg, "vocab_size", None), "tokenizer": type(self.tokenizer).__name__,
                "do_lower_case": getattr(self.tokenizer, "do_lower_case", None), "layer": "last_hidden_state",
                "device": str(self.device)}

    def tokenize(self, transcript) -> dict:
        return tokenize_with_words(transcript, self.tokenizer, self.max_len)

    def encode(self, toks: list[dict], batch_size: int = 16) -> np.ndarray:
        """toks: tokenize() 结果列表 → float32 (n, max_len, hidden)，填充位置整行为 0。"""
        import torch

        out = []
        with torch.no_grad():
            for b in range(0, len(toks), batch_size):
                chunk = toks[b:b + batch_size]
                ids = torch.as_tensor(np.stack([t["input_ids"] for t in chunk]), device=self.device)
                mask = torch.as_tensor(np.stack([t["attention_mask"] for t in chunk]), device=self.device)
                seg = torch.as_tensor(np.stack([t["token_type_ids"] for t in chunk]), device=self.device)
                h = self.model(input_ids=ids, attention_mask=mask, token_type_ids=seg).last_hidden_state
                h = h * mask.unsqueeze(-1).to(h.dtype)
                out.append(h.float().cpu().numpy())
        if not out:
            return np.zeros((0, self.max_len, self.hidden_size), np.float32)
        return np.concatenate(out, axis=0).astype(np.float32)
