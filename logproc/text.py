"""判定用テキスト（影テキスト）の作成。原文と文字位置が1対1で対応する。"""
from __future__ import annotations

import unicodedata

# NFKC で数字に化けると区切りの判定ができなくなる丸数字などは残す
_KEEP_RANGES = ((0x2460, 0x24FF), (0x2776, 0x2793))
EMPTY_TOKENS = {"", "-", "－", "ー", "―", "—", "‐", "−", "n/a", "N/A", "なし", "無し", "特になし", "特に無し"}


def clean_log_text(text) -> str:
    """セル値をログ処理用の文字列にする（_x000D_ と CR を除く。他は変えない）。"""
    if text is None:
        return ""
    s = str(text).replace("_x000D_", "")
    return s.replace("\r\n", "\n").replace("\r", "\n")


def shadow(text: str) -> str:
    """1文字ずつ NFKC をかけた写し。結果が1文字にならない文字は元のまま残し、長さを保つ。"""
    out = []
    for ch in text:
        code = ord(ch)
        if any(lo <= code <= hi for lo, hi in _KEEP_RANGES) or code < 0x80:
            out.append(ch)
            continue
        n = unicodedata.normalize("NFKC", ch)
        out.append(n if len(n) == 1 else ch)
    return "".join(out)


def is_empty_log(text: str) -> bool:
    return shadow(text).strip() in EMPTY_TOKENS


def nfkc(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "")
