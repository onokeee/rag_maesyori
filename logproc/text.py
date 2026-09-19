"""判定用テキスト（影テキスト）の作成。原文と文字位置が1対1で対応する。"""
from __future__ import annotations

import unicodedata
from functools import lru_cache

# NFKC で数字に化けると区切りの判定ができなくなる丸数字などは残す
_KEEP_RANGES = ((0x2460, 0x24FF), (0x2776, 0x2793))
EMPTY_TOKENS = {"", "-", "－", "ー", "―", "—", "‐", "−", "n/a", "N/A", "なし", "無し", "特になし", "特に無し"}


def clean_log_text(text) -> str:
    """セル値をログ処理用の文字列にする（_x000D_ と CR を除く。他は変えない）。"""
    if text is None:
        return ""
    s = str(text).replace("_x000D_", "")
    return s.replace("\r\n", "\n").replace("\r", "\n")


class _ShadowTable(dict):
    """str.translate 用の1文字の写しの表。初めて出た文字だけ NFKC を計算して覚える（文字の種類は有限）。"""

    def __missing__(self, code: int) -> str:
        if code < 0x80 or any(lo <= code <= hi for lo, hi in _KEEP_RANGES):
            ch = chr(code)
        else:
            ch = chr(code)
            n = unicodedata.normalize("NFKC", ch)
            ch = n if len(n) == 1 else ch
        self[code] = ch
        return ch


_SHADOW_TABLE = _ShadowTable()


@lru_cache(maxsize=4096)
def _shadow_cached(text: str) -> str:
    return text.translate(_SHADOW_TABLE)


def shadow(text: str) -> str:
    """1文字ずつ NFKC をかけた写し。結果が1文字にならない文字は元のまま残し、長さを保つ。

    同じ行のセルを何度も（空判定・マスク・区切り・識別子の抜き出しで）写すので、結果を覚えておく。
    """
    if text.isascii():
        return text
    return _shadow_cached(text)


def is_empty_log(text: str) -> bool:
    return shadow(text).strip() in EMPTY_TOKENS


def nfkc(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "")
