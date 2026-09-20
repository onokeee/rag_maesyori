"""出力ファイル名（Windows と LightRAG の両方で安全な名前）。

LightRAG のファイル名ヒント（`.[legacy-R(...)]`）は付けない。ヒントはサーバー側の取り込み設定より
優先されてしまう上に、そのサーバーが知らない書き方だと取り込みが HTTP 400 で断られる。
名前は「安定・意味が分かる・重複しない」だけを満たし、チャンクへの耐性は本文の作り方（記録の分割）で確保する。
"""
from __future__ import annotations

import re
import unicodedata

_UNSAFE = re.compile(r'[\\/:*?"<>|\[\]\s\x00-\x1f\x7f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def safe_filename_part(text, max_len: int = 60) -> str:
    """ファイル名の1部品。NFKC、禁止文字・空白・角かっこを _ に、'.[' を除去、前後の . _ を除去。

    ここは本文と違って NFKC をそのままかける（囲み文字は残さない）。ファイル名は同じ文書に対して
    安定・一意であればよく、①などを残すと OS や LightRAG 側での扱いが揺れるため。
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = s.replace(".[", "")  # LightRAG のヒント記法と誤認されないように
    s = _UNSAFE.sub("_", s)
    s = re.sub(r"_+", "_", s).strip("._")
    s = s[:max_len].strip("._")
    if s.split(".")[0].upper() in _WINDOWS_RESERVED:
        s = f"{s}_"
    return s


def md_filename(parts: list[str]) -> str:
    """部品を _ で連結した .md ファイル名。空の部品は飛ばす。"""
    safe = [p for p in (safe_filename_part(x) for x in parts) if p]
    return f"{'_'.join(safe) or '無題'}.md"
