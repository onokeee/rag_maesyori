"""出力ファイル名（Windows と LightRAG の両方で安全な名前）。"""
from __future__ import annotations

import re
import unicodedata

# LightRAG のファイル名ヒント（一覧表の記録ファイル用。既定では付けない）
LIGHTRAG_HINT_RECORDS = "legacy-R(chunk_ts=800,chunk_ol=0)"

_UNSAFE = re.compile(r'[\\/:*?"<>|\[\]\s\x00-\x1f\x7f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def safe_filename_part(text, max_len: int = 60) -> str:
    """ファイル名の1部品。NFKC、禁止文字・空白・角かっこを _ に、'.[' を除去、前後の . _ を除去。"""
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


def md_filename(parts: list[str], hint: str | None = None) -> str:
    """部品を _ で連結した .md ファイル名。空の部品は飛ばす。hint があれば '.[hint]' を付ける。"""
    safe = [p for p in (safe_filename_part(x) for x in parts) if p]
    base = "_".join(safe) or "無題"
    return f"{base}.[{hint}].md" if hint else f"{base}.md"
