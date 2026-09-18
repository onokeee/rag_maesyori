"""出力ファイル名（Windows と LightRAG の両方で安全な名前）。"""
from __future__ import annotations

import re
import unicodedata

# LightRAG のファイル名ヒント（一覧表の記録ファイル用。新しい取り込み設定では既定で付ける）
# chunk_ts は記録1件がまるごと1チャンクに収まる大きさにする（オフライン評価: 800 では T1 の 63%・T5 の 43%・T2 の 12% が途中で切れた）
LIGHTRAG_HINT_RECORDS = "legacy-R(chunk_ts=1500,chunk_ol=0)"

_CHUNK_TS = re.compile(r"chunk_ts\s*=\s*(\d+)")
_UNSAFE = re.compile(r'[\\/:*?"<>|\[\]\s\x00-\x1f\x7f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def hint_chunk_tokens(hint: str | None = None) -> int:
    """ヒントの chunk_ts（1チャンクの最大トークン数）。読めなければ 0。記録の大きさの上限はこれから決める。"""
    m = _CHUNK_TS.search(hint if hint is not None else LIGHTRAG_HINT_RECORDS)
    return int(m.group(1)) if m else 0


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


def md_filename(parts: list[str], hint: str | None = None) -> str:
    """部品を _ で連結した .md ファイル名。空の部品は飛ばす。hint があれば '.[hint]' を付ける。"""
    safe = [p for p in (safe_filename_part(x) for x in parts) if p]
    base = "_".join(safe) or "無題"
    return f"{base}.[{hint}].md" if hint else f"{base}.md"
