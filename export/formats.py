"""読み取り結果から、保管用のJSONと RAG 投入用の Markdown を生成する。

Markdown は LightRAG 調査の指針（docs/design.md 6章）に従う。
  - 1帳票だけで意味が通るように、種類・識別番号・設備・日付をタイトルと本文に書く
  - 定型文・種類の版・DBの文書ID・セル座標は出さない（JSON側に残す）
  - 値は NFKC＋空白の畳み込み、数値は単位付き。人名の項目と「出さない」項目は省く
  - 同じ入力からは同じバイト列になる（生成日時などを書かない）
"""
from __future__ import annotations

import math
import re
import unicodedata
from pathlib import Path

from excel.text import nfkc_value
from pattern.dictionary import is_person_field
from pattern.model import DEFAULT_MD_OPTIONS, DEFAULT_TITLE_KEYS

try:  # WP-core の共通実装があればそれを使う
    from core import mdtext as _core_mdtext
except ImportError:  # pragma: no cover - core 未導入時
    _core_mdtext = None
try:
    from core import naming as _core_naming
except ImportError:  # pragma: no cover - core 未導入時
    _core_naming = None

# 長文項目の見出しに識別子を入れるのは、推定トークン数がこれを超える帳票だけ
HEADING_IDENTIFIER_TOKENS = 1000
AI_MARK = "（AI入力）"

_ISO_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")


# ---- JSON ---------------------------------------------------------------------

def build_json(doc: dict, extraction: dict) -> dict:
    return {
        "document_id": doc["id"],
        "pattern": extraction["pattern"],
        "values": extraction["values"],
        "fields": [
            {k: f.get(k) for k in ("field_name", "display_name", "data_type", "value", "unit", "rag_output", "sheet",
                                   "label_cell", "value_cell", "edited", "ai_filled")}
            for f in extraction["fields"]
        ],
        "missing_required": extraction["missing_required"],
        "attachments": extraction["attachments"],
        "source": {
            "file_name": doc["file_name"],
            "file_hash": doc["file_hash"],
            "sheets": extraction["sheets"],
            "url": f"/forms/{doc['id']}/original",
        },
    }


# ---- Markdown -----------------------------------------------------------------

def build_markdown(doc: dict, extraction: dict) -> str:
    """RAG（LightRAG）に投入する Markdown。書式は docs/design.md 6.1。"""
    pattern = extraction["pattern"]
    type_name = _one_line(pattern["name"])
    shown = _shown_fields(extraction)
    filled = [f for f in shown if not _is_blank(f["value"])]
    title_fields = _title_fields(pattern, extraction["fields"], shown)

    title_texts = _title_texts(title_fields, heading=False)
    title = f"{type_name} {'｜'.join(title_texts)}" if title_texts else f"{type_name} {_one_line(Path(doc['file_name']).stem)}"
    identifier = "／".join(_title_texts(title_fields, heading=True))

    head = [f"# {_escape_line(title)}"]
    basics = [f"- 帳票の種類: {type_name}"]
    basics += _basic_lines(filled)
    tail = []
    attachments = extraction.get("attachments") or []
    if attachments:
        tail.append(f"- 添付画像: {len(attachments)}枚")
    tail.append(f"- 出典: {_source_text(doc, title_fields)}")

    long_fields = [f for f in filled if f["data_type"] == "text"]

    def render(with_identifier: bool) -> str:
        blocks = [head, basics]
        for f in long_fields:
            heading = _one_line(f["display_name"])
            if with_identifier and identifier:
                heading += f"（{identifier}）"
            lines = [line for line in _format_value(f).split("\n") if line.strip()]
            blocks.append([f"## {_escape_line(heading)}", *(_escape_line(line) for line in lines)])
        blocks.append(tail)
        return _join_blocks(blocks)

    text = render(False)
    if long_fields and identifier and _estimate_tokens(text) > HEADING_IDENTIFIER_TOKENS:
        text = render(True)
    return text


def markdown_filename(doc: dict, extraction: dict) -> str:
    """{種類名}_{タイトル項目値...}.md。タイトル項目が空なら {種類名}_{file_hash先頭8}.md。"""
    pattern = extraction["pattern"]
    shown = _shown_fields(extraction)
    values = [_plain_value(f) for f in _title_fields(pattern, extraction["fields"], shown)]
    values = [v for v in values if _safe_filename_part(v)]
    if not values:
        values = [str(doc.get("file_hash") or "")[:8] or Path(doc["file_name"]).stem]
    return _md_filename([pattern["name"], *values])


# ---- 項目の選別・整形 ------------------------------------------------------------

def _md_options(extraction: dict) -> dict:
    return {**DEFAULT_MD_OPTIONS, **(extraction["pattern"].get("md_options") or {})}


def _shown_fields(extraction: dict) -> list[dict]:
    """Markdown に出す項目（「出さない」項目と、設定により人名の項目を除く）。"""
    omit_person = bool(_md_options(extraction).get("omit_person_fields", True))
    return [
        f for f in extraction["fields"]
        if f.get("rag_output", "show") != "omit"
        and not (omit_person and is_person_field(f["field_name"], f.get("display_name", "")))
    ]


def _title_fields(pattern: dict, all_fields: list[dict], shown: list[dict]) -> list[dict]:
    """タイトルに使う項目（値のあるもの）。未設定なら報告番号・設備・発生日の辞書キー順。"""
    by_name = {f["field_name"]: f for f in shown}
    configured = [n for n in (pattern.get("title_fields") or []) if n in {f["field_name"] for f in all_fields}]
    keys = configured or list(DEFAULT_TITLE_KEYS)
    return [by_name[k] for k in dict.fromkeys(keys) if k in by_name and not _is_blank(by_name[k]["value"])]


def _title_texts(fields: list[dict], heading: bool) -> list[str]:
    """設備番号と設備名が両方あれば1つにまとめる（タイトル: 名前（番号）、見出し: 番号 名前）。"""
    names = {f["field_name"]: f for f in fields}
    pair = "equipment_id" in names and "equipment_name" in names
    texts: list[str] = []
    done_pair = False
    for f in fields:
        if pair and f["field_name"] in ("equipment_id", "equipment_name"):
            if not done_pair:
                eq_id = _one_line(_plain_value(names["equipment_id"]))
                eq_name = _one_line(_plain_value(names["equipment_name"]))
                texts.append(f"{eq_id} {eq_name}" if heading else f"{eq_name}（{eq_id}）")
                done_pair = True
            continue
        texts.append(_one_line(_plain_value(f)))
    return [t for t in texts if t]


def _basic_lines(filled: list[dict]) -> list[str]:
    short = [f for f in filled if f["data_type"] != "text"]
    names = {f["field_name"]: f for f in short}
    pair = "equipment_id" in names and "equipment_name" in names
    lines: list[str] = []
    done_pair = False
    for f in short:
        if pair and f["field_name"] in ("equipment_id", "equipment_name"):
            if not done_pair:
                eq_id, eq_name = names["equipment_id"], names["equipment_name"]
                value = f"{_plain_value(eq_name)}（{_plain_value(eq_id)}）"
                if eq_id.get("ai_filled") or eq_name.get("ai_filled"):
                    value += AI_MARK
                lines += _md_bullet("設備", value)
                done_pair = True
            continue
        lines += _md_bullet(_one_line(f["display_name"]), _format_value(f))
    return lines


def _format_value(f: dict) -> str:
    """本文用の値。日付は「2026-09-14（2026年9月）」、数値は単位付き、AI入力には印を付ける。"""
    value = f["value"]
    text = _plain_value(f) if f["data_type"] != "text" else nfkc_value(value)
    if f["data_type"] == "date":
        m = _ISO_DATE.match(text)
        if m:
            text = f"{text}（{int(m[1])}年{int(m[2])}月）"
    elif f["data_type"] == "number" and _is_number(value) and f.get("unit"):
        text = f"{text}{nfkc_value(f['unit'])}"
    if f.get("ai_filled"):
        text += AI_MARK
    return text


def _plain_value(f: dict) -> str:
    """タイトル・ファイル名用の値（印や単位なし）。"""
    value = f["value"]
    if _is_number(value):
        return _number_text(value)
    if f["data_type"] == "text":
        return _one_line(value)
    return nfkc_value(value)


def _source_text(doc: dict, title_fields: list[dict]) -> str:
    """出典: 元ファイル名（報告番号 R2026-00123）。報告番号がなければ最初の文字列のタイトル項目。"""
    file_name = _one_line(doc["file_name"])
    ident = next((f for f in title_fields if f["field_name"] == "report_id"), None)
    ident = ident or next((f for f in title_fields if f["data_type"] == "string"
                           and f["field_name"] not in ("equipment_id", "equipment_name")), None)
    if ident is None:
        return file_name
    return f"{file_name}（{_one_line(ident['display_name'])} {_plain_value(ident)}）"


def _is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _number_text(value) -> str:
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.10f}".rstrip("0").rstrip(".")
    return str(value)


def _one_line(value) -> str:
    return " ".join(nfkc_value(value).split())


# ---- Markdown テキスト処理（core/mdtext・core/naming があればそれを使う） -------------------

_MD_BLOCK_START = re.compile(r"^(\s*)(#|>|[-*+](?=\s|$)|\d+[.)](?=\s|$)|```|~~~)")
_MD_RULE = re.compile(r"^\s*([-=_*])\1{2,}\s*$")


def _escape_line(line: str) -> str:
    if _core_mdtext is not None and hasattr(_core_mdtext, "escape_md_line"):
        return _core_mdtext.escape_md_line(line)
    if _MD_RULE.match(line):
        return "\\" + line.lstrip()
    m = _MD_BLOCK_START.match(line)
    if not m:
        return line
    token = m[2]
    if token[0].isdigit():  # 「1. 」→「1\. 」
        return line[: m.end() - 1] + "\\" + line[m.end() - 1:]
    return line[: m.start(2)] + "\\" + line[m.start(2):]


def _md_bullet(label: str, value: str) -> list[str]:
    if _core_mdtext is not None and hasattr(_core_mdtext, "md_bullet"):
        return _core_mdtext.md_bullet(label, value)
    lines = [line.strip() for line in str(value).split("\n") if line.strip()]
    if not lines:
        return []
    if len(lines) == 1:
        return [f"- {label}: {lines[0]}"]
    return [f"- {label}:", *(f"  {_escape_line(line)}" for line in lines)]


def _estimate_tokens(text: str) -> int:
    if _core_mdtext is not None and hasattr(_core_mdtext, "estimate_tokens"):
        return _core_mdtext.estimate_tokens(text)
    ascii_chars = sum(1 for ch in text if ord(ch) < 128)
    return (len(text) - ascii_chars) + math.ceil(ascii_chars / 3)


def _join_blocks(blocks: list[list[str]]) -> str:
    if _core_mdtext is not None and hasattr(_core_mdtext, "join_blocks"):
        return _core_mdtext.join_blocks(blocks)
    parts = ["\n".join(ln.rstrip() for ln in b if ln.strip()) for b in blocks]
    parts = [p for p in parts if p]
    return "\n\n".join(parts) + "\n" if parts else ""


_FILENAME_UNSAFE = re.compile(r'[\\/:*?"<>|\[\]\s\x00-\x1f\x7f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def _safe_filename_part(text: str, max_len: int = 60) -> str:
    if _core_naming is not None and hasattr(_core_naming, "safe_filename_part"):
        return _core_naming.safe_filename_part(text, max_len)
    s = unicodedata.normalize("NFKC", str(text or "")).replace(".[", "")
    s = re.sub(r"_+", "_", _FILENAME_UNSAFE.sub("_", s)).strip("._")
    s = s[:max_len].strip("._")
    return f"{s}_" if s.split(".")[0].upper() in _WINDOWS_RESERVED else s


def _md_filename(parts: list[str], hint: str | None = None) -> str:
    if _core_naming is not None and hasattr(_core_naming, "md_filename"):
        return _core_naming.md_filename(parts, hint)
    safe = [p for p in (_safe_filename_part(part) for part in parts) if p]
    name = "_".join(safe) or "無題"
    return name + (f".[{hint}]" if hint else "") + ".md"
