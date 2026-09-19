"""読み取り結果から、保管用のJSONと RAG 投入用の Markdown を生成する。

Markdown は LightRAG 調査の指針（docs/design.md 6章）に従う。
  - 1帳票だけで意味が通るように、種類・識別番号・設備・日付をタイトルと本文に書く
  - 定型文・種類の版・DBの文書ID・セル座標は出さない（JSON側に残す）
  - 値は NFKC＋空白の畳み込み、数値は単位付き。人名の項目と「出さない」項目は省く
  - 値が別の欄の見出し語そのもの（読み取り誤り。「- 数量: 単価」）の項目は出さない（JSON には残す）
  - 明細表は見出しの下に1行1明細で「- 品番: X／品名: Y／数量: 2」と書く（パイプ表は使わない）
  - 同じ入力からは同じバイト列になる（生成日時などを書かない）
"""
from __future__ import annotations

import math
import re
import unicodedata
from pathlib import Path

from excel.tables import is_table_value, table_row_items
from excel.text import nfkc_value as _excel_nfkc_value, normalize_label as _normalize_label
from pattern.dictionary import is_person_field, is_person_label
from pattern.model import DEFAULT_MD_OPTIONS, DEFAULT_TITLE_KEYS

try:  # WP-core の共通実装があればそれを使う
    from core import mdtext as _core_mdtext
except ImportError:  # pragma: no cover - core 未導入時
    _core_mdtext = None
try:
    from core import naming as _core_naming
except ImportError:  # pragma: no cover - core 未導入時
    _core_naming = None

# 長文項目の見出しに識別子を入れるのは、推定トークン数がこれを超える帳票だけ。
# LightRAG が帳票を2つ以上の断片に切りうる大きさ（サーバー既定の固定窓 1,200トークン）に合わせる。
# 推定式が実トークン以上になったので、1,200 未満の帳票＝1断片に収まる帳票には識別子を入れない。
HEADING_IDENTIFIER_TOKENS = 1200
# 値が見出し語かどうかを見るのは短い値だけ（長い本文にたまたま同じ語が入っていても消さない）
MAX_LABEL_VALUE_CHARS = 20
AI_MARK = "（AI入力）"

_ISO_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?: \d{2}:\d{2})?$")  # 時刻付き（発生日時）も
# 明細表の合計行の先頭（「合計」「小計」「部品費計」）
_TOTAL_LABEL = re.compile(r"^(?:合計|小計|総計|計|.{1,6}計)$")


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
    heading_texts = _title_texts(title_fields, heading=True)
    stem = _one_line(Path(doc["file_name"]).stem)
    if not title_texts:
        title = f"{type_name} {stem}"
    else:
        if _has_equipment_only(title_fields):
            # 設備だけでは同じ設備の帳票が同じタイトル・同じ見出しになる。番号らしい項目・日付、
            # それも無ければ元ファイル名を足して、1帳票だけで見分けられるようにする（design.md 6章）
            extra = _fallback_identifier(filled, title_fields) or stem
            title_texts, heading_texts = [*title_texts, extra], [*heading_texts, extra]
        elif not _has_identifier(title_fields):
            # 識別番号も設備も入らないタイトルは、元ファイル名を足して帳票を特定できるようにする
            title_texts = [*title_texts, stem]
        title = f"{type_name} {'｜'.join(title_texts)}"
    identifier = "／".join(heading_texts)

    head = [f"# {_escape_line(title)}"]
    basics = [f"- 帳票の種類: {type_name}"]
    basics += _basic_lines(filled)
    tail = []
    attachments = extraction.get("attachments") or []
    if attachments:
        tail.append(f"- 添付画像: {len(attachments)}枚")
    tail.append(f"- 出典: {_source_text(doc, title_fields)}")

    long_fields = [f for f in filled if f["data_type"] in ("text", "table")]
    omit_person = bool(_md_options(extraction).get("omit_person_fields", True))

    def render(with_identifier: bool) -> str:
        blocks = [head, basics]
        for f in long_fields:
            heading = _one_line(f["display_name"])
            if with_identifier and identifier:
                heading += f"（{identifier}）"
            if f["data_type"] == "table":
                lines = table_markdown_lines(f["value"], omit_person)
                if f.get("ai_filled") and lines:
                    lines[-1] += AI_MARK
                blocks.append([f"## {_escape_line(heading)}", *lines])
                continue
            lines = [line for line in _format_value(f).split("\n") if line.strip()]
            blocks.append([f"## {_escape_line(heading)}", *(_escape_line(line) for line in lines)])
        blocks.append(tail)
        return _join_blocks(blocks)

    text = render(False)
    if long_fields and identifier and _estimate_tokens(text) > HEADING_IDENTIFIER_TOKENS:
        text = render(True)
    return text


def markdown_filename(doc: dict, extraction: dict) -> str:
    """{種類名}_{タイトル項目値...}.md。識別番号が入らないときは {file_hash先頭8} を足して一意にする。

    LightRAG 1.5.x は文書IDがファイル名の MD5 なので、同名のファイルは2件目が HTTP 409 で入らない
    （オフライン評価 3章: 見本 150 件中 3 件が同名だった）。
    """
    pattern = extraction["pattern"]
    shown = _shown_fields(extraction)
    title_fields = _title_fields(pattern, extraction["fields"], shown)
    values = [_plain_value(f) for f in title_fields]
    values = [v for v in values if _safe_filename_part(v)]
    hash8 = str(doc.get("file_hash") or "")[:8]
    if not values:
        return _md_filename([pattern["name"], hash8 or Path(doc["file_name"]).stem])
    if not any(f["field_name"] == "report_id" for f in title_fields) and hash8:
        values.append(hash8)
    return _md_filename([pattern["name"], *values])


# ---- 項目の選別・整形 ------------------------------------------------------------

def _md_options(extraction: dict) -> dict:
    return {**DEFAULT_MD_OPTIONS, **(extraction["pattern"].get("md_options") or {})}


def _shown_fields(extraction: dict) -> list[dict]:
    """Markdown に出す項目（「出さない」項目、設定により人名の項目、値が見出し語だけの項目を除く）。"""
    omit_person = bool(_md_options(extraction).get("omit_person_fields", True))
    labels = set(extraction["pattern"].get("labels") or ())
    people = _person_values(extraction["fields"]) if omit_person else set()
    return [
        f for f in extraction["fields"]
        if f.get("rag_output", "show") != "omit"
        and not (omit_person and is_person_field(f["field_name"], f.get("display_name", "")))
        and not _is_label_value(f, labels)
        and not _is_person_name_field(f, people)
    ]


def _person_values(fields: list[dict]) -> set[str]:
    """人名の項目（作成・確認・承認・報告者…）に読み取った値（正規化したもの）。"""
    return {_normalize_label(_one_line(f["value"])) for f in fields
            if f["data_type"] == "string" and not _is_blank(f["value"])
            and is_person_field(f["field_name"], f.get("display_name", ""))} - {""}


def _is_person_name_field(f: dict, people: set[str]) -> bool:
    """押印欄の人名（「斎藤」）を見出しにした項目か、値が人名の項目の値と同じ短い項目か。

    帳票の種類を作るときに人名が見出しとして選ばれると、人名の項目を省いても「- 斎藤: 森」が出てしまう。
    """
    if not people or f["data_type"] != "string" or is_person_field(f["field_name"], f.get("display_name", "")):
        return False
    if _normalize_label(f.get("display_name", "")) in people:
        return True
    text = _one_line(f["value"])
    return bool(text) and len(text) <= MAX_LABEL_VALUE_CHARS and _normalize_label(text) in people


def _is_label_value(f: dict, labels: set[str]) -> bool:
    """値が別の欄（またはこの欄）の見出し語そのものか。読み取り誤りで「- 品名: 品番」になった行を出さないための判定。

    LightRAG オフライン評価 3章: `- 品名: 品番`（F4 30/30）など、値が見出し語のままの行が LLM のエンティティになっていた。
    """
    if not labels or f["data_type"] in ("table", "date", "number") or f.get("edited"):
        return False
    text = _one_line(f["value"])
    if not text or len(text) > MAX_LABEL_VALUE_CHARS or "\n" in str(f["value"] or ""):
        return False
    return _normalize_label(text) in labels


def _title_fields(pattern: dict, all_fields: list[dict], shown: list[dict]) -> list[dict]:
    """タイトルに使う項目（値のあるもの）。未設定なら報告番号・設備・発生日の辞書キー順。明細表は使わない。"""
    by_name = {f["field_name"]: f for f in shown if f["data_type"] != "table"}
    configured = [n for n in (pattern.get("title_fields") or []) if n in {f["field_name"] for f in all_fields}]
    keys = configured or list(DEFAULT_TITLE_KEYS)
    return [by_name[k] for k in dict.fromkeys(keys) if k in by_name and not _is_blank(by_name[k]["value"])]


def _has_identifier(fields: list[dict]) -> bool:
    """タイトル項目に、その帳票を見分けられるもの（識別番号・設備）があるか。"""
    return any(f["field_name"] in ("report_id", "equipment_id", "equipment_name") for f in fields)


def _has_equipment_only(fields: list[dict]) -> bool:
    """タイトル項目で帳票を見分ける手がかりが設備だけか（識別番号も日付も無い）。

    同じ設備の点検記録は何十件もあるので、設備だけでは帳票を見分けられない。
    """
    return (any(f["field_name"] in ("equipment_id", "equipment_name") for f in fields)
            and not any(f["field_name"] == "report_id" or f["data_type"] == "date" for f in fields))


# 「作業No.」「管理番号」「点検№」のような番号らしい見出し（設備番号は設備なので除く）
_NUMBER_LABEL = re.compile(r"(?:No\.?|NO\.?|番号|№)\s*$")


def _fallback_identifier(filled: list[dict], title_fields: list[dict]) -> str:
    """タイトル項目で見分けられないときに足す値: 番号らしい項目 → 日付の項目の順。無ければ ""。"""
    used = {f["field_name"] for f in title_fields}
    rest = [f for f in filled if f["field_name"] not in used
            and f["field_name"] not in ("equipment_id", "equipment_name")]
    numbered = next((f for f in rest if f["data_type"] in ("string", "number")
                     and _NUMBER_LABEL.search(_one_line(f.get("display_name")))), None)
    dated = next((f for f in rest if f["data_type"] == "date"), None)
    found = numbered or dated
    return _one_line(_plain_value(found)) if found else ""


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
    short = [f for f in filled if f["data_type"] not in ("text", "table")]
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


def table_markdown_lines(value, omit_person: bool = False) -> list[str]:
    """明細表の1行を「- 品番: X／品名: Y」の1行にする。空のセルは書かない。合計行は「- 合計: 投入数: 50枚／…」。

    omit_person: 人名の列見出し（担当・氏名など）の列を出さない（本文の人名の項目と同じ扱い）。
    """
    lines = []
    if not is_table_value(value):
        return lines
    for row in value["rows"]:
        items = [(_one_line(k), _one_line(v)) for k, v in table_row_items(value, row)]
        items = [(k, v) for k, v in items if v]
        if omit_person:
            items = [(k, v) for k, v in items if not is_person_label(k)]
        if not items:
            continue
        if _TOTAL_LABEL.match(items[0][1]):
            # 数字の無い合計行（「合計」だけの行）は記録として意味がないので出さない
            if len(items) > 1:
                lines.append(f"- {items[0][1]}: " + "／".join(f"{k}: {v}" for k, v in items[1:]))
            continue
        lines.append("- " + "／".join(f"{k}: {v}" for k, v in items))
    return lines


def _is_blank(value) -> bool:
    if isinstance(value, dict):
        return not value.get("rows")
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

def nfkc_value(text) -> str:
    """値の正規化（NFKC＋空白の畳み込み）。丸数字などの囲み文字（①②Ⓐ㋐）は原文どおり残す。

    ①→1 にすると「①破損ウェーハ片を回収」が「1破損ウェーハ片を回収」になり、番号と本文の区切りが消える。
    ㈱→(株)、⑴→(1) のように区切りが残る表記は今までどおり正規化する（core/mdtext.nfkc_keep_enclosed）。
    """
    s = "" if text is None else str(text).replace("_x000D_", "")
    if _core_mdtext is not None and hasattr(_core_mdtext, "nfkc_value"):
        return _core_mdtext.nfkc_value(s)
    return _excel_nfkc_value(s)


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
