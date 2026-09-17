"""セル値の正規化と型変換。"""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, time

from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel

_EDGE_CHARS = "[]()<>【】〈〉《》「」『』■□◆◇●○・*※:;."
_INLINE_SEP = re.compile(r"[:：]")
_DATE_RE = re.compile(r"(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")
_ERA_RE = re.compile(r"(?:令和|R)\s*(\d{1,2}|元)\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")
_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")
_SPACES_RE = re.compile(r"[^\S\n]+")
# 見出し末尾の単位: 「作業時間(h)」「停止時間（分）」「金額[円]」
_HEADER_UNIT_RE = re.compile(r"^(.+?)\s*[(\[]\s*([^()\[\]\d]{1,6})\s*[)\]]$")
# 値の末尾の単位: 「1.5時間」「95分」「120 min」
_VALUE_UNIT_RE = re.compile(r"^-?\d+(?:\.\d+)?\s*([^\d\s.,\-]{1,4})$")
UNIT_ALIASES = {"h": "時間", "hr": "時間", "hrs": "時間", "hour": "時間", "hours": "時間",
                "min": "分", "mins": "分", "sec": "秒", "yen": "円", "¥": "円"}

MAX_LABEL_LENGTH = 20


def normalize_label(text) -> str:
    """ラベル比較用の正規化。全角/半角・空白・前後の記号・大文字小文字の揺れを吸収する。

    例: 「設備№：」「 設備 No. 」「【設備NO】」→ "設備no"
    """
    s = unicodedata.normalize("NFKC", str(text))
    s = re.sub(r"\s+", "", s)
    return s.strip(_EDGE_CHARS).lower()


def normalize_sheet_name(name) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(name))).lower()


def nfkc_value(text) -> str:
    """値の正規化（Markdown 出力用）。NFKC＋行内の連続空白を1つに畳む。改行と日本語間の空白は残す。

    ラベル比較用の normalize_label とは別物（こちらは空白を消さない）。
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text)).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [_SPACES_RE.sub(" ", line).strip() for line in s.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def normalize_unit(unit: str) -> str:
    u = unicodedata.normalize("NFKC", unit or "").strip()
    return UNIT_ALIASES.get(u.lower(), u)


def split_label_unit(label: str) -> tuple[str, str]:
    """「作業時間(h)」→ ("作業時間", "時間")。単位がなければ (label, "")。"""
    s = unicodedata.normalize("NFKC", str(label or "")).strip()
    m = _HEADER_UNIT_RE.match(s)
    if not m or not m[1].strip():
        return str(label or "").strip(), ""
    return m[1].strip(), normalize_unit(m[2])


def value_unit(text) -> str:
    """「1.5時間」→ "時間"。数値＋単位の形でなければ ""。"""
    m = _VALUE_UNIT_RE.match(unicodedata.normalize("NFKC", str(text or "")).strip())
    return normalize_unit(m[1]) if m else ""


def cell_text(value) -> str:
    """セル値を表示用の文字列にする。"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        if value.time() == time(0):
            return value.date().isoformat()
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def split_inline(text: str) -> tuple[str, str] | None:
    """「設備番号：EQ-001」のように1セルにラベルと値が入っている場合に分割する。

    戻り値: (正規化済みラベル, 値テキスト)
    """
    m = _INLINE_SEP.search(text)
    if not m:
        return None
    label, value = text[: m.start()], text[m.end():].strip()
    if "\n" in label or not value:
        return None
    label_norm = normalize_label(label)
    if not label_norm or len(label_norm) > MAX_LABEL_LENGTH:
        return None
    return label_norm, value


def to_date(raw, text: str, date1904: bool = False) -> tuple[str | None, str | None]:
    """日付に変換する。戻り値: (ISO形式の日付 or 元テキスト, 警告)

    date1904: ブックが1904年基準（Mac版Excel由来）のときシリアル値の起点を変える。
    """
    if isinstance(raw, datetime):
        return raw.date().isoformat(), None
    if isinstance(raw, date):
        return raw.isoformat(), None
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        low = 20000 - 1462 if date1904 else 20000
        if low < raw < 80000:
            epoch = MAC_EPOCH if date1904 else WINDOWS_EPOCH
            return from_excel(raw, epoch=epoch).date().isoformat(), None

    s = unicodedata.normalize("NFKC", text or "")
    m = _ERA_RE.search(s)
    if m:
        year = 2018 + (1 if m[1] == "元" else int(m[1]))
        parsed = _safe_date(year, int(m[2]), int(m[3]))
        if parsed:
            return parsed, None
    m = _DATE_RE.search(s)
    if m:
        parsed = _safe_date(int(m[1]), int(m[2]), int(m[3]))
        if parsed:
            return parsed, None
    return (text or None), "日付として解釈できません"


def _safe_date(y: int, m: int, d: int) -> str | None:
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def to_number(raw, text: str) -> tuple[int | float | str | None, str | None]:
    """数値に変換する。「2.5時間」のような単位付きは数値部分を取り出して警告を付ける。"""
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return (int(raw) if float(raw).is_integer() else raw), None

    s = unicodedata.normalize("NFKC", text or "").replace(",", "").strip()
    m = _NUM_RE.search(s)
    if not m:
        return (text or None), "数値として解釈できません"
    number = float(m[0])
    value = int(number) if number.is_integer() else number
    warning = None if m[0] == s else f"「{text}」から数値部分を抽出しました"
    return value, warning
