"""帳票（1ファイル＝1件）の読み取り。

Excel のセル構造の読み込み（旧 excel/）、ラベル探索による値の読み取り（旧 excel/extractor）、
帳票の種類の定義・候補・クリックからの項目作り（旧 pattern/）、帳票の Markdown（旧 export/formats）を
1ファイルにまとめてある。
"""
from __future__ import annotations

import io
import posixpath
import re
import unicodedata
import warnings
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field, asdict, replace
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel

from app import core



# ====================================================================================================
# 元 excel/text.py
# セル値の正規化と型変換。
# ====================================================================================================

_EDGE_CHARS = "[]()<>【】〈〉《》「」『』■□◆◇●○・*※:;."
_BRACKETS = {"[": "]", "(": ")", "<": ">", "【": "】", "〈": "〉", "《": "》", "「": "」", "『": "』"}
_CLOSERS = {close: open_ for open_, close in _BRACKETS.items()}
_EDGE_MARKS = "".join(ch for ch in _EDGE_CHARS if ch not in _BRACKETS and ch not in _CLOSERS)
# 見出しの先頭の項番: 「3.」「3)」「3、」「(3)」「D3」（8D の D1〜D8）「A.機構部」。丸数字は section_stripped で別に見る
_SECTION_NO_RE = re.compile(r"^(?:\d{1,2}[.)、](?!\d)|\(\d{1,2}\)|d[1-8](?=\D)|[a-z][.)、](?=[^\x00-\x7f]))")
_TRAILING_PAREN_RE = re.compile(r"\([^()]{1,12}\)$")
_INLINE_SEP = re.compile(r"[:：]")
_DATE_RE = re.compile(r"(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")
# 年の無い日付（「2/12」「2/12 3時17分」「2月12日 14:05」）。年は推測せず、原文のまま残して警告だけ出す。
# 「24/8/25」のような2桁の年は、年の欄が空なのか2桁で書いたのか分からないのでこの形には含めない
_NO_YEAR_RE = re.compile(r"^\s*(\d{1,2})\s*(?:/|月|-|\.)\s*(\d{1,2})\s*日?(?:\s|$|[^\d/\-.])")
# 日付の直後の時刻「2026/2/12 14:05」「2024年7月29日 13:41」「2023-07-10T23:08」「2023年7月10日 23時08分」。
# 日付との間の「日」・空白・曜日「(月)」・「T」は飛ばす
_CLOCK_LEAD_RE = re.compile(r"日?\s*(?:\([月火水木金土日]\)\s*)?T?\s*")
_CLOCK_RE = re.compile(r"(\d{1,2})\s*(?::(\d{2})(?::\d{2})?|時\s*(\d{1,2})\s*分?)(?!\d)")
# 和暦「令和5年11月16日」「R5.11.16」「平成30年4月1日」「H30.4.1」。年の起点は元号ごと
_ERA_RE = re.compile(r"(?<![A-Za-z])(令和|R|平成|H)\s*(\d{1,2}|元)\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")
_ERA_BASE = {"令和": 2018, "R": 2018, "平成": 1988, "H": 1988}
# 「.5」「約.5時間」のような整数部の無い小数も読む（「5」と読まない）
_NUM_RE = re.compile(r"-?(?:\d+(?:\.\d+)?|\.\d+)")
_DURATION_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:時間|hours?|hrs?|h)\s*(\d+(?:\.\d+)?)\s*(?:分|mins?|m)(?![A-Za-z])",
                          re.IGNORECASE)
# 時刻の形の時間「2:45」「25:30」（時間の欄に h:mm で書いた値。Excel の [h]:mm のセルもこの形の文字にする）
_HMM_RE = re.compile(r"^(\d{1,4}):([0-5]\d)(?::([0-5]\d))?$")
# 数値の範囲「10～20分」「1〜2時間」（NFKC で「～」は「~」になる）
_NUM_RANGE_RE = re.compile(r"(?:\d+(?:\.\d+)?|\.\d+)\s*[^\d\s~〜]{0,4}?\s*[~〜]\s*(\d+(?:\.\d+)?|\.\d+)")
# 時刻の範囲「09:30-12:45」「9:30～13:55」（作業時間の欄によくある書き方。先頭の 9 を数値として読まない）
# 「12/24 21:53-12/25 11:09」のように各時刻の前に月日が付いた書き方も範囲として見る（先頭の月を数値として読まない）
_TIME_RANGE_RE = re.compile(r"(?:\d{1,2}/\d{1,2}\s*)?(\d{1,2})\s*:\s*(\d{2})\s*[-~〜ー―]\s*"
                            r"(?:\d{1,2}/\d{1,2}\s*)?(\d{1,2})\s*:\s*(\d{2})")
# 時刻の範囲に添えた時間数「（3.2h）」「（工数 8.75h）」「(195分)」
_TIME_AMOUNT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(時間|hours?|hrs?|h|分|mins?)(?![A-Za-z])", re.IGNORECASE)
_SPACES_RE = re.compile(r"[^\S\n]+")
# 見出し末尾の単位: 「作業時間(h)」「停止時間（分）」「金額[円]」
_HEADER_UNIT_RE = re.compile(r"^(.+?)\s*[(\[]\s*([^()\[\]\d]{1,6})\s*[)\]]$")
# 値の末尾の単位: 「1.5時間」「95分」「120 min」
_VALUE_UNIT_RE = re.compile(r"^-?(?:\d+(?:\.\d+)?|\.\d+)\s*([^\d\s.,\-]{1,4})$")
# 前後に言葉や括弧書きのある値（「約90分」「595分（9.9h）」）の、最初の数値の直後の単位
_UNIT_AFTER_RE = re.compile(r"\s*([^\d\s.,\-:/~()\[\]{}<>、。・]{1,4})(?![^\d\s.,\-:/~()\[\]{}<>、。・])")
# 会計の書き方の負号「▲50万円」「△0.8」（一覧表の側と同じく負の数として読む）
_MINUS_MARK_RE = re.compile(r"^[△▲]\s*(?=\d)")
# 見えない文字（ゼロ幅スペース・語結合子・BOM・方向制御・ソフトハイフン）。Web や Teams から貼った文字に混ざり、
# NFKC でも消えないので、ラベルが見つからない・同じ値が別の値になる。セルの文字とラベル・値の正規化で消す
_INVISIBLE_RE = re.compile("[\u200b-\u200d\u2060\ufeff\u202a-\u202e\u2066-\u2069\u00ad]")
# 見出しの括弧書きのうち単位とみなす和文（「発生原因（推定）」「担当（記入）」の括弧書きは単位でない）。
# 英字・記号の単位（h, min, mm, %, ℃）と1文字の和文（分・円・枚・個）はこの一覧に無くても単位とみなす
_JA_UNITS = {"時間", "千円", "万円", "百万円", "人日", "人時", "日間", "ヶ月", "か月", "カ月", "箇所", "ケ所"}
UNIT_ALIASES = {"h": "時間", "hr": "時間", "hrs": "時間", "hour": "時間", "hours": "時間",
                "min": "分", "mins": "分", "sec": "秒", "yen": "円", "¥": "円"}

MAX_LABEL_LENGTH = 20

# Excel のエラー値（数式の結果）。値として取り込まない（一覧表の側 tables/normalize.py と同じ一覧）
EXCEL_ERRORS = {"#N/A", "#DIV/0!", "#REF!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#GETTING_DATA"}
EXCEL_ERROR_WARNING = "Excelのエラー値"


def excel_error(text) -> str:
    """セルの文字が Excel のエラー値（「#REF!」「#N/A」）ならその文字、違えば ""。"""
    s = str(text or "").strip().upper()
    return s if s in EXCEL_ERRORS else ""


# チェックボックス表記「■重大　□大　□中」「☑同型機　□類似設備」
CHECKED_MARKS = "■☑☒✓✔"
_CHECKBOX_RE = re.compile(r"([■□☑☐☒✓✔])\s*([^■□☑☐☒✓✔\s]+)")


def normalize_label(text) -> str:
    """ラベル比較用の正規化。全角/半角・空白・前後の記号・大文字小文字の揺れを吸収する。

    例: 「設備№：」「 設備 No. 」「【設備NO】」→ "設備no"
    """
    s = unicodedata.normalize("NFKC", strip_invisible(str(text)))
    s = re.sub(r"\s+", "", s)
    return _strip_edges(s).lower()


def strip_invisible(text: str) -> str:
    """見えない文字（ゼロ幅スペースなど）を消す。多くの文字列には無いので、先に確かめてから置き換える。"""
    return _INVISIBLE_RE.sub("", text) if _INVISIBLE_RE.search(text) else text


def _strip_edges(s: str) -> str:
    """前後の記号を除く。括弧は対になっているときだけ外す（「停止時間(分)」の「)」だけを消さない）。"""
    prev = None
    while s and s != prev:
        prev = s
        s = s.strip(_EDGE_MARKS)
        if not s:
            break
        close = _BRACKETS.get(s[0])
        if close and s.endswith(close) and len(s) > 1:
            s = s[1:-1]
            continue
        if close and close not in s[1:]:
            s = s[1:]
        if s and s[-1] in _CLOSERS and _CLOSERS[s[-1]] not in s[:-1]:
            s = s[:-1]
    return s


def section_stripped(text) -> str:
    """見出しの先頭の項番を除いた正規化ラベル。「３．暫定対策」「①何が」「(2)原因」「D2 問題の記述」→ 項番なし。

    項番が無い、または除くと文字が残らない（「1.5」など）ときは "" を返す。
    「2号機」のような数字始まりの値は項番とみなさない（区切りの「.」「)」「、」か丸数字が必要）。
    """
    raw = str(text).strip()
    if raw[:1] and "①" <= raw[0] <= "⑳":
        rest = normalize_label(raw[1:])
    else:
        compact = re.sub(r"\s+", "", unicodedata.normalize("NFKC", raw)).lower()
        m = _SECTION_NO_RE.match(compact)
        rest = normalize_label(compact[m.end():]) if m else ""
    return rest if any(ch.isalpha() for ch in rest) else ""


def label_parts(text) -> tuple[str, ...]:
    """「ライン／工程」のように「／」で2〜3個のラベルをまとめた見出しの、各部分の正規化ラベル。

    各部分が2文字以上で文字を含むときだけ分ける（「L/min」は分けない）。
    """
    s = unicodedata.normalize("NFKC", str(text or "")).strip()
    if "\n" in s or "/" not in s:
        return ()
    parts = [normalize_label(p) for p in s.split("/")]
    if not 2 <= len(parts) <= 3 or any(len(p) < 2 or not any(ch.isalpha() for ch in p) for p in parts):
        return ()
    return tuple(parts)


def split_combined_value(text, count: int) -> list[str] | None:
    """「L6／STI-CMP」を count 個に分ける。区切りの数が合わなければ None。"""
    parts = [p.strip() for p in re.split(r"[/／]", str(text or ""))]
    return parts if len(parts) == count and all(parts) else None


# 設備番号らしい記号: 英字と数字を両方含む空白なしの英数字（CMP-108, ROB-821, EQ001）
_CODE_TOKEN = r"(?=[A-Za-z0-9\-_.#]*[A-Za-z])(?=[A-Za-z0-9\-_.#]*\d)[A-Za-z0-9][A-Za-z0-9\-_.#]{2,19}"
_CODE_NAME_PATTERNS = (
    (re.compile(rf"^({_CODE_TOKEN})\s*\((.+)\)$"), 1, 2),        # CVD-202（TEOS CVD 2号機）
    (re.compile(rf"^(.+?)\s*\(({_CODE_TOKEN})\)$"), 2, 1),        # TEOS CVD 2号機（CVD-202）
    (re.compile(rf"^({_CODE_TOKEN})(?:\s*[/:]\s*|\s+)(.+)$"), 1, 2),  # CMP-108　STI-CMP 8号機 / CMP-108：STI-CMP
    (re.compile(rf"^(.+?)(?:\s*[/:]\s*|\s+)({_CODE_TOKEN})$"), 2, 1),  # W-CMP 3号機　CMP-103
)
_CODE_ONLY = re.compile(rf"^{_CODE_TOKEN}$")


def split_code_name(text) -> tuple[str, str] | None:
    """「CMP-108　STI-CMP 8号機」「ROB-821（ウェーハソーター 1号機）」→ (設備番号, 設備名)。分けられなければ None。

    1つのセルに「対象設備」「使用設備」として番号と名前をまとめて書く帳票用。名前の側が番号だけ（「CMP-101 / CMP-102」）なら分けない。
    """
    s = unicodedata.normalize("NFKC", str(text or "")).strip()
    if not s or "\n" in s:
        return None
    for pattern, code_group, name_group in _CODE_NAME_PATTERNS:
        m = pattern.match(s)
        if m:
            code, name = m[code_group].strip(), m[name_group].strip()
            if name and not _CODE_ONLY.match(name) and not all(_CODE_ONLY.match(p) for p in re.split(r"[\s,、/]+", name) if p):
                return code, name
    return None


def paren_stripped(norm: str) -> str:
    """末尾の括弧書きを除いた正規化ラベル。「停止時間(分)」→「停止時間」。括弧が無ければ ""。"""
    m = _TRAILING_PAREN_RE.search(norm)
    if not m or m.start() == 0:
        return ""
    return norm[: m.start()]


_LABEL_SUFFIXES = ("内容", "欄", "日時")


def label_base(norm: str) -> str:
    """表記の揺れを比べるための見出しの基本形。末尾の括弧書きと「内容」「欄」「日時」を除く。変わらなければ ""。

    例: 「発生原因(なぜ起きたか)」→「発生原因」、「応急処置内容」→「応急処置」、「復旧完了日時」→「復旧完了」。
    """
    s = paren_stripped(norm) or norm
    for suffix in _LABEL_SUFFIXES:
        if s.endswith(suffix) and len(s) - len(suffix) >= 2:
            s = s[: -len(suffix)]
            break
    return s if s != norm else ""


def normalize_sheet_name(name) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", strip_invisible(str(name)))).lower()


def nfkc_value(text) -> str:
    """値の正規化（Markdown 出力用）。NFKC＋行内の連続空白を1つに畳む。改行と日本語間の空白は残す。

    ラベル比較用の normalize_label とは別物（こちらは空白を消さない）。
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", strip_invisible(str(text))).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
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
    if not m or not m[1].strip() or not _unit_like(m[2].strip()):
        return str(label or "").strip(), ""
    return m[1].strip(), normalize_unit(m[2])


def _unit_like(text: str) -> bool:
    if text.isascii() or text in _JA_UNITS or len(text) == 1:
        return True
    return normalize_unit(text) != text  # 別名の一覧にあるもの


def _number_text(text) -> str:
    """数値として読む前の正規化: NFKC・桁区切りのカンマを除く・「▲」「△」の負号を「-」にする。"""
    s = unicodedata.normalize("NFKC", str(text or "")).replace(",", "").strip()
    return _MINUS_MARK_RE.sub("-", s)


def is_plain_number(text) -> bool:
    """「626」「1,032」「▲5」のような数字だけの値か。"""
    return re.fullmatch(r"-?(?:\d+(?:\.\d+)?|\.\d+)", _number_text(text)) is not None


def value_unit(text) -> str:
    """「1.5時間」→ "時間"。数値＋単位の形でなければ ""。"""
    m = _VALUE_UNIT_RE.match(_number_text(text))
    return normalize_unit(m[1]) if m else ""


def written_unit(text) -> str:
    """値に書かれた単位。「1.5時間」に加えて「約90分」「595分（9.9h）」のような前後に言葉のある値からも、
    最初の数値（to_number が読む数値）の直後の単位を返す。単位らしくなければ ""。

    「3時間40分」「09:30-12:45」は to_number が換算するので、ここでは単位を返さない。
    """
    s = _number_text(text)
    whole = value_unit(s)
    if whole or _DURATION_RE.search(s) or _TIME_RANGE_RE.search(s) or _HMM_RE.match(s):
        return whole
    r = _NUM_RANGE_RE.search(s)
    if r:
        # 「10〜20分」: 範囲の後ろの数値の単位を見る（先頭の「10」の直後の「〜」を単位としない）
        m = _UNIT_AFTER_RE.match(s, r.end())
        return normalize_unit(m[1]) if m and _unit_like(m[1]) else ""
    first = _NUM_RE.search(s)
    m = _UNIT_AFTER_RE.match(s, first.end()) if first else None
    if not m or not _unit_like(m[1]):
        return ""
    return normalize_unit(m[1])


_FORMAT_LITERAL_RE = re.compile(r'"([^"]*)"|\\(.)')


def format_unit(number_format) -> str:
    """セルの表示形式に書かれた単位（「#,##0"分"」→ 分、「"¥"#,##0」→ 円）。単位らしい文字が無ければ ""。

    画面では「3,095分」と見えるのに、セルの値は 3095 だけのことがある。その単位を値に書かれた単位として扱うために使う。
    """
    fmt = str(number_format or "")
    first = fmt.split(";", 1)[0]  # 正の数の書式だけを見る
    literal = "".join(a or b for a, b in _FORMAT_LITERAL_RE.findall(first)).strip()
    literal = unicodedata.normalize("NFKC", literal).strip()
    if (not literal or any(ch.isdigit() for ch in literal) or len(literal) > 4 or not _unit_like(literal)
            or not any(ch.isalpha() or ch in "¥$€℃%" for ch in literal)):
        return ""
    return normalize_unit(literal)


def numeric_unit(text, unit: str = "") -> str:
    """数値項目の単位。帳票の種類の設定があればそれ、無ければ書かれた値（「390分」「1,032分」「3時間40分」）から。

    どちらからも決まらなければ ""（単位不明）。単位を勝手に決めない（既定値は持たない）。
    """
    u = normalize_unit(unit)
    if u:
        return u
    s = _number_text(text)
    if _DURATION_RE.search(s) or _TIME_RANGE_RE.search(s) or _HMM_RE.match(s):
        return "分"  # to_number が「3時間40分」「09:30-12:45」「2:45」を分に換算する
    return written_unit(s)


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
    if isinstance(value, time):
        return value.strftime("%H:%M" if value.second == 0 else "%H:%M:%S")
    if isinstance(value, timedelta):
        # Excel の [h]:mm のセル。「1 day, 1:30:00」でなく画面の表示どおり「25:30」にする
        seconds = round(value.total_seconds())
        sign, seconds = ("-" if seconds < 0 else ""), abs(seconds)
        hours, rest = divmod(seconds, 3600)
        minutes, secs = divmod(rest, 60)
        return f"{sign}{hours}:{minutes:02d}" + (f":{secs:02d}" if secs else "")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = strip_invisible(str(value)).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def pick_checked(text) -> tuple[str | None, str | None] | None:
    """「□重大　■大　□中」→ ("大", None)。チェックボックス表記でなければ None。

    1行に □/■ などの印が2つ以上あるときだけ解釈する。複数チェックは「、」でつなぐ。
    どれにもチェックが無ければ (None, 警告)。
    """
    if not isinstance(text, str) or "\n" in text.strip():
        return None
    options = _CHECKBOX_RE.findall(text)
    if len(options) < 2:
        return None
    picked = [option for mark, option in options if mark in CHECKED_MARKS]
    if not picked:
        return None, "チェックの入った選択肢がありません"
    return "、".join(picked), None


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
        # 時刻が入っていれば残す（「発生日時」の 23:08 を黙って落とさない）
        return (raw.date().isoformat() if raw.time() == time(0) else raw.strftime("%Y-%m-%d %H:%M")), None
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
        year = _ERA_BASE[m[1]] + (1 if m[2] == "元" else int(m[2]))
        parsed = _safe_date(year, int(m[3]), int(m[4]))
        if parsed:
            return _with_clock(parsed, s, m.end())
    m = _DATE_RE.search(s)
    if m:
        parsed = _safe_date(int(m[1]), int(m[2]), int(m[3]))
        if parsed:
            return _with_clock(parsed, s, m.end())
    m = _NO_YEAR_RE.match(s)
    if m and 1 <= int(m[1]) <= 12 and 1 <= int(m[2]) <= 31:
        # 年は補わない（同じ帳票の別の項目やファイル名から推すと、外れたときに誤った日付を Markdown に書くことになる）
        return (text or None), "年が書かれていません。元のファイルを確かめて、2026-02-12 のように年から書いてください"
    return (text or None), "日付として解釈できません"


_TRAILING_MARKS = " )]」』】.。、"


def _with_clock(parsed: str, s: str, pos: int) -> tuple[str, str | None]:
    """日付の直後の時刻（「14:05」「13時41分」）を付ける。日付と時刻のあとに別の文字（「～7/12」）が続けば警告を付ける。

    時刻や範囲の終わりを黙って落とさない（落としたことが分かるよう要確認にする）。
    """
    lead = _CLOCK_LEAD_RE.match(s, pos)
    t = _CLOCK_RE.match(s, lead.end())
    if t:
        hour, minute = int(t[1]), int(t[2] or t[3])
        if hour < 24 and minute < 60:
            parsed += f" {hour:02d}:{minute:02d}"  # 日付の直後の「14:05」は残す
            pos = t.end()
        else:
            t = None
    if not t:
        pos = lead.end()
    rest = s[pos:].strip()
    if rest.strip(_TRAILING_MARKS):
        return parsed, f"日付のあとに「{rest}」が続いています。範囲や補足は元のファイルで確かめてください"
    return parsed, None


def _safe_date(y: int, m: int, d: int) -> str | None:
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def to_number(raw, text: str, unit: str = "") -> tuple[int | float | str | None, str | None]:
    """数値に変換する。「2.5時間」のような単位付きは数値部分を取り出して警告を付ける。

    値が「3時間40分」の形なら、unit（「分」「時間」）に換算する（220 / 3.67）。unit が空なら分にする。
    """
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return (int(raw) if float(raw).is_integer() else raw), None
    if isinstance(raw, time) and not isinstance(raw, datetime):
        return _minutes_number(text, raw.hour * 60 + raw.minute + raw.second / 60, unit)
    if isinstance(raw, timedelta):
        return _minutes_number(text, raw.total_seconds() / 60, unit)

    s = _number_text(text)
    hm = _HMM_RE.match(s)
    if hm:
        # 「2:45」は2時間45分（先頭の 2 だけを読まない）
        return _minutes_number(text, int(hm[1]) * 60 + int(hm[2]) + int(hm[3] or 0) / 60, unit)
    m = _DURATION_RE.search(s)
    if m and normalize_unit(unit) in ("分", "時間", ""):
        # 単位の決まっていない項目では分にする（「3時間40分」を 3 と読まない）
        target = normalize_unit(unit) or "分"
        minutes = float(m[1]) * 60 + float(m[2])
        number = minutes if target == "分" else round(minutes / 60, 2)
        value = int(number) if float(number).is_integer() else number
        return value, f"「{text}」を{target}に換算しました"
    r = _TIME_RANGE_RE.search(s)
    if r:
        return _time_range_number(text, s, r, unit)
    target = normalize_unit(unit)
    if target and _NUM_RANGE_RE.search(s) and written_unit(s) not in ("", target):
        # 「時間」の欄に「10～20分」: 先頭の 10 を時間として読まない
        return (text or None), f"「{text}」は範囲で、単位も項目の単位（{target}）と違います。数値として読み取れません"
    m = _NUM_RE.search(s)
    if not m:
        return (text or None), "数値として読み取れません"
    number = float(m[0])
    value = int(number) if number.is_integer() else number
    warning = None if m[0] == s else f"「{text}」から数値の部分だけを読み取りました"
    return value, warning


def _minutes_number(text: str, minutes: float, unit: str) -> tuple[int | float | str | None, str | None]:
    """時:分で書かれた時間（時刻のセル・[h]:mm のセル・「2:45」）を、項目の単位（分・時間。無ければ分）の数値にする。"""
    target = normalize_unit(unit) or "分"
    if target not in ("分", "時間"):
        return (text or None), "時:分の形です。数値として読み取れません"
    number = round(minutes, 2) if target == "分" else round(minutes / 60, 2)
    value = int(number) if float(number).is_integer() else number
    return value, f"「{text}」を{target}に換算しました"


def _time_range_number(text: str, s: str, r: re.Match, unit: str) -> tuple[int | float | str | None, str | None]:
    """「09:30-12:45（3.2h）」の形。添えた時間数があればそれを、無ければ範囲の長さを、項目の単位で返す。

    開始時刻（9）を数値として返さない。単位の決まっていない項目では分にする（「3時間40分」と同じ）。
    """
    target = normalize_unit(unit) or "分"
    if target not in ("分", "時間"):
        return (text or None), "時刻の範囲です。数値として読み取れません"
    rest = s[:r.start()] + " " + s[r.end():]
    m = _TIME_AMOUNT_RE.search(rest)
    if not m and "/" in r[0]:
        # 月日の付いた範囲は日をまたぐ日数が分からないので、終了−開始では出さない
        return (text or None), "時刻の範囲です。時間数が書かれていないので数値として読み取れません"
    if m:
        amount, written = float(m[1]), normalize_unit(m[2])
        minutes = amount * 60 if written == "時間" else amount
        note = f"「{text}」の時間数（{m[1]}{written}）を読み取りました"
    else:
        start = int(r[1]) * 60 + int(r[2])
        end = int(r[3]) * 60 + int(r[4])
        minutes = end - start if end >= start else end + 24 * 60 - start  # 日をまたぐ作業
        note = f"「{text}」は時刻の範囲なので、開始から終了までの時間を出しました"
    number = minutes if target == "分" else round(minutes / 60, 2)
    value = int(number) if float(number).is_integer() else number
    return value, f"{note}（{target}）。元のファイルと照らして確かめてください"


# ====================================================================================================
# 元 excel/image_detector.py
# Excel内の画像の「存在と位置」だけを検出する（画像の内容は解析しない）。
#
# openpyxlの画像読み込みはPillowに依存し、グループ化された図などを取りこぼすため、
# xlsx(zip)内の drawing XML を直接読む。
# ====================================================================================================

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
ANCHOR_TAGS = {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}


def detect_images(path: str | Path) -> list[dict]:
    """戻り値: [{"type": "image", "sheet": "...", "location": "H10:M20", "name": "..."}]"""
    results: list[dict] = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        workbook_part = next(
            (target for rtype, target in _rels(zf, names, "").values() if rtype.endswith("/officeDocument")),
            "xl/workbook.xml",
        )
        if workbook_part not in names:
            return results
        workbook = ET.fromstring(zf.read(workbook_part))
        workbook_rels = _rels(zf, names, workbook_part)
        # 複数のシートが同じ drawing を指すことがあるので、drawing ごとに1回だけ読む
        parsed: dict[str, ET.Element] = {}

        for sheet in workbook.findall("main:sheets/main:sheet", NS):
            sheet_part = workbook_rels.get(sheet.get(R_ID), ("", ""))[1]
            if sheet_part not in names:
                continue
            for rtype, drawing_part in _rels(zf, names, sheet_part).values():
                if not rtype.endswith("/drawing") or drawing_part not in names:
                    continue
                if drawing_part not in parsed:
                    parsed[drawing_part] = ET.fromstring(zf.read(drawing_part))
                results.extend(_images_in_drawing(parsed[drawing_part], sheet.get("name")))
    return results


def _rels(zf: zipfile.ZipFile, names: set[str], part: str) -> dict[str, tuple[str, str]]:
    """パーツのリレーションを {rId: (type, 解決済みパス)} で返す。"""
    folder, filename = posixpath.split(part)
    rels_path = posixpath.join(folder, "_rels", filename + ".rels")
    if rels_path not in names:
        return {}
    out = {}
    for rel in ET.fromstring(zf.read(rels_path)).findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out[rel.get("Id")] = (rel.get("Type", ""), resolved)
    return out


def _images_in_drawing(root: ET.Element, sheet_name: str) -> list[dict]:
    images = []
    for anchor in root:
        if anchor.tag.rsplit("}", 1)[-1] not in ANCHOR_TAGS:
            continue
        location = _anchor_location(anchor)
        for pic in anchor.findall(".//xdr:pic", NS):
            props = pic.find("xdr:nvPicPr/xdr:cNvPr", NS)
            images.append({
                "type": "image",
                "sheet": sheet_name,
                "location": location,
                "name": props.get("name", "") if props is not None else "",
            })
    return images


def _anchor_location(anchor: ET.Element) -> str:
    start = _marker(anchor.find("xdr:from", NS))
    end = _marker(anchor.find("xdr:to", NS))
    if start and end and end != start:
        return f"{start}:{end}"
    return start or ""


def _marker(marker: ET.Element | None) -> str:
    if marker is None:
        return ""
    col, row = marker.find("xdr:col", NS), marker.find("xdr:row", NS)
    if col is None or row is None:
        return ""
    return f"{get_column_letter(int(col.text) + 1)}{int(row.text) + 1}"


# ====================================================================================================
# 元 excel/workbook.py
# Excelの構造解析。結合セルを考慮した「値のあるセル」のグリッドを作る。
# ====================================================================================================

@dataclass
class Cell:
    """値を持つセル。結合セルの場合は左上セルが範囲全体を代表する。"""

    row: int
    col: int
    max_row: int
    max_col: int
    value: object
    text: str
    norm: str
    inline: tuple[str, str] | None
    bold: bool = False
    filled: bool = False
    fill: str = ""  # 塗りつぶし色の識別キー（"rgb:FFDDEEFF" / "theme:4:0.6" など。塗りなしは ""）
    alt_norm: str = ""  # 先頭の項番を除いたラベル（「３．暫定対策」→「暫定対策」）。無ければ ""
    part_norms: tuple[str, ...] = ()  # 「ライン／工程」のように2つのラベルをまとめた見出しの各部分
    fmt_unit: str = ""  # 数値のセルの表示形式に書かれた単位（「#,##0"分"」→「分」）。無ければ ""

    @property
    def coord(self) -> str:
        start = f"{get_column_letter(self.col)}{self.row}"
        if (self.max_row, self.max_col) == (self.row, self.col):
            return start
        return f"{start}:{get_column_letter(self.max_col)}{self.max_row}"

    @property
    def label_keys(self) -> tuple[str, ...]:
        """セル全体をラベルとみなすときの比較キー。"""
        return (self.norm, self.alt_norm) if self.alt_norm else (self.norm,)

    def matches(self, label_norms: set[str]) -> bool:
        """セル全体がラベル候補に一致するか（セル内ラベル「設備番号：EQ-001」は含まない）。"""
        return self.norm in label_norms or (bool(self.alt_norm) and self.alt_norm in label_norms)

    def is_label(self, label_norms: set[str]) -> bool:
        return self.matches(label_norms) or (self.inline is not None and self.inline[0] in label_norms)


class SheetGrid:
    def __init__(self, ws):
        self.name: str = ws.title
        self.hidden: bool = getattr(ws, "sheet_state", "visible") != "visible"
        self._bounds: dict[tuple[int, int], tuple[int, int, int, int]] = {}
        for rng in ws.merged_cells.ranges:
            bounds = (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    self._bounds[(r, c)] = bounds

        self.cells: dict[tuple[int, int], Cell] = {}
        self.max_row = self.max_col = 0
        for (r, c), xl in ws._cells.items():
            text = cell_text(xl.value)
            if not text:
                continue
            top, left, bottom, right = self.bounds(r, c)
            if (top, left) != (r, c):
                continue
            self.cells[(r, c)] = Cell(
                row=r, col=c, max_row=bottom, max_col=right,
                value=xl.value, text=text, norm=normalize_label(text), inline=split_inline(text),
                bold=bool(xl.font and xl.font.b),
                filled=bool(xl.fill and xl.fill.fill_type == "solid"),
                fill=fill_key(xl),
                alt_norm=section_stripped(text) if _label_like(xl.value, text) else "",
                part_norms=label_parts(text) if _label_like(xl.value, text) else (),
                fmt_unit=(format_unit(xl.number_format)
                          if isinstance(xl.value, (int, float)) and not isinstance(xl.value, bool) else ""),
            )
            self.max_row = max(self.max_row, bottom)
            self.max_col = max(self.max_col, right)

        self._by_norm: dict[str, list[Cell]] = defaultdict(list)
        self._by_inline: dict[str, list[Cell]] = defaultdict(list)
        self._by_base: dict[str, list[str]] = defaultdict(list)  # 括弧書きを除いたラベル → 元のキー
        self._by_variant: dict[str, list[str]] = defaultdict(list)  # 括弧書き・「内容」を除いたラベル → 元のキー
        self._by_part: dict[str, list[Cell]] = defaultdict(list)  # 「ライン／工程」の「ライン」「工程」
        for cell in self.text_cells():
            for key in cell.part_norms:
                self._by_part[key].append(cell)
            for key in cell.label_keys:
                self._by_norm[key].append(cell)
            if cell.inline:
                self._by_inline[cell.inline[0]].append(cell)
            for key in (*cell.label_keys, *(cell.inline[:1] if cell.inline else ())):
                base = paren_stripped(key)
                if base:
                    self._by_base[base].append(key)
                # 表記違いの照合は見出しらしいセル（塗りつぶし・太字・「ラベル：値」）だけ。値の「完了」を「完了日時」と見ない
                if cell.filled or cell.bold or (cell.inline and key == cell.inline[0]):
                    self._by_variant[label_base(key) or key].append(key)

    def bounds(self, row: int, col: int) -> tuple[int, int, int, int]:
        """(top, left, bottom, right) を返す。結合されていなければ自セルのみ。"""
        return self._bounds.get((row, col), (row, col, row, col))

    def cell_at(self, row: int, col: int) -> Cell | None:
        top, left, _, _ = self.bounds(row, col)
        return self.cells.get((top, left))

    def text_cells(self) -> list[Cell]:
        return [self.cells[key] for key in sorted(self.cells)]

    def resolve_labels(self, label_norms: set[str]) -> set[str]:
        """シートで探すラベルの集合。候補どおりのラベルが無いときだけ、末尾の括弧書きの違いを許す。

        例: 候補「停止時間」でシートに「停止時間(分)」だけがある（または逆）。「原因(推定)」と「原因(確定)」の
        ように括弧書きで区別しているラベルは、候補どおりのものがあればそちらだけを使う。
        """
        if any(n in self._by_norm or n in self._by_inline or n in self._by_part for n in label_norms):
            return label_norms
        bases = {paren_stripped(n) or n for n in label_norms}
        extra = {key for base in bases for key in self._by_base.get(base, [])}
        extra |= {base for base in bases if base in self._by_norm or base in self._by_inline}
        return label_norms | extra if extra else label_norms

    def label_variants(self, label_norms: set[str]) -> set[str]:
        """シートにある、候補と表記だけが違うラベル（候補どおりのラベルから値が取れなかったときに使う）。

        「発生原因」と「発生原因（なぜ起きたか）」、「応急処置」と「応急処置内容」。
        括弧書きどうしが違うもの（「原因（推定）」と「原因（確定）」）は別の項目なので含めない。
        """
        found = set()
        for norm in label_norms:
            qualified = bool(paren_stripped(norm))
            for key in self._by_variant.get(label_base(norm) or norm, []):
                if key in label_norms or (qualified and paren_stripped(key)):
                    continue
                found.add(key)
        return found

    def find_labels(self, label_norms: set[str]) -> list[Cell]:
        """ラベル候補に一致するセルを上→下、左→右の順で返す（セル内ラベルも含む）。

        明細表の列見出し（時系列表の「対応内容」、部品表の「備考」など）は、同じラベルが表の外にもあれば後回しにする。
        """

        label_norms = self.resolve_labels(label_norms)
        hits = {}
        for norm in label_norms:
            for cell in self._by_norm.get(norm, []) + self._by_inline.get(norm, []) + self._by_part.get(norm, []):
                hits[(cell.row, cell.col)] = cell
        if len(hits) > 1:
            headers = table_header_keys(self)
            return [hits[key] for key in sorted(hits, key=lambda k: (k in headers, k))]
        return [hits[key] for key in sorted(hits)]


def _label_like(value, text: str) -> bool:
    return isinstance(value, str) and "\n" not in text and len(text) <= MAX_LABEL_LENGTH + 4


def fill_key(xl) -> str:
    """セルの塗りつぶし色を比較用の文字列にする（ラベル欄と同じ色かの判定に使う）。"""
    fill = getattr(xl, "fill", None)
    if fill is None or fill.fill_type != "solid":
        return ""
    color = fill.fgColor
    try:
        if color.type == "theme":
            return f"theme:{color.theme}:{round(color.tint or 0, 3)}"
        return f"{color.type}:{color.value}"
    except (AttributeError, TypeError, ValueError):
        return "solid"


@dataclass
class WorkbookInfo:
    path: Path | None   # 読んだファイル（メモリから読んだときは None）
    grids: dict[str, SheetGrid]
    images: list[dict]
    date1904: bool = False  # 1904年基準のブック（日付シリアル値の起点が違う）
    # 計算結果が保存されていない数式のセル {シート名: {(行, 列)}}（openpyxl などで書いたブック）。
    # data_only で読むと空になるので、「値が空」でなく「数式の結果が無い」と知らせるのに使う
    uncached_formulas: dict[str, set[tuple[int, int]]] = field(default_factory=dict)

    @property
    def sheet_names(self) -> list[str]:
        return list(self.grids)

    def images_in(self, sheet_names: list[str]) -> list[dict]:
        return [img for img in self.images if img["sheet"] in sheet_names]


def load_workbook_info(path: str | Path | io.BytesIO) -> WorkbookInfo:
    """ブックを読む。パスのほか、メモリの中のブック（BytesIO）も渡せる。

    読み方は同じで、どこから読むかだけが違う。帳票登録は見本の Excel をサーバーに置かずに読むので
    BytesIO を渡す（2026-09-21 の利用者の指示「見本のExcelは置かずに、設定だけ保持する」）。
    """
    def source():
        if isinstance(path, (str, Path)):
            return path
        path.seek(0)   # 同じブックを3回読む（openpyxl・画像・数式）ので、そのつど先頭へ戻す
        return path

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(source(), data_only=True)
    try:
        grids = {ws.title: SheetGrid(ws) for ws in wb.worksheets}
        date1904 = wb.epoch == MAC_EPOCH
    finally:
        wb.close()
    return WorkbookInfo(path=Path(path) if isinstance(path, (str, Path)) else None, grids=grids,
                        images=detect_images(source()), date1904=date1904,
                        uncached_formulas=uncached_formula_cells(source()))


def uncached_formula_cells(path: str | Path) -> dict[str, set[tuple[int, int]]]:
    """数式（<f>）があって計算結果（<v>）が無いセルを、シートごとに返す。読めないブックは空。"""
    out: dict[str, set[tuple[int, int]]] = {}
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            workbook_part = next((target for rtype, target in _rels(zf, names, "").values()
                                  if rtype.endswith("/officeDocument")), "xl/workbook.xml")
            if workbook_part not in names:
                return out
            workbook_rels = _rels(zf, names, workbook_part)
            for sheet in ET.fromstring(zf.read(workbook_part)).findall("main:sheets/main:sheet", NS):
                part = workbook_rels.get(sheet.get(R_ID), ("", ""))[1]
                if part not in names:
                    continue
                data = zf.read(part)
                if b"<f" not in data:  # 数式の無いシート（ほとんど）は読み直さない
                    continue
                cells = _uncached_in_sheet(data)
                if cells:
                    out[sheet.get("name")] = cells
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, ValueError):
        return {}
    return out


def _uncached_in_sheet(xml: bytes) -> set[tuple[int, int]]:
    cells: set[tuple[int, int]] = set()
    main = "{" + NS["main"] + "}"
    for _, elem in ET.iterparse(io.BytesIO(xml)):
        if elem.tag != main + "c":
            continue
        formula, value = elem.find(main + "f"), elem.find(main + "v")
        # 空文字の結果（t="str" の空の <v>）は計算済み
        if formula is not None and (value is None or (not value.text and elem.get("t") != "str")) and elem.get("r"):
            letters, row = coordinate_from_string(elem.get("r"))
            cells.add((row, column_index_from_string(letters)))
        elem.clear()
    return cells


# ====================================================================================================
# 元 excel/tables.py
# 帳票の中の明細表（交換部品・時系列・チェックシートなど）を読む。
#
# 見出しのセル（アンカー: 「■ 交換部品」「使用部品」など）の下、または縦に結合した見出しの右にある
# 「列見出しの行」を見つけ、その下の行を空行・見出し欄（列見出しと同じ塗りつぶし色のセル）まで読む。
# 1行 = 1明細。値は {"columns": [列見出し...], "rows": [[セルの文字列...], ...]} で持つ。
# ====================================================================================================

MAX_TABLE_ROWS = 200
MAX_HEADER_CHARS = 30
MIN_COLUMNS = 2
# 自動検出（見本からの候補づくり・見出しの優先順位）で、アンカーの無い表とみなす最小の列数・行数
MIN_COLUMNS_WITHOUT_ANCHOR = 3
MIN_ROWS_WITHOUT_ANCHOR = 2
# 列見出しが「同じ表」とみなす一致率（共通の列見出し ÷ どちらかにある列見出し）
SAME_COLUMNS_RATIO = 0.5

_NUMBER_LIKE = re.compile(r"^[\d\s,.\-/:+%]+$")
# 判定記号と、その凡例の1組（「◎：主要因」「○=要因の可能性あり」「－：対象外」）
_JUDGE_MARKS = "◎○◯〇△▲▽×✕✖"
_LEGEND_PAIR = re.compile(rf"[{_JUDGE_MARKS}ー―－\-]\s*[:：=＝]")
MAX_LEGEND_CHARS = 60
_SEQ_HEADERS = {"no", "№", "#", "項番", "番号", "順", "順番"}
_TOTAL_WORDS = {"合計", "小計", "総計", "計", "合計数", "総合計"}
# 合計行の先頭（「合計」「小計」「部品費計」「工数計」「部品費合計」）。
# 「〜計」を全部合計とみなすと「温度計」「圧力計」「設計」などの行で読み取りが止まり、Markdownでも消えるので、
# 合計の言い方だけにする（export.formats も同じ決まりを使う）
TOTAL_LABEL_RE = re.compile(r"^(?:合計|小計|総計|総合計|合計数|計|.{0,6}(?:合計|小計)|"
                            r".{0,6}(?:費|工数|個数|件数|台数|本数|枚数|金額|額)計)$")
_NONE_MARKS = {"なし", "無し", "該当なし", "特になし", "-", "ー", "―", "‐", "/", "〃"}
# 「上の行と同じ」の記号（″ は NFKC で ′′ になる）。明細表の値では上の行の同じ列の値に置き換える
_DITTO_MARKS = {"〃", "′′", "同上", "仝"}
# 見出しらしい書き出し: 「■ 交換部品」「【時系列】」「▼ 回答欄」「1. 時系列」「A. 機構部」
_HEADING_START = re.compile(r"^(?:[■□◆◇●▼▽▶►【\[]|\d{1,2}[.)、](?!\d)|[a-z][.)、](?![a-z])|d[1-8](?=\D))")


@dataclass
class Table:
    anchor: Cell | None
    header: list[Cell]
    rows: list[list[Cell | None]]
    blank_rows: int = 0  # 飛ばした「No だけの空き行」の数

    @property
    def has_room(self) -> bool:
        """行を足せる明細表の形か（2行以上ある／空き行がある／「■」「1.」の見出しの下／縦結合の見出しに余りの行がある）。

        「影響｜停止時間｜影響ロット」のような見出しの行＋値の行1つは、明細表でなく項目の並びとみなす。
        """
        if len(self.rows) >= 2 or self.blank_rows:
            return True
        anchor = self.anchor
        if anchor is None:
            return False
        if _HEADING_START.match(_compact(anchor.text)):
            return True
        return anchor.max_row > anchor.row and anchor.max_row > self.last_row

    @property
    def last_row(self) -> int:
        """読み取った範囲の最後の行。"""
        return max([c.max_row for row in self.rows for c in row if c is not None]
                   or [max(h.max_row for h in self.header)])

    @property
    def coord(self) -> str:
        first = self.header[0]
        right = max(c.max_col for c in self.header)
        from openpyxl.utils import get_column_letter

        return f"{get_column_letter(first.col)}{first.row}:{get_column_letter(right)}{self.last_row}"

    def to_value(self) -> dict | None:
        """保存用の値。連番だけの No 列は落とす。行が無ければ None。"""
        columns = [" ".join(h.text.split()) for h in self.header]
        rows = [[c.text if c is not None else "" for c in row] for row in self.rows]
        if len(columns) > 1 and _is_seq_column(columns[0], [r[0] for r in rows]):
            columns = columns[1:]
            rows = [_drop_seq_cell(r) for r in rows]
        # 空の行と「なし」「－」だけの行（明細なしの記入）は落とす
        rows = [r for r in rows if any(v.strip() and _compact(v) not in _NONE_MARKS for v in r)]
        if not rows:
            return None
        _resolve_ditto(rows)
        return {"columns": columns, "rows": rows}


def _resolve_ditto(rows: list[list[str]]) -> None:
    """「〃」「同上」のセルを、上の行の同じ列の値にする（縦結合のセルと同じ扱い。1行だけ読んでも意味が通るように）。

    上の行の値が空か、最初の行なら書かれたままにする。
    """
    for above, row in zip(rows, rows[1:]):
        for i, v in enumerate(row):
            if _compact(v) in _DITTO_MARKS and i < len(above) and above[i].strip()                     and _compact(above[i]) not in _DITTO_MARKS:
                row[i] = above[i]


def find_table(grid: SheetGrid, anchor: Cell, direction: str = "auto", stop_labels: set[str] | None = None) -> Table | None:
    """アンカーのセルから明細表を探す。

    - 縦に結合したアンカー（「使用部品」が4行分など）は、右隣の列見出しの行を先に見る（読む行はアンカーの範囲内）
    - それ以外は、アンカーの下2行以内で、アンカーの列から始まる列見出しの行を探す
    """
    stop_labels = stop_labels or set()
    vertical = anchor.max_row > anchor.row
    if direction in ("auto", "right") and vertical:
        header = header_cells(grid, anchor.row, anchor.max_col + 1)
        if len(header) >= MIN_COLUMNS:
            return read_table(grid, anchor, header, stop_labels, last_row=anchor.max_row)
    if direction in ("auto", "below"):
        marked = bool(_HEADING_START.match(_compact(anchor.text)))
        row, last = anchor.max_row + 1, anchor.max_row + 2
        while row <= last:
            starts = [c for c in _cells_in_row(grid, row) if anchor.col <= c.col <= anchor.max_col]
            if starts:
                header = header_cells(grid, row, starts[0].col)
                if len(header) >= MIN_COLUMNS:
                    return read_table(grid, anchor, header, stop_labels)
                if marked and last == anchor.max_row + 2 and _label_value_row(grid, row):
                    # 「■ 水平展開」の下に「展開区分｜☑同型機…」の1行があり、その下に列見出しが来る様式
                    last, row = last + 1, row + 1
                    continue
                break
            if any(c.col <= anchor.max_col and c.max_col >= anchor.col for c in _cells_in_row(grid, row)):
                break
            row += 1
    return None


def header_cells(grid: SheetGrid, row: int, start_col: int) -> list[Cell]:
    """row 行目の start_col から右へ、すき間なく並ぶ見出しらしいセル（高さが先頭のセルと同じもの）。"""
    cells: list[Cell] = []
    col = start_col
    while col <= grid.max_col:
        cell = grid.cells.get((row, col))
        if cell is None or not _header_like(cell) or (cells and cell.max_row != cells[0].max_row):
            break
        cells.append(cell)
        col = cell.max_col + 1
    return cells


def read_table(grid: SheetGrid, anchor: Cell | None, header: list[Cell], stop_labels: set[str],
               last_row: int | None = None) -> Table:
    """列見出しの下の行を読む。空行、見出し欄（列見出しと同じ色の文字列・ラベル）、合計行の後で止める。

    縦に結合したデータセル（点検部位など）は、結合範囲の各行に同じ値を入れる。
    連番（No）だけが書かれた空き行（様式の固定行数の余り）は飛ばして続ける。
    """
    fills = {h.fill for h in header if h.fill}
    row = max(h.max_row for h in header) + 1
    end = min(grid.max_row, last_row or grid.max_row, row + MAX_TABLE_ROWS - 1)
    rows: list[list[Cell | None]] = []
    blank_rows = 0
    seq_header = _norm_header(header[0]) in _SEQ_HEADERS
    while row <= end:
        cells: list[Cell | None] = []
        seen: set[tuple[int, int]] = set()
        for h in header:
            found = None
            for col in range(h.col, h.max_col + 1):
                found = grid.cell_at(row, col)
                if found is not None:
                    break
            if found is not None and (found.row, found.col) in seen:
                found = None
            if found is not None:
                seen.add((found.row, found.col))
            cells.append(found)
        fresh = [c for c in cells if c is not None and c.row == row]
        if not fresh:
            break
        total = any(_is_total(c) for c in fresh)
        if not total and any(_is_stop(c, fills, stop_labels, header[0].col) for c in fresh):
            break
        if seq_header and all(c is cells[0] for c in fresh) and _NUMBER_LIKE.match(fresh[0].norm):
            row += 1  # No だけの空き行
            blank_rows += 1
            continue
        rows.append(cells)
        if total and any(c.filled for c in fresh):
            break
        row = min(c.max_row for c in fresh) + 1
    return Table(anchor, header, rows, blank_rows)


def detect_tables(grid: SheetGrid) -> list[Table]:
    """シート内の明細表を自動で見つける（見本からの候補づくり・見出しの優先順位に使う）。

    列見出し = 塗りつぶしのある短い文字列セルが2つ以上すき間なく並ぶ行。
    アンカー（真上の見出し、または左の縦結合の見出し）があれば1行以上、無ければ3列以上・2行以上の表だけを採る。
    """
    cached = getattr(grid, "_tables", None)
    if cached is not None:
        return cached
    tables: list[Table] = []
    used: set[tuple[int, int]] = set()
    for cell in grid.text_cells():
        if (cell.row, cell.col) in used or not cell.filled or not _header_like(cell):
            continue
        # 行の先頭の縦結合セル（「使用部品」）は列見出しにならず（高さが違う）、右隣からの並びのアンカーになる
        header = _contiguous([c for c in header_cells(grid, cell.row, cell.col) if c.filled])
        if len(header) < MIN_COLUMNS:
            continue
        used.update((c.row, c.col) for c in header)
        anchor, last_row = _find_anchor(grid, header)
        table = read_table(grid, anchor, header, set(), last_row=last_row)
        if anchor is not None and table.rows:
            tables.append(table)
        elif len(header) >= MIN_COLUMNS_WITHOUT_ANCHOR and len(table.rows) >= MIN_ROWS_WITHOUT_ANCHOR:
            tables.append(table)
    grid._tables = tables
    return tables


def find_table_by_columns(grid: SheetGrid, columns: list[str], keep=None) -> Table | None:
    """列見出しの半分以上が columns と同じ、行のある明細表（最も似ているもの）。keep(表) が偽の表は使わない。"""

    wanted = {normalize_label(c) for c in columns} - {""}
    if len(wanted) < MIN_COLUMNS:
        return None
    best, best_score = None, 0.0
    for table in detect_tables(grid):
        have = {h.norm for h in table.header}
        score = len(have & wanted) / len(have | wanted)
        if (score >= SAME_COLUMNS_RATIO and score > best_score and table.to_value() is not None
                and (keep is None or keep(table))):
            best, best_score = table, score
    return best


def stacked_tables(grid: SheetGrid, table: Table, stop_labels: set[str] | None = None,
                   limit: int = 10) -> list[Table]:
    """読み取った範囲のすぐ下に続く、同じ形（同じ列位置・同じ塗りつぶし色）の列見出しの行とその行。

    特性要因図のように「人｜機械」「材料｜方法」「測定｜環境」と同じ形の小さな表が縦に並ぶ帳票では、
    組ごとに列見出しが変わるため、1つの表として読むと2組目から先が落ちる。組ごとに読んで
    merge_table_values でまとめる（design.md 6.1「積み重なった列見出し」）。
    """
    stop_labels = stop_labels or set()
    if not table.header:
        return []
    fills = {h.fill for h in table.header if h.fill}
    shape = [(h.col, h.max_col) for h in table.header]
    last = table.last_row
    blocks: list[Table] = []
    while len(blocks) < limit:
        header = None
        for row in (last + 1, last + 2):
            if row > grid.max_row:
                break
            found = _contiguous([c for c in header_cells(grid, row, shape[0][0]) if c.filled])
            if [(h.col, h.max_col) for h in found] == shape \
                    and (not fills or {h.fill for h in found if h.fill} & fills):
                header = found
                break
        if header is None:
            break
        block = read_table(grid, None, header, stop_labels)
        blocks.append(block)
        last = max(block.last_row, header[0].max_row)
    return blocks


def merge_table_values(values: list[dict | None]) -> dict | None:
    """積み重なった組（同じ形の列見出しが縦に並ぶ表）の値を、1つの明細表の値にまとめる。

    列見出しは出てきた順に並べ、同じ見出しは同じ列にそろえる（組ごとに見出しが変わる特性要因図は
    「人｜機械｜材料｜方法｜測定｜環境」の6列になり、各行は自分の組の列だけ埋まる）。
    同じ見出しが続くだけの組（ページをまたいで見出しを繰り返す様式）は、そのまま行が増える。
    """
    parts = [v for v in values if is_table_value(v) and v.get("rows")]
    if not parts:
        return next((v for v in values if is_table_value(v)), None)
    if len(parts) == 1:
        return parts[0]

    columns: list[str] = []
    keys: list[str] = []
    rows: list[list[str]] = []
    for part in parts:
        used: set[int] = set()
        slots: list[int] = []
        for i, name in enumerate(part["columns"]):
            key = normalize_label(name) or f"\x00{len(keys)}"  # 空の見出しは他とまとめない
            slot = next((n for n, k in enumerate(keys) if k == key and n not in used), None)
            if slot is None:
                keys.append(key)
                columns.append(name)
                slot = len(keys) - 1
            used.add(slot)
            slots.append(slot)
        for row in part["rows"]:
            cells = [""] * len(keys)
            for i, cell in enumerate(row):
                if i < len(slots):
                    cells[slots[i]] = cell
            rows.append(cells)
    return {"columns": columns, "rows": [row + [""] * (len(columns) - len(row)) for row in rows]}


# ---- 区切りの見出し（区画）------------------------------------------------------------------
# 発行側と回答側で同じ意味の欄が並ぶ帳票（「処置内容」と「▼ 回答欄」の下の「暫定対策（処置）」）で、
# 項目がどちら側の欄かを見分けるための区画。区画 = 「■」「▼」「【】」「1.」などで始まる見出しのセルから、
# 右は同じ高さにある次の見出しの手前まで（無ければシートの右端まで）、下は次の見出しまで。
MAX_SECTION_CHARS = 60
_SECTION_TAIL = re.compile(r"[(（].*$")


@dataclass
class Section:
    key: str        # 区画の名前（比較用。「▼ 回答欄（宛先部署にて記入…）」「【回答欄】」→「回答」）
    cell: Cell
    right: int      # 区画の右端の列


def section_key(cell: Cell) -> str:
    """区画の見出しのセルの比較用の名前（section_name）。"""
    return section_name(cell.text)


def section_name(text) -> str:
    """区画の見出しの比較用の名前。印・項番・括弧書き・末尾の「欄」「内容」を除く（版ごとの書き方の違いを吸収する）。

    「▼ 回答欄（宛先部署にて記入し…）」「【回答欄】」「回答欄」→「回答」。帳票の種類の画面で入力された区画もこれでそろえる。
    何度通しても同じ結果になる（「■ 処置内容欄」も「処置内容」も「処置」）。
    """

    title = _title_text(text)
    norm = normalize_label(_SECTION_TAIL.sub("", title)) or normalize_label(title)
    norm = section_stripped(norm) or norm
    # 「処置内容欄」→「処置内容」→「処置」のように末尾の語が重なることがあるので、変わらなくなるまで除く。
    # 保存した区画をもう一度この関数に通しても同じ名前になる（見出しと保存値が必ずそろう）。
    while base := label_base(norm):
        norm = base
    return norm


def _section_start(cell: Cell) -> bool:
    """区画の見出しか（塗りつぶし・太字で、「■」「▼」「【】」「1.」などの印で始まる1行の短い文字列）。"""
    if not isinstance(cell.value, str) or "\n" in cell.text or cell.inline is not None:
        return False
    if len(cell.norm) > MAX_SECTION_CHARS or pick_checked(cell.text) is not None:
        return False
    return (cell.filled or cell.bold) and bool(_HEADING_START.match(_compact(cell.text)))


def sections(grid: SheetGrid) -> list[Section]:
    """シートの区画（見出しの上→下、左→右の順）。"""
    cached = getattr(grid, "_sections", None)
    if cached is not None:
        return cached
    heads = [c for c in grid.text_cells() if _section_start(c)]
    out: list[Section] = []
    for head in heads:
        # 同じ高さ（行が重なる）で右にある次の見出しの手前までを、この区画の横幅にする
        right_heads = [h.col for h in heads if h.col > head.max_col and h.row <= head.max_row and h.max_row >= head.row]
        # 上の行で右側に始まった区画（右上の「【回答欄】」の下に左の「■ 発行部署記入欄」がある版）の列は、
        # その区画のまま（左の区画の横幅に入れない）。右上の区画の列に、あとから別の見出しが無いときだけ
        for h in heads:
            if h.col > head.max_col and h.max_row < head.row and not any(
                    h.max_row < k.row < head.row and k.max_col >= h.col for k in heads):
                right_heads.append(h.col)
        out.append(Section(section_key(head), head, min(right_heads, default=grid.max_col + 1) - 1))
    grid._sections = out
    return out


def _enclosing(grid: SheetGrid, cell: Cell) -> list[Section]:
    """セルを含む区画（見出しがセルより上か同じ行で、横幅にセルの列が入るもの）。近い見出し（下・右）から順に。"""
    found = [sec for sec in sections(grid) if sec.cell.row <= cell.row and sec.cell.col <= cell.col <= sec.right]
    return sorted(found, key=lambda sec: (sec.cell.row, sec.cell.col), reverse=True)


def section_of(grid: SheetGrid, cell: Cell) -> str:
    """セルが入っている区画の名前。どの区画にも入らなければ ""。区画の見出しのセル自身はその区画に入る。"""
    found = _enclosing(grid, cell)
    return found[0].key if found else ""


def _heading_level(cell: Cell) -> int:
    """区画の見出しの階層。「■」「▼」「【】」などの印=0、「1.」「D1」=1、「a.」=2（数字の小見出しは印の見出しの中）。"""
    head = _compact(cell.text)
    if re.match(r"^[a-z][.)、]", head):
        return 2
    return 1 if re.match(r"^(?:\d|d[1-8])", head) else 0


def sections_of(grid: SheetGrid, cell: Cell) -> list[str]:
    """セルが入っている区画の名前を、内側（一番近い見出し）から外側へ。

    「▼ 回答欄」の下の「1. 暫定対策」の中のセルは ["暫定対策", "回答"]。外側の見出しは、それより内側の見出しより
    上の階層（印の見出しは数字の小見出しの外側）のものだけ。同じ階層の見出しが下にあれば、上の区画はそこで終わっている。
    """
    out: list[str] = []
    level = None
    for sec in _enclosing(grid, cell):
        lv = _heading_level(sec.cell)
        if level is None or lv < level:
            out.append(sec.key)
            level = lv
    return out


def table_header_keys(grid: SheetGrid) -> set[tuple[int, int]]:
    """明細表の列見出しのセル位置。"""
    return {(c.row, c.col) for t in detect_tables(grid) for c in t.header}


def list_header_keys(grid: SheetGrid) -> set[tuple[int, int]]:
    """行を足せる明細表の列見出しのうち、その列だけに収まる値が2行以上並ぶもののセル位置。1つの値の項目のラベルにはしない。

    押印欄（承認｜確認｜作成 の下に印と日付、その下に横長の注記）は含めない。見出し（アンカー）の無い表は、
    押印と日付の2行と区別するため3行以上並ぶときだけ。
    """
    keys = set()
    for table in detect_tables(grid):
        if not table.has_room:
            continue
        # 見出しの無い表でも、日付だけの行の無い表（水平展開先の「確認｜対象設備｜設備名…」が2行）は一覧とみなす
        need = 2 if table.anchor is not None or not _has_date_row(table) else 3
        for i, h in enumerate(table.header):
            cells = [row[i] for row in table.rows if i < len(row) and row[i] is not None]
            if sum(1 for c in cells if c.col >= h.col and c.max_col <= h.max_col) >= need:
                keys.add((h.row, h.col))
    return keys


def _has_date_row(table: Table) -> bool:
    """押印欄（承認｜確認｜作成 の下に印、その下に「6/5」のような日付）らしい、日付・数字だけの行があるか。"""
    for row in table.rows:
        cells = [c for c in row if c is not None]
        if cells and all(not isinstance(c.value, str) or _NUMBER_LIKE.match(c.norm) for c in cells):
            return True
    return False


def seq_header(cell: Cell) -> bool:
    """連番の列見出し（No / № / 項番）。表の中にあれば、報告書の「No.」欄とは別物。"""
    return _norm_header(cell) in _SEQ_HEADERS


def section_heading(cell: Cell) -> bool:
    """塗りつぶしのある「■ 現象・処置」「▼ 回答欄（…）」のような区切りの見出し。値にはしない。"""
    return (cell.filled and isinstance(cell.value, str) and "\n" not in cell.text and cell.inline is None
            and bool(_HEADING_START.match(_compact(cell.text))) and pick_checked(cell.text) is None)


def heading_like(cell: Cell) -> bool:
    """見出しらしいセル（塗りつぶし・太字・「■」「1.」などで始まる短い文字列）。"""
    if not isinstance(cell.value, str) or len(cell.norm) > MAX_HEADER_CHARS + 10:
        return False
    marked = bool(_HEADING_START.match(_compact(cell.text)))
    if cell.inline and not marked:  # 「区分：同型機」はラベルと値。「D1：チームの結成」は見出し
        return False
    return cell.filled or cell.bold or marked


def legend_text(text) -> bool:
    """判定記号の凡例か（「◎：主要因　○：影響あり　×：検証の結果 要因でない」「（◎主要因／○寄与要因／×否定）」）。

    値でも見出しでもない飾りなので、項目の値の候補にも、表の見出しを探すときの「右にある値」にもしない。
    """
    s = " ".join(str(text or "").split())
    if not s or len(s) > MAX_LEGEND_CHARS:
        return False
    return sum(s.count(m) for m in _JUDGE_MARKS) >= 3 or len(_LEGEND_PAIR.findall(s)) >= 2


def table_title(cell: Cell) -> str:
    """アンカーの表示名。「■ 経緯（時系列）」→「経緯（時系列）」。"""
    return _title_text(cell.text)


def _title_text(raw) -> str:
    text = _one_line(raw)
    text = re.sub(r"^[■□◆◇●▼▽▶►・*※\s]+", "", text)
    if text.startswith(("【", "[")) and text.endswith(("】", "]")):
        text = text[1:-1]
    return text.strip(" :：") or _one_line(raw)


# ---- 値（{"columns": [...], "rows": [[...]]}）の扱い -----------------------------------------

def is_table_value(value) -> bool:
    return isinstance(value, dict) and isinstance(value.get("rows"), list)


def clean_table_value(value) -> dict | None:
    """画面から戻ってきた表の値をそろえる。列数に合わせて行を詰め、空の行を落とす。

    行が無くても列見出しがあれば {"columns": [...], "rows": []} を返す（確認・修正画面の表を残し、行を入れ直せるように）。
    列見出しも行も無ければ None。Markdown 側は行の無い明細表を出さない。
    """
    if not is_table_value(value):
        return None
    columns = [str(c if c is not None else "").strip() for c in (value.get("columns") or [])]
    rows = []
    for row in value["rows"]:
        if not isinstance(row, list):
            continue
        cells = ["" if c is None else str(c).replace("\r\n", "\n").strip() for c in row]
        while len(cells) > len(columns):
            columns.append(f"列{len(columns) + 1}")
        cells += [""] * (len(columns) - len(cells))
        if any(cells):
            rows.append(cells)
    if not rows and not any(c for c in columns):
        return None
    return {"columns": columns, "rows": rows}


def drop_seq_column(value):
    """連番だけの No 列を落とした表の値（読み取りの Table.to_value と同じ決まり）。

    読み取れなかった明細表は、帳票の種類の列見出し（No を含む）のまま入力欄に出るので、
    手で入れた行だけ Markdown が「- No.: 1／…」になってしまう。出すときに読み取った表と形をそろえる。
    値が 1,2,3… の連番でない No 列（本当の品番・項番）は、読み取り側と同じく残す。
    """
    if not is_table_value(value):
        return value
    rows = [r for r in value["rows"] if isinstance(r, list)]
    columns = [str(c if c is not None else "") for c in (value.get("columns") or [])]
    if len(columns) < 2 or not rows or any(len(r) < 2 for r in rows):
        return value
    cells = [[str(c if c is not None else "") for c in r] for r in rows]
    if not _is_seq_column(columns[0], [r[0] for r in cells]):
        return value
    return {**value, "columns": columns[1:], "rows": [_drop_seq_cell(r) for r in cells]}


def parse_table_text(text: str) -> tuple[dict | None, str | None]:
    """確認画面の入力（JSON）を表の値にする。戻り値: (値, 警告)。空なら (None, None)。"""
    import json

    if not (text or "").strip():
        return None, None
    try:
        data = json.loads(text)
    except ValueError:
        return None, "表の形式が正しくないため、変更を反映できませんでした"
    if not is_table_value(data):
        return None, "表の形式が正しくないため、変更を反映できませんでした"
    return clean_table_value(data), None


def table_row_items(value, row: list) -> list[tuple[str, str]]:
    """1行を (列見出し, 値) の組にする。空のセルは除く。"""
    columns = list(value.get("columns") or [])
    items = []
    for i, cell in enumerate(row):
        text = "" if cell is None else str(cell).strip()
        if text:
            items.append((columns[i] if i < len(columns) and columns[i] else f"列{i + 1}", text))
    return items


def table_text_lines(value) -> list[str]:
    """画面表示用: 1行を「品番: X／品名: Y」の1行にする。"""
    if not is_table_value(value):
        return []
    return ["／".join(f"{k}: {' '.join(v.split())}" for k, v in table_row_items(value, row))
            for row in value["rows"]]


# ---- 内部 ---------------------------------------------------------------------------

def _find_anchor(grid: SheetGrid, header: list[Cell]) -> tuple[Cell | None, int | None]:
    first, right = header[0], header[-1].max_col
    row = first.row
    if first.col > 1:
        left = grid.cell_at(row, first.col - 1)
        if left is not None and left.max_col == first.col - 1 and left.max_row > row and left.row <= row \
                and heading_like(left):
            return left, left.max_row
    above, last = row - 1, row - 2
    skipped = False
    while above >= max(1, last):
        # 判定記号の凡例（「◎：主要因　○：影響あり　×：否定」）は飾りなので、見出しの右にあっても数に入れない
        in_span = [c for c in _cells_in_row(grid, above)
                   if c.col <= right and c.max_col >= first.col and not legend_text(c.text)]
        if not in_span:
            above -= 1
            continue
        cell = in_span[0]
        marked = bool(_HEADING_START.match(_compact(cell.text)))
        # 右に値が並ぶセル（「区分｜☑同型機」）はアンカーにしない。「■」「1.」で始まる見出しは右に凡例があってもよい
        alone = len(in_span) == 1 or marked
        if alone and cell.col <= first.col + 1 and heading_like(cell) and (marked or not skipped) \
                and len(header_cells(grid, above, cell.col)) < MIN_COLUMNS:
            return cell, None
        if not skipped and _label_value_row(grid, above):
            # 「■ 水平展開」と列見出しの間に「展開区分｜☑同型機…」の1行がある様式: その上の「■」「1.」の見出しを見る
            skipped, last, above = True, last - 1, above - 1
            continue
        break
    return None, None


def _label_value_row(grid: SheetGrid, row: int) -> bool:
    """「展開区分｜☑同型機　□類似設備…」のような、見出し欄1つとその値だけの1行か（表の見出しと列見出しの間に挟まる行）。"""
    cells = _cells_in_row(grid, row)
    if len(cells) != 2:
        return False
    label, value = cells
    return (heading_like(label) and label.filled and not _HEADING_START.match(_compact(label.text))
            and not value.filled and value.col == label.max_col + 1 and value.max_row == label.max_row)


def _cells_in_row(grid: SheetGrid, row: int) -> list[Cell]:
    """row 行目から始まるセル（左上がその行にあるもの）を左から順に。"""
    by_row = getattr(grid, "_cells_by_row", None)
    if by_row is None:
        by_row = {}
        for cell in grid.text_cells():
            by_row.setdefault(cell.row, []).append(cell)
        grid._cells_by_row = by_row
    return by_row.get(row, [])


def _header_like(cell: Cell) -> bool:
    if not isinstance(cell.value, str) or cell.inline is not None:
        return False
    if len(cell.norm) > MAX_HEADER_CHARS or _NUMBER_LIKE.match(cell.norm):
        return False
    if legend_text(cell.text):  # 見出しの右に置かれた判定記号の凡例は列見出しではない
        return False
    return pick_checked(cell.text) is None


def _contiguous(cells: list[Cell]) -> list[Cell]:
    out: list[Cell] = []
    for c in cells:
        if out and c.col != out[-1].max_col + 1:
            break
        out.append(c)
    return out


def _is_stop(cell: Cell, fills: set[str], stop_labels: set[str], first_col: int) -> bool:
    """表の終わりを示すセル: 列見出しと同じ色の文字列、表の左端の色付きの短い文字列（「確認期間」などの項目欄）、
    色付き・太字のラベル。表の途中の色付きの値（NG の強調など）では止めない。"""
    if not isinstance(cell.value, str):
        return False
    if cell.fill and cell.fill in fills:
        return True
    if cell.filled and cell.col <= first_col and len(cell.norm) <= MAX_HEADER_CHARS:
        return True
    return (cell.filled or cell.bold) and cell.is_label(stop_labels)


def _is_total(cell: Cell) -> bool:
    if not isinstance(cell.value, str):
        return False
    return cell.norm in _TOTAL_WORDS or (cell.filled and len(cell.norm) <= 8 and bool(TOTAL_LABEL_RE.match(cell.norm)))


def _drop_seq_cell(row: list[str]) -> list[str]:
    """No 列を落とした行。合計行の「合計」が No 列にあれば、次の列が空のときだけそこへ移す。"""
    first, rest = row[0].strip(), row[1:]
    if first and not first.isdigit() and not rest[0].strip():
        return [first, *rest[1:]]
    return rest


def _is_seq_column(header: str, values: list[str]) -> bool:
    if _compact(header).strip(".:：") not in _SEQ_HEADERS:
        return False
    numbers = [v.strip() for v in values if v.strip() and _compact(v) not in _TOTAL_WORDS]
    if not numbers or not all(n.isdigit() for n in numbers):
        return False
    return [int(n) for n in numbers] == list(range(1, len(numbers) + 1))


def _norm_header(cell: Cell) -> str:
    return cell.norm.strip(".:：")


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text))).lower()


def _one_line(text) -> str:
    return " ".join(unicodedata.normalize("NFKC", str(text or "")).split())


# ====================================================================================================
# 元 excel/extractor.py
# テンプレート定義に従ってExcelから項目を抽出する。
#
# セル番地を固定せず、「ラベル候補に一致するセルを探す → その右 or 下にある値を取る」
# という方式なので、帳票ごとにセル位置がズレていても同じJSONになる。
# ====================================================================================================

MAX_RIGHT_STEPS = 10
MAX_TEXT_ROWS = 30
# 様式番号の記載（「様式MT-031(1) Rev.1」「製造部 設備保全課 様式MT-031 Rev.2」）。欄外の注記で、項目の値ではない。
# 先頭に無くても注記なので途中でも見るが、長い本文の中の「様式」で値を打ち切らないよう短い1行のセルだけ。
_FORM_NUMBER_RE = re.compile(r"様式\s*[:：]?\s*[A-Za-z0-9]")
MAX_FORM_NUMBER_CHARS = 60
# 日付だけの値（時刻をつなぐ対象）と、時刻だけのセル（「12:07」「12:07:30」「9時5分」「9時」）
_DATE_ONLY_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_CLOCK_ONLY_RE = re.compile(r"^\s*(\d{1,2})\s*(?::\s*(\d{2})(?::\d{2})?|時(?:\s*(\d{1,2})\s*分)?)\s*$")


@dataclass
class FieldResult:
    field_name: str
    display_name: str
    data_type: str
    value: object = None
    sheet: str | None = None
    label_cell: str | None = None
    value_cell: str | None = None
    label_found: bool = False
    warning: str | None = None
    edited: bool = False
    unit: str = ""
    # 帳票の種類で決めた単位。unit は読み取った値（「14.9h」）で書き替わることがあるので、
    # 手で直したときに元の単位で判断し直せるよう別に持つ
    spec_unit: str = ""
    rag_output: str = "show"
    # 明細表の列見出し。値が空でも確認・修正画面で表として入力できるようにするために持つ
    table_columns: list[str] = field(default_factory=list)
    # 明細表: まとめて読んだ「同じ形の列見出しの組」の数（1 なら普通の表。2以上は確認画面で知らせる）
    table_blocks: int = 1


def extract_document(info: WorkbookInfo, pattern: PatternDef, sheet_names: list[str]) -> dict:
    stop_labels = pattern.label_norms() | DICTIONARY_NORMS
    results = [_extract_field(info, fd, sheet_names, stop_labels) for fd in pattern.fields]
    extraction = {
        "pattern": {
            "id": pattern.id,
            "name": pattern.name,
            "version": pattern.version,
            "version_no": pattern.version_no,
            "image_processing": pattern.image_processing,
            "title_fields": list(pattern.title_fields),
            "md_options": dict(pattern.md_options or {}),
            # 「- 数量: 単価」のように値が別の欄の見出し語になった行を Markdown に出さないための照合用
            "labels": sorted(pattern.output_label_norms()),
        },
        "sheets": sheet_names,
        "fields": [asdict(r) for r in results],
        "attachments": info.images_in(sheet_names),
    }
    refresh_summary(extraction)
    return extraction


def is_blank_value(value) -> bool:
    """値が空か。明細表は行が無ければ空（列見出しだけ残した値も空とみなす）。"""
    if is_table_value(value):
        return not value.get("rows")
    return value is None or (isinstance(value, str) and not value.strip())


def refresh_summary(extraction: dict) -> None:
    """fields から values を作り直す。"""
    extraction["values"] = {f["field_name"]: f["value"] for f in extraction["fields"]}


def apply_manual_values(extraction: dict, form) -> None:
    """プレビュー画面で人が修正した値を反映する。"""
    for f in extraction["fields"]:
        key = f"value-{f['field_name']}"
        if key not in form:
            continue
        text = form[key].replace("\r\n", "\n").strip()
        if f["data_type"] == "table":
            value, warning = parse_table_text(text)
            if warning:  # 形が壊れた入力では値を変えない
                f["warning"] = warning
                continue
        elif f["data_type"] == "date" and text:
            value, warning = to_date(text, text)
        elif f["data_type"] == "number" and text:
            # 帳票の種類の単位で判断し直す（読み取りで書き替わった unit を引き継がない）。古いデータは unit で代用
            spec = f.get("spec_unit", f.get("unit") or "") or ""
            current = f.get("unit") or ""
            if current and current != spec and not written_unit(text) and is_plain_number(text):
                # 「625分」と読んだ項目（種類の単位は時間）の数字だけを直した: 画面の入力欄に単位は出ないので、
                # 読み取った単位のまま（書かれた単位として）判断する。種類の単位に黙って変えると60倍の値になる
                text = f"{text}{current}"
            value, warning = to_number(text, text, spec)
            unit, warning = number_unit(value, text, spec, f["field_name"], f["display_name"], warning)
            # 値と単位が同じでも、入力に警告（「1.5～3時間」の範囲など）が出たとき・前の警告が消えるときは直したものとする
            if value != f["value"] or unit != (f.get("unit") or "") or warning or f.get("warning"):
                f["value"], f["unit"], f["warning"], f["edited"] = value, unit, warning, True
            continue
        else:
            value, warning = (text or None), None
        # 日付も同じ: 「2026-09-14 24:30」は日付だけ同じでも時刻が不正なので、警告を残して手で修正にする
        if value != f["value"] or (f["data_type"] == "date" and value is not None and (warning or f.get("warning"))):
            f["value"], f["warning"], f["edited"] = value, warning, True
    refresh_summary(extraction)


def _sheet_order(fd: FieldDef, sheet_names: list[str]) -> list[str]:
    """その項目を登録したシートを先に見る。

    同じ見出し語が2枚のシートに並ぶ帳票（1次報告／2次報告、8D報告など）で、渡された順に
    シートを回ると、2枚目に登録した項目が1枚目の同名の欄の値を読んでしまう。
    読むシート（sheet_names）に入っているときだけ先頭に回す（画面で外したシートは見ない）。
    """
    own = getattr(fd, "sheet_name", "") or ""
    names = [own] if own in sheet_names else []
    return names + [n for n in sheet_names if n not in names]


def _extract_field(info: WorkbookInfo, fd: FieldDef, sheet_names: list[str], stop_labels: set[str]) -> FieldResult:
    result = FieldResult(fd.field_name, fd.display_name, fd.data_type,
                         unit=fd.unit, spec_unit=fd.unit or "", rag_output=fd.rag_output)
    if fd.data_type == "table":
        result.table_columns = [c for c in (fd.table_columns or []) if str(c).strip()]
        return _extract_table_field(info, fd, sheet_names, stop_labels, result)
    if getattr(fd, "cell", "") and not any(str(c).strip() for c in fd.candidates):
        # 見出しのない「値だけ」の項目: 見本でクリックしたセルの番地から読む
        if not _extract_at_cell(info, fd, sheet_names, result):
            result.warning = "値のセルが空です"
        return result
    for name in _sheet_order(fd, sheet_names):
        grid = info.grids.get(name)
        if grid is None:
            continue
        label, values, inline_value = locate_value(grid, fd, stop_labels)
        if label is None:
            continue
        if not result.label_found:
            result.label_found, result.sheet, result.label_cell = True, name, label.coord
            empty_label = (name, label)
        if not values:
            continue
        result.sheet, result.label_cell = name, label.coord
        result.value_cell = values[0].coord if inline_value is None else label.coord
        error = excel_error(inline_value if inline_value is not None else values[0].text)
        if error and (inline_value is not None or len(values) == 1):
            # 「#REF!」「#DIV/0!」を値にしない（数値の項目で 0 と単位「!」に化けない）。要確認にする
            result.warning = f"{EXCEL_ERROR_WARNING}（{error}）です。元のファイルで数式を確かめて、値を入力してください"
            return result
        result.value, result.warning = _convert(fd.data_type, values, inline_value, info.date1904, fd.unit)
        if fd.data_type == "date" and inline_value is None and len(values) == 1:
            _join_clock_beside(grid, values[0], stop_labels, result)
        if fd.data_type == "number":
            text = inline_value if inline_value is not None else values[0].text
            if inline_value is None and values[0].fmt_unit and not written_unit(text):
                # 表示形式「#,##0"分"」のセル: 画面に見えている単位を、値に書かれた単位として扱う
                text = f"{text}{values[0].fmt_unit}"
            _apply_number_unit(result, fd, text)
        return result
    if getattr(fd, "cell", "") and _extract_at_cell(info, fd, sheet_names, result, stop_labels):
        # 見出しが見つからない帳票では、見本でクリックしたセルの番地を控えとして読む
        return result
    if not result.label_found:
        result.warning = "ラベルが見つかりません"
    elif _uncached_beside(info, *empty_label):
        result.warning = UNCACHED_FORMULA_WARNING
    else:
        result.warning = "ラベルはありますが値が空です"
    return result


UNCACHED_FORMULA_WARNING = "数式の計算結果が保存されていません。Excelで開いて保存し直すか、値を入力してください"


def _extract_at_cell(info: WorkbookInfo, fd: FieldDef, sheet_names: list[str], result: FieldResult,
                     stop_labels: set[str] = frozenset()) -> bool:
    """見本でクリックしたセルの番地から値を読む（見出しで見つからなかったときの控え）。

    様式が違う帳票では同じ番地に別の欄が来るので、見出しらしいセル（他の欄の見出し）や、
    その型として読めない文字は値にしない。読めないときは空のままにする（当て推量で埋めない）。
    """

    # その項目を登録したシートがブックに在れば、そのシートだけを見る。空だったからといって
    # 別のシートの同じ番地を読むと、まったく関係のない欄の値が入ってしまう。
    # ほかのシートを見るのは、シートの名前が変わっていて見つからないときだけ
    own = getattr(fd, "sheet_name", "") or ""
    names = [own] if own in info.grids else list(sheet_names)
    for name in names:
        grid = info.grids.get(name)
        if grid is None:
            continue
        key = cell_key(grid, fd.cell)
        cell = grid.cells.get(key) if key else None
        if cell is None or not cell.text.strip():
            continue
        if stop_labels and ({cell.norm, cell.alt_norm} - {""}) & stop_labels:
            continue  # その番地には別の欄の見出しが来ている（様式が違う）
        value, warning = _convert(fd.data_type, [cell], None, info.date1904, fd.unit)
        if stop_labels and warning:
            continue  # その型として読めない＝別の欄の値
        result.sheet, result.value_cell = name, cell.coord
        result.value, result.warning = value, warning
        if fd.data_type == "number":
            _apply_number_unit(result, fd, cell.text)
        return True
    return False


def _uncached_beside(info: WorkbookInfo, sheet: str, label: Cell) -> bool:
    """ラベルのすぐ右かすぐ下のセルが、計算結果の保存されていない数式か（空に見えるが値が無いのではない）。"""
    cells = getattr(info, "uncached_formulas", {}).get(sheet)
    return bool(cells) and ((label.row, label.max_col + 1) in cells or (label.max_row + 1, label.col) in cells)


def _join_clock_beside(grid: SheetGrid, cell: Cell, stop_labels: set[str], result: FieldResult) -> None:
    """日付のセルのすぐ右のセルに時刻だけが書かれた様式（「2023-05-23｜12:07」）は、時刻を日付につなぐ。

    つなぐのは紛れの無いときだけ: 日付だけ読めた（警告なし）、すぐ右（すき間なし・同じ高さ）のセルが時刻だけ、
    その右が空か見出し欄（範囲「9:00｜～｜17:00」や「9:00｜17:00」はつながない）。
    """
    if not isinstance(result.value, str) or result.warning or not _DATE_ONLY_RE.match(result.value):
        return
    beside = grid.cells.get((cell.row, cell.max_col + 1))
    if beside is None or beside.max_row != cell.max_row:
        return
    clock = _clock_only(beside)
    if clock is None:
        return
    after = scan_right(grid, beside, set())  # その右で最初に値のあるセル
    if after and not (is_stop_cell(after[0], beside, stop_labels) or (after[0].filled and isinstance(after[0].value, str)
                                                                    and len(after[0].norm) <= MAX_LABEL_LENGTH)):
        return
    result.value = f"{result.value} {clock}"
    result.value_cell = f"{cell.coord.split(':')[0]}:{beside.coord.split(':')[-1]}"


def _clock_only(cell: Cell) -> str | None:
    """時刻だけのセル（時刻の値、「12:07」「9時5分」）なら「HH:MM」。違えば None。"""
    if isinstance(cell.value, time):
        return f"{cell.value.hour:02d}:{cell.value.minute:02d}"
    if not isinstance(cell.value, str):
        return None
    m = _CLOCK_ONLY_RE.match(unicodedata.normalize("NFKC", cell.value))
    if not m:
        return None
    hour, minute = int(m[1]), int(m[2] if m[2] is not None else (m[3] or 0))
    return f"{hour:02d}:{minute:02d}" if hour < 24 and minute < 60 else None


def number_unit(value, text: str, spec_unit: str, field_name: str, display_name: str,
                warning: str | None) -> tuple[str, str | None]:
    """数値項目の単位と警告を決める。戻り値: (単位, 警告)。

    帳票の種類に単位が無ければ書かれた値（「390分」）から補い、どちらにも無い項目のうち
    「単位で意味が変わる」ものは要確認にする（単位は勝手に決めない）。
    読み取り・AI補完・手修正のどの経路からでも同じ判断になるよう、ここ1か所で決める。
    """
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return spec_unit, warning  # 数値として読めなかった値は、別の警告が出ている
    unit = numeric_unit(text, spec_unit)
    # 「約90分」「595分（9.9h）」のように前後に言葉がある値も、書かれた単位を見る（種類の単位で上書きしない）
    written = written_unit(text)
    if written and unit and written != unit:
        # 「停止時間（分）」の欄に「14.9h」と書かれた帳票。勝手に換算せず、書かれたとおりの単位で出して要確認にする
        return written, (f"この項目の単位は「{unit}」ですが、値には「{written}」と書かれています。"
                         f"書かれたとおり「{written}」として出します。どちらが正しいか確かめてください")
    if written and written == unit and value_unit(text) and warning and "数値の部分だけ" in warning:
        warning = None  # 「390分」の「分」は単位として取り込んだので、読み落としではない
    if unit or warning:
        return unit, warning
    hint = ambiguous_unit_hint(field_name, display_name)
    if hint:
        # 先頭の「単位が書かれていません」は views/forms._UNIT_WARNING_PREFIXES が見ているので変えない
        example = hint.split("か", 1)[0]  # 「分か時間かで…」→「分」
        warning = (f"単位が書かれていません（{hint}）。値に単位を付けて入力してください（例: {value}{example}）。"
                   "これから読み取る帳票のためには「帳票の種類」の画面でこの項目の単位を決めてください"
                   "（読み取り済みの帳票には反映されません）")
    return unit, warning


def _apply_number_unit(result: FieldResult, fd: FieldDef, text: str) -> None:
    result.unit, result.warning = number_unit(result.value, text, fd.unit, fd.field_name, fd.display_name,
                                              result.warning)


def _extract_table_field(info: WorkbookInfo, fd: FieldDef, sheet_names: list[str], stop_labels: set[str],
                         result: FieldResult) -> FieldResult:
    """明細表: 見出し（アンカー）の下か右の列見出しから、行を読む。"""
    header_found = False
    for name in _sheet_order(fd, sheet_names):
        grid = info.grids.get(name)
        if grid is None:
            continue
        anchor, table = locate_table(grid, fd, stop_labels)
        if anchor is None:
            continue
        if not result.label_found:
            result.label_found, result.sheet, result.label_cell = True, name, anchor.coord
        header_found = header_found or table is not None
        if table is None:
            continue
        # すぐ下に同じ形の列見出しが続く表（特性要因図の「人｜機械」「材料｜方法」…）は、組ごとに読んで1つにまとめる
        blocks = stacked_tables(grid, table, stop_labels - grid.resolve_labels(fd.label_norms()))
        value = merge_table_values([table.to_value(), *(b.to_value() for b in blocks)])
        if value is None or not value.get("rows"):
            continue
        result.sheet, result.label_cell, result.value_cell = name, anchor.coord, table.coord
        result.value = value
        result.table_blocks = 1 + len(blocks)
        return result
    if not result.label_found:
        result.warning = "ラベルが見つかりません"
    elif not header_found:
        result.warning = "見出しはありますが、その下（または右）に表の列見出しが見つかりません"
    else:
        result.warning = "表に行がありません"
    return result


def _columns_differ(table: Table, columns: list[str]) -> bool:
    """読み取った表の列見出しが、見本で見た列見出し（fd.table_columns）とほとんど重ならないか。

    「探す見出し」が表そのものの列見出しと同じ語（F3 の D6 の「実施内容」）だと、そのセルを見出しとみなして
    1行下のデータ行を列見出しとして読んでしまう（先頭の明細と一部の列が落ちる）。列がずれていることは
    この比較で分かる。
    """

    wanted = {normalize_label(c) for c in columns} - {""}
    if len(wanted) < MIN_COLUMNS or not table.header:
        return False
    have = {h.norm for h in table.header}
    return len(have & wanted) / len(have | wanted) < SAME_COLUMNS_RATIO


def locate_table(grid: SheetGrid, fd: FieldDef, stop_labels: set[str]) -> tuple[Cell | None, Table | None]:
    """戻り値: (見出しのセル, 表)。行のある表を優先し、無ければ最初に見つかった見出しと表（行なし）を返す。"""
    norms = grid.resolve_labels(fd.label_norms())
    first: tuple[Cell | None, Table | None] = (None, None)
    labels = grid.find_labels(norms)
    if fd.section:
        # 区画（例: 回答欄）が決めてあれば、区画の中の見出しを先に見る（無ければ今までどおりシート全体）
        labels = sorted(labels, key=lambda c: not _in_section(grid, fd, c))
    for cell in labels:
        table = find_table(grid, cell, fd.direction, stop_labels - norms)
        if table is not None and table.to_value() is not None:
            if _columns_differ(table, fd.table_columns):
                # 列が見本と合わない＝見出し行を1行取り違えている。見本の列見出しに合う表を優先する
                better = _table_by_columns(grid, fd)
                if better is not None:
                    return cell, better
            return cell, table
        if first[0] is None or (first[1] is None and table is not None):
            first = (cell, table)
    # 見出しの書き方が違う帳票: 見本で見た列見出しと並びが似た表を探す
    table = _table_by_columns(grid, fd)
    if table is not None:
        return table.anchor or table.header[0], table
    return first


def _in_section(grid: SheetGrid, fd: FieldDef, cell: Cell) -> bool:
    """セルが項目の区画（fd.section）の中か（区画の中の小見出しの中も含む）。"""
    return fd.section in sections_of(grid, cell)


def _table_by_columns(grid: SheetGrid, fd: FieldDef) -> Table | None:
    """見本の列見出しに合う表。区画が決めてあれば、列見出しが区画の中にある表を先に探す。"""
    if fd.section:
        table = find_table_by_columns(grid, fd.table_columns, keep=lambda t: _in_section(grid, fd, t.header[0]))
        if table is not None:
            return table
    return find_table_by_columns(grid, fd.table_columns)


def locate_value(grid: SheetGrid, fd: FieldDef, stop_labels: set[str]) -> tuple[Cell | None, list[Cell], str | None]:
    """戻り値: (ラベルセル, 値セルのリスト, セル内ラベルの場合の値テキスト)

    明細表の項目は、行のある表が見つかったときだけ値セルのリストに見出しのセルを入れる（有無の判定用）。
    """
    if fd.data_type == "table":
        anchor, table = locate_table(grid, fd, stop_labels)
        found = anchor is not None and table is not None and table.to_value() is not None
        return anchor, ([anchor] if found else []), None
    headers = table_header_keys(grid)
    first_label = None
    for label_set, cells in _ordered_label_sets(_limit_to_section(grid, fd, _label_sets(grid, fd)), headers):
        # それぞれの組で、すぐ隣に値があるラベルを先に見る（押印欄の縦書き「発信部署」の2行下の日付を、「発信部署」の値にしない）
        for allow_gap in (False, True):
            for cell in cells:
                first_label = first_label or cell
                found = _value_at(grid, fd, cell, label_set, stop_labels, headers, allow_gap)
                if found is not None:
                    return found
    return first_label, [], None


def label_hits(grid: SheetGrid, fd: FieldDef, stop_labels: set[str]) -> list[Cell]:
    """項目の見出しに当たるセルのうち、値が読めるものすべて（locate_value が見る順。区画では絞らない）。

    見本から帳票の種類を作るときに、同じ意味の欄がシートの複数の区画にあるか（発行側と回答側）を調べるのに使う。
    """
    headers = table_header_keys(grid)
    hits: list[Cell] = []
    for label_set, cells in _ordered_label_sets(_label_sets(grid, fd), headers):
        for cell in cells:
            if cell not in hits and any(_value_at(grid, fd, cell, label_set, stop_labels, headers, gap) is not None
                                        for gap in (False, True)):
                hits.append(cell)
    return hits


def _label_sets(grid: SheetGrid, fd: FieldDef) -> list[tuple[set[str], list[Cell]]]:
    """探すラベルの組と、それぞれに当たるセル。

    クリックした（＝見本で見た）見出し → 辞書が足した言い換えの見出し
    → 表記だけ違うラベル（「発生原因（なぜ起きたか）」「応急処置内容」）の順。
    """

    norms = grid.resolve_labels(fd.label_norms())
    headers = table_header_keys(grid)
    lists = list_header_keys(grid)
    # 候補の1つ目は、画面でクリックした見出し（見本で実際に見た見出し）。辞書が足した言い換え
    # （「担当者」に対する「報告者」「記入者」…）より先に探す。同じシートに両方が並ぶ帳票で、
    # 先に出てくる言い換えの欄（「報告者」）の値を、クリックした欄（「担当者」）の値にしないため。
    primary = {normalize_label(fd.search_labels()[0])} & norms
    # (ラベルの組, 表記違いか) の順に見る。表記違いの組では区切りの見出しを使わない
    label_sets = [(primary, False), (norms - primary, False),
                  (grid.label_variants(fd.label_norms()) - norms, True)]
    if fd.field_name in COMBINED_EQUIPMENT_PARTS:
        # 設備番号・設備名: 最後に「対象設備：CMP-108　STI-CMP 8号機」のような番号と名前をまとめた欄も見る
        label_sets.append((COMBINED_EQUIPMENT_NORMS - norms - label_sets[2][0], True))
    found_sets: list[tuple[set[str], list[Cell]]] = []
    for label_set, variant in label_sets:
        # 行を足せる明細表の列見出し（「設備No｜設備名」の下に何行も並ぶ表）は、1つの値のラベルにしない。
        # 表記違いのラベルでは、区切りの見出し（「■ 承認欄」）も使わない
        # 明細表の連番の列見出し（No）も使わない（報告書の「No.」欄と取り違えない）
        cells = [c for c in grid.find_labels(label_set)
                 if (c.row, c.col) not in lists and not (variant and section_heading(c))
                 and not ((c.row, c.col) in headers and seq_header(c))] if label_set else []
        found_sets.append((label_set, cells))
    return found_sets


def _ordered_label_sets(found_sets: list[tuple[set[str], list[Cell]]],
                        headers: set[tuple[int, int]]) -> list[tuple[set[str], list[Cell]]]:
    """見る順に並べる。明細表の列見出しにあるラベルは、表の外のラベル（表記違い・まとめ欄を含む）をすべて見た後に使う。"""
    attempts = [(s, [c for c in cells if (c.row, c.col) not in headers]) for s, cells in found_sets]
    in_headers = [(s, [c for c in cells if (c.row, c.col) in headers]) for s, cells in found_sets]
    return attempts + in_headers


def _limit_to_section(grid: SheetGrid, fd: FieldDef,
                      found_sets: list[tuple[set[str], list[Cell]]]) -> list[tuple[set[str], list[Cell]]]:
    """項目に区画（fd.section。例: 回答欄）が決めてあり、その区画の中に探す見出しがあれば、区画の中の見出しだけにする。

    回答欄の見出しの値が空（未回答）でも、発行側の同じ意味の欄の値は読まない（どちら側の欄かは帳票の種類で決める）。
    区画の無い版・区画の中に見出しが無い版は、今までどおりシート全体の見出しを使う。
    """
    if not fd.section:
        return found_sets
    # 区画の中の小見出し（「▼ 回答欄」の下の「1. 暫定対策」）の中の見出しも、回答欄の中とみなす
    inside = [(s, [c for c in cells if fd.section in sections_of(grid, c)]) for s, cells in found_sets]
    return inside if any(cells for _, cells in inside) else found_sets


def _value_at(grid: SheetGrid, fd: FieldDef, cell: Cell, norms: set[str], stop_labels: set[str],
              headers: set[tuple[int, int]], allow_gap: bool) -> tuple[Cell, list[Cell], str | None] | None:
    """ラベルのセル1つについて値を探す。見つからなければ None。"""
    part = next((i for i, key in enumerate(cell.part_norms) if key in norms), None)
    if cell.matches(norms) or part is not None:
        if fd.direction == "same_cell":
            return None
        multi = fd.data_type == "text" and part is None
        values: list[Cell] = []
        # 見本でクリックした向き（右・下）は当てにする順で、見つからなければもう一方も見る。
        # 同じ種類でも版によって値の位置が変わる帳票があるため（今までの「自動」は 右→下）
        first_below = fd.direction == "below"
        if not first_below:
            values = scan_right(grid, cell, stop_labels)
        if not values:
            values = scan_below(grid, cell, stop_labels, multi=multi, allow_gap=allow_gap)
        if not values and first_below:
            values = scan_right(grid, cell, stop_labels)
        if values and (values[0].row, values[0].col) in headers:
            return None  # 明細表の見出し（「■ 暫定対策」の下の No｜処置内容…）を値にしない
        if values and _combined_equipment(fd, cell.label_keys):
            values = _code_name_part(fd, values[0])
        if values and part is not None:
            # 「ライン／工程」→「L6／STI-CMP」: 値も同じ数に分かれるときだけ、その部分を値にする
            pieces = split_combined_value(values[0].text, len(cell.part_norms))
            if pieces is None:
                return None
            values = [replace(values[0], value=pieces[part], text=pieces[part])]
        return (cell, values, None) if values else None
    # 「設備番号　：IMP-603」のように1つのセルに見出しと値が入った版もある。見本で右・下をクリックして
    # いても、その向きに値が無ければセル内の値を読む（版によって書き方が変わる帳票があるため）
    if not allow_gap and cell.inline is not None and cell.inline[0] in norms:
        if _combined_equipment(fd, cell.inline[:1]):
            pieces = split_code_name(cell.inline[1])
            return (cell, [cell], pieces[COMBINED_EQUIPMENT_PARTS[fd.field_name]]) if pieces else None
        return cell, [cell], cell.inline[1]
    return None


def _combined_equipment(fd: FieldDef, keys) -> bool:
    """設備番号・設備名の項目を、番号と名前をまとめた欄（「対象設備」）から読むところか。"""
    return fd.field_name in COMBINED_EQUIPMENT_PARTS and any(k in COMBINED_EQUIPMENT_NORMS for k in keys)


def _code_name_part(fd: FieldDef, cell: Cell) -> list[Cell]:
    """「CMP-108　STI-CMP 8号機」から項目に当たる部分（番号 or 名前）。分けられない値なら []（このラベルは使わない）。"""
    pieces = split_code_name(cell.text)
    if pieces is None:
        return []
    piece = pieces[COMBINED_EQUIPMENT_PARTS[fd.field_name]]
    return [replace(cell, value=piece, text=piece)]


def scan_right(grid: SheetGrid, label: Cell, stop_labels: set[str]) -> list[Cell]:
    """ラベルの右側で最初に値があるセルを探す。別のラベルに当たったら値なし。"""
    row, col = label.row, label.max_col + 1
    for _ in range(MAX_RIGHT_STEPS):
        if col > grid.max_col:
            break
        top, left, _, right = grid.bounds(row, col)
        cell = grid.cells.get((top, left))
        if cell is not None:
            return [] if is_stop_cell(cell, label, stop_labels) else [cell]
        col = right + 1
    return []


def scan_below(grid: SheetGrid, label: Cell, stop_labels: set[str], multi: bool, allow_gap: bool = True) -> list[Cell]:
    """ラベルの下の値を探す。文章型(multi)は空行か別ラベルまでの連続セルをまとめて取る。

    allow_gap: ラベルの直下が空のとき、1行空けた下を見るか。
    """
    col, row = label.col, label.max_row + 1
    values: list[Cell] = []
    skipped_empty = False
    while row <= grid.max_row and len(values) < MAX_TEXT_ROWS:
        _, _, bottom, _ = grid.bounds(row, col)
        cell = grid.cell_at(row, col)
        if cell is None:
            if values or skipped_empty or not allow_gap:
                break
            skipped_empty = True
        else:
            if is_stop_cell(cell, label, stop_labels) or _value_of_left_label(grid, cell, label, stop_labels):
                break
            values.append(cell)
            if not multi:
                break
        row = bottom + 1
    return values


def _value_of_left_label(grid: SheetGrid, cell: Cell, label: Cell, stop_labels: set[str]) -> bool:
    """ラベルの下で見つけたセルが、ラベルより左から結合された別のラベルの値の欄か。

    「クローズ判定」（空欄）の下に、左の「品証コメント」の値の欄（横長の結合セル）が来る帳票で、
    品証コメントをクローズ判定の値にしない。
    """
    if cell.col >= label.col:
        return False
    left = grid.cell_at(cell.row, cell.col - 1) if cell.col > 1 else None
    if left is None or left.max_col != cell.col - 1:
        return False
    # 見出し欄 = 項目のラベル、または塗りつぶしのある短い文字列（見本で使わなかった欄の見出しも含む）
    return is_stop_cell(left, None, stop_labels) or (
        left.filled and isinstance(left.value, str) and len(left.norm) <= MAX_LABEL_LENGTH)


def is_form_number_note(text) -> bool:
    """欄外の様式番号の注記（「様式MT-031 Rev.1」「製造部 設備保全課 様式MT-031 Rev.2」）か。項目の値にしない。"""
    s = str(text or "")
    return "\n" not in s and len(s.strip()) <= MAX_FORM_NUMBER_CHARS and bool(_FORM_NUMBER_RE.search(s))


def is_stop_cell(cell: Cell, label: Cell | None, stop_labels: set[str]) -> bool:
    """値の探索を打ち切るセルか。別のラベル、区切りの見出し（「▼ 回答欄」）、様式番号の注記、
    またはラベルと同じ塗りつぶし色の短い文字列（見出し欄）。

    押印欄（承認｜確認｜作成）や、見出しが横に並んで値がその下の行にある帳票では、右隣も見出しなので
    値として取らずに下を探す。長文の下にある色付きの見出し（「写真」など）で文章の取り込みも止める。
    """
    if cell.is_label(stop_labels) or section_heading(cell) or is_form_number_note(cell.text):
        return True
    return (label is not None and bool(label.fill) and cell.fill == label.fill
            and isinstance(cell.value, str) and len(cell.norm) <= MAX_LABEL_LENGTH)


def _convert(data_type: str, cells: list[Cell], inline_value: str | None, date1904: bool = False, unit: str = ""):
    if inline_value is not None:
        raw, texts = inline_value, [inline_value]
    else:
        raw, texts = cells[0].value, [c.text for c in cells]
    if data_type in ("string", "text") and len(texts) == 1:
        checked = pick_checked(texts[0])
        if checked is not None:
            return checked
    if data_type == "text":
        return "\n".join(texts), None
    if data_type == "date":
        return to_date(raw, texts[0], date1904)
    if data_type == "number":
        return to_number(raw, texts[0], unit)
    return texts[0], None


# ====================================================================================================
# 元 pattern/model.py
# 帳票の種類（旧: テンプレート／パターン）の定義。DBやFlaskに依存しない。
# ====================================================================================================

DATA_TYPES = {
    "string": "文字列",
    "text": "文章（複数行）",
    "date": "日付",
    "number": "数値",
    "table": "明細表（列見出しと行）",
}

DIRECTIONS = {
    "auto": "自動（右→下）",
    "right": "ラベルの右",
    "below": "ラベルの下",
    "same_cell": "同じセル（設備番号：EQ-001）",
}

# 画像の扱い。今は枚数だけ Markdown に書く（"vision" は未実装。DB に残っている古い値は読める）
IMAGE_PROCESSING = {
    "none": "読み取らない（枚数だけ Markdown に書きます）",
}

# Markdown に出すかどうか
RAG_OUTPUTS = {
    "show": "出す",
    "omit": "出さない",
}

# タイトル項目が未設定のときに使う「識別らしい項目」（この順で並べる）
DEFAULT_TITLE_KEYS = ("report_id", "equipment_id", "equipment_name", "occurred_date")

# Markdown の作り方の追加設定。入力を削る設定は置かない（人名の項目も必ず出す）。
DEFAULT_MD_OPTIONS: dict = {}


@dataclass
class SheetDef:
    sheet_name: str


@dataclass
class FieldDef:
    field_name: str
    display_name: str
    candidates: list[str] = field(default_factory=list)
    data_type: str = "string"
    direction: str = "auto"
    unit: str = ""
    rag_output: str = "show"
    # 明細表: 見本で見た列見出し。見出し（アンカー）の書き方が違う帳票でも、列見出しの並びが似た表を探すのに使う
    table_columns: list[str] = field(default_factory=list)
    # 探す区画（区切りの見出しの名前。excel.tables.section_key の形。例: 「回答」）。空ならシート全体。
    # 発行側と回答側で同じ意味の欄が並ぶ帳票で、どちら側の欄を読むかを決める。シートにその区画があり、
    # その中に探す見出しがあるときだけ区画の中で探す（区画の無い版・区画の中に見出しの無い版は全体で探す）
    section: str = ""
    # クリックで作った項目の控え: 見本でクリックしたシート名・見出しセル・値セルの番地。
    # 見出しで探して見つからなかったときだけ、この番地のセルを読む（見出しのない「値だけ」の項目もここで読む）
    sheet_name: str = ""
    label_cell: str = ""
    cell: str = ""
    # 見出しを手で直した項目。別の欄を足したときに、手で付けた名前へ勝手に戻さないための目印
    renamed: bool = False

    def search_labels(self) -> list[str]:
        labels = [c.strip() for c in self.candidates if c.strip()] or [self.display_name]
        return list(dict.fromkeys(labels))

    def label_norms(self) -> set[str]:
        return {normalize_label(label) for label in self.search_labels()} - {""}


@dataclass
class PatternDef:
    name: str
    version: str = "v1"
    description: str = ""
    image_processing: str = "none"
    status: str = "draft"
    sheets: list[SheetDef] = field(default_factory=list)
    fields: list[FieldDef] = field(default_factory=list)
    id: int | None = None
    title_fields: list[str] = field(default_factory=list)
    md_options: dict = field(default_factory=lambda: dict(DEFAULT_MD_OPTIONS))
    version_no: int = 1

    @property
    def label(self) -> str:
        return f"{self.name} {self.version}"

    def label_norms(self) -> set[str]:
        norms: set[str] = set()
        for fd in self.fields:
            norms |= fd.label_norms()
        return norms

    def output_label_norms(self) -> set[str]:
        """この帳票の見出し語（候補ラベル・表示名・明細表の列見出し）の正規化形。値がこれに当たる行は Markdown に出さない。"""
        norms = self.label_norms()
        for fd in self.fields:
            norms |= {normalize_label(x) for x in (fd.display_name, *(fd.table_columns or []))}
        return norms - {""}


# ====================================================================================================
# 元 pattern/dictionary.py
# 帳票でよく使われる項目名の辞書。
#
# サンプルExcelからテンプレートを作るときの候補提示と、
# 値の探索時に「隣のセルが別のラベルかどうか」の判定に使う。
# ====================================================================================================

@dataclass(frozen=True)
class StandardField:
    field_name: str
    display_name: str
    data_type: str
    synonyms: tuple[str, ...]


STANDARD_FIELDS = [
    StandardField("report_id", "報告番号", "string",
                  ("報告番号", "報告No", "報告書番号", "報告書No", "修理番号", "修理No", "管理番号", "管理No", "連絡票No",
                   "連絡No", "連絡番号", "帳票No", "Report No", "No")),
    StandardField("subject", "件名", "string", ("件名", "タイトル", "表題")),
    StandardField("equipment_id", "設備番号", "string",
                  ("設備番号", "装置番号", "設備No", "装置No", "設備コード", "装置コード", "設備ID", "機番", "号機",
                   "Equipment ID")),
    StandardField("equipment_name", "設備名", "string", ("設備名", "装置名", "設備名称", "装置名称", "機器名")),
    StandardField("process", "工程", "string", ("工程", "発生工程", "工程名")),
    StandardField("location", "発生場所", "string", ("発生場所", "設置場所", "場所", "ライン", "生産ライン", "設置ライン")),
    StandardField("department", "部署", "string", ("部署", "所属", "担当部署", "部署名", "起票部署", "発行部署", "発行元",
                   "発信部署")),
    StandardField("occurred_date", "発生日", "date", ("発生日", "発生日時", "故障発生日", "発生年月日", "故障発生日時", "不具合発生日時",
                   "異常発生日時")),
    StandardField("completed_date", "修理完了日", "date", ("修理完了日", "完了日", "復旧日", "修理日", "復旧日時", "完了日時", "復旧完了日時",
                   "生産復帰日時")),
    StandardField("reporter", "報告者", "string", ("報告者", "作成者", "記入者", "担当者", "作業者", "起票者", "発信者", "連絡者",
                   "発行者")),
    StandardField("approver", "承認者", "string", ("承認者", "承認", "確認者", "完了承認者")),
    StandardField("work_hours", "作業時間", "number", ("作業時間", "修理時間", "対応時間")),
    StandardField("downtime", "停止時間", "number", ("停止時間", "ダウンタイム", "設備停止時間", "設備停止")),
    StandardField("symptom", "故障内容", "text", ("故障内容", "症状", "不具合内容", "異常内容", "現象", "故障状況", "異常の内容",
                   "不具合現象", "発生現象", "問題の記述", "問題の概要")),
    StandardField("severity", "重要度", "string", ("重要度", "重大度", "影響度", "ランク")),
    StandardField("alarm", "アラーム", "string", ("アラーム", "アラーム内容", "エラー", "エラーコード", "エラー内容", "アラームNo",
                   "アラームコード", "ALM No", "発生アラーム", "発報アラーム")),
    StandardField("investigation", "原因調査", "text", ("原因調査", "調査内容", "調査結果", "確認・調査結果", "調査・検証結果")),
    StandardField("cause", "原因", "text", ("原因", "故障原因", "推定原因", "発生原因", "発生要因", "真因", "根本原因", "原因(確定)")),
    StandardField("repair", "修理内容", "text", ("修理内容", "処置内容", "処置", "対応内容", "作業内容", "対策", "応急処置", "暫定対策",
                   "復旧処置")),
    StandardField("result", "修理結果", "text", ("修理結果", "結果", "確認結果", "処置結果", "復旧確認", "効果判定", "効果の確認結果")),
    StandardField("prevention", "再発防止策", "text", ("再発防止策", "再発防止対策", "恒久対策", "再発防止", "歯止め")),
    StandardField("parts", "使用部品", "text", ("使用部品", "交換部品", "部品")),
    StandardField("remarks", "備考", "text", ("備考", "特記事項", "その他")),
]

LOOKUP: dict[str, StandardField] = {
    normalize_label(synonym): sf for sf in STANDARD_FIELDS for synonym in sf.synonyms
}
DICTIONARY_NORMS: set[str] = set(LOOKUP)
BY_FIELD_NAME: dict[str, StandardField] = {sf.field_name: sf for sf in STANDARD_FIELDS}

# 設備番号と設備名を1つのセルにまとめて書く見出し（「対象設備：CMP-108　STI-CMP 8号機」）。
# 値が「番号＋名前」に分けられるときだけ、設備番号（番号の部分）と設備名（名前の部分）のラベルとして使う
COMBINED_EQUIPMENT_LABELS = ("対象設備", "使用設備", "対象装置", "使用装置", "設備", "装置")
COMBINED_EQUIPMENT_NORMS: set[str] = {normalize_label(label) for label in COMBINED_EQUIPMENT_LABELS}
# 分けた値のどちらを使うか（0: 番号, 1: 名前）
COMBINED_EQUIPMENT_PARTS = {"equipment_id": 0, "equipment_name": 1}

# 人名が入る項目（Markdown で既定では出さない）
PERSON_FIELD_NAMES = {"reporter", "approver", "worker", "person", "inspector", "creator", "checker", "author"}
_PERSON_LABEL_SUFFIXES = ("者", "担当", "氏名", "名前", "検印", "承認", "確認印")
# 押印欄の見出し（「確認」「作成」）。前方の語を含む「効果確認」「作成日」まで人名にしないよう、完全一致だけで見る
_PERSON_LABEL_EXACT = {"確認", "作成", "立会", "立ち会い", "審査", "署名", "サイン", "印"}


def is_person_field(field_name: str, display_name: str = "") -> bool:
    """報告者・承認者・担当者など人名の項目か。「担当部署」「設備名」「効果確認」は含めない。"""
    if field_name in PERSON_FIELD_NAMES:
        return True
    label = normalize_label(display_name)
    if not label:
        return False
    return label in _PERSON_LABEL_EXACT or is_person_label(display_name)


# 単位が書かれていないと意味が変わる数値項目（見出しの語 → 画面に出す説明）。
# 「件数」「回数」「人数」「枚数」のように数えるだけの項目は、単位が無いのが普通なので入れない
# （ここに無い項目は単位が空でも要確認にしない＝確認画面が警告だらけにならないようにする）。
AMBIGUOUS_UNIT_HINTS: tuple[tuple[tuple[str, ...], str], ...] = (
    (("時間", "工数", "所要", "ダウンタイム", "期間", "リードタイム", "タクト"), "分か時間かで意味が変わります"),
    (("金額", "費用", "コスト", "価格", "単価", "原価", "経費", "予算", "損失額", "修理費", "部品費"),
     "円か千円かで意味が変わります"),
    (("長さ", "寸法", "距離", "厚み", "厚さ", "板厚", "幅", "直径", "外径", "内径", "隙間", "クリアランス", "変位", "摩耗量"),
     "mmかmかで意味が変わります"),
    (("重量", "質量", "重さ"), "gかkgかで意味が変わります"),
    (("温度", "温度差"), "℃かKかで意味が変わります"),
    (("圧力", "真空度"), "MPaかkPaかで意味が変わります"),
    (("流量", "風量"), "L/minかm3/hかで意味が変わります"),
    (("電流", "電圧", "電力", "消費電力"), "AかmA（VかmV）かで意味が変わります"),
)
# 単位の付かない数え方・割合の項目（「不良件数」のように上の語を含んでいても、こちらが優先）
_COUNT_SUFFIXES = ("件数", "回数", "人数", "枚数", "個数", "台数", "本数", "点数", "数量", "員数", "率", "割合", "%")


def ambiguous_unit_hint(field_name: str, display_name: str = "") -> str:
    """単位が書かれていないと意味が変わる数値項目か。当てはまれば画面に出す説明、当てはまらなければ ""。

    単位の無い数値をすべて要確認にすると「件数」「回数」「人数」で確認画面が埋まるので、
    辞書が「単位で意味が変わる」と知っている項目だけを要確認にする。
    """
    standard = BY_FIELD_NAME.get(field_name)
    labels = [normalize_label(x) for x in (display_name, standard.display_name if standard else "")]
    labels = [x for x in labels if x]
    if not labels or any(x.endswith(_COUNT_SUFFIXES) for x in labels):
        return ""
    for words, hint in AMBIGUOUS_UNIT_HINTS:
        if any(normalize_label(w) in label for label in labels for w in words):
            return hint
    return ""


def is_person_label(display_name: str) -> bool:
    """見出しの語だけで人名の欄とわかるか（明細表の列見出し「担当」「氏名」用）。

    押印欄の「確認」「作成」は、明細表では判定記号の列でもありうるので含めない。
    """
    label = normalize_label(display_name)
    return bool(label) and label.endswith(_PERSON_LABEL_SUFFIXES)


# ====================================================================================================
# 元 pattern/builder.py
# 見本ファイル（1〜数ファイル）から帳票の種類の候補を作る。
#
# ここで作るのはあくまで「候補」で、人が画面で確認・修正してから使用開始する。
# 複数サンプルを渡すと、セル位置が違っても同じラベルが全サンプルにあるかで確度を上げる。
# ====================================================================================================

_VALUE_LIKE = re.compile(r"^[\x20-\x7e]+$")  # 英数字記号のみ（EQ-001, CMP, 2026/09/14 など）
# 注記・凡例の書き出し（「※判定　○」「※故障発生日時…は事後保全時に記入」）
_NOTE_START = re.compile(r"^\s*[※＊*]")
# 「－」だけの値（記入なしの印）
_DASH_ONLY = {"-", "－", "ー", "―", "‐", "/", "／"}
# 「8D No.」「QA番号」のような短い接頭辞付きの識別番号の見出し（辞書の「報告No」と同じ扱い）
_ID_LABEL_RE = re.compile(r"^[a-z0-9]{1,3}(?:no|番号)$")
# 日付らしい見出しの末尾（見本の値が日付として読めるときだけ日付型にする）
_DATE_LABEL_SUFFIXES = ("日", "日付", "日時", "期限", "予定日", "年月日")
# 値そのものが日付だけで書かれているか（見出しからは日付と分からない「発生」欄などのため）。
# 日付を含むだけの文（「2026-09-14 に発生」「ロットNo. 2026-09-14-3」）を日付にしないよう、値の全体が
# 日付（＋時刻・曜日）の表記のときだけとする。年の無い「2/12」は to_date が警告を返すので日付にしない。
_DATE_VALUE_RE = re.compile(r"^(?:令和|平成|昭和|[RHS])?\s*(?:\d{1,4}|元)\s*[年/\-.]\s*\d{1,2}\s*[月/\-.]\s*\d{1,2}\s*日?"
                            r"(?:\s*\([日月火水木金土]\))?"
                            r"(?:\s*T?\s*\d{1,2}\s*(?::\d{2}(?::\d{2})?|時\s*\d{1,2}\s*分?)\s*)?$")
# 見出しの先頭の項番（表示名から除く）: 「1. 時系列」「D1 チーム編成」「A. 機構部」
_SECTION_PREFIX = re.compile(r"^(?:\d{1,2}\s*[.)、．]|[A-Za-zＡ-Ｚ]\s*[.．)、]|D[1-8](?![0-9])\s*[:：]?)\s*")


@dataclass
class _Suggestion:
    field_name: str | None
    display_name: str
    data_type: str
    labels: list[str] = field(default_factory=list)
    synonyms: tuple[str, ...] = ()
    samples: set[int] = field(default_factory=set)
    sheets: set[str] = field(default_factory=set)
    examples: list[str] = field(default_factory=list)
    score: int = 0  # 明細表では、見本での最大の行数
    unit: str = ""
    table_labels: tuple[str, ...] = ()  # 明細表の見出しの別名（辞書の同義語。use の判定には使わない）
    columns: frozenset[str] = frozenset()  # 明細表の列見出し（正規化済み）
    table_room: bool = False  # 行を足せる明細表の形か（excel.tables.Table.has_room）
    column_labels: list[str] = field(default_factory=list)  # 明細表の列見出し（表記のまま）


def suggest_rows(infos: list[WorkbookInfo]) -> tuple[list[dict], list[dict]]:
    """画面表示用の (シート行, 項目行) を返す。"""
    sheet_rows = _suggest_sheets(infos)
    selected = {normalize_sheet_name(r["sheet_name"]) for r in sheet_rows if r["use"]}
    n = len(infos)

    suggestions: dict[str, _Suggestion] = {}
    for index, info in enumerate(infos):
        for grid in info.grids.values():
            for key, sug in _label_candidates(grid).items():
                if sug.data_type == "table" and key not in suggestions:
                    key = _same_table_key(suggestions, sug, index) or key
                merged = suggestions.setdefault(key, sug)
                if merged is not sug:
                    merged.labels.extend(l for l in sug.labels if l not in merged.labels)
                    merged.examples.extend(sug.examples)
                    merged.score = max(merged.score, sug.score)
                    merged.unit = merged.unit or sug.unit
                    merged.columns = merged.columns | sug.columns
                    merged.table_room = merged.table_room or sug.table_room
                    known = {normalize_label(c) for c in merged.column_labels}
                    merged.column_labels.extend(c for c in sug.column_labels if normalize_label(c) not in known)
                merged.samples.add(index)
                merged.sheets.add(normalize_sheet_name(grid.name))

    rows, used_names = [], set()
    for sug in suggestions.values():
        # 「修理報告書」を選んでいれば「修理報告書(2)」「設備修理報告書」も同じシートとみなす（matcher と同じ）
        in_selected_sheet = any(sel in sheet for sheet in sug.sheets for sel in selected)
        in_dictionary = bool(sug.synonyms)
        if not in_selected_sheet and not in_dictionary:
            continue
        if sug.data_type == "table":
            # 明細表は、行を足せる形のときだけ使う（1行だけの「見出しの行＋値の行」は項目として読む）
            use = in_selected_sheet and sug.table_room
        else:
            use = in_selected_sheet and (in_dictionary or (len(sug.samples) == n and sug.score >= 1)) \
                and not _note_or_legend(sug)
        field_name = sug.field_name
        if not field_name or field_name in used_names:
            field_name = _next_field_name(used_names)
        used_names.add(field_name)
        rows.append({
            "use": use,
            "field_name": field_name,
            "display_name": sug.display_name,
            "candidates": "\n".join(dict.fromkeys([*sug.labels, *sug.synonyms, *sug.table_labels])),
            "data_type": sug.data_type,
            "direction": "auto",
            "unit": sug.unit,
            "rag_output": "show",
            "table_columns": "\n".join(sug.column_labels),
            "examples": list(dict.fromkeys(sug.examples))[:3],
            "seen": f"{len(sug.samples)}/{n}",
            "section": "",
        })
    _learn_sections(infos, selected, rows)
    rows.sort(key=lambda r: not r["use"])
    return sheet_rows, rows


def _learn_sections(infos: list[WorkbookInfo], selected: set[str], rows: list[dict]) -> None:
    """発行側と回答側で同じ意味の欄が並ぶ帳票で、項目を読む区画（「▼ 回答欄」など）を見本から決める。

    項目の見出しに当たる欄（値のあるもの）が1つの見本の中で2つ以上の区画にあり、どの見本にも共通してある区画が
    名前のある区画1つだけのとき、その区画を項目の区画にする。例: 「処置内容（発行側）」と「暫定対策（処置）（回答欄）」が
    両方ある版と、「応急処置（回答欄）」だけの版を見本にすると、処置の項目は「回答欄」の中で探す。
    どの見本にも共通する区画が2つ以上ある（区画の外にも共通してある）ときは決めない（今までどおり上にある欄を読む）。
    """
    used = [r for r in rows if r["use"] and r["data_type"] != "table"]
    if not used:
        return
    stop_labels = DICTIONARY_NORMS.union(*(_row_field(r).label_norms() for r in used))
    for row in used:
        fd = _row_field(row)
        per_sample: list[set[str]] = []
        ambiguous = False
        for info in infos:
            keys: set[str] = set()
            for name, grid in info.grids.items():
                if not any(sel in normalize_sheet_name(name) for sel in selected):
                    continue
                keys |= {section_of(grid, cell) for cell in label_hits(grid, fd, stop_labels)}
            if keys:
                per_sample.append(keys)
                ambiguous = ambiguous or len(keys) > 1
        if not ambiguous:
            continue
        common = set.intersection(*per_sample)
        if len(common) == 1 and "" not in common:
            row["section"] = common.pop()


def _row_field(row: dict) -> FieldDef:
    return FieldDef(row["field_name"], row["display_name"], row["candidates"].splitlines(),
                    data_type=row["data_type"])


def suggest_title_fields(field_rows: list[dict]) -> list[str]:
    """タイトル項目の候補。使う項目のうち 報告番号→設備番号→設備名→発生日 の順で並べる。"""
    used = {r["field_name"] for r in field_rows if r.get("use")}
    return [key for key in DEFAULT_TITLE_KEYS if key in used]


def _suggest_sheets(infos: list[WorkbookInfo]) -> list[dict]:
    rows: dict[str, dict] = {}
    for info in infos:
        # 辞書ラベルの数が多いシートを主シートとする（同数なら値のあるセルが多い方）
        score = {name: (sum(1 for c in grid.text_cells() if c.is_label(DICTIONARY_NORMS)), len(grid.cells))
                 for name, grid in info.grids.items()}
        main = max(score, key=score.get, default=None)
        best_hits = score[main][0] if main else 0
        for name in info.grids:
            key = normalize_sheet_name(name)
            row = rows.setdefault(key, {"sheet_name": name, "use": False, "required": True, "seen": 0})
            row["seen"] += 1
            if name == main and score[name][1] or score[name][0] >= max(2, best_hits * 0.5):
                row["use"] = True

    # 「修理報告書」と「修理報告書(2)」のような重複はサンプルごとの揺れなので短い方だけ選ぶ
    selected: list[str] = []
    for key in sorted(rows, key=len):
        if rows[key]["use"]:
            if any(s in key for s in selected):
                rows[key]["use"] = False
            else:
                selected.append(key)
    n = len(infos)
    return [{**r, "seen": f"{r['seen']}/{n}"} for r in rows.values()]


def _label_candidates(grid: SheetGrid) -> dict[str, _Suggestion]:
    # 背景色付きの短いセルは帳票上のラベル欄とみなし、値探索の打ち切り対象にする
    stop_labels = DICTIONARY_NORMS | {c.norm for c in grid.text_cells() if c.filled and len(c.norm) <= MAX_LABEL_LENGTH}
    tables = _table_candidates(grid)
    # 明細表の見出し・列見出しは、1つの値の項目の候補にしない（列見出しの項目は1行目の値しか読めない）
    table_cells = {(c.row, c.col) for t, _ in tables if t.has_room for c in (t.anchor, *t.header)}
    # 表の連番の列見出し（No）は、報告書の「No.」欄の候補にしない
    table_cells |= {key for key in table_header_keys(grid) if seq_header(grid.cells[key])}
    found: list[tuple[Cell, _Suggestion, list[Cell]]] = []
    for cell in grid.text_cells():
        if (cell.row, cell.col) in table_cells:
            continue
        if cell.inline:
            label_text = re.split(r"[:：]", cell.text, maxsplit=1)[0].strip()
            sug = _make_suggestion(cell.inline[0], label_text, cell.inline[1], cell)
            found.append((cell, sug, []))
            found.extend((cell, part, []) for part in _equipment_parts(cell.inline[0], label_text, cell.inline[1], cell))
            continue
        if "\n" in cell.text or len(cell.norm) > MAX_LABEL_LENGTH or not isinstance(cell.value, str):
            continue
        if not any(ch.isalpha() for ch in cell.norm):  # 「×」「≦80%」「28℃」などの記号・数値はラベルにしない
            continue
        entry = _lookup(cell)
        if entry is None and _VALUE_LIKE.match(cell.norm):
            continue
        values = scan_right(grid, cell, stop_labels) or scan_below(grid, cell, stop_labels, multi=False)
        if not values:
            continue
        label_text = cell.text.rstrip(":：").strip()
        found.append((cell, _make_suggestion(cell.norm, label_text, values[0].value, cell), values))
        found.extend((cell, part, values) for part in _equipment_parts(cell.norm, label_text, values[0].text, cell))

    # 辞書ラベルの「値」として使われたセルは、ラベル候補から外す。人名の欄（作成・確認・承認）の値も同じ:
    # 押印欄の「斎藤」を見出しにすると、人名の項目を省いても見出しと値（隣の人名）が Markdown に出てしまう
    value_cells = {(v.row, v.col) for _, sug, values in found
                   if sug.synonyms or is_person_field(sug.field_name or "", sug.display_name) for v in values}
    result: dict[str, _Suggestion] = {}
    for cell, sug, _ in found:
        if not sug.synonyms and (cell.row, cell.col) in value_cells:
            continue
        key = sug.field_name or f"label:{cell.inline[0] if cell.inline else cell.norm}"
        result.setdefault(key, sug)
    for table, sug in tables:
        result.setdefault(f"table:{sug.field_name or table.anchor.alt_norm or table.anchor.norm}", sug)
    return result


def _same_table_key(suggestions: dict[str, _Suggestion], sug: _Suggestion, sample_index: int) -> str | None:
    """見本ごとに見出しの書き方が違う同じ明細表（「D1 チーム編成」「D1：チームの結成」）を1つの候補にまとめる。

    列見出しの半分以上が同じで、まだその見本から入っていない明細表の候補があれば、そのキーを返す。
    """
    best, best_score = None, 0.0
    for key, other in suggestions.items():
        if other.data_type != "table" or sample_index in other.samples or not other.columns:
            continue
        score = len(sug.columns & other.columns) / len(sug.columns | other.columns)
        if score >= SAME_COLUMNS_RATIO and score > best_score:
            best, best_score = key, score
    return best


def _table_candidates(grid: SheetGrid) -> list[tuple[Table, _Suggestion]]:
    """見出し（アンカー）のある明細表を、明細表の項目の候補にする。"""
    out = []
    for table in detect_tables(grid):
        if table.anchor is None:
            continue
        value = table.to_value()
        anchor = table.anchor
        label = " ".join(anchor.text.split())
        entry = _lookup(anchor)
        display = _SECTION_PREFIX.sub("", table_title(anchor)).strip() or table_title(anchor)
        lines = table_text_lines(value)
        sug = _Suggestion(entry.field_name if entry else None, display, "table", [label],
                          examples=[lines[0]] if lines else [], score=len(value["rows"]) if value else 0,
                          table_room=table.has_room,
                          table_labels=entry.synonyms if entry else (),
                          columns=frozenset(h.norm for h in table.header),
                          column_labels=[" ".join(h.text.split()) for h in table.header])
        out.append((table, sug))
    return out


def _note_or_legend(sug: _Suggestion) -> bool:
    """既定でチェックを外す候補か（注記・判定の凡例・「－」だけの値）。候補としては残すので、人が画面で選べる。

    design.md 6.1「値にしないもの」: 凡例（「○:良好 △:要観察」）、欄外の注記（「※…は事後保全時に記入」）。
    """
    if _NOTE_START.match(sug.display_name):
        return True
    examples = [e.strip() for e in sug.examples if e and e.strip()]
    if not examples:
        return False
    # 凡例は、どれか1つの見本で見つかればその欄は値を書く欄ではない（別の見本ではチェック印だけが入る）
    if any(legend_text(e) for e in examples):
        return True
    return all(_NOTE_START.match(e) or e in _DASH_ONLY for e in examples)


def _lookup(cell: Cell):
    """辞書の項目。「３．暫定対策」のような項番付きの見出しは項番を除いても引く。"""
    return _dict_entry(cell.norm) or (_dict_entry(cell.alt_norm) if cell.alt_norm else None)


def _dict_entry(norm: str):
    """正規化ラベルから辞書の項目を引く。

    「8D No.」「QA番号」のように短い英数字の接頭辞が付いた見出しは、辞書の「報告No」と同じ識別番号として扱う
    （帳票ごとに接頭辞が違うので、辞書に並べきれない）。
    """
    if not norm:
        return None
    entry = LOOKUP.get(norm)
    if entry is None and _ID_LABEL_RE.match(norm):
        entry = BY_FIELD_NAME["report_id"]
    return entry


def _equipment_parts(norm: str, label_text: str, value_text: str, cell: Cell) -> list[_Suggestion]:
    """「対象設備：CMP-108　STI-CMP 8号機」のような番号と名前をまとめた欄から、設備番号・設備名の候補を作る。"""
    if norm not in COMBINED_EQUIPMENT_NORMS:
        return []
    pieces = split_code_name(value_text)
    if pieces is None:
        return []
    out = []
    for field_name, index in COMBINED_EQUIPMENT_PARTS.items():
        entry = next(sf for sf in LOOKUP.values() if sf.field_name == field_name)
        out.append(_Suggestion(entry.field_name, entry.display_name, entry.data_type, [label_text], entry.synonyms,
                               examples=[pieces[index]], score=3))
    return out


def _make_suggestion(norm: str, label_text: str, value, cell: Cell) -> _Suggestion:
    # 「作業時間(h)」のような見出しは単位を分けて辞書を引く（候補ラベルは元の表記のまま残す）
    base, unit = split_label_unit(label_text)
    entry = _dict_entry(norm) or (_dict_entry(cell.alt_norm) if cell.alt_norm and not cell.inline else None)
    entry = entry or (_dict_entry(normalize_label(base)) if unit else None)
    display = base if unit else label_text
    example = cell_text(value)
    if entry:
        sug = _Suggestion(entry.field_name, display, entry.data_type, [label_text], entry.synonyms,
                          examples=[example], score=3, unit=unit)
    else:
        score = int(cell.bold) + int(cell.filled) + int(cell.text.rstrip().endswith((":", "：")))
        sug = _Suggestion(None, display, _guess_type(value, display), [label_text], examples=[example], score=score,
                          unit=unit)
    if not sug.unit and sug.data_type == "number":
        # 値が「1.5時間」のように単位付きで書かれていれば、その単位を候補にする
        sug.unit = value_unit(example)
    return sug


def _guess_type(value, label: str = "") -> str:
    if isinstance(value, date):
        return "date"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return "number"
    text = str(value)
    # 見本が文字列で日付を持つ帳票（「令和5年12月5日」「2025/12/19(金)」）も日付にして、本文を ISO にそろえる
    if _date_label(label) and to_date(value, cell_text(value))[1] is None:
        return "date"
    # 「発生」のように見出しからは日付と分からない欄も、値が日付だけで書かれていれば日付にする
    # （文字列のままだと ISO＋「（2024年7月）」の注記が付かず、同じ md の中で日付の書き方が混ざる）
    if _date_value(value):
        return "date"
    return "text" if "\n" in text or len(text) > 40 else "string"


def _date_value(value) -> bool:
    """値そのものが日付だけで書かれているか（年月日がそろい、警告なしに ISO にできる）。

    見本ごとの多数決はしない（この候補を作ったセルの値だけで見る。別シートの同名ラベルの値が混ざるため）。
    """
    text = cell_text(value)
    return bool(_DATE_VALUE_RE.match(nfkc_value(text))) and to_date(value, text)[1] is None


def _date_label(label: str) -> bool:
    """日付らしい見出しか（「作業日」「回答日」「次回点検予定日」「発生日時」）。「曜日」は値が日付にならないので残らない。"""
    return normalize_label(label).endswith(_DATE_LABEL_SUFFIXES)


def _next_field_name(used: set[str]) -> str:
    i = 1
    while f"field_{i}" in used:
        i += 1
    return f"field_{i}"


# ====================================================================================================
# 元 pattern/clicks.py
# クリックしたセルから読み取る項目を作る。
#
# 画面では「見出しのセル → 値のセル」の順にクリックするだけ。項目のキー名・型・単位・探す見出しは、
# クリックした2つのセルとその値からここで決める（画面には出さない）。
# ====================================================================================================

BASE_ROW = {
    "use": True, "rag_output": "show", "table_columns": "", "section": "",
    "unit": "", "direction": "auto", "sheet_name": "", "label_cell": "", "cell": "", "examples": [],
    # 辞書で分かった項目名（「設備番号」など）。別の見本で書き方の違う同じ欄をクリックしたときに、
    # 新しい項目にせず、その項目の探す見出しに足すために使う（保存はしない）
    "base_name": "",
}


def cell_key(grid: SheetGrid, coord) -> tuple[int, int] | None:
    """セル番地（"C5" / "C5:D6"）を、結合を考えた左上セルの (行, 列) にする。読めなければ None。"""
    text = str(coord or "").split(":")[0].strip()
    if not text:
        return None
    try:
        letters, row = coordinate_from_string(text.upper())
        col = column_index_from_string(letters)
    except Exception:
        return None
    if row < 1 or col < 1:
        return None
    top, left, _, _ = grid.bounds(row, col)
    return (top, left)


def coord_of(key: tuple[int, int]) -> str:
    return f"{get_column_letter(key[1])}{key[0]}"


def click_field(grid: SheetGrid, label_coord, value_coord="", used_names=()) -> tuple[dict | None, str]:
    """クリックした2つのセルから項目行を作る。戻り値: (項目行, エラーメッセージ)。

    - 見出しのセルだけ: 明細表の見出し（列見出しの1つ目）なら明細表の項目、そうでなければ値が空のままの項目
    - 同じセルを2回: 「設備番号：EQ-001」ならセル内の見出しと値、そうでなければ見出しのない「値だけ」の項目
    - 別のセル: 見出しと値（右・下・それ以外）
    """
    used = set(used_names)
    label_key = cell_key(grid, label_coord)
    if label_key is None:
        return None, "セルを選び直してください"
    label = grid.cells.get(label_key)
    value_key = cell_key(grid, value_coord) if value_coord else None

    if value_key is None:
        table = _table_row(grid, label_key, used)
        if table is not None:
            return table, ""
        if label is None or not label.text.strip():
            return None, "文字のないセルは見出しに選べません"
        return _label_row(grid, label, None, "auto", used), ""

    if value_key == label_key:
        if label is not None and label.inline:
            # 「設備番号：EQ-001」のように1つのセルに見出しと値が入っている
            return _label_row(grid, label, label_key, "same_cell", used), ""
        if label is None or not label.text.strip():
            return None, "空のセルは項目にできません"
        return _value_only_row(grid, label_key, used), ""

    if label is None or not label.text.strip():
        return None, "文字のないセルは見出しに選べません"
    return _label_row(grid, label, value_key, _direction(grid, label_key, value_key), used), ""


def _direction(grid: SheetGrid, label_key: tuple[int, int], value_key: tuple[int, int]) -> str:
    top, left, bottom, right = grid.bounds(*label_key)
    vtop, vleft, vbottom, vright = grid.bounds(*value_key)
    if vtop <= bottom and vbottom >= top and vleft > right:
        return "right"
    if vleft <= right and vright >= left and vtop > bottom:
        return "below"
    return "auto"


def _label_row(grid: SheetGrid, label, value_key, direction: str, used: set[str]) -> dict:
    value_cell = grid.cells.get(value_key) if value_key else None
    if direction == "same_cell":
        label_text = re.split(r"[:：]", label.text, maxsplit=1)[0].strip()
        sug = _make_suggestion(label.inline[0], label_text, label.inline[1], label)
        value_text = label.inline[1]
    else:
        label_text = label.text.rstrip(":：").strip()
        value = value_cell.value if value_cell is not None else ""
        sug = _make_suggestion(label.norm, label_text, value, label)
        value_text = value_cell.text if value_cell is not None else ""
    field_name = _field_name(sug.field_name, used)
    # 同じ名前の項目がすでにあるとき（「担当者」の次に「報告者」を押したとき）は、辞書の名前ではなく
    # クリックした見出しをそのまま項目の名前にする（どちらの欄か見分けられるように）
    taken = bool(sug.field_name) and field_name != sug.field_name
    return {
        **BASE_ROW,
        "field_name": field_name,
        "base_name": sug.field_name or "",
        "display_name": label_text if taken else (sug.display_name or label_text),
        # クリックした見出しと、辞書で分かる同じ意味の見出し（「設備番号」に対する「設備No」など）。
        # 書き方の違う帳票でも同じ欄を読めるようにする（値は書き替えない）
        "candidates": "\n".join(dict.fromkeys([label_text, *sug.synonyms])),
        "data_type": sug.data_type,
        "unit": sug.unit,
        "direction": direction,
        # クリックした見出しが入っている区画（「▼ 回答欄」）。発行側と回答側に同じ欄がある帳票で、
        # 人が指したほうの欄を読むための目印（その区画のある帳票でだけ効く）
        "section": section_of(grid, label),
        "sheet_name": grid.name,
        "label_cell": coord_of((label.row, label.col)),
        "cell": coord_of(value_key) if value_key else "",
        "examples": [value_text] if value_text else [],
    }


def split_rows(row: dict, used_names=()) -> list[dict]:
    """クリックで作った項目行を、必要なら複数の項目行にする。

    「使用設備：ROB-821（ウェーハソーター 1号機）」のように番号と名前を1つのセルにまとめた欄は、
    設備番号・設備名の2項目にする。文字は1つも消さず、同じセルの読む場所を分けるだけ
    （帳票ごとに「対象設備」「使用設備」「設備」と書き方が違うので、探す見出しは全部入れる）。
    """
    used = set(used_names)
    if row["data_type"] == "table":
        return [row]
    labels = [l for l in row["candidates"].splitlines() if l.strip()]
    if not any(normalize_label(l) in COMBINED_EQUIPMENT_NORMS for l in labels):
        return [row]
    pieces = split_code_name((row.get("examples") or [""])[0])
    if pieces is None:
        return [row]
    out = []
    for field_name, index in COMBINED_EQUIPMENT_PARTS.items():
        entry = BY_FIELD_NAME[field_name]
        out.append({**row,
                    "field_name": _field_name(field_name, used),
                    "base_name": field_name,
                    "display_name": entry.display_name,
                    "candidates": "\n".join(dict.fromkeys([*labels, *COMBINED_EQUIPMENT_LABELS, *entry.synonyms])),
                    "data_type": entry.data_type,
                    "unit": "",
                    "examples": [pieces[index]]})
    return out


def _value_only_row(grid: SheetGrid, key: tuple[int, int], used: set[str]) -> dict:
    """見出しのない「値だけ」の項目（クリックしたセルの番地で読む）。"""
    cell = grid.cells[key]
    coord = coord_of(key)
    return {
        **BASE_ROW,
        "field_name": _field_name(None, used),
        "display_name": f"値（{coord}）",
        "candidates": "",
        "data_type": _guess_type(cell.value),
        "sheet_name": grid.name,
        "cell": coord,
        "examples": [cell.text],
    }


def _table_row(grid: SheetGrid, key: tuple[int, int], used: set[str]) -> dict | None:
    """クリックしたセルが明細表の見出し（または列見出しの1つ目）なら、その表を1つの項目にする。"""
    table = _table_at(grid, key)
    if table is None:
        return None
    anchor = table.anchor
    head = anchor if anchor is not None else table.header[0]
    label_text = " ".join(head.text.split())
    entry = _lookup(head)
    display = (_SECTION_PREFIX.sub("", table_title(head)).strip() or table_title(head)) if anchor else label_text
    value = table.to_value()
    lines = table_text_lines(value) if value else []
    return {
        **BASE_ROW,
        "field_name": _field_name(entry.field_name if entry else None, used),
        "base_name": entry.field_name if entry else "",
        "display_name": display,
        "candidates": "\n".join(dict.fromkeys([label_text, *(entry.synonyms if entry else ())])),
        "data_type": "table",
        "section": section_of(grid, head),
        "table_columns": "\n".join(" ".join(h.text.split()) for h in table.header),
        "sheet_name": grid.name,
        "label_cell": coord_of((head.row, head.col)),
        "examples": lines[:1],
    }


def _table_at(grid: SheetGrid, key: tuple[int, int]):
    """クリックしたセルを見出し（アンカー）か列見出しに持つ明細表。

    見本での記入が1行だけの表（「使用部品」に部品1つ）も、列見出しのすぐ上にある見出しを指したときは
    明細表にする（人が「この表」と指しているので、見本の行数では決めない）。列見出しを指したときは
    今までどおり、行が並ぶ形に見える表だけ（押印欄の「承認｜確認｜作成」を表にしないため）。
    """
    for table in detect_tables(grid):
        if _table_heading(table) == key:
            return table
        if table.has_room and any((h.row, h.col) == key for h in table.header):
            return table
    return None


def _table_heading(table) -> tuple[int, int] | None:
    """列見出しのすぐ上にある、その表の見出しセル。離れていれば None（帳票のタイトルを表の見出しにしない）。"""
    anchor = table.anchor
    if anchor is None or not table.header or anchor.max_row + 1 != min(h.row for h in table.header):
        return None
    return (anchor.row, anchor.col)


def _field_name(name: str | None, used: set[str]) -> str:
    """まだ使っていない項目の名前。辞書の名前がふさがっていれば「reporter_2」、名前が無ければ「field_1」。"""
    if not name:
        name = _next_field_name(used)
    elif name in used:
        i = 2
        while f"{name}_{i}" in used:
            i += 1
        name = f"{name}_{i}"
    used.add(name)
    return name


def table_cells(grid: SheetGrid) -> set[str]:
    """1回のクリックで明細表の項目になるセル（見出しと列見出し）の番地。`_table_at` と同じ判断。"""
    out: set[str] = set()
    for table in detect_tables(grid):
        heading = _table_heading(table)
        if heading is not None:
            out.add(coord_of(heading))
        if table.has_room:
            out.update(coord_of((h.row, h.col)) for h in table.header)
    return out


def merge_target(field_rows: list[dict], row: dict, grid: SheetGrid | None = None) -> dict | None:
    """クリックで作った項目行が、登録済みのどの項目と同じ欄か（別の見本で書き方が違うだけの欄）。

    辞書で同じ項目と分かるもの（「設備番号」と「設備No」）と、列見出しがほとんど同じ明細表。
    見つかれば、新しい項目にせずその項目の探す見出しに足す。
    ただし、その項目の見出しが同じシートの別のセルにあるとき（`same_sheet_field`）は、
    同じ見本の中の別の欄を押したということなので足さない（別の項目にする）。
    """
    if same_sheet_field(field_rows, row, grid) is not None:
        return None
    return next(_same_field_rows(field_rows, row), None)


def same_sheet_field(field_rows: list[dict], row: dict, grid: SheetGrid | None = None) -> dict | None:
    """いま押したシートに、自分の見出しのセルを別に持っている「同じ欄らしい」登録済みの項目。

    「担当者」を登録した見本で「報告者」を押したときのように、同じシートの別のセルを指している。
    辞書では同じ名前になるが人が読みたい欄は2つなので、見出しを足さずに別の項目にする。
    別の見本で書き方の違う同じ欄（「設備No」と「設備番号」）を押したときは、その項目の見出しのセルが
    このシートに無い（または押したセルそのもの）ので None になり、今までどおり見出しを足す。
    """
    if grid is None:
        return None
    for other in _same_field_rows(field_rows, row):
        if _label_cell_here(grid, other, row):
            return other
    return None


def _same_field_rows(field_rows: list[dict], row: dict):
    """同じ欄かもしれない登録済みの項目（辞書で同じ名前の項目／列見出しがほとんど同じ明細表）。"""
    base = row.get("base_name") or ""
    if base:
        same = next((r for r in field_rows
                     if r["field_name"] == base and r["data_type"] == row["data_type"]), None)
        if same is not None:
            yield same
    if row["data_type"] != "table":
        return
    columns = _column_norms(row)
    if not columns:
        return
    for other in field_rows:
        if other["data_type"] != "table":
            continue
        have = _column_norms(other)
        if have and len(columns & have) / len(columns | have) >= SAME_COLUMNS_RATIO:
            yield other


def _label_cell_here(grid: SheetGrid, other: dict, row: dict) -> bool:
    """登録済みの項目の見出しが、いま押したシートの別のセルにそのまま書かれているか。"""
    if (other.get("sheet_name") or "") != grid.name:
        return False
    key = cell_key(grid, other.get("label_cell") or "")
    if key is None or coord_of(key) == (row.get("label_cell") or ""):
        return False
    cell = grid.cells.get(key)
    if cell is None or not cell.text.strip():
        return False
    norms = {normalize_label(l) for l in (other.get("candidates") or "").splitlines() if l.strip()} - {""}
    return bool(norms) and (cell.inline[0] if cell.inline else cell.norm) in norms


def separate_names(twin: dict, row: dict) -> None:
    """同じシートの別の欄として登録するとき、2つの項目を見分けられる名前にする。

    辞書の名前が同じだけの別の欄（「担当者」と「報告者」）なので、どちらもクリックした見出しを名前にする。
    見出しまで同じ（区画違いの同じ名前の欄）ときは、新しいほうにセル番地を添える。
    ただし手で直した見出しはそのまま残し、新しいほうの名前を変えて見分けられるようにする。
    """
    if row["display_name"] != twin["display_name"]:
        return
    twin_label, row_label = _first_label(twin), _first_label(row)
    if twin.get("renamed"):
        if row_label and row_label != twin["display_name"]:
            row["display_name"] = row_label
        elif row.get("label_cell"):
            row["display_name"] = f"{row['display_name']}（{row['label_cell']}）"
        return
    if twin_label and row_label and twin_label != row_label:
        twin["display_name"], row["display_name"] = twin_label, row_label
    elif row.get("label_cell"):
        row["display_name"] = f"{row['display_name']}（{row['label_cell']}）"


def _first_label(row: dict) -> str:
    return next((l.strip() for l in (row.get("candidates") or "").splitlines() if l.strip()), "")


def merge_labels(target: dict, row: dict) -> None:
    """同じ欄と分かった項目に、クリックした見出しを足す（値や型は変えない）。"""
    labels = target["candidates"].splitlines() + row["candidates"].splitlines()
    target["candidates"] = "\n".join(dict.fromkeys(l for l in labels if l.strip()))
    if row["data_type"] == "table":
        columns = target["table_columns"].splitlines() + row["table_columns"].splitlines()
        target["table_columns"] = "\n".join(dict.fromkeys(c for c in columns if c.strip()))


def _column_norms(row: dict) -> set[str]:
    return {normalize_label(c) for c in (row.get("table_columns") or "").splitlines()} - {""}


# ====================================================================================================
# 元 pattern/forms.py
# 帳票の種類の編集画面の行（dict）と PatternDef の相互変換。
# ====================================================================================================

def pattern_to_rows(pattern: PatternDef) -> tuple[list[dict], list[dict]]:
    sheet_rows = [{"use": True, "sheet_name": s.sheet_name} for s in pattern.sheets]
    field_rows = [
        {
            "use": True,
            "field_name": f.field_name,
            "display_name": f.display_name,
            "candidates": "\n".join(f.candidates),
            "data_type": f.data_type,
            "direction": f.direction,
            "unit": f.unit,
            "rag_output": f.rag_output,
            "table_columns": "\n".join(f.table_columns),
            "section": f.section,
            "sheet_name": getattr(f, "sheet_name", "") or "",
            "label_cell": getattr(f, "label_cell", "") or "",
            "cell": getattr(f, "cell", "") or "",
            "renamed": bool(getattr(f, "renamed", False)),
        }
        for f in pattern.fields
    ]
    return sheet_rows, field_rows


def pattern_to_meta(pattern: PatternDef) -> dict:
    """編集画面の初期値用のメタ情報（rows_to_pattern に渡す meta と同じ形）。"""
    return {
        "name": pattern.name,
        "version": pattern.version,
        "description": pattern.description,
        "image_processing": pattern.image_processing,
        "title_fields": list(pattern.title_fields),
        "md_options": {**DEFAULT_MD_OPTIONS, **(pattern.md_options or {})},
        "version_no": pattern.version_no,
    }


def rows_to_pattern(pattern_id: int, meta: dict, sheet_rows: list[dict], field_rows: list[dict]) -> PatternDef:
    fields = [
        FieldDef(
            field_name=r["field_name"],
            display_name=r["display_name"],
            # 「値だけ」の項目（クリックで作った、見出しの無い項目）は探す見出しを持たない
            candidates=r["candidates"].splitlines() or ([] if r.get("cell") else [r["display_name"]]),
            data_type=r["data_type"],
            direction=r["direction"],
            unit=r.get("unit", "") or "",
            rag_output=r.get("rag_output", "show") if r.get("rag_output") in RAG_OUTPUTS else "show",
            table_columns=(r.get("table_columns") or "").splitlines() if r["data_type"] == "table" else [],
            section=_section_value(r.get("section")),
            sheet_name=r.get("sheet_name", "") or "",
            label_cell=r.get("label_cell", "") or "",
            cell=r.get("cell", "") or "",
            renamed=bool(r.get("renamed")),
        )
        for r in field_rows
        if r["use"]
    ]
    names = {f.field_name for f in fields}
    return PatternDef(
        id=pattern_id,
        name=meta["name"],
        version=meta.get("version", "v1"),
        description=meta.get("description", ""),
        image_processing=meta.get("image_processing", "none"),
        sheets=[SheetDef(r["sheet_name"]) for r in sheet_rows if r["use"]],
        fields=fields,
        title_fields=[n for n in dict.fromkeys(meta.get("title_fields") or []) if n in names],
        md_options={**DEFAULT_MD_OPTIONS, **(meta.get("md_options") or {})},
        version_no=int(meta.get("version_no") or 1),
    )


def _section_value(text) -> str:
    """探す区画（「回答欄」「▼ 回答欄」のどちらで入力しても、見出しと同じ比較用の名前にする）。"""

    return section_name(text) if str(text or "").strip() else ""


# ====================================================================================================
# 元 pattern/matcher.py
# アップロードされたExcelがどのテンプレートに合うかを判定する（ルールベース）。
#
# AI判定を追加する場合は rank_patterns の結果（上位候補が僅差のとき等）をAIに渡して並べ替える。
# ====================================================================================================

MIN_CONTINUATION_FIELDS = 2


@dataclass
class PatternMatch:
    pattern: PatternDef
    confidence: int          # 0〜100。見つかった項目の割合が主（画面には出さない。評価の JSON に残る）
    sheet_names: list[str]
    found_fields: int
    total_fields: int
    sheet_score: float = 0.0  # 種類のシート名がブックのシート名と合う度合い（1.0 = 同じ名前があった）
    score: float = 0.0        # 順位付けの点（match_score）。confidence と違い、項目の少ない種類が有利にならない


# 順位付けの点 = (見つかった項目数 + シート名の一致 × SHEET_WORTH) ÷ (項目数 + PHANTOM_FIELDS)。
# 「見つかった割合」だけで並べると、4項目の種類が 4/4 で 100% になり、33項目中32項目が見つかった本物の
# 種類（97%）に勝ってしまう。分母に PHANTOM_FIELDS 個の「見つからなかったことにする項目」を足すと、
# 項目の少ない種類ほど点が伸びず（4/4 → 4/14 = 0.29、20/20 → 20/30 = 0.67、32/33 → 32/43 = 0.74）、
# 割合が同じなら見つかった数の多い方が上になる。シート名が同じなら、項目が SHEET_WORTH 個見つかったのと
# 同じだけ足す（名前が違う版もあるので、シート名だけで順位が決まるほどは重くしない）。
PHANTOM_FIELDS = 10
SHEET_WORTH = 2


def match_score(found: int, total: int, sheet_score: float = 0.0) -> float:
    """帳票の種類の順位付けの点（0〜1 弱）。found: 見つかった項目数、total: 種類の項目数。"""
    if total <= 0:
        return 0.0
    return (found + SHEET_WORTH * sheet_score) / (total + PHANTOM_FIELDS)


def rank_patterns(info: WorkbookInfo, patterns: list[PatternDef]) -> list[PatternMatch]:
    """1つのブックに合う順に種類を並べる（点が同じなら見つかった割合、それも同じなら登録順）。"""
    matches = [match_pattern(info, p) for p in patterns]
    return sorted(matches, key=lambda m: (m.score, m.confidence), reverse=True)


@dataclass
class BatchMatch:
    """まとめて置いたファイル全部に対する、1つの種類の合い具合。"""
    pattern: PatternDef
    matches: list[PatternMatch]   # ファイルごと（置いた順）
    votes: int                    # この種類が最も合ったファイルの数
    score: float                  # ファイルごとの点の平均
    found_min: int
    found_max: int
    total_fields: int
    sheet_names: list[str]        # どれかのファイルで選ばれたシート（出てきた順）


def rank_batch(ranked: list[dict[int, PatternMatch]]) -> list[BatchMatch]:
    """置かれたファイル全部をまとめた種類の順位（ranked: ファイルごとの {種類の id: PatternMatch}）。

    まず「何件のファイルでその種類が最も合ったか」（票）、同じなら点の平均で並べる。同じフォームの帳票を
    まとめて置く前提なので、多数のファイルが合意した種類を先に出し、1件だけ極端に高い点の種類には引きずられない。
    1つのファイルで点が並んだ種類には、どちらにも票を入れる。
    """
    if not ranked:
        return []
    best_scores = [max((m.score for m in r.values()), default=0.0) for r in ranked]
    out = []
    for pattern_id in ranked[0]:
        ms = [r[pattern_id] for r in ranked if pattern_id in r]
        if not ms:
            continue
        sheets: list[str] = []
        for m in ms:
            sheets.extend(n for n in m.sheet_names if n not in sheets)
        votes = sum(1 for r, top in zip(ranked, best_scores) if pattern_id in r and r[pattern_id].score >= top)
        out.append(BatchMatch(
            pattern=ms[0].pattern, matches=ms, votes=votes, score=sum(m.score for m in ms) / len(ms),
            found_min=min(m.found_fields for m in ms), found_max=max(m.found_fields for m in ms),
            total_fields=ms[0].total_fields, sheet_names=sheets,
        ))
    out.sort(key=lambda b: (b.votes, b.score), reverse=True)
    return out


def match_pattern(info: WorkbookInfo, pattern: PatternDef) -> PatternMatch:
    found_by_sheet = {
        name: {fd.field_name for fd in pattern.fields if grid.find_labels(fd.label_norms())}
        for name, grid in info.grids.items()
    }
    total = len(pattern.fields) or 1

    chosen: list[str] = []
    if pattern.sheets:
        name_scores = []
        for sd in pattern.sheets:
            # 名前がそのまま同じシートがあれば、そのシートにする。項目の数で競わせると、
            # 項目の少ない2枚目のシート（別紙など）が1枚目に負けて選ばれず、読まれなくなる
            best = next((n for n in info.grids if _sheet_name_score(sd.sheet_name, n) == 1.0), None) or max(
                info.grids,
                key=lambda n: _sheet_name_score(sd.sheet_name, n) * 0.5 + len(found_by_sheet[n]) / total,
                default=None,
            )
            score = _sheet_name_score(sd.sheet_name, best) if best else 0.0
            if best is None or (score == 0 and not found_by_sheet[best]):
                name_scores.append(0.0)
                continue
            name_scores.append(score)
            if best not in chosen:
                chosen.append(best)
        sheet_score = sum(name_scores) / len(name_scores)
    else:
        best = max(info.grids, key=lambda n: len(found_by_sheet[n]), default=None)
        if best and found_by_sheet[best]:
            chosen.append(best)
        sheet_score = 1.0 if chosen else 0.0

    if chosen:
        _add_continuation_sheets(info, pattern, chosen, found_by_sheet)
    found = set().union(*(found_by_sheet[n] for n in chosen)) if chosen else set()
    field_score = len(found) / len(pattern.fields) if pattern.fields else 0.0
    confidence = round(100 * (0.25 * sheet_score + 0.75 * field_score))
    return PatternMatch(pattern, confidence, chosen, len(found), len(pattern.fields),
                        sheet_score=sheet_score, score=match_score(len(found), len(pattern.fields), sheet_score))


def _add_continuation_sheets(info: WorkbookInfo, pattern: PatternDef, chosen: list[str],
                             found_by_sheet: dict[str, set[str]]) -> None:
    """「8D報告(1)」「8D報告(2)」のように1件の帳票が複数シートに分かれている場合、続きのシートも選ぶ。

    選んだシートに無い項目の値が2つ以上（項目数の1割以上）あるシートを足す。一覧表らしいシート・非表示シートと、
    見出しだけで値の無いシート（記入要領など）は足さない。選んだシートはブック内の順に並べ直す。
    """
    stop_labels = pattern.label_norms() | DICTIONARY_NORMS

    def with_value(name: str) -> set[str]:
        grid = info.grids[name]
        return {fd.field_name for fd in pattern.fields
                if fd.field_name in found_by_sheet[name] and locate_value(grid, fd, stop_labels)[1]}

    covered = set().union(*(with_value(n) for n in chosen))
    need = max(MIN_CONTINUATION_FIELDS, 0.1 * len(pattern.fields))
    candidates = []
    for name, grid in info.grids.items():
        if name in chosen or grid.hidden or len(found_by_sheet[name] - covered) < need:
            continue
        candidates.append((name, with_value(name)))
    for name, values in sorted(candidates, key=lambda c: -len(c[1])):
        if len(values - covered) >= need and not _is_table_like(info.grids[name]):
            chosen.append(name)
            covered |= values
    order = {name: i for i, name in enumerate(info.grids)}
    chosen.sort(key=order.get)


def _sheet_name_score(expected: str, actual: str) -> float:
    e, a = normalize_sheet_name(expected), normalize_sheet_name(actual)
    if e == a:
        return 1.0
    if e and a and (e in a or a in e):
        return 0.6
    return 0.0


# ---- 一覧表らしさ（帳票フローから「一覧表の取り込みへ」を提案する判定） ----

MIN_TABLE_ROWS = 10
MIN_HEADER_CELLS = 3
DOMINANT_TABLE_ROWS = 30
SIMILAR_ROW_RATIO = 0.6


def table_like_sheets(info: WorkbookInfo, sheet_names: list[str] | None = None) -> list[str]:
    names = sheet_names if sheet_names is not None else info.sheet_names
    return [n for n in names if n in info.grids and _is_table_like(info.grids[n])]


def _is_table_like(grid: SheetGrid) -> bool:
    if getattr(grid, "hidden", False):
        return False
    run, text_rows = _table_run_length(grid)
    return run >= MIN_TABLE_ROWS and (run >= DOMINANT_TABLE_ROWS or run * 2 >= text_rows)


def _table_run_length(grid: SheetGrid) -> tuple[int, int]:
    """戻り値: (見出し行の直下から続く「同じ形の行」の最大行数, 値のある行数)"""
    rows: dict[int, set[int]] = {}
    header_like: set[int] = set()
    by_row: dict[int, list[Cell]] = {}
    for cell in grid.text_cells():
        by_row.setdefault(cell.row, []).append(cell)
    for r, cells in by_row.items():
        rows[r] = {c.col for c in cells}
        strings = [c for c in cells if isinstance(c.value, str) and not _NUMBER_LIKE.match(c.norm)]
        if len(cells) >= MIN_HEADER_CELLS and len(strings) >= len(cells) * 0.8:
            header_like.add(r)

    best = 0
    for header_row in sorted(header_like):
        start = header_row + 1
        if start not in rows or start in header_like:
            continue
        base, run, r = rows[start], 0, start
        # 空行をはさまず、列の組が先頭データ行と似ている行を数える
        while r in rows and len(rows[r]) >= 2 and _similar(rows[r], base):
            run += 1
            r += 1
        best = max(best, run)
    return best, len(rows)


def _similar(a: set[int], b: set[int]) -> bool:
    return len(a & b) >= SIMILAR_ROW_RATIO * len(a | b)


# ====================================================================================================
# 元 export/formats.py
# 読み取り結果から、RAG 投入用の Markdown を生成する。
#
# Markdown は LightRAG 調査の指針（docs/design.md 6章）に従う。
#   - 1帳票だけで意味が通るように、種類・識別番号・設備・日付をタイトルと本文に書く
#   - 定型文・種類の版・DBの文書ID・セル座標は出さない（JSON側に残す）
#   - 値は NFKC＋空白の畳み込み、数値は単位付き。「出さない」にした項目だけ省く（人名の項目も出す）
#   - 値が別の欄の見出し語そのもの（読み取り誤り。「- 数量: 単価」）の項目は出さない（JSON には残す）
#   - 明細表は見出しの下に1行1明細で「- 品番: X／品名: Y／数量: 2」と書く（パイプ表は使わない）
#   - 同じ入力からは同じバイト列になる（生成日時などを書かない）
# ====================================================================================================

# 長文項目の見出しに識別子を入れるのは、推定トークン数がこれを超える帳票だけ。
# LightRAG が帳票を2つ以上の断片に切りうる大きさ（サーバー既定の固定窓 1,200トークン）に合わせる。
# 推定式が実トークン以上になったので、1,200 未満の帳票＝1断片に収まる帳票には識別子を入れない。
HEADING_IDENTIFIER_TOKENS = 1200
# 明細表の1節（`## 見出し` から次の見出しまで）の推定トークン数の上限。
# これを超えると、明細表の途中で切れた断片に識別番号も設備名も1文字も入らないことがあるので、
# 識別子を入れる帳票では「（続き）」の見出しで分ける（固定窓 1,200 に対して余裕を取った値）。
TABLE_SECTION_TOKENS = 800
# 値が見出し語かどうかを見るのは短い値だけ（長い本文にたまたま同じ語が入っていても消さない）
MAX_LABEL_VALUE_CHARS = 20

_ISO_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?: \d{2}:\d{2})?$")  # 時刻付き（発生日時）も
# 明細表の合計行の先頭（読み取り側と同じ決まり。excel.tables.TOTAL_LABEL_RE）
_TOTAL_LABEL = TOTAL_LABEL_RE


# ---- Markdown -----------------------------------------------------------------

def build_markdown(doc: dict, extraction: dict) -> str:
    """RAG（LightRAG）に投入する Markdown。書式は docs/design.md 6.1。"""
    pattern = extraction["pattern"]
    type_name = _md_one_line(pattern["name"])
    shown = _shown_fields(extraction)
    filled = [f for f in shown if not _is_blank(f["value"])]
    title_fields = _title_fields(pattern, extraction["fields"], shown)

    title_texts = _title_texts(title_fields, heading=False)
    heading_texts = _title_texts(title_fields, heading=True)
    stem = _md_one_line(Path(doc["file_name"]).stem)
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
    tail.append(f"- 出典: {_source_text(doc, title_fields, filled)}")

    long_fields = [f for f in filled if f["data_type"] in ("text", "table")]

    def render(with_identifier: bool) -> str:
        blocks = [head, basics]
        for f in long_fields:
            heading = _md_one_line(f["display_name"])
            if with_identifier and identifier:
                heading += f"（{identifier}）"
            if f["data_type"] == "table":
                lines = table_markdown_lines(f["value"])
                if not (with_identifier and identifier):
                    blocks.append([f"## {_escape_line(heading)}", *lines])
                    continue
                # 長い明細表は「（続き）」の見出しで分ける（どの断片にも識別子が入るように）
                base = _md_one_line(f["display_name"])
                for i, part in enumerate(_split_table_lines(lines)):
                    part_heading = heading if i == 0 else f"{base}（続き）（{identifier}）"
                    blocks.append([f"## {_escape_line(part_heading)}", *part])
                continue
            lines = [line for line in _format_value(f).split("\n") if line.strip()]
            blocks.append([f"## {_escape_line(heading)}", *(_escape_line(line) for line in lines)])
        blocks.append(tail)
        return _join_blocks(blocks)

    text = render(False)
    if long_fields and identifier and _estimate_tokens(text) > HEADING_IDENTIFIER_TOKENS:
        text = render(True)
    return text


def _split_table_lines(lines: list[str]) -> list[list[str]]:
    """明細表の行を、1節が TABLE_SECTION_TOKENS に収まるまとまりに分ける。

    分けないと、明細表の行だけで埋まった断片（LightRAG のチャンク）ができ、
    その断片の中に識別番号も設備名も日付も1文字も無くなる（docs/research.md 2章（LightRAG オフライン評価）8.5）。
    """
    parts: list[list[str]] = []
    current: list[str] = []
    tokens = 0
    for line in lines:
        n = _estimate_tokens(line)
        if current and tokens + n > TABLE_SECTION_TOKENS:
            parts.append(current)
            current, tokens = [], 0
        current.append(line)
        tokens += n
    if current or not parts:
        parts.append(current)
    return parts


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

def _shown_fields(extraction: dict) -> list[dict]:
    """Markdown に出す項目（「出さない」にした項目と、値が見出し語だけの項目を除く）。

    人名の項目も出す（読み取った内容は削らない。docs/design.md 6章）。
    """
    labels = set(extraction["pattern"].get("labels") or ())
    return [f for f in extraction["fields"]
            if f.get("rag_output", "show") != "omit" and not _is_label_value(f, labels)]


def _is_label_value(f: dict, labels: set[str]) -> bool:
    """値が別の欄（またはこの欄）の見出し語そのものか。読み取り誤りで「- 品名: 品番」になった行を出さないための判定。

    LightRAG オフライン評価 3章: `- 品名: 品番`（F4 30/30）など、値が見出し語のままの行が LLM のエンティティになっていた。
    """
    if not labels or f["data_type"] in ("table", "date", "number") or f.get("edited"):
        return False
    text = _md_one_line(f["value"])
    if not text or len(text) > MAX_LABEL_VALUE_CHARS or "\n" in str(f["value"] or ""):
        return False
    return normalize_label(text) in labels


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
    found = _fallback_field(filled, title_fields)
    return _md_one_line(_plain_value(found)) if found else ""


def _fallback_field(filled: list[dict], title_fields: list[dict]) -> dict | None:
    """_fallback_identifier で使う項目（番号らしい項目 → 日付の項目）。"""
    used = {f["field_name"] for f in title_fields}
    rest = [f for f in filled if f["field_name"] not in used
            and f["field_name"] not in ("equipment_id", "equipment_name")]
    numbered = next((f for f in rest if f["data_type"] in ("string", "number")
                     and _NUMBER_LABEL.search(_md_one_line(f.get("display_name")))), None)
    dated = next((f for f in rest if f["data_type"] == "date"), None)
    return numbered or dated


def _title_texts(fields: list[dict], heading: bool) -> list[str]:
    """設備番号と設備名が両方あれば1つにまとめる（タイトル: 名前（番号）、見出し: 番号 名前）。"""
    names = {f["field_name"]: f for f in fields}
    pair = "equipment_id" in names and "equipment_name" in names
    texts: list[str] = []
    done_pair = False
    for f in fields:
        if pair and f["field_name"] in ("equipment_id", "equipment_name"):
            if not done_pair:
                eq_id = _md_one_line(_plain_value(names["equipment_id"]))
                eq_name = _md_one_line(_plain_value(names["equipment_name"]))
                texts.append(f"{eq_id} {eq_name}" if heading else f"{eq_name}（{eq_id}）")
                done_pair = True
            continue
        texts.append(_md_one_line(_plain_value(f)))
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
                lines += _md_bullet("設備", value)
                done_pair = True
            continue
        lines += _md_bullet(_md_one_line(f["display_name"]), _format_value(f))
    return lines


def _format_value(f: dict) -> str:
    """本文用の値。日付は「2026-09-14（2026年9月）」、数値は単位付き。"""
    value = f["value"]
    text = _plain_value(f) if f["data_type"] != "text" else md_value_text(value)
    if f["data_type"] == "date":
        m = _ISO_DATE.match(text)
        if m:
            text = f"{text}（{int(m[1])}年{int(m[2])}月）"
    elif f["data_type"] == "number" and _is_number(value) and f.get("unit"):
        text = f"{text}{md_value_text(f['unit'])}"
    return text


def _plain_value(f: dict) -> str:
    """タイトル・ファイル名用の値（印や単位なし）。"""
    value = f["value"]
    if _is_number(value):
        return _md_number_text(value)
    if f["data_type"] == "text":
        return _md_one_line(value)
    return md_value_text(value)


def _source_text(doc: dict, title_fields: list[dict], filled: list[dict] | None = None) -> str:
    """出典: 元ファイル名（報告番号 R2026-00123）。報告番号がなければ最初の文字列のタイトル項目。

    タイトル項目が設備だけのときは、タイトルに足した項目（作業No.などの番号 → 日付）を書く。
    """
    file_name = _md_one_line(doc["file_name"])
    ident = next((f for f in title_fields if f["field_name"] == "report_id"), None)
    ident = ident or next((f for f in title_fields if f["data_type"] == "string"
                           and f["field_name"] not in ("equipment_id", "equipment_name")), None)
    if ident is None and filled and _has_equipment_only(title_fields):
        ident = _fallback_field(filled, title_fields)
    if ident is None:
        return file_name
    return f"{file_name}（{_md_one_line(ident['display_name'])} {_plain_value(ident)}）"


def table_markdown_lines(value) -> list[str]:
    """明細表の1行を「- 品番: X／品名: Y」の1行にする。空のセルは書かない。合計行は「- 合計: 投入数: 50枚／…」。"""
    lines = []
    if not is_table_value(value):
        return lines
    # 手で入れた行の「No: 1」を、読み取った表と同じように落とす（読み取り側の Table.to_value と同じ決まり）
    value = drop_seq_column(value)
    for row in value["rows"]:
        items = [(_md_one_line(k), _md_one_line(v)) for k, v in table_row_items(value, row)]
        items = [(k, v) for k, v in items if v]
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


def _md_number_text(value) -> str:
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.10f}".rstrip("0").rstrip(".")
    return str(value)


def _md_one_line(value) -> str:
    return " ".join(md_value_text(value).split())


# ---- Markdown テキスト処理（core/mdtext・core/naming があればそれを使う） -------------------

def md_value_text(text) -> str:
    """値の正規化（NFKC＋空白の畳み込み）。丸数字などの囲み文字（①②Ⓐ㋐）は原文どおり残す。

    ①→1 にすると「①破損ウェーハ片を回収」が「1破損ウェーハ片を回収」になり、番号と本文の区切りが消える。
    ㈱→(株)、⑴→(1) のように区切りが残る表記は今までどおり正規化する（core/mdtext.nfkc_keep_enclosed）。
    """
    s = "" if text is None else str(text).replace("_x000D_", "")
    return core.nfkc_value(s)


def _escape_line(line: str) -> str:
    return core.escape_md_line(line)


def _md_bullet(label: str, value: str) -> list[str]:
    return core.md_bullet(label, value)


def _estimate_tokens(text: str) -> int:
    return core.estimate_tokens(text)


def _join_blocks(blocks: list[list[str]]) -> str:
    return core.join_blocks(blocks)


def _safe_filename_part(text: str, max_len: int = 60) -> str:
    return core.safe_filename_part(text, max_len)


def _md_filename(parts: list[str]) -> str:
    return core.md_filename(parts)
