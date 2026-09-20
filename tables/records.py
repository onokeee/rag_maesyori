"""記録の値の見方（設備の列、日付の月、数値の書き方）。

`tables/summaries.py`（月次集計・設備別年度集計・データセット説明）から、記録ファイルの作成に必要な
共通の処理だけを残したもの。集計は 2026-09-20 に外した（docs/design.md 6.3）。
書き方（桁区切り・単位）を1か所にまとめて、同じデータから同じ md が出るようにする。
"""
from __future__ import annotations

import math
import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal

from tables.spec import TableSpec


# ---- 書式 -----------------------------------------------------------------------------

def round_half_up(value: float, digits: int = 0) -> float | int:
    if not math.isfinite(value):
        return value  # 桁あふれした値（inf）は丸めない（Decimal の quantize が落ちるため）
    q = Decimal(1).scaleb(-digits)
    d = Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP)
    return int(d) if digits == 0 else float(d)


def fmt_number(value, max_decimals: int = 2) -> str:
    """桁区切り付き。小数は最大2桁（末尾の0は削る）。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(int(value))
    if isinstance(value, int):
        return f"{value:,}"
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if not math.isfinite(f):
        return str(value)  # 桁あふれした値などで、丸め（Decimal）を落とさない
    if f.is_integer():
        return f"{int(f):,}"
    r = round_half_up(f, max_decimals)
    text = f"{r:,.{max_decimals}f}".rstrip("0").rstrip(".")
    return text


def unit_label(unit: str) -> str:
    u = (unit or "").strip()
    return {"h": "時間", "H": "時間", "hr": "時間", "min": "分"}.get(u, u)


# ---- 月 -------------------------------------------------------------------------------

def month_label(month: str) -> str:
    return f"{int(month[:4])}年{int(month[5:7])}月"


def month_first_day(month: str) -> str:
    return f"{month}-01"


def month_last_day(month: str) -> str:
    # 翌月を経由しない（9999-12 の翌月は日付にできない。9999/12/31 は「期限なし」によく使われる）
    import calendar
    from datetime import date

    y, m = int(month[:4]), int(month[5:7])
    return date(y, m, calendar.monthrange(y, m)[1]).isoformat()


def is_month(text) -> bool:
    s = str(text or "")
    return len(s) >= 7 and s[4] == "-" and s[:4].isdigit() and s[5:7].isdigit() and 1 <= int(s[5:7]) <= 12


# ---- 設備の列 --------------------------------------------------------------------------

def entity_columns(spec: TableSpec):
    entity = spec.first_role("entity")
    label = spec.first_role("entity_label")
    return entity, label


_CODE_NAME_PAREN = re.compile(r"^([0-9A-Za-z][0-9A-Za-z\-_/.]{0,19})[ 　]*[（(][ 　]*(.+?)[ 　]*[)）]$")
_CODE_NAME_SPACE = re.compile(r"^([0-9A-Za-z][0-9A-Za-z\-_/.]{0,19})[ 　]+(\S.*)$")
# 「名前（番号）」の並び（「Oxideエッチャ 2号機（ETC-302）」）。番号が後ろにあっても同じ設備として扱う。
_NAME_CODE_PAREN = re.compile(r"^(.+?)[ 　]*[（(][ 　]*([0-9A-Za-z][0-9A-Za-z\-_/.]{0,19})[ 　]*[)）]$")
_HAS_DIGIT = re.compile(r"\d")
_HAS_ALPHA = re.compile(r"[A-Za-z]")
_CODE_ONLY = re.compile(r"^[0-9A-Za-z][0-9A-Za-z\-_/.]*$")
# 番号の後ろの括弧が名前ではなく但し書きのとき（「IMP-602（推定）」）。番号だけを残す
_QUALIFIERS = {"推定", "仮", "予定", "調査中", "不明", "未定", "確認中", "暫定", "候補", "要確認", "代替", "予備", "旧", "新"}
# 設備の列の値がこれだけなら、設備が決まっていない（設備ごとのファイル分けに入れない）
_PLACEHOLDER_ENTITIES = {"推定", "仮", "予定", "調査中", "不明", "未定", "確認中", "暫定", "要確認",
                         "〃", "′′", "同上", "々", "仝"}  # 補えなかった「上と同じ」の記号も設備ではない


def split_entity_code(text) -> tuple[str, str]:
    """「ETC-302(OXIDEエッチャ 2号機)」「CVD-203 W-CVD 3号機」「Oxideエッチャ 2号機（ETC-302）」→ (設備番号, 名前)。

    「IMP-602（推定）」のように括弧が但し書きなら (番号, "")。分けられなければ (原文, "")。設備名の列がない台帳で、同じ設備が「番号だけ」「番号＋名前」「名前＋番号」と
    揺れるとファイル分けが割れるため、番号にそろえる。
    """
    s = " ".join(str(text or "").split())
    # 後ろの括弧の番号を先に見る: 「Fab1 OHTシステム（OHT-801）」の番号は Fab1 ではなく OHT-801
    for pattern in (_NAME_CODE_PAREN, _CODE_NAME_PAREN, _CODE_NAME_SPACE):
        m = pattern.match(s)
        if not m:
            continue
        if pattern is _NAME_CODE_PAREN:
            code, name = m.group(2), m.group(1).strip()
        else:
            code, name = m.group(1), m.group(2).strip()
        if not name or not _HAS_DIGIT.search(code) or not _HAS_ALPHA.search(code):
            continue
        if pattern is _CODE_NAME_PAREN and name in _QUALIFIERS:
            return code, ""  # 番号は残し、但し書きは設備名にしない
        if all(_CODE_ONLY.match(p) for p in re.split(r"[\s,、/]+", name) if p):
            continue  # 名前の側も番号だけ（「CMP-101 / CMP-102」）なら分けない
        return code, name
    return s, ""


def entity_value(values: dict, spec: TableSpec) -> tuple[str, str]:
    """(設備番号, 設備名)。設備名の列がない code 列では「番号(名前)」を分ける（ファイル分けの単位をそろえる）。"""
    entity, label = entity_columns(spec)
    if entity is None:
        return "", ""
    raw = " ".join(str(values.get(entity.key) or "").split())
    if unicodedata.normalize("NFKC", raw).strip("（）() ") in _PLACEHOLDER_ENTITIES:
        return "", ""  # 「調査中」だけの値は設備ではない（記録の本文には原文のまま出る）
    if label is not None:
        return raw, " ".join(str(values.get(label.key) or "").split())
    if entity.type == "code" and raw:
        code, name = split_entity_code(raw)
        if name or code != raw:
            return code, name
    return raw, ""


def entity_display(values: dict, spec: TableSpec) -> tuple[str, str, str]:
    """(設備番号, 設備名, 表示「設備名（設備番号）」)。"""
    eid, name = entity_value(values, spec)
    if eid and name and name != eid:
        return eid, name, f"{name}（{eid}）"
    return eid, name, eid or name


__all__ = ["entity_columns", "entity_display", "entity_value", "fmt_number", "is_month", "month_first_day",
           "month_label", "month_last_day", "round_half_up", "split_entity_code", "unit_label"]
