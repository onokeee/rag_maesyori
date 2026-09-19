"""集計（月次・設備別年度）。数値はすべてコードで計算する（AI は使わない）。

並び順と数値の書き方を固定して、同じデータから同じ結果が出るようにする:
桁区切りあり、平均は四捨五入（整数の列は整数、小数の列は小数1桁）、内訳は件数の多い順→名前順。
"""
from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter, defaultdict
from decimal import ROUND_HALF_UP, Decimal

from tables.spec import SummarySpec, TableSpec

MAX_COVERAGE_MONTHS = 600  # 取り込み範囲がこれより長ければ、記録のある月だけを使う
# 記録のある月どうしがこれより離れていれば、間の0件の月は作らない（打ち間違えた遠い年の1件で、
# 空の月次集計が何百ファイルもできるのを防ぐ）
MAX_EMPTY_GAP_MONTHS = 12


# ---- 書式 -----------------------------------------------------------------------------

def round_half_up(value: float, digits: int = 0) -> float | int:
    if not math.isfinite(value):
        return value  # 桁あふれした合計（inf）は丸めない（Decimal の quantize が落ちるため）
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
        return str(value)  # 桁あふれした合計などで、丸め（Decimal）を落とさない
    if f.is_integer():
        return f"{int(f):,}"
    r = round_half_up(f, max_decimals)
    text = f"{r:,.{max_decimals}f}".rstrip("0").rstrip(".")
    return text


def unit_label(unit: str) -> str:
    u = (unit or "").strip()
    return {"h": "時間", "H": "時間", "hr": "時間", "min": "分"}.get(u, u)


def fmt_measure(value, unit: str, with_hours: bool = False) -> str:
    """「2,460分（41.0時間）」のような書き方。"""
    label = unit_label(unit)
    text = f"{fmt_number(value)}{label}"
    if with_hours and label == "分" and isinstance(value, (int, float)) and abs(value) >= 60:
        hours = round_half_up(float(value) / 60, 1)
        text += f"（{hours:,.1f}時間）"
    return text


def fmt_average(total: float, count: int, integer_values: bool) -> int | float:
    if not count:
        return 0
    avg = total / count
    if integer_values:
        rounded = round_half_up(avg, 0)
        if rounded != 0 or total == 0:
            return rounded  # 整数だけの列でも、0 に丸まるときは小数1桁で出す（0 と書くと事実と違う）
    return round_half_up(avg, 1)


def month_label(month: str) -> str:
    return f"{int(month[:4])}年{int(month[5:7])}月"


def month_add(month: str, n: int) -> str:
    y, m = int(month[:4]), int(month[5:7])
    total = y * 12 + (m - 1) + n
    if total // 12 > 9999:
        return "9999-12"  # 日付にできる最後の月で止める（9999年度の終わりなど）
    return f"{total // 12:04d}-{total % 12 + 1:02d}"


def month_range(start: str, end: str) -> list[str]:
    out = []
    cur = start
    while cur <= end and len(out) <= MAX_COVERAGE_MONTHS:
        out.append(cur)
        nxt = month_add(cur, 1)
        if nxt == cur:
            break
        cur = nxt
    return out


def month_first_day(month: str) -> str:
    return f"{month}-01"


def month_last_day(month: str) -> str:
    # 翌月を経由しない（9999-12 の翌月は日付にできない。9999/12/31 は「期限なし」によく使われる）
    import calendar
    from datetime import date

    y, m = int(month[:4]), int(month[5:7])
    return date(y, m, calendar.monthrange(y, m)[1]).isoformat()


def fiscal_year_of(month: str, start_month: int) -> int:
    y, m = int(month[:4]), int(month[5:7])
    return y if m >= start_month else y - 1


def is_month(text) -> bool:
    s = str(text or "")
    return len(s) >= 7 and s[4] == "-" and s[:4].isdigit() and s[5:7].isdigit() and 1 <= int(s[5:7]) <= 12


# ---- 列の役割 ---------------------------------------------------------------------------

def measure_columns(spec: TableSpec) -> list[tuple[str, str, str]]:
    """集計する数値列 (キー, 表示名, 単位)。単価は合計しても意味がないので除く。"""
    out = []
    for col in spec.columns:
        if col.type == "number" and col.role == "measure" and "単価" not in col.display and not col.key.startswith("unit_price"):
            out.append((col.key, col.display, col.unit))
    return out


def resolve_metrics(summary: SummarySpec, spec: TableSpec) -> dict:
    """{"count": bool, "sum": [key], "avg": [key], "max": [key]}。count だけの指定は数値列の合計（年度は平均も）に広げる。"""
    measures = [k for k, _d, _u in measure_columns(spec)]
    out = {"count": False, "sum": [], "avg": [], "max": []}
    for metric in summary.metrics or ["count"]:
        if metric == "count":
            out["count"] = True
            continue
        kind, _, key = metric.partition(":")
        if kind in ("sum", "avg", "max") and key and key not in out[kind]:
            out[kind].append(key)
    if not out["sum"] and not out["avg"] and not out["max"]:
        out["sum"] = list(measures)
        if summary.id == "entity_fiscal_year":
            out["avg"] = list(measures)
    out["count"] = True
    return out


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
# 設備の列の値がこれだけなら、設備が決まっていない（設備別の集計・ファイル分けに入れない）
_PLACEHOLDER_ENTITIES = {"推定", "仮", "予定", "調査中", "不明", "未定", "確認中", "暫定", "要確認"}


def split_entity_code(text) -> tuple[str, str]:
    """「ETC-302(OXIDEエッチャ 2号機)」「CVD-203 W-CVD 3号機」「Oxideエッチャ 2号機（ETC-302）」→ (設備番号, 名前)。

    「IMP-602（推定）」のように括弧が但し書きなら (番号, "")。分けられなければ (原文, "")。設備名の列がない台帳で、同じ設備が「番号だけ」「番号＋名前」「名前＋番号」と
    揺れると集計とファイル分けが割れるため、番号にそろえる。
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
    """(設備番号, 設備名)。設備名の列がない code 列では「番号(名前)」を分ける（集計・ファイル分けの単位をそろえる）。"""
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


def category_column(spec: TableSpec):
    """内訳に使う区分の列。故障区分があればそれ、なければ最初の category 役割の列。

    記録に出していない列（md=omit。意味の分からないコード値など）は内訳に使わない。
    記録のどこにも出ていない値で「F01: 3件」と集計しても、資料群の中で意味を確かめられないため。
    """
    preferred = spec.column("failure_category")
    if preferred is not None and preferred.role == "category" and preferred.md != "omit":
        return preferred
    for col in spec.columns:
        if col.role == "category" and col.md != "omit":
            return col
    return None


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _month(values: dict, spec: TableSpec) -> str | None:
    v = values.get(spec.date_key)
    s = str(v or "")
    return s[:7] if is_month(s) else None


def coverage_months(coverage: dict | None, records: list[dict], spec: TableSpec) -> list[str]:
    months = sorted({m for m in (_month(r.get("values", {}), spec) for r in records) if m})
    start = (coverage or {}).get("start") or (months[0] if months else None)
    end = (coverage or {}).get("end") or (months[-1] if months else None)
    if not start or not end or not is_month(start) or not is_month(end) or start > end:
        return months
    rng = month_range(start, end)
    if len(rng) > MAX_COVERAGE_MONTHS:
        return months
    inside = [m for m in months if start <= m <= end]
    skip: set[str] = set()
    for a, b in zip(inside, inside[1:]):
        gap = month_range(month_add(a, 1), month_add(b, -1))
        if len(gap) > MAX_EMPTY_GAP_MONTHS:
            skip.update(gap)
    return sorted((set(rng) - skip) | set(inside))


def _breakdown(rows: list[dict], cat_key: str | None, sum_key: str | None) -> list[dict]:
    if not cat_key:
        return []
    counts: Counter[str] = Counter()
    sums: dict[str, float] = defaultdict(float)
    for values in rows:
        name = str(values.get(cat_key) or "（空欄）")
        counts[name] += 1
        if sum_key and _num(values.get(sum_key)) is not None:
            sums[name] += values[sum_key]
    items = sorted(counts.items(), key=lambda t: (-t[1], t[0]))
    return [{"name": name, "count": n, "sum": _clean(sums.get(name, 0)) if sum_key else None} for name, n in items]


def _clean(v):
    if isinstance(v, float):
        v = round(v, 9)
        if v.is_integer():
            return int(v)
    return v


def _integer_values(values_list: list) -> bool:
    return all(isinstance(v, int) or (isinstance(v, float) and v.is_integer()) for v in values_list)


# ---- 月次集計 ---------------------------------------------------------------------------

def month_summaries(records: list[dict], spec: TableSpec, coverage: dict | None, summary: SummarySpec) -> list[dict]:
    """月ごと（全設備）の件数・数値の合計・上位の設備・区分の内訳。取り込み範囲の月は0件でも作る。"""
    metrics = resolve_metrics(summary, spec)
    by_month: dict[str, list[dict]] = defaultdict(list)
    for rec in records:
        values = rec.get("values", {})
        m = _month(values, spec)
        if m:
            by_month[m].append(values)
    cat = category_column(spec)
    entity, _label = entity_columns(spec)
    sum_keys = metrics["sum"]
    rank_key = sum_keys[0] if sum_keys else None
    out = []
    for month in coverage_months(coverage, records, spec):
        rows = by_month.get(month, [])
        sums = {k: _clean(sum(v[k] for v in rows if _num(v.get(k)) is not None)) for k in sum_keys}
        top = []
        if entity is not None and rows:
            groups: dict[str, list[dict]] = defaultdict(list)
            for v in rows:
                eid = entity_value(v, spec)[0]
                if eid:
                    groups[eid].append(v)
            items = []
            for eid, grows in groups.items():
                total = _clean(sum(g[rank_key] for g in grows if _num(g.get(rank_key)) is not None)) if rank_key else None
                main = ""
                if cat is not None:
                    cats = Counter(str(g.get(cat.key)) for g in grows if g.get(cat.key))
                    if cats:
                        main = sorted(cats.items(), key=lambda t: (-t[1], t[0]))[0][0]
                _eid, _name, display = entity_display(grows[0], spec)
                items.append({"entity": eid, "display": display, "count": len(grows), "sum": total, "main_category": main})
            items.sort(key=lambda t: (-(t["sum"] or 0) if rank_key else 0, -t["count"], t["entity"]))
            top = items[: max(1, int(summary.top_n or 5))]
        out.append({
            "month": month, "count": len(rows), "sums": sums, "rank_key": rank_key, "top": top,
            "categories": _breakdown(rows, cat.key if cat else None, rank_key),
        })
    return out


# ---- 設備別年度集計 ------------------------------------------------------------------------

def entity_fiscal_year_summaries(records: list[dict], spec: TableSpec, coverage: dict | None,
                                 summary: SummarySpec) -> list[dict]:
    """設備×年度の件数・合計・平均・最大、月別（0件の月も）、区分の内訳。記録のある設備・年度だけ作る。"""
    entity, label = entity_columns(spec)
    if entity is None:
        return []
    metrics = resolve_metrics(summary, spec)
    fs = int(spec.fiscal_year_start_month or 4)
    cov_months = coverage_months(coverage, records, spec)
    groups: dict[tuple[str, int], list[dict]] = defaultdict(list)
    for rec in records:
        values = rec.get("values", {})
        m = _month(values, spec)
        eid = entity_value(values, spec)[0]
        if not m or not eid:
            continue
        groups[(eid, fiscal_year_of(m, fs))].append(values)
    cat = category_column(spec)
    keys = sorted(set(metrics["sum"]) | set(metrics["avg"]) | set(metrics["max"]),
                  key=lambda k: [c[0] for c in measure_columns(spec)].index(k) if k in [c[0] for c in measure_columns(spec)] else 99)
    out = []
    for (eid, fy) in sorted(groups):
        rows = groups[(eid, fy)]
        fy_start = f"{fy:04d}-{fs:02d}"
        fy_end = month_add(fy_start, 11)
        months = [m for m in cov_months if fy_start <= m <= fy_end]
        by_month: dict[str, list[dict]] = defaultdict(list)
        for v in rows:
            by_month[_month(v, spec)].append(v)
        month_items = []
        for m in months:
            mrows = by_month.get(m, [])
            month_items.append({
                "month": m, "count": len(mrows),
                "sums": {k: _clean(sum(v[k] for v in mrows if _num(v.get(k)) is not None)) for k in keys},
                "has_value": {k: any(_num(v.get(k)) is not None for v in mrows) for k in keys},
            })
        stats = {}
        for k in keys:
            vals = [v[k] for v in rows if _num(v.get(k)) is not None]
            total = _clean(sum(vals)) if vals else 0
            max_item = None
            if vals:
                best_row = max((v for v in rows if _num(v.get(k)) is not None),
                               key=lambda v: (v[k], -int((_month(v, spec) or "0000-00").replace("-", ""))))
                max_item = {"month": _month(best_row, spec), "value": best_row[k]}
            n = len(vals)
            avg_base = n  # 平均は値のある記録だけで割る（空欄は「データなし」で 0 ではない）
            stats[k] = {
                "sum": total, "n": n, "rows": len(rows),
                "avg": fmt_average(float(total), avg_base, _integer_values(vals)) if avg_base else None,
                "max": max_item,
            }
        first = rows[0]
        _eid, name, display = entity_display(first, spec)
        covered = list(months)
        out.append({
            "entity": eid, "name": name, "display": display, "fiscal_year": fy, "count": len(rows),
            "months": month_items, "stats": stats, "metrics": metrics, "keys": keys,
            "range": (covered[0], covered[-1]) if covered else (fy_start, fy_end),
            "categories": _breakdown(rows, cat.key if cat else None, keys[0] if keys and metrics["sum"] else None),
        })
    return out


def dataset_counts(records: list[dict], spec: TableSpec) -> dict:
    """データセット説明用: 件数、日付の範囲、設備の数。"""
    entity, _label = entity_columns(spec)
    months = sorted({m for m in (_month(r.get("values", {}), spec) for r in records) if m})
    dates = sorted(str(r.get("values", {}).get(spec.date_key))[:10] for r in records
                   if is_month(r.get("values", {}).get(spec.date_key)))
    entities = {entity_value(r.get("values", {}), spec)[0] for r in records} - {""} if entity is not None else set()
    return {"records": len(records), "months": months, "date_min": dates[0] if dates else None,
            "date_max": dates[-1] if dates else None, "entities": len(entities)}


__all__ = [
    "dataset_counts", "entity_display", "entity_fiscal_year_summaries", "entity_value", "fiscal_year_of", "fmt_measure",
    "fmt_number", "measure_columns", "month_label", "month_range", "month_summaries", "resolve_metrics",
    "split_entity_code",
]
