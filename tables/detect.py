"""見出し帯・データ範囲・行の分類・表の種類の判定。

- 見出し行: アンカー語 → 自動の点数付け（文字列の割合・辞書の語・下の行の型がそろっているか など）
- 2段見出し: 横に結合された上段を右へ埋めて「上_下」でつなぐ（CSVは上段の空欄を右へ埋める）
- データの終わり: 空行3行、合計行（含む）、注記行・表題行・見出しの再出現（含まない）
- MVPでは1シート1表。下や右に別の表があれば警告する
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from typing import Iterator

from tables.source import NA_TOKENS, SourceRow, value_kind

HEAD_ROWS = 80  # 見出し探索のために読む行数
HEADER_SEARCH_ROWS = 40  # 見出し行の候補にする先頭行数
PREVIEW_ROWS = 60  # row_classes に全行を載せる先頭行数
MAX_ROW_CLASSES = 500
BLANK_ROWS_END = 3
LOOKAHEAD_ROWS = 300
KEY_SAMPLE_ROWS = 50

_BLOCK_TITLE_RE = re.compile(r"^\s*[■□◆◇●○▼▽★☆]")
_NOTE_RE = re.compile(r"^\s*(※|\*\d|（注|\(注|注[記意]?\s*[\d０-９]*\s*[:：)）.．、\s])")
_TOTAL_RE = re.compile(r"(総合計|合計|総計|累計)")
_SUBTOTAL_RE = re.compile(r"(小計|計|平均)$")
_PAREN_TAIL_RE = re.compile(r"\s*[（(][^()（）]*[)）]\s*$")
_UNIT_RE = re.compile(r"^(.*?)\s*[（(\[［]\s*([^()（）\[\]［］]{1,12})\s*[)）\]］]\s*$")
_KNOWN_UNITS = {
    "分", "秒", "時間", "時", "日", "日数", "h", "hr", "hrs", "H", "min", "sec", "s", "円", "千円", "万円", "百万円",
    "個", "件", "回", "本", "枚", "台", "式", "人", "名", "箇所", "ヶ所", "%", "％", "mm", "cm", "m", "km", "μm", "um",
    "nm", "g", "kg", "t", "l", "L", "ml", "mL", "kPa", "Pa", "MPa", "V", "A", "W", "kW", "kWh", "℃", "°C", "pcs",
    "ppm", "rpm", "L/min", "sccm", "slm", "Torr", "件数", "人時", "工数", "h/人",
}
_MONTH_RES = [
    re.compile(r"^(\d{1,2})月(分)?$"),
    re.compile(r"^(\d{4})[/\-.年](\d{1,2})月?$"),
    re.compile(r"^(R|H|令和|平成)(\d{1,2}|元)[./年](\d{1,2})月?$"),
    re.compile(r"^(\d{2})/(\d{1,2})$"),
    re.compile(r"^(\d{4})-(\d{2})-01$"),  # 月初日の日付セルを見出しにした表
    re.compile(r"^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?$", re.I),
]


@dataclass
class RowClass:
    index: int
    kind: str  # title/header/data/continuation/subtotal/note/blank/excluded
    reason: str = ""


@dataclass
class LayoutGuess:
    sheet: str
    table_kind: str  # list/crosstab/form_like/unknown
    header_rows: list[int]
    data_start: int
    data_end: int
    headers: list[str]  # 2段は "上_下" で結合、結合セルは右へ埋める
    row_classes: list[RowClass]  # 先頭60行＋データ範囲の要約
    confidence: float
    warnings: list[str]
    header_levels: list[list[str]] = field(default_factory=list)  # 見出しの段ごとの元の文字列（表示用）
    key_columns: list[int] = field(default_factory=list)  # 継続行の判定に使う列（0始まり）
    counts: dict[str, int] = field(default_factory=dict)  # データ範囲内の行の種類ごとの件数


# ---- 公開関数 ----

def split_header_unit(header: str) -> tuple[str, str]:
    """「停止時間(分)」→ ("停止時間", "分")。単位らしくない括弧は分けない。"""
    text = str(header or "").strip()
    m = _UNIT_RE.match(text)
    if not m or not m.group(1).strip():
        return text, ""
    unit = unicodedata.normalize("NFKC", m.group(2)).strip()
    if unit in _KNOWN_UNITS or unit.lower() in {u.lower() for u in _KNOWN_UNITS if u.isascii()}:
        return m.group(1).strip(), unit
    return text, ""


def is_month_label(text: str) -> bool:
    """見出しが月・年月として読めるか（「4月」「2026/04」「R8.4」「2025年_4月」など）。"""
    s = unicodedata.normalize("NFKC", str(text or "")).strip().replace(" ", "")
    if not s:
        return False
    last = s.split("_")[-1]
    return any(r.match(last) for r in _MONTH_RES) or any(r.match(s) for r in _MONTH_RES)


def guess_layout(source, sheet, anchors: list[str] | None = None, header_row: int | None = None,
                 data_end: int | None = None, header_rows: list[int] | None = None,
                 max_scan_rows: int | None = None) -> LayoutGuess:
    """表の見出し帯・データ範囲・種類を推定する。

    header_rows を渡すとそのまま使う。header_row だけなら2段見出しかを自動で確かめる。
    max_scan_rows を渡すと、データの終わりの探索をその行数で打ち切る（簡易判定用）。
    """
    head = list(source.rows(sheet, 1, HEAD_ROWS))
    warnings: list[str] = []
    if not head or all(r.is_blank for r in head):
        return LayoutGuess(sheet, "unknown", [], 1, 0, [], [RowClass(r.index, "blank") for r in head[:PREVIEW_ROWS]],
                           0.0, ["表が見つかりません（先頭に値のある行がありません）"])
    by_index = {r.index: r for r in head}
    is_csv = getattr(source, "kind", "") == "csv"

    score = 1.0
    if header_rows:
        rows_h = sorted(header_rows)
    else:
        if header_row:
            best = header_row
        else:
            best, score = _choose_header_row(head, anchors)
            if best is None:
                return _no_table(sheet, head, "見出し行が見つかりません（2列以上に文字が並ぶ行がありません）")
        rows_h = _header_band(by_index, best, is_csv)

    levels, headers = _build_headers(by_index, rows_h, is_csv)
    width, split_warning = _table_width(by_index, rows_h, levels)
    if split_warning:
        warnings.append(split_warning)
    headers = _dedupe(headers[:width])
    levels = [lv[:width] for lv in levels]
    data_start = rows_h[-1] + 1

    header_norms = {_norm(h) for h in headers if h and not h.startswith("列")}
    ctx = _Ctx(width=width, header_norms=header_norms, key_cols=[], is_csv=is_csv, header_rows=rows_h)
    ctx.key_cols = _auto_key_columns(source, sheet, data_start, ctx)

    classes, counts, end, scan_warnings = _scan(
        source, sheet, head, ctx, data_start, data_end, max_scan_rows)
    warnings.extend(scan_warnings)

    months = sum(1 for h in headers if is_month_label(h))
    data_rows = counts.get("data", 0)
    if months >= 3 and months >= 0.3 * max(1, len([h for h in headers if not h.startswith("列")])):
        table_kind = "crosstab"
    elif _labels_down_first_column(head, rows_h, data_start, end):
        table_kind = "form_like"
    elif data_rows >= 3 and score >= 0.4:
        table_kind = "list"
    elif _form_like(head):
        table_kind = "form_like"
    else:
        table_kind = "unknown"
    if table_kind in ("form_like", "unknown"):
        warnings.append("一覧表の形に見えません（見出しの下に同じ形の行が続いていません）")

    hidden_cols = _hidden_columns(source, sheet, width)
    if hidden_cols:
        warnings.append(f"非表示の列があります（{', '.join(hidden_cols)}列）")

    confidence = round(max(0.0, min(1.0, score)) * min(1.0, data_rows / 5), 3)
    return LayoutGuess(
        sheet=sheet, table_kind=table_kind, header_rows=rows_h, data_start=data_start, data_end=end,
        headers=headers, row_classes=classes, confidence=confidence, warnings=warnings,
        header_levels=levels, key_columns=ctx.key_cols, counts=counts,
    )


def classify_rows(source, sheet, layout: LayoutGuess, key_columns: list[int] | None = None
                  ) -> Iterator[tuple[SourceRow, RowClass]]:
    """データ範囲（data_start〜data_end）の行を分類しながら返す。"""
    ctx = _ctx_from_layout(layout)
    ctx.is_csv = getattr(source, "kind", "") == "csv"
    if key_columns is not None:
        ctx.key_cols = list(key_columns)
    if layout.data_end < layout.data_start:
        return
    limit = layout.data_end - layout.data_start + 1
    for row in source.rows(sheet, layout.data_start, limit):
        yield row, _classify(row, ctx)


def sample_data_rows(source, sheet, layout: LayoutGuess, n: int = 200) -> list[SourceRow]:
    """列の対応づけ候補に使う、先頭のデータ行（継続行・小計などを除く）。"""
    out: list[SourceRow] = []
    for row, rc in classify_rows(source, sheet, layout):
        if rc.kind == "data":
            out.append(row)
            if len(out) >= n:
                break
    return out


def looks_like_list(source, sheet) -> bool:
    """一覧表らしいか（見出し行の下に同じ形の行が10行以上）。帳票フローからも使う。"""
    try:
        layout = guess_layout(source, sheet, max_scan_rows=200)
    except Exception:
        return False
    return layout.table_kind in ("list", "crosstab") and layout.counts.get("data", 0) >= 10


# ---- 内部: 行の事実 ----

@dataclass
class _Ctx:
    width: int
    header_norms: set[str]
    key_cols: list[int]
    is_csv: bool
    header_rows: list[int]


def _ctx_from_layout(layout: LayoutGuess) -> _Ctx:
    width = len(layout.headers)
    norms = {_norm(h) for h in layout.headers if h and not h.startswith("列")}
    return _Ctx(width=width, header_norms=norms, key_cols=list(layout.key_columns),
                is_csv=False, header_rows=list(layout.header_rows))


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(text or ""))).lower()


def _is_header_cell(cell) -> bool:
    if not cell.text:
        return False
    return isinstance(cell.value, str) and value_kind(cell.value, cell.text) in ("string", "code") \
        and len(cell.text) <= 40 and "\n" not in cell.text


def _headerish(row: SourceRow) -> bool:
    cells = [c for c in row.cells if c.text]
    if len(cells) < 2:
        return False
    return sum(1 for c in cells if _is_header_cell(c)) / len(cells) >= 0.8


def _last_col(row: SourceRow) -> int:
    for i in range(len(row.cells) - 1, -1, -1):
        c = row.cells[i]
        if c.text or (c.merged_anchor and c.merged_anchor[0] == row.index and c.merged_anchor[1] <= i + 1
                      and row.cells[c.merged_anchor[1] - 1].text):
            return i + 1
    return 0


def _covered_positions(row: SourceRow) -> int:
    """値のあるセル＋同じ行で横に結合された範囲の数。"""
    n = 0
    for i, c in enumerate(row.cells):
        if c.text:
            n += 1
        elif c.merged_anchor and c.merged_anchor[0] == row.index and c.merged_anchor[1] - 1 < len(row.cells) \
                and row.cells[c.merged_anchor[1] - 1].text:
            n += 1
    return n


# ---- 内部: 見出し行の選択 ----

def _choose_header_row(head: list[SourceRow], anchors: list[str] | None) -> tuple[int | None, float]:
    from tables.dictionary import lookup_header

    scores: dict[int, float] = {}
    candidates = head[:HEADER_SEARCH_ROWS]
    for pos, row in enumerate(candidates):
        cells = [c for c in row.cells if c.text]
        if len(cells) < 2:
            continue
        texts = [c.text for c in cells]
        hdr_ratio = sum(1 for c in cells if _is_header_cell(c)) / len(cells)
        unique = len(set(texts)) / len(texts)
        below = [r for r in head[pos + 1:pos + 13] if not r.is_blank]
        width_below = max([_last_col(r) for r in below] or [0])
        coverage = min(1.0, _covered_positions(row) / max(1, _last_col(row), width_below))
        if below:
            consist = sum(1 for r in below if sum(1 for c in r.cells if c.text) >= max(2, 0.3 * len(cells))) / len(below)
            contrast = sum(1 for r in below if not _headerish(r)) / len(below)
        else:
            consist = contrast = 0.0
        dict_hits = 0
        for t in texts:
            if len(t) <= 20:
                found = lookup_header(t)
                if found and found[1] == "dictionary":
                    dict_hits += 1
        bold = sum(1 for c in cells if c.bold) / len(cells)
        score = (0.3 * hdr_ratio + 0.1 * unique + 0.15 * coverage + 0.15 * consist + 0.15 * contrast
                 + 0.1 * min(1.0, dict_hits / 3) + 0.05 * bold)
        if _BLOCK_TITLE_RE.match(texts[0]) or _NOTE_RE.match(texts[0]):
            score -= 0.3
        if len(cells) == 2 and len(texts[0]) <= 10 and not _is_header_cell(cells[1]):
            score -= 0.2  # 「出力日時, 2026/09/14」のような前置き行
        # 上にも同じ形の行が並んでいるなら、見出しではなくデータの途中
        above = [r for r in head[max(0, pos - 10):pos] if not r.is_blank]
        if above:
            same_shape = sum(1 for r in above if sum(1 for c in r.cells if c.text) >= 0.6 * len(cells))
            score -= 0.3 * same_shape / max(len(above), 3)
        scores[row.index] = score

    if not scores:
        return None, 0.0
    if anchors:
        anchor_norms = [_norm(a) for a in anchors if a]
        best_hits = 0
        best_row = None
        for row in candidates:
            norms = [_norm(c.text) for c in row.cells if c.text]
            hits = sum(1 for a in anchor_norms if any(a and a in n for n in norms))
            if hits > best_hits or (hits == best_hits and hits and best_row is not None
                                    and scores.get(row.index, 0) > scores.get(best_row, 0)):
                best_hits, best_row = hits, row.index
        if best_row is not None:
            return best_row, max(scores.get(best_row, 0.5), 0.5 + 0.5 * best_hits / max(1, len(anchor_norms)))
    best = max(scores, key=lambda i: (round(scores[i], 3), -i))
    return best, scores[best]


def _header_band(by_index: dict[int, SourceRow], best: int, is_csv: bool) -> list[int]:
    """選んだ行と上下の行から、2段見出しかどうかを決める。"""
    row = by_index.get(best)
    above = by_index.get(best - 1)
    below = by_index.get(best + 1)
    after_below = by_index.get(best + 2)
    after = by_index.get(best + 1)
    if row is None:
        return [best]
    if above is not None and not above.is_blank and _pair_is_header(above, row, after, is_csv):
        return [best - 1, best]
    if below is not None and not below.is_blank and _pair_is_header(row, below, after_below, is_csv):
        return [best, best + 1]
    return [best]


def _pair_is_header(top: SourceRow, bottom: SourceRow, after: SourceRow | None, is_csv: bool) -> bool:
    if after is not None and not after.is_blank and _headerish(after):
        return False
    top_cells = [c for c in top.cells if c.text]
    bottom_cells = [c for c in bottom.cells if c.text]
    if len(top_cells) < 2 or len(bottom_cells) < 2:
        return False  # 上段が1セルだけならタイトル行とみなす
    if _BLOCK_TITLE_RE.match(top_cells[0].text) or _NOTE_RE.match(top_cells[0].text):
        return False
    if not all(_is_header_cell(c) for c in top_cells) or not _headerish(bottom):
        return False
    horizontal = any(
        c.merged_anchor and c.merged_anchor[0] == top.index and c.merged_anchor[1] != i + 1
        for i, c in enumerate(top.cells)
    )
    vertical = any(c.merged_anchor and c.merged_anchor[0] == top.index for c in bottom.cells)
    if horizontal or vertical:
        return True
    if is_csv:
        # 結合のないCSV: 上段がまばらで、上段の値の真下にも下段の値がある
        if len(top_cells) >= 2 and len(top_cells) <= 0.6 * len(bottom_cells):
            positions = [i for i, c in enumerate(top.cells) if c.text]
            return all(i < len(bottom.cells) and bottom.cells[i].text for i in positions)
    return False


def _build_headers(by_index: dict[int, SourceRow], rows_h: list[int], is_csv: bool
                   ) -> tuple[list[list[str]], list[str]]:
    rows = [by_index.get(r) for r in rows_h]
    ncols = max((len(r.cells) for r in rows if r is not None), default=0)
    levels: list[list[str]] = []
    for level, row in enumerate(rows):
        values = [""] * ncols
        if row is None:
            levels.append(values)
            continue
        last_text = ""
        last_pos = -1
        for i in range(ncols):
            cell = row.cells[i] if i < len(row.cells) else None
            text = _clean_header(cell.text) if cell else ""
            anchor = cell.merged_anchor if cell else None
            if text:
                values[i] = text
                last_text, last_pos = text, i
                continue
            if anchor and anchor[0] == row.index and anchor[1] - 1 != i:
                src = row.cells[anchor[1] - 1] if anchor[1] - 1 < len(row.cells) else None
                values[i] = _clean_header(src.text) if src else ""  # 横の結合は右へ埋める
            elif anchor and anchor[0] != row.index:
                values[i] = ""  # 上の段からの縦の結合は重ねない
            elif is_csv and level < len(rows) - 1 and last_pos >= 0:
                lower = rows[-1]
                if lower is not None and i < len(lower.cells) and lower.cells[i].text:
                    values[i] = last_text  # CSVの2段見出しは上段の空欄を右へ埋める
        levels.append(values)
    headers = []
    for i in range(ncols):
        parts: list[str] = []
        for lv in levels:
            if lv[i] and (not parts or parts[-1] != lv[i]):
                parts.append(lv[i])
        headers.append("_".join(parts))
    return levels, headers


def _clean_header(text: str) -> str:
    return re.sub(r"[ \t]+", " ", (text or "").replace("\n", "")).strip()


def _dedupe(headers: list[str]) -> list[str]:
    seen: Counter[str] = Counter()
    out = []
    for i, h in enumerate(headers):
        name = h or f"列{i + 1}"
        seen[name] += 1
        out.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    return out


def _table_width(by_index: dict[int, SourceRow], rows_h: list[int], levels: list[list[str]]) -> tuple[int, str]:
    """見出しの右端。空の見出し列を挟んで右に別の表があれば、左の表だけにする。"""
    ncols = len(levels[0]) if levels else 0
    filled = [any(lv[i] for lv in levels) for i in range(ncols)]
    if not any(filled):
        return 0, ""
    last = max(i for i, f in enumerate(filled) if f) + 1
    first = min(i for i, f in enumerate(filled) if f)
    data_rows = [by_index[r] for r in sorted(by_index) if r > rows_h[-1]][:30]
    data_rows = [r for r in data_rows if not r.is_blank]
    for gap in range(first + 1, last):
        if filled[gap]:
            continue
        right = sum(1 for i in range(gap + 1, last) if filled[i])
        if right < 2:
            continue
        used = sum(1 for r in data_rows if gap < len(r.cells) and r.cells[gap].text)
        if data_rows and used > 0.1 * len(data_rows):
            continue
        from openpyxl.utils import get_column_letter

        return gap, (f"見出しの右側（{get_column_letter(gap + 2)}列以降）に別の表があるようです。"
                     "最初の表だけを読み取ります。右側の表は範囲を指定して取り込んでください")
    return last, ""


# ---- 内部: 行の分類 ----

def _aggregate_label(row: SourceRow, width: int) -> str | None:
    """小計・合計行なら "subtotal" / "total"。"""
    cells = row.cells[:width] if width else row.cells
    nonempty = [c for c in cells if c.text]
    found: list[tuple[int, str]] = []
    for pos, c in enumerate(nonempty[:3]):
        if not isinstance(c.value, str) or len(c.text) > 30 or "\n" in c.text:
            continue
        if "計" not in c.text and "平均" not in c.text:
            continue
        label = _PAREN_TAIL_RE.sub("", unicodedata.normalize("NFKC", c.text)).replace(" ", "")
        if label and len(label) <= 25:
            found.append((pos, label))
    if not found:
        return None
    strings = sum(1 for c in nonempty if c.text not in NA_TOKENS and value_kind(c.value, c.text) != "number")
    for pos, label in found:
        if _TOTAL_RE.search(label) and strings <= 3:
            return "subtotal" if "小計" in label else "total"
        if pos == 0 and _SUBTOTAL_RE.search(label) and strings <= 2:
            return "subtotal"
    return None


def _classify(row: SourceRow, ctx: _Ctx) -> RowClass:
    width = ctx.width or len(row.cells)
    cells = row.cells[:width]
    nonempty = [c for c in cells if c.text]
    if not nonempty:
        if any(c.text for c in row.cells[width:]):
            return RowClass(row.index, "excluded", "見出しの範囲外にだけ値がある行")
        return RowClass(row.index, "blank")
    if row.hidden:
        return RowClass(row.index, "excluded", "非表示の行")
    strikes = sum(1 for c in nonempty if c.strike)
    if strikes and (strikes >= 0.5 * len(nonempty) or any(k < len(cells) and cells[k].strike for k in ctx.key_cols)):
        return RowClass(row.index, "excluded", "取り消し線の行")
    first = nonempty[0].text
    if ctx.header_norms and len(first) <= 40 and _norm(first) in ctx.header_norms:
        norms = {_norm(c.text) for c in nonempty}
        hits = len(norms & ctx.header_norms)
        if hits >= max(2, 0.6 * len(ctx.header_norms)):
            return RowClass(row.index, "header", "見出しの再出現")
    if len(nonempty) <= 2 and _NOTE_RE.match(first):
        return RowClass(row.index, "note", "注記")
    if len(nonempty) <= 2 and _BLOCK_TITLE_RE.match(first):
        return RowClass(row.index, "title", "表題")
    agg = _aggregate_label(row, width)
    if agg:
        return RowClass(row.index, "subtotal", "合計" if agg == "total" else "小計")
    if ctx.is_csv and width >= 4 and len(row.cells) < 0.5 * width and len(nonempty) <= 2:
        return RowClass(row.index, "excluded", "列数が見出しと合わない")
    if ctx.key_cols and all(not _filled(row, k) for k in ctx.key_cols):
        return RowClass(row.index, "continuation", "キー列が空で文章だけの行")
    return RowClass(row.index, "data")


def _filled(row: SourceRow, col: int) -> bool:
    cell = row.cell(col)
    if cell is None:
        return False
    return bool(cell.text) or (cell.merged_anchor is not None and cell.merged_anchor != (row.index, col + 1))


def _auto_key_columns(source, sheet, data_start: int, ctx: _Ctx) -> list[int]:
    """継続行の判定に使うキー列。値がほぼ埋まった短い列（日付・コードを優先）を最大3列。"""
    rows = []
    blanks = 0
    for row in source.rows(sheet, data_start, KEY_SAMPLE_ROWS * 3):
        rc = _classify(row, ctx)
        if rc.kind == "blank":
            blanks += 1
            if blanks >= BLANK_ROWS_END:
                break
            continue
        blanks = 0
        if rc.kind in ("title", "note", "header"):
            break
        if rc.kind == "data":
            rows.append(row)
            if len(rows) >= KEY_SAMPLE_ROWS:
                break
    if len(rows) < 3:
        return []
    candidates = []
    for col in range(ctx.width):
        filled = sum(1 for r in rows if _filled(r, col))
        if filled < 0.9 * len(rows):
            continue
        kinds = Counter(value_kind(r.cells[col].value, r.cells[col].text) for r in rows if r.cell(col) and r.cells[col].text)
        if kinds.get("text", 0) > 0.2 * len(rows):
            continue
        priority = 0 if (kinds.get("date", 0) + kinds.get("datetime", 0) + kinds.get("code", 0)) >= 0.5 * len(rows) else 1
        candidates.append((priority, col))
    return [col for _, col in sorted(candidates)[:3]]


def _scan(source, sheet, head: list[SourceRow], ctx: _Ctx, data_start: int, data_end: int | None,
          max_scan_rows: int | None):
    """先頭からデータの終わりまで分類する。戻り値: (row_classes, counts, data_end, warnings)"""
    classes: list[RowClass] = []
    warnings: list[str] = []
    counts: Counter[str] = Counter()
    extra = 0

    def record(rc: RowClass):
        nonlocal extra
        if rc.index <= PREVIEW_ROWS:
            classes.append(rc)
        elif rc.kind not in ("data", "blank") and extra < MAX_ROW_CLASSES:
            classes.append(rc)
            extra += 1

    for row in head:
        if row.index >= data_start:
            break
        if row.index in ctx.header_rows:
            kind = "header"
        elif row.is_blank:
            kind = "blank"
        elif _NOTE_RE.match(next(c.text for c in row.cells if c.text)):
            kind = "note"
        else:
            kind = "title"
        record(RowClass(row.index, kind, "表の上" if kind in ("title", "note") else ""))

    last_nonblank = data_start - 1
    blank_run = 0
    blank_total = 0
    end: int | None = None
    stop_reason = ""
    scanned = 0
    after_rows: list[SourceRow] = []
    for row in source.rows(sheet, data_start):
        if end is not None:
            after_rows.append(row)
            if len(after_rows) >= LOOKAHEAD_ROWS:
                break
            continue
        scanned += 1
        rc = _classify(row, ctx)
        if data_end is not None:
            if row.index > data_end:
                end = data_end
                after_rows.append(row)
                continue
            counts[rc.kind] += 1
            record(rc)
            continue
        if rc.kind == "blank":
            blank_run += 1
            record(rc)
            if blank_run >= BLANK_ROWS_END:
                end = last_nonblank
                stop_reason = "blank"
            continue
        if rc.kind in ("note", "title") or (rc.kind == "header" and rc.reason == "見出しの再出現"):
            end = last_nonblank
            after_rows.append(row)
            stop_reason = rc.kind
            continue
        blank_total += blank_run  # 途中の空行はデータ範囲に数える
        blank_run = 0
        counts[rc.kind] += 1
        record(rc)
        last_nonblank = row.index
        if rc.kind == "subtotal" and rc.reason == "合計":
            end = row.index
            stop_reason = "total"
            continue
        if max_scan_rows is not None and scanned >= max_scan_rows:
            end = row.index
            stop_reason = "limit"
            break
    if end is None:
        end = data_end if data_end is not None else last_nonblank

    if data_end is None:
        counts["blank"] = blank_total
    if not counts.get("blank"):
        counts.pop("blank", None)
    # 終わりより後の行（先頭60行の表示用と、別の表があるかの確認）
    other_table_row = None
    for row in after_rows:
        rc = _classify(row, ctx)
        if rc.kind == "blank":
            kind, reason = "blank", ""
        elif rc.kind in ("note", "title"):
            kind, reason = rc.kind, "表の下"
        else:
            kind, reason = "excluded", f"データの範囲外（{end}行目で終了）"
        if row.index <= PREVIEW_ROWS or (kind != "blank" and extra < 20):
            if row.index > PREVIEW_ROWS:
                extra += 1
            classes.append(RowClass(row.index, kind, reason))
        if other_table_row is None and (
            (rc.kind == "title" and row.index > end)
            or (rc.kind not in ("blank", "note") and sum(1 for c in row.cells if c.text) >= 3)
        ):
            other_table_row = row.index
    if stop_reason != "limit" and other_table_row is not None:
        warnings.append(
            f"データの終わり（{end}行目）より下にも表らしい行があります（{other_table_row}行目）。"
            "最初の表だけを読み取ります。残りは範囲を指定して1つずつ取り込んでください"
        )
    if stop_reason == "header":
        warnings.append(f"{end + 1}行目で見出しが再び出たので、そこで読み取りを止めました")
    if counts.get("excluded"):
        warnings.append(f"除外した行が{counts['excluded']}行あります（非表示・取り消し線など）")
    classes.sort(key=lambda rc: rc.index)
    return classes, dict(counts), end, warnings


# ---- 内部: その他 ----

def _no_table(sheet: str, head: list[SourceRow], message: str) -> LayoutGuess:
    classes = [RowClass(r.index, "blank" if r.is_blank else "title") for r in head[:PREVIEW_ROWS]]
    kind = "form_like" if _form_like(head) else "unknown"
    return LayoutGuess(sheet, kind, [], 1, 0, [], classes, 0.0, [message])


def _labels_down_first_column(head: list[SourceRow], rows_h: list[int], data_start: int, end: int) -> bool:
    """見出し行が2列以下で、下の行の先頭列に項目名（辞書の語）が並ぶなら帳票（ラベル: 値）。"""
    from tables.dictionary import lookup_header

    header = next((r for r in head if r.index == rows_h[-1]), None)
    if header is None or sum(1 for c in header.cells if c.text) > 3:
        return False
    firsts = [next(c for c in r.cells if c.text) for r in head if data_start <= r.index <= end and not r.is_blank]
    labels = sum(1 for c in firsts if _is_header_cell(c) and (lookup_header(c.text) or ("", ""))[1] == "dictionary")
    return len(firsts) >= 3 and labels >= max(3, 0.5 * len(firsts))


def _form_like(head: list[SourceRow]) -> bool:
    """「ラベル: 値」が並ぶ帳票らしい形か。"""
    pairs = 0
    for row in head:
        cells = [c for c in row.cells if c.text]
        if 2 <= len(cells) <= 6 and _is_header_cell(cells[0]) and len(cells[0].text) <= 12:
            pairs += 1
        elif len(cells) == 1 and re.search(r"[:：]\s*\S", cells[0].text) and len(cells[0].text) <= 40:
            pairs += 1
    return pairs >= 5


def _hidden_columns(source, sheet, width: int) -> list[str]:
    if getattr(source, "kind", "") != "excel":
        return []
    from openpyxl.utils import get_column_letter

    for info in source.sheets():
        if info.name == sheet:
            return [get_column_letter(c) for c in info.hidden_columns if c <= width]
    return []

