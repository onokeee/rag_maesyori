"""テンプレート定義に従ってExcelから項目を抽出する。

セル番地を固定せず、「ラベル候補に一致するセルを探す → その右 or 下にある値を取る」
という方式なので、帳票ごとにセル位置がズレていても同じJSONになる。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import asdict, dataclass, field, replace
from datetime import time

from excel.tables import (MIN_COLUMNS, SAME_COLUMNS_RATIO, Table, find_table, find_table_by_columns, is_table_value,
                          list_header_keys, merge_table_values, parse_table_text, section_heading, sections_of,
                          seq_header, stacked_tables, table_header_keys)
from excel.text import (EXCEL_ERROR_WARNING, MAX_LABEL_LENGTH, excel_error, numeric_unit, pick_checked, split_code_name, split_combined_value,
                        is_plain_number, to_date, to_number, value_unit, written_unit)
from excel.workbook import Cell, SheetGrid, WorkbookInfo
from pattern.dictionary import (COMBINED_EQUIPMENT_NORMS, COMBINED_EQUIPMENT_PARTS, DICTIONARY_NORMS,
                                ambiguous_unit_hint)
from pattern.model import FieldDef, PatternDef

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
    required: bool
    value: object = None
    sheet: str | None = None
    label_cell: str | None = None
    value_cell: str | None = None
    label_found: bool = False
    warning: str | None = None
    edited: bool = False
    ai_filled: bool = False
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
    """fields から values / missing_required を再計算する。"""
    fields = extraction["fields"]
    extraction["values"] = {f["field_name"]: f["value"] for f in fields}
    extraction["missing_required"] = [
        f["display_name"] for f in fields if f["required"] and is_blank_value(f["value"])
    ]


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
                f["value"], f["unit"], f["warning"], f["edited"], f["ai_filled"] = value, unit, warning, True, False
            continue
        else:
            value, warning = (text or None), None
        # 日付も同じ: 「2026-09-14 24:30」は日付だけ同じでも時刻が不正なので、警告を残して手で修正にする
        if value != f["value"] or (f["data_type"] == "date" and value is not None and (warning or f.get("warning"))):
            f["value"], f["warning"], f["edited"], f["ai_filled"] = value, warning, True, False
    refresh_summary(extraction)


def _extract_field(info: WorkbookInfo, fd: FieldDef, sheet_names: list[str], stop_labels: set[str]) -> FieldResult:
    result = FieldResult(fd.field_name, fd.display_name, fd.data_type, fd.required,
                         unit=fd.unit, spec_unit=fd.unit or "", rag_output=fd.rag_output)
    if fd.data_type == "table":
        result.table_columns = [c for c in (fd.table_columns or []) if str(c).strip()]
        return _extract_table_field(info, fd, sheet_names, stop_labels, result)
    for name in sheet_names:
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
    if not result.label_found:
        result.warning = "ラベルが見つかりません"
    elif _uncached_beside(info, *empty_label):
        result.warning = UNCACHED_FORMULA_WARNING
    else:
        result.warning = "ラベルはありますが値が空です"
    return result


UNCACHED_FORMULA_WARNING = "数式の計算結果が保存されていません。Excelで開いて保存し直すか、値を入力してください"


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
    for name in sheet_names:
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
    from excel.text import normalize_label

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

    候補どおりのラベル → 表記だけ違うラベル（「発生原因（なぜ起きたか）」「応急処置内容」）の順。
    """
    norms = grid.resolve_labels(fd.label_norms())
    headers = table_header_keys(grid)
    lists = list_header_keys(grid)
    label_sets = [norms, grid.label_variants(fd.label_norms()) - norms]
    if fd.field_name in COMBINED_EQUIPMENT_PARTS:
        # 設備番号・設備名: 最後に「対象設備：CMP-108　STI-CMP 8号機」のような番号と名前をまとめた欄も見る
        label_sets.append(COMBINED_EQUIPMENT_NORMS - norms - label_sets[1])
    found_sets: list[tuple[set[str], list[Cell]]] = []
    for n, label_set in enumerate(label_sets):
        # 行を足せる明細表の列見出し（「設備No｜設備名」の下に何行も並ぶ表）は、1つの値のラベルにしない。
        # 表記違いのラベルでは、区切りの見出し（「■ 承認欄」）も使わない
        # 明細表の連番の列見出し（No）も使わない（報告書の「No.」欄と取り違えない）
        cells = [c for c in grid.find_labels(label_set)
                 if (c.row, c.col) not in lists and not (n and section_heading(c))
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
        if fd.direction in ("auto", "right"):
            values = scan_right(grid, cell, stop_labels)
        if not values and fd.direction in ("auto", "below"):
            values = scan_below(grid, cell, stop_labels, multi=multi, allow_gap=allow_gap)
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
    if not allow_gap and cell.inline is not None and cell.inline[0] in norms and fd.direction in ("auto", "same_cell"):
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
