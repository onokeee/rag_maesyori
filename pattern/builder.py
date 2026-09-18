"""見本ファイル（1〜数ファイル）から帳票の種類の候補を作る。

ここで作るのはあくまで「候補」で、人が画面で確認・修正してから使用開始する。
複数サンプルを渡すと、セル位置が違っても同じラベルが全サンプルにあるかで確度を上げる。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

from excel.extractor import scan_below, scan_right
from excel.tables import (SAME_COLUMNS_RATIO, Table, detect_tables, legend_text, seq_header, table_header_keys,
                          table_text_lines, table_title)
from excel.text import (MAX_LABEL_LENGTH, cell_text, normalize_label, normalize_sheet_name, split_code_name,
                        split_label_unit, to_date, value_unit)
from excel.workbook import Cell, SheetGrid, WorkbookInfo
from pattern.dictionary import (BY_FIELD_NAME, COMBINED_EQUIPMENT_NORMS, COMBINED_EQUIPMENT_PARTS,
                                DICTIONARY_NORMS, LOOKUP)
from pattern.forms import pattern_to_rows
from pattern.model import DEFAULT_TITLE_KEYS, FieldDef, PatternDef

_VALUE_LIKE = re.compile(r"^[\x20-\x7e]+$")  # 英数字記号のみ（EQ-001, CMP, 2026/09/14 など）
# 注記・凡例の書き出し（「※判定　○」「※故障発生日時…は事後保全時に記入」）
_NOTE_START = re.compile(r"^\s*[※＊*]")
# 「－」だけの値（記入なしの印）
_DASH_ONLY = {"-", "－", "ー", "―", "‐", "/", "／"}
# 「8D No.」「QA番号」のような短い接頭辞付きの識別番号の見出し（辞書の「報告No」と同じ扱い）
_ID_LABEL_RE = re.compile(r"^[a-z0-9]{1,3}(?:no|番号)$")
# 日付らしい見出しの末尾（見本の値が日付として読めるときだけ日付型にする）
_DATE_LABEL_SUFFIXES = ("日", "日付", "日時", "期限", "予定日", "年月日")
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
            "required": False,
            "direction": "auto",
            "unit": sug.unit,
            "rag_output": "show",
            "table_columns": "\n".join(sug.column_labels),
            "examples": list(dict.fromkeys(sug.examples))[:3],
            "seen": f"{len(sug.samples)}/{n}",
        })
    rows.sort(key=lambda r: not r["use"])
    return sheet_rows, rows


def suggest_title_fields(field_rows: list[dict]) -> list[str]:
    """タイトル項目の候補。使う項目のうち 報告番号→設備番号→設備名→発生日 の順で並べる。"""
    used = {r["field_name"] for r in field_rows if r.get("use")}
    return [key for key in DEFAULT_TITLE_KEYS if key in used]


def merge_with_existing(pattern: PatternDef, sheet_rows: list[dict], field_rows: list[dict]) -> tuple[list[dict], list[dict]]:
    """登録済みテンプレートに、サンプルから見つかった新しいラベルを追加候補として足す。"""
    existing_sheets, existing_fields = pattern_to_rows(pattern)
    known_sheets = {normalize_sheet_name(r["sheet_name"]) for r in existing_sheets}
    for row in sheet_rows:
        if normalize_sheet_name(row["sheet_name"]) not in known_sheets:
            existing_sheets.append({**row, "use": False})

    norms_by_row = [FieldDef("", "", r["candidates"].splitlines()).label_norms() for r in existing_fields]
    for row in field_rows:
        norms = FieldDef("", "", row["candidates"].splitlines()).label_norms()
        target = next(
            (er for er, en in zip(existing_fields, norms_by_row) if er["field_name"] == row["field_name"] or en & norms),
            None,
        )
        if target:
            labels = target["candidates"].splitlines() + row["candidates"].splitlines()
            target["candidates"] = "\n".join(dict.fromkeys(l for l in labels if l.strip()))
            target["examples"], target["seen"] = row["examples"], row["seen"]
            target["unit"] = target.get("unit") or row.get("unit", "")
            if target.get("data_type") == "table" and row.get("table_columns"):
                columns = (target.get("table_columns") or "").splitlines() + row["table_columns"].splitlines()
                target["table_columns"] = "\n".join(dict.fromkeys(c for c in columns if c.strip()))
        else:
            used = {r["field_name"] for r in existing_fields}
            if row["field_name"] in used:
                row["field_name"] = _next_field_name(used)
            existing_fields.append({**row, "use": False})
    return existing_sheets, existing_fields


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

    # 辞書ラベルの「値」として使われたセルは、ラベル候補から外す
    value_cells = {(v.row, v.col) for _, sug, values in found if sug.synonyms for v in values}
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
    return "text" if "\n" in text or len(text) > 40 else "string"


def _date_label(label: str) -> bool:
    """日付らしい見出しか（「作業日」「回答日」「次回点検予定日」「発生日時」）。「曜日」は値が日付にならないので残らない。"""
    return normalize_label(label).endswith(_DATE_LABEL_SUFFIXES)


def _next_field_name(used: set[str]) -> str:
    i = 1
    while f"field_{i}" in used:
        i += 1
    return f"field_{i}"
