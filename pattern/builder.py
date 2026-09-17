"""見本ファイル（1〜数ファイル）から帳票の種類の候補を作る。

ここで作るのはあくまで「候補」で、人が画面で確認・修正してから使用開始する。
複数サンプルを渡すと、セル位置が違っても同じラベルが全サンプルにあるかで確度を上げる。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

from excel.extractor import scan_below, scan_right
from excel.text import MAX_LABEL_LENGTH, cell_text, normalize_label, normalize_sheet_name, split_label_unit, value_unit
from excel.workbook import Cell, SheetGrid, WorkbookInfo
from pattern.dictionary import DICTIONARY_NORMS, LOOKUP
from pattern.forms import pattern_to_rows
from pattern.model import DEFAULT_TITLE_KEYS, FieldDef, PatternDef

_VALUE_LIKE = re.compile(r"^[\x20-\x7e]+$")  # 英数字記号のみ（EQ-001, CMP, 2026/09/14 など）


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
    score: int = 0
    unit: str = ""


def suggest_rows(infos: list[WorkbookInfo]) -> tuple[list[dict], list[dict]]:
    """画面表示用の (シート行, 項目行) を返す。"""
    sheet_rows = _suggest_sheets(infos)
    selected = {normalize_sheet_name(r["sheet_name"]) for r in sheet_rows if r["use"]}
    n = len(infos)

    suggestions: dict[str, _Suggestion] = {}
    for index, info in enumerate(infos):
        for grid in info.grids.values():
            for key, sug in _label_candidates(grid).items():
                merged = suggestions.setdefault(key, sug)
                if merged is not sug:
                    merged.labels.extend(l for l in sug.labels if l not in merged.labels)
                    merged.examples.extend(sug.examples)
                    merged.score = max(merged.score, sug.score)
                    merged.unit = merged.unit or sug.unit
                merged.samples.add(index)
                merged.sheets.add(normalize_sheet_name(grid.name))

    rows, used_names = [], set()
    for sug in suggestions.values():
        in_selected_sheet = bool(sug.sheets & selected)
        in_dictionary = bool(sug.synonyms)
        if not in_selected_sheet and not in_dictionary:
            continue
        use = in_selected_sheet and (in_dictionary or (len(sug.samples) == n and sug.score >= 1))
        field_name = sug.field_name
        if not field_name or field_name in used_names:
            field_name = _next_field_name(used_names)
        used_names.add(field_name)
        rows.append({
            "use": use,
            "field_name": field_name,
            "display_name": sug.display_name,
            "candidates": "\n".join(dict.fromkeys([*sug.labels, *sug.synonyms])),
            "data_type": sug.data_type,
            "required": False,
            "direction": "auto",
            "unit": sug.unit,
            "rag_output": "show",
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
    found: list[tuple[Cell, _Suggestion, list[Cell]]] = []
    for cell in grid.text_cells():
        if cell.inline:
            label_text = re.split(r"[:：]", cell.text, maxsplit=1)[0].strip()
            sug = _make_suggestion(cell.inline[0], label_text, cell.inline[1], cell)
            found.append((cell, sug, []))
            continue
        if "\n" in cell.text or len(cell.norm) > MAX_LABEL_LENGTH or not isinstance(cell.value, str):
            continue
        entry = LOOKUP.get(cell.norm)
        if entry is None and _VALUE_LIKE.match(cell.norm):
            continue
        values = scan_right(grid, cell, stop_labels) or scan_below(grid, cell, stop_labels, multi=False)
        if not values:
            continue
        found.append((cell, _make_suggestion(cell.norm, cell.text.rstrip(":：").strip(), values[0].value, cell), values))

    # 辞書ラベルの「値」として使われたセルは、ラベル候補から外す
    value_cells = {(v.row, v.col) for _, sug, values in found if sug.synonyms for v in values}
    result: dict[str, _Suggestion] = {}
    for cell, sug, _ in found:
        if not sug.synonyms and (cell.row, cell.col) in value_cells:
            continue
        key = sug.field_name or f"label:{cell.inline[0] if cell.inline else cell.norm}"
        result.setdefault(key, sug)
    return result


def _make_suggestion(norm: str, label_text: str, value, cell: Cell) -> _Suggestion:
    # 「作業時間(h)」のような見出しは単位を分けて辞書を引く（候補ラベルは元の表記のまま残す）
    base, unit = split_label_unit(label_text)
    entry = LOOKUP.get(norm) or (LOOKUP.get(normalize_label(base)) if unit else None)
    display = base if unit else label_text
    example = cell_text(value)
    if entry:
        sug = _Suggestion(entry.field_name, display, entry.data_type, [label_text], entry.synonyms,
                          examples=[example], score=3, unit=unit)
    else:
        score = int(cell.bold) + int(cell.filled) + int(cell.text.rstrip().endswith((":", "：")))
        sug = _Suggestion(None, display, _guess_type(value), [label_text], examples=[example], score=score, unit=unit)
    if not sug.unit and sug.data_type == "number":
        # 値が「1.5時間」のように単位付きで書かれていれば、その単位を候補にする
        sug.unit = value_unit(example)
    return sug


def _guess_type(value) -> str:
    if isinstance(value, date):
        return "date"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return "number"
    text = str(value)
    return "text" if "\n" in text or len(text) > 40 else "string"


def _next_field_name(used: set[str]) -> str:
    i = 1
    while f"field_{i}" in used:
        i += 1
    return f"field_{i}"
