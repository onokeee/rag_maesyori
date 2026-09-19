"""帳票サンプル（F1〜F5）で、現行アプリの「見本から帳票の種類を作る → 抽出」の精度を測る。

    python -m scripts.samples.evaluate_forms                 # samples/forms を評価して結果を表示
    python -m scripts.samples.evaluate_forms --samples 3 --json samples/forms/_evaluation.json
    python -m scripts.samples.evaluate_forms --samples 2 --offset 5 --json out.json   # 別の見本で汎化を確認

アプリのコードは変更せず、画面操作と同じ順に関数を呼ぶ:
  1. 帳票フォルダごとに見本ファイルを3つ選ぶ（layout_version が混ざるように各版から順に取る）
  2. pattern.builder.suggest_rows で候補を作り、use=True の項目・シートをそのまま採用
     （キー名は辞書の field_name / field_N のまま）→ pattern.forms.rows_to_pattern で PatternDef にする
  3. 全ファイルについて pattern.matcher.match_pattern でシートを選び、excel.extractor.extract_document で抽出
  4. _expected.jsonl の values と、両方に存在する項目だけを比較する

項目の対応付け:
  - 辞書の項目（report_id, symptom など）は FIELD_MAP で正解データのキーに対応させる
  - 辞書にない項目（field_N）は、候補ラベルが正解データの labels_used のどのキーのラベルと一致するかで対応させる
    （テンプレート内の全ファイルで最も多く一致したキー。表の列見出し parts.qty などは対象外）

比較の正規化: NFKC（全角半角）・空白除去・英字小文字化。日付項目は日付部分（年月日）だけを比較し、
数値項目は数値として比較する。一致しなくても、片方がもう片方を含む場合は「部分一致」として別に数える。
「なし」「－」「／」などの記入は空欄と同じに扱う（正解が空で抽出が「なし」でも誤りにしない）。
押印欄の「佐々木⏎5/17」は1行目（氏名）が正解と同じなら一致とする（押印の日付は正解データに無い）。
低い一致率はサンプルの不具合ではなく、現行アプリの抽出方式の限界を示す情報として扱う。

明細表（正解が行のリスト）の扱い:
  - 1つの値の項目（文字列・文章など）とは比較しない（表を1つの文字列と比べても意味がないため）。
  - 代わりに「明細表」の集計を別に出す。対象は正解に行のリストがあるキーすべて（分母を減らさない）。
    明細表の項目（data_type="table"）が対応付いていれば行ごとに照合し、無ければ全行を「読めなかった行」に数える。
  - 行の一致: 正解の行の空でないセル値（*_norm などの派生キーと連番 No は除く）が、抽出した1行の中にすべてあること。
  - 明細表の項目と正解キーの対応は、全ファイルで一致した行が最も多いキー（見本ごとに見出しの書き方が違い、
    同じ表が複数の項目に分かれることがあるため、1つのキーに複数の項目を対応させ、ファイルごとに行を合わせて照合する）。
  - 1つの値の正解キーのうち、どの項目にも対応しないが、そのファイルの明細表の列見出しがラベル（labels_used）と
    同じものは、その列の値と比較する（合計行があれば合計行、1行だけならその行、複数行なら改行でつないだ値）。
    明細表に移った「ロットNo.」「投入数」などを比較から外さないため。項目名は「表の項目名.列見出し」で出す。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from excel.extractor import extract_document  # noqa: E402
from excel.text import normalize_label, to_date, to_number  # noqa: E402
from excel.workbook import load_workbook_info  # noqa: E402
from pattern.builder import suggest_rows  # noqa: E402
from pattern.forms import rows_to_pattern  # noqa: E402
from pattern.matcher import match_pattern  # noqa: E402

FORMS_ROOT = PROJECT_ROOT / "samples" / "forms"

# 辞書の field_name → 正解データ（_expected.jsonl の values）のキー候補。先頭から、その帳票に存在するものを使う
FIELD_MAP: dict[str, tuple[str, ...]] = {
    "report_id": ("report_id",),
    "subject": ("title", "subject"),
    "equipment_id": ("equipment_id",),
    "equipment_name": ("equipment_name",),
    "process": ("process",),
    "location": ("line", "location"),
    "department": ("department", "issuing_dept"),
    "occurred_date": ("occurred_at",),
    "completed_date": ("recovered_at", "completed_at"),
    "reporter": ("reporter",),
    "approver": ("approver",),
    "work_hours": ("work_hours",),
    "downtime": ("downtime_min",),
    "severity": ("severity",),
    "symptom": ("symptom",),
    "alarm": ("alarm",),
    "investigation": ("investigation",),
    "cause": ("cause",),
    "repair": ("action",),
    "result": ("result",),
    "prevention": ("prevention",),
    "parts": ("parts",),
    "remarks": ("remarks",),
}

_WS = re.compile(r"\s+")
_TRAIL = "。．.、,"
# 複数値の区切り（正解データはリスト、Excel上は「、」や改行で並べていることがある）
_SEP = re.compile(r"[、,;/\n]")


# ---------------------------------------------------------------------------
# 値の正規化と比較
# ---------------------------------------------------------------------------
def flatten(value) -> str:
    """正解データの値（文字列・リスト・表）を比較用の1つの文字列にする。"""
    if value is None:
        return ""
    if isinstance(value, list):
        return "\n".join(flatten(v) for v in value)
    if isinstance(value, dict):
        return " ".join(flatten(v) for v in value.values())
    return str(value)


def norm_text(value) -> str:
    s = unicodedata.normalize("NFKC", flatten(value)).replace("_x000D_", "")
    s = s.replace("〜", "~").replace("～", "~").replace("−", "-").replace("―", "-").replace("‐", "-")
    return _WS.sub("", s).strip(_TRAIL).lower()


def loose_text(value) -> str:
    """区切り記号（、 , / 改行）の違いも無視した正規化。複数値の並べ方の揺れを吸収する。"""
    s = unicodedata.normalize("NFKC", flatten(value)).replace("_x000D_", "")
    return _WS.sub("", _SEP.sub("", s)).strip(_TRAIL).lower()


def as_date(value) -> str | None:
    text = flatten(value)
    if not text:
        return None
    parsed, warning = to_date(text, text)
    if warning is None and parsed:
        return parsed[:10]  # 時刻付き（「2023-07-10 23:08」）も日付の一致で採点する
    # 「24/8/25」「6/19 0時43分」のような年が2桁・無しの表記は日付として扱わない（比較は文字列で行う）
    return None


def as_number(value, unit: str = ""):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = flatten(value)
    if not text:
        return None
    num, _ = to_number(text, text, unit)  # 「1時間32分」は項目の単位（分）に換算して比べる
    return float(num) if isinstance(num, (int, float)) else None


# 空欄と同じ意味の記入（「該当なし」「－」など）
EMPTY_MARKS = {"なし", "無し", "該当なし", "特になし", "-", "－", "ー", "―", "‐", "/", "／"}
_EMPTY_NORMS = {""}  # norm_text の定義後に埋める
# 押印欄の2行目（「5/17」「11/18」「2024/5/17」）
_STAMP_DATE = re.compile(r"^\s*(?:\d{2,4}[/.-])?\d{1,2}[/.-]\d{1,2}\s*$")


def is_empty_value(value) -> bool:
    return value is None or norm_text(value) in _EMPTY_NORMS


def compare(data_type: str, extracted, expected, unit: str = "") -> str:
    """戻り値: exact / normalized / date / number / partial / wrong / missing / empty_ok / spurious"""
    exp_empty = is_empty_value(expected)
    ext_empty = is_empty_value(extracted)
    if exp_empty:
        return "empty_ok" if ext_empty else "spurious"
    if ext_empty:
        return "missing"
    if flatten(extracted) == flatten(expected):
        return "exact"
    if norm_text(extracted) == norm_text(expected) or loose_text(extracted) == loose_text(expected):
        return "normalized"
    if isinstance(extracted, str) and "\n" in extracted and not isinstance(expected, (list, dict)):
        first, rest = extracted.split("\n", 1)
        if _STAMP_DATE.match(rest) and norm_text(first) == norm_text(expected):
            return "normalized"
    if data_type == "date" or as_date(expected):
        d_ext, d_exp = as_date(extracted), as_date(expected)
        if d_ext and d_exp:
            return "date" if d_ext == d_exp else "wrong"
    if data_type == "number":
        n_ext, n_exp = as_number(extracted, unit), as_number(expected, unit)
        if n_ext is not None and n_exp is not None and abs(n_ext - n_exp) < 1e-6:
            return "number"
    a, b = norm_text(extracted), norm_text(expected)
    if len(a) >= 2 and (a in b or b in a):
        return "partial"
    return "wrong"


MATCH = {"exact", "normalized", "date", "number"}
_EMPTY_NORMS |= {norm_text(m) for m in EMPTY_MARKS}


# ---------------------------------------------------------------------------
# 明細表（行のリスト）の照合
# ---------------------------------------------------------------------------
_ROW_SKIP_KEYS = {"no", "table"}


def is_table_value(value) -> bool:
    return isinstance(value, list) and any(isinstance(v, dict) for v in value)


def table_keys(entries: list[dict]) -> set[str]:
    return {k for e in entries for k, v in e["values"].items() if is_table_value(v)}


_DIGIT_COMMA = re.compile(r"(?<=\d),(?=\d{3})")


_DATE_ONLY = re.compile(r"^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}$")


def cell_norm(value) -> str:
    """表のセルの比較用の正規化（norm_text に加え、数値の桁区切り「2,450,000」を外し、日付だけのセルは年月日にそろえる）。"""
    text = _DIGIT_COMMA.sub("", norm_text(value))
    if _DATE_ONLY.match(text):
        text = as_date(text) or text
    return text


def expected_rows(value) -> list[list[str]]:
    """正解の行（セル値の正規化リスト）。文字列のリストは1列の表とみなす。"""
    if not isinstance(value, list):
        return [] if is_empty_value(value) else [[norm_text(value)]]
    rows = []
    for item in value:
        if isinstance(item, dict):
            cells = []
            for k, v in item.items():
                if k.lower() in _ROW_SKIP_KEYS or k.endswith(("_norm", "_as_written")):
                    continue
                v = item.get(f"{k}_as_written", v)  # 帳票に書かれたままの値（「〃」など）がある列はそちらで照合
                if not is_empty_value(v):
                    cells.append(cell_norm(v))
        else:
            cells = [] if is_empty_value(item) else [cell_norm(item)]
        if cells:
            rows.append(cells)
    return rows


def extracted_rows(value) -> list[list[str]]:
    if isinstance(value, dict) and isinstance(value.get("rows"), list):
        return [[cell_norm(c) for c in row if not is_empty_value(c)] for row in value["rows"]]
    return []


def _row_matches(exp: list[str], got: list[str]) -> bool:
    joined = "".join(got)
    return all((c in got) if len(c) < 3 else (c in joined) for c in exp)


def compare_table(extracted, expected) -> dict:
    exp_rows, got_rows = expected_rows(expected), extracted_rows(extracted)
    used: set[int] = set()
    matched = 0
    for row in exp_rows:
        hit = next((i for i, got in enumerate(got_rows) if i not in used and _row_matches(row, got)), None)
        if hit is not None:
            used.add(hit)
            matched += 1
    return {"files": 1, "expected_rows": len(exp_rows), "extracted_rows": len(got_rows), "matched_rows": matched,
            "exact_tables": int(matched == len(exp_rows) == len(got_rows))}


# ---------------------------------------------------------------------------
# 評価
# ---------------------------------------------------------------------------
def pick_samples(entries: list[dict], n: int, offset: int = 0) -> list[dict]:
    """layout_version が混ざるように、版ごとのファイルから順に n 個取る（決定的）。

    offset: 版ごとに何番目のファイルから取り始めるか（見本の選び方を変えて、特定のファイルへの過適合を確かめる）。
    版の並び順も offset だけ回すので、n が版の数より少ないときは使う版も変わる。
    """
    by_version: dict[str, list[dict]] = defaultdict(list)
    for e in entries:
        by_version[e["layout_version"]].append(e)
    versions = sorted(by_version)
    if versions:
        k = offset % len(versions)
        versions = versions[k:] + versions[:k]
    for v in versions:
        k = offset % len(by_version[v])
        by_version[v] = by_version[v][k:] + by_version[v][:k]
    picked, i = [], 0
    while len(picked) < min(n, len(entries)):
        for v in versions:
            if i < len(by_version[v]) and len(picked) < n:
                picked.append(by_version[v][i])
        i += 1
    return picked


def map_fields(pattern, entries: list[dict]) -> dict[str, tuple[str, str]]:
    """PatternDef の field_name → (正解キー, 対応付けの根拠)。1つの値の項目だけ（明細表は map_table_fields）。

    行のリストが正解のキーには対応させない。
    """
    all_keys = set().union(*(e["values"].keys() for e in entries))
    tables = table_keys(entries)

    def allowed(fd, key: str) -> bool:
        return (key in tables) == (fd.data_type == "table")

    mapping: dict[str, tuple[str, str]] = {}
    for fd in pattern.fields:
        if fd.data_type == "table":
            continue
        if fd.field_name in FIELD_MAP:
            key = next((k for k in FIELD_MAP[fd.field_name] if k in all_keys and allowed(fd, k)), None)
            if key:
                mapping[fd.field_name] = (key, "dictionary")
    taken = {k for k, _ in mapping.values()}
    for fd in pattern.fields:
        if fd.field_name in mapping or fd.field_name in FIELD_MAP or fd.data_type == "table":
            continue
        norms = fd.label_norms()
        hits: Counter = Counter()
        order: dict[str, int] = {}   # 同数のときは帳票の上の方（labels_used の先頭側）に出てくるキーを優先
        for e in entries:
            for pos, (key, label) in enumerate(e.get("labels_used", {}).items()):
                if "." in key or key not in e["values"] or key in taken or not allowed(fd, key):
                    continue
                if normalize_label(label) in norms:
                    hits[key] += 1
                    order[key] = min(order.get(key, pos), pos)
        if hits:
            key = sorted(hits.items(), key=lambda kv: (-kv[1], order[kv[0]], kv[0]))[0][0]
            mapping[fd.field_name] = (key, "label")
            taken.add(key)
    return mapping


def map_table_fields(pattern, entries: list[dict], extractions: dict[str, dict]) -> dict[str, list[str]]:
    """正解キー（行のリスト）→ 明細表の項目名のリスト。各項目を、一致した行が最も多いキーに対応させる。"""
    keys = sorted(table_keys(entries))
    mapping: dict[str, list[str]] = defaultdict(list)
    for fd in pattern.fields:
        if fd.data_type != "table":
            continue
        hits: Counter = Counter()
        for e in entries:
            value = extractions[e["file"]][fd.field_name]["value"]
            if value is None:
                continue
            for key in keys:
                hits[key] += compare_table(value, e["values"].get(key))["matched_rows"]
        best = sorted(hits.items(), key=lambda kv: (-kv[1], kv[0]))
        if best and best[0][1] > 0:
            mapping[best[0][0]].append(fd.field_name)
    return dict(mapping)


def _joined_table(values: list) -> dict | None:
    rows = [row for v in values if isinstance(v, dict) for row in v.get("rows", [])]
    return {"columns": [], "rows": rows} if rows else None


_TOTAL_CELL = re.compile(r"^(?:合計|小計|総計|計|.{1,6}計)$")


def column_value(value, label: str):
    """明細表の値から、列見出しが label の列の値（合計行 → 1行 → 改行でつないだ値）。列が無ければ False。"""
    if not isinstance(value, dict):
        return False
    norms = {normalize_label(c) for c in [label]}
    columns = [normalize_label(c) for c in value.get("columns", [])]
    if not any(c in norms for c in columns):
        return False
    i = next(i for i, c in enumerate(columns) if c in norms)
    rows = [r for r in value.get("rows", []) if i < len(r)]
    totals = [r for r in rows if any(_TOTAL_CELL.match(norm_text(c)) for c in r if c)]
    data = [r for r in rows if r not in totals]
    if totals and totals[0][i] and not _TOTAL_CELL.match(norm_text(totals[0][i])):
        return totals[0][i]
    cells = [r[i] for r in data if r[i]]
    return "\n".join(cells) if cells else None


def evaluate_template(folder: Path, n_samples: int, offset: int = 0) -> dict:
    entries = [json.loads(line) for line in (folder / "_expected.jsonl").open(encoding="utf-8") if line.strip()]
    samples = pick_samples(entries, n_samples, offset)
    infos = {e["file"]: load_workbook_info(folder / e["file"]) for e in entries}

    sheet_rows, field_rows = suggest_rows([infos[e["file"]] for e in samples])
    meta = {"name": folder.name, "version": "eval"}
    pattern = rows_to_pattern(1, meta, sheet_rows, field_rows)
    mapping = map_fields(pattern, entries)
    unused_dictionary = [r["field_name"] for r in field_rows if not r["use"] and r["field_name"] in FIELD_MAP]

    per_field: dict[str, Counter] = defaultdict(Counter)
    per_version: dict[str, Counter] = defaultdict(Counter)
    per_reason: Counter = Counter()
    tables: dict[str, Counter] = {k: Counter() for k in sorted(table_keys(entries))}
    matches, extractions = {}, {}
    for e in entries:
        info = infos[e["file"]]
        matches[e["file"]] = match_pattern(info, pattern)
        extraction = extract_document(info, pattern, matches[e["file"]].sheet_names)
        extractions[e["file"]] = {f["field_name"]: f for f in extraction["fields"]}
    table_fields = map_table_fields(pattern, entries, extractions)
    mapped_keys = {k for k, _ in mapping.values()} | set(tables)
    table_names = [fd.field_name for fd in pattern.fields if fd.data_type == "table"]
    column_fields: dict[str, str] = {}
    failures: list[dict] = []
    files = []
    for e in entries:
        match = matches[e["file"]]
        by_name = extractions[e["file"]]
        file_counts: Counter = Counter()
        for key, counts in tables.items():
            value = _joined_table([by_name[n]["value"] for n in table_fields.get(key, [])])
            counts.update(compare_table(value, e["values"].get(key)))
        pairs = [(name, key, by_name[name]) for name, (key, _src) in mapping.items()]
        # どの項目にも対応しない1つの値のキーは、明細表の列（ラベルと同じ列見出し）と比較する
        for key, label in e.get("labels_used", {}).items():
            if "." in key or key in mapped_keys or key not in e["values"] or is_table_value(e["values"][key]):
                continue
            for name in table_names:
                got = column_value(by_name[name]["value"], label)
                if got is not False:
                    pseudo = f"{name}.{label}"
                    column_fields[pseudo] = key
                    f = {**by_name[name], "value": got, "data_type": "string"}
                    pairs.append((pseudo, key, f))
                    break
        for field_name, key, f in pairs:
            outcome = compare(f["data_type"], f["value"], e["values"].get(key), f.get("unit") or "")
            per_field[field_name][outcome] += 1
            per_version[e["layout_version"]][outcome] += 1
            file_counts[outcome] += 1
            if outcome in ("wrong", "missing", "partial", "spurious"):
                reason = failure_reason(outcome, f)
                per_reason[reason] += 1
                failures.append({
                    "file": e["file"], "layout_version": e["layout_version"], "field": field_name, "key": key,
                    "outcome": outcome, "reason": reason, "expected": flatten(e["values"].get(key))[:80],
                    "extracted": flatten(f["value"])[:80], "cell": f.get("value_cell") or f.get("label_cell"),
                    "warning": f.get("warning"),
                })
        files.append({"file": e["file"], "layout_version": e["layout_version"], "confidence": match.confidence,
                      "sheets": match.sheet_names, "counts": dict(file_counts)})

    total = sum(per_field.values(), Counter())
    return {
        "template": folder.name,
        "samples": [s["file"] for s in samples],
        "sample_versions": [s["layout_version"] for s in samples],
        "pattern_sheets": [s.sheet_name for s in pattern.sheets],
        "pattern_fields": [{"field_name": f.field_name, "display_name": f.display_name, "data_type": f.data_type,
                            "mapped_to": mapping.get(f.field_name, (None, None))[0]
                            or next((k for k, names in table_fields.items() if f.field_name in names), None),
                            "mapping": mapping.get(f.field_name, (None, None))[1]
                            or ("table_rows" if any(f.field_name in names for names in table_fields.values()) else None)}
                           for f in pattern.fields]
        + [{"field_name": name, "display_name": name, "data_type": "string", "mapped_to": key,
            "mapping": "table_column"} for name, key in column_fields.items()],
        "unused_dictionary_suggestions": unused_dictionary,
        "files": files,
        "totals": dict(total),
        "per_field": {k: dict(v) for k, v in per_field.items()},
        "per_version": {k: dict(v) for k, v in per_version.items()},
        "per_reason": dict(per_reason),
        "tables": {k: {"field": ",".join(table_fields.get(k, [])), **dict(v)} for k, v in tables.items()},
        "failures": failures,
    }


def failure_reason(outcome: str, f: dict) -> str:
    """不一致の大まかな分類（どこで失敗したか）。"""
    if outcome == "spurious":
        return "正解は空なのに値を読んだ"
    if not f.get("label_found"):
        return "見出しが見つからない"
    if outcome == "missing":
        return "見出しはあるが値が空"
    if outcome == "partial":
        return "部分一致（範囲の過不足）"
    return "別の値を読んだ"


def table_rate(results: list[dict]) -> tuple[int, int, int]:
    """(一致した行, 正解の行, 抽出した行)"""
    m = e = x = 0
    for r in results:
        for c in r.get("tables", {}).values():
            m += c.get("matched_rows", 0)
            e += c.get("expected_rows", 0)
            x += c.get("extracted_rows", 0)
    return m, e, x


def rate(counts: dict, include_empty: bool = False) -> tuple[int, int, float]:
    """(一致数, 比較数, 一致率)。既定では正解が空の項目（empty_ok / spurious）を分母に入れない。"""
    keys = ("exact", "normalized", "date", "number", "partial", "wrong", "missing")
    if include_empty:
        keys += ("empty_ok", "spurious")
    n = sum(counts.get(k, 0) for k in keys)
    ok = sum(counts.get(k, 0) for k in MATCH) + (counts.get("empty_ok", 0) if include_empty else 0)
    return ok, n, (ok / n if n else 0.0)


def print_report(results: list[dict]) -> None:
    for r in results:
        ok, n, pct = rate(r["totals"])
        partial = r["totals"].get("partial", 0)
        print(f"\n=== {r['template']}  一致 {ok}/{n} = {pct:.1%}（部分一致 {partial}）")
        print(f"  見本: {', '.join(f'{f}[{v}]' for f, v in zip(r['samples'], r['sample_versions']))}")
        print(f"  採用シート: {r['pattern_sheets']}  項目数: {len(r['pattern_fields'])}"
              f"（比較対象 {sum(1 for f in r['pattern_fields'] if f['mapped_to'])}）")
        if r["unused_dictionary_suggestions"]:
            print(f"  候補にあるが use=False の辞書項目: {r['unused_dictionary_suggestions']}")
        print("  版ごと: " + "  ".join(f"{v} {rate(c)[2]:.0%}" for v, c in sorted(r["per_version"].items())))
        if r.get("per_reason"):
            print("  不一致の分類: " + "  ".join(f"{k} {v}" for k, v in sorted(r["per_reason"].items(), key=lambda kv: -kv[1])))
        for key, c in r.get("tables", {}).items():
            print(f"  明細表 {key:<26} 項目={c.get('field') or '（なし）':<12} 行 {c.get('matched_rows', 0)}/{c.get('expected_rows', 0)}"
                  f"（抽出 {c.get('extracted_rows', 0)}行, 全行一致 {c.get('exact_tables', 0)}/{c.get('files', 0)}ファイル）")
        for field_name, counts in sorted(r["per_field"].items(), key=lambda kv: rate(kv[1])[2]):
            fdef = next(f for f in r["pattern_fields"] if f["field_name"] == field_name)
            ok, n, pct = rate(counts)
            detail = ", ".join(f"{k}={v}" for k, v in sorted(counts.items()))
            print(f"    {field_name:<16}→ {fdef['mapped_to']:<22} {ok:>2}/{n:<2} {pct:>5.0%}  {detail}")


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass
    parser = argparse.ArgumentParser(description="帳票サンプルで現行アプリの抽出精度を評価する")
    parser.add_argument("--forms", default=str(FORMS_ROOT), help="帳票フォルダのルート（既定: samples/forms）")
    parser.add_argument("--samples", type=int, default=3, help="見本にするファイル数（既定: 3）")
    parser.add_argument("--offset", type=int, default=0,
                        help="見本の選び方をずらす（版ごとに何番目から取るか。既定: 0 = 先頭から）")
    parser.add_argument("--only", help="対象フォルダ名の先頭（例: F1,F3）")
    parser.add_argument("--json", help="結果を書き出す JSON のパス（既定: <forms>/_evaluation.json）")
    args = parser.parse_args(argv)

    root = Path(args.forms)
    folders = sorted(d for d in root.iterdir() if (d / "_expected.jsonl").exists())
    if args.only:
        prefixes = [p.strip() for p in args.only.split(",") if p.strip()]
        folders = [d for d in folders if any(d.name.startswith(p) for p in prefixes)]
    t0 = time.perf_counter()
    results = [evaluate_template(d, args.samples, args.offset) for d in folders]
    print_report(results)
    ok = sum(rate(r["totals"])[0] for r in results)
    n = sum(rate(r["totals"])[1] for r in results)
    spurious = sum(r["totals"].get("spurious", 0) for r in results)
    m, e, x = table_rate(results)
    print(f"\n全体: 1つの値の項目 {ok}/{n} = {ok / n if n else 0:.1%}（正解が空なのに読んだ {spurious}）"
          f" ／ 明細表の行 {m}/{e} = {m / e if e else 0:.1%}（抽出 {x}行）")
    out = Path(args.json) if args.json else root / "_evaluation.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    print(f"\n{len(results)} 帳票フォルダ, {time.perf_counter() - t0:.1f}s, 詳細: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
