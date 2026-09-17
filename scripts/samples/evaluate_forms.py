"""帳票サンプル（F1〜F5）で、現行アプリの「見本から帳票の種類を作る → 抽出」の精度を測る。

    python -m scripts.samples.evaluate_forms                 # samples/forms を評価して結果を表示
    python -m scripts.samples.evaluate_forms --samples 3 --json samples/forms/_evaluation.json

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
低い一致率はサンプルの不具合ではなく、現行アプリの抽出方式の限界を示す情報として扱う。
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
        return parsed
    # 「24/8/25」「6/19 0時43分」のような年が2桁・無しの表記は日付として扱わない（比較は文字列で行う）
    return None


def as_number(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = flatten(value)
    if not text:
        return None
    num, _ = to_number(text, text)
    return float(num) if isinstance(num, (int, float)) else None


def compare(data_type: str, extracted, expected) -> str:
    """戻り値: exact / normalized / date / number / partial / wrong / missing / empty_ok / spurious"""
    exp_empty = norm_text(expected) == ""
    ext_empty = extracted is None or norm_text(extracted) == ""
    if exp_empty:
        return "empty_ok" if ext_empty else "spurious"
    if ext_empty:
        return "missing"
    if flatten(extracted) == flatten(expected):
        return "exact"
    if norm_text(extracted) == norm_text(expected) or loose_text(extracted) == loose_text(expected):
        return "normalized"
    if data_type == "date" or as_date(expected):
        d_ext, d_exp = as_date(extracted), as_date(expected)
        if d_ext and d_exp:
            return "date" if d_ext == d_exp else "wrong"
    if data_type == "number":
        n_ext, n_exp = as_number(extracted), as_number(expected)
        if n_ext is not None and n_exp is not None and abs(n_ext - n_exp) < 1e-6:
            return "number"
    a, b = norm_text(extracted), norm_text(expected)
    if len(a) >= 2 and (a in b or b in a):
        return "partial"
    return "wrong"


MATCH = {"exact", "normalized", "date", "number"}


# ---------------------------------------------------------------------------
# 評価
# ---------------------------------------------------------------------------
def pick_samples(entries: list[dict], n: int) -> list[dict]:
    """layout_version が混ざるように、版ごとの先頭ファイルから順に n 個取る（決定的）。"""
    by_version: dict[str, list[dict]] = defaultdict(list)
    for e in entries:
        by_version[e["layout_version"]].append(e)
    versions = sorted(by_version)
    picked, i = [], 0
    while len(picked) < min(n, len(entries)):
        for v in versions:
            if i < len(by_version[v]) and len(picked) < n:
                picked.append(by_version[v][i])
        i += 1
    return picked


def map_fields(pattern, entries: list[dict]) -> dict[str, tuple[str, str]]:
    """PatternDef の field_name → (正解キー, 対応付けの根拠)。"""
    all_keys = set().union(*(e["values"].keys() for e in entries))
    mapping: dict[str, tuple[str, str]] = {}
    for fd in pattern.fields:
        if fd.field_name in FIELD_MAP:
            key = next((k for k in FIELD_MAP[fd.field_name] if k in all_keys), None)
            if key:
                mapping[fd.field_name] = (key, "dictionary")
    taken = {k for k, _ in mapping.values()}
    for fd in pattern.fields:
        if fd.field_name in mapping or fd.field_name in FIELD_MAP:
            continue
        norms = fd.label_norms()
        hits: Counter = Counter()
        order: dict[str, int] = {}   # 同数のときは帳票の上の方（labels_used の先頭側）に出てくるキーを優先
        for e in entries:
            for pos, (key, label) in enumerate(e.get("labels_used", {}).items()):
                if "." in key or key not in e["values"] or key in taken:
                    continue
                if normalize_label(label) in norms:
                    hits[key] += 1
                    order[key] = min(order.get(key, pos), pos)
        if hits:
            key = sorted(hits.items(), key=lambda kv: (-kv[1], order[kv[0]], kv[0]))[0][0]
            mapping[fd.field_name] = (key, "label")
            taken.add(key)
    return mapping


def evaluate_template(folder: Path, n_samples: int) -> dict:
    entries = [json.loads(line) for line in (folder / "_expected.jsonl").open(encoding="utf-8") if line.strip()]
    samples = pick_samples(entries, n_samples)
    infos = {e["file"]: load_workbook_info(folder / e["file"]) for e in entries}

    sheet_rows, field_rows = suggest_rows([infos[e["file"]] for e in samples])
    meta = {"name": folder.name, "version": "eval"}
    pattern = rows_to_pattern(1, meta, sheet_rows, field_rows)
    mapping = map_fields(pattern, entries)
    unused_dictionary = [r["field_name"] for r in field_rows if not r["use"] and r["field_name"] in FIELD_MAP]

    per_field: dict[str, Counter] = defaultdict(Counter)
    per_version: dict[str, Counter] = defaultdict(Counter)
    failures: list[dict] = []
    files = []
    for e in entries:
        info = infos[e["file"]]
        match = match_pattern(info, pattern)
        extraction = extract_document(info, pattern, match.sheet_names)
        by_name = {f["field_name"]: f for f in extraction["fields"]}
        file_counts: Counter = Counter()
        for field_name, (key, _src) in mapping.items():
            f = by_name[field_name]
            outcome = compare(f["data_type"], f["value"], e["values"].get(key))
            per_field[field_name][outcome] += 1
            per_version[e["layout_version"]][outcome] += 1
            file_counts[outcome] += 1
            if outcome in ("wrong", "missing", "partial", "spurious"):
                failures.append({
                    "file": e["file"], "layout_version": e["layout_version"], "field": field_name, "key": key,
                    "outcome": outcome, "expected": flatten(e["values"].get(key))[:80],
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
                            "mapped_to": mapping.get(f.field_name, (None, None))[0],
                            "mapping": mapping.get(f.field_name, (None, None))[1]} for f in pattern.fields],
        "unused_dictionary_suggestions": unused_dictionary,
        "files": files,
        "totals": dict(total),
        "per_field": {k: dict(v) for k, v in per_field.items()},
        "per_version": {k: dict(v) for k, v in per_version.items()},
        "failures": failures,
    }


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
    parser.add_argument("--only", help="対象フォルダ名の先頭（例: F1,F3）")
    parser.add_argument("--json", help="結果を書き出す JSON のパス（既定: <forms>/_evaluation.json）")
    args = parser.parse_args(argv)

    root = Path(args.forms)
    folders = sorted(d for d in root.iterdir() if (d / "_expected.jsonl").exists())
    if args.only:
        prefixes = [p.strip() for p in args.only.split(",") if p.strip()]
        folders = [d for d in folders if any(d.name.startswith(p) for p in prefixes)]
    t0 = time.perf_counter()
    results = [evaluate_template(d, args.samples) for d in folders]
    print_report(results)
    out = Path(args.json) if args.json else root / "_evaluation.json"
    out.write_text(json.dumps(results, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    print(f"\n{len(results)} 帳票フォルダ, {time.perf_counter() - t0:.1f}s, 詳細: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
