"""一覧表パイプライン（spec / normalize / checks / summaries / markdown / outputs / pipeline / store）のテスト。"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import random
import zipfile
from collections import Counter
from datetime import datetime, time as dtime, timedelta
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Font

from tables import outputs, pipeline, store
from tables.checks import Issue, has_blocking, run_checks
from tables.detect import guess_layout, sample_data_rows
from tables.mapping import suggest_columns
from tables.markdown import render_all
from tables.normalize import ConvertContext, convert_cell, find_year_context, read_records
from tables.source import open_source
from tables.spec import (
    ColumnSpec, resolve_columns, spec_from_dict, spec_from_suggestions, spec_hash, spec_to_dict, validate_spec,
)

SAMPLES = Path(__file__).resolve().parent.parent / "samples" / "tables"
HINT = ".[legacy-R(chunk_ts=800,chunk_ol=0)].md"


# ---- 合成データ ---------------------------------------------------------------------------

LIST_HEADERS = ["管理No", "発生日", "設備番号", "設備名", "故障区分", "現象", "処置", "停止時間(h)", "費用(円)", "担当者", "対応内容"]


def make_list_book(path: Path, include_tr004: bool = True, tr001_hours=1.5) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "2026年8月"
    ws["A1"] = "トラブル対応一覧 2026年度"
    ws.merge_cells("A1:K1")
    for i, h in enumerate(LIST_HEADERS, start=1):
        ws.cell(row=3, column=i, value=h).font = Font(bold=True)
    rows = [
        (4, ["TR-001", datetime(2026, 8, 3), "ｃｍｐ－101", "CMP研磨装置1号機", "機械", "スラリー流量低下", "フィルター交換",
             tr001_hours, "1,234", "田中", "8/3 10:00 田中：ライン停止の連絡あり。\n8/3 10:20 原点復帰で復旧。"]),
        (5, ["TR-002", "R8.8.10", "CMP-101", None, "電気", "#1 アラーム", "- 再起動", "90分", "▲500", "佐藤", None]),
        (6, [None, None, None, None, None, None, "続き：ケーブル交換", None, None, None, None]),
        (7, ["TR-003", "8/15", "CVD-201", "CVD 1号機", "電気", "真空度低下", "Oリング交換", 2, "300-", "高橋", None]),
        (8, ["TR-004", 20260820, "CVD-201", None, "機械", "+停止(セル先頭が記号)", "@cmd", 0.5, 100, "鈴木", None]),
        (9, ["TR-005", datetime(2026, 8, 21), "CVD-201", "CVD 1号機", "機械", "非表示の行", "x", 9, 0, "", None]),
        (10, ["TR-006", datetime(2026, 8, 22), "CVD-201", "CVD 1号機", "機械", "取り消し線の行", "x", 9, 0, "", None]),
        (11, ["TR-007", datetime(2026, 7, 30), "ETC-301", "エッチャー1号機", "その他", "7月の記録", "確認", 1, 0, "", None]),
    ]
    if not include_tr004:
        rows = [r for r in rows if r[1][0] != "TR-004"]
    hours = sum(1.5 if r[1][0] == "TR-002" else r[1][7] for r in rows
                if r[1][0] not in (None, "TR-005", "TR-006"))
    for row_no, values in rows:
        for col, v in enumerate(values, start=1):
            if v is not None:
                ws.cell(row=row_no, column=col, value=v)
    ws.merge_cells("D4:D5")  # 設備名の縦結合
    ws.row_dimensions[9].hidden = True
    for col in range(1, 12):
        ws.cell(row=10, column=col).font = Font(strike=True)
    ws.cell(row=12, column=1, value="小計")
    ws.cell(row=12, column=8, value=hours)
    ws.cell(row=13, column=1, value="合計")
    ws.cell(row=13, column=8, value=hours)
    ws.cell(row=15, column=1, value="※停止時間は時間で記入")
    wb.save(path)
    return path


def list_spec_dict(**markdown) -> dict:
    md = {"file_prefix": "トラブル対応一覧", "group_by": "month", "max_records_per_file": 300}
    md.update(markdown)
    return {
        "name": "トラブル対応一覧",
        "description": "保全課のトラブル対応記録",
        "columns": [
            {"key": "record_no", "display": "管理No", "headers": ["管理No"], "type": "code", "role": "key", "required": True},
            {"key": "occurred_at", "display": "発生日", "headers": ["発生日"], "type": "date", "role": "date", "required": True},
            {"key": "equipment_id", "display": "設備番号", "headers": ["設備番号"], "type": "code", "role": "entity",
             "normalize": ["nfkc", "upper"]},
            {"key": "equipment_name", "display": "設備名", "headers": ["設備名"], "type": "string", "role": "entity_label",
             "fill_down_blank": True},
            {"key": "failure_category", "display": "故障区分", "headers": ["故障区分"], "type": "enum", "role": "category",
             "allowed": ["機械", "電気"]},
            {"key": "symptom", "display": "現象", "headers": ["現象"], "type": "text", "role": "text", "md": "body"},
            {"key": "action", "display": "処置", "headers": ["処置"], "type": "text", "role": "text", "md": "body"},
            {"key": "downtime", "display": "停止時間", "headers": ["停止時間"], "type": "number", "role": "measure", "unit": "分"},
            {"key": "cost", "display": "費用", "headers": ["費用"], "type": "number", "role": "measure", "unit": "円"},
            {"key": "worker", "display": "担当者", "headers": ["担当者"], "type": "string", "role": "person"},
            {"key": "response_log", "display": "対応内容", "headers": ["対応内容"], "type": "text", "role": "log", "md": "body"},
        ],
        "record": {"key": ["record_no"], "fallback_key": ["occurred_at", "equipment_id", "symptom:20"]},
        "period": {"date_column": "occurred_at"},
        "log_stage": {"column": "response_log", "mask": []},
        "markdown": md,
    }


def read_list(path: Path, spec):
    source = open_source(path, path.name)
    layout = guess_layout(source, "2026年8月")
    return read_records(source, {"sheet": "2026年8月"}, layout, spec)


def august(records) -> list[dict]:
    return [r.to_dict() for r in records if str(r.values.get("occurred_at", "")).startswith("2026-08")]


def _no_blank_inside_records(text: str) -> bool:
    lines = text.split("\n")
    for i, line in enumerate(lines[:-1]):
        if line == "":
            nxt = lines[i + 1]
            prev = lines[i - 1] if i else ""
            if not (nxt.startswith("## ") or prev.startswith("# ")):
                return False
    return True


# ---- 値の変換 -------------------------------------------------------------------------------

@pytest.mark.parametrize("type_, value, text, expected, flag", [
    ("date", None, "令和8年8月3日", "2026-08-03", None),
    ("date", None, "R8.8.3", "2026-08-03", None),
    ("date", None, "H31.4.30", "2019-04-30", None),
    ("date", None, "昭和64年1月7日", "1989-01-07", None),
    ("date", None, "令和元年5月1日", "2019-05-01", None),
    ("date", None, "20260803", "2026-08-03", None),
    ("date", 20260803, "20260803", "2026-08-03", None),
    ("date", None, "2026/8/3(月)", "2026-08-03", None),
    ("date", None, "８/５", "2026-08-05", "year_inferred"),
    ("date", None, "2/5", "2027-02-05", "year_inferred"),
    ("date", 46237.0, "46237", "2026-08-03", None),
    ("datetime", None, "2026/08/03 14:20:00", "2026-08-03 14:20", None),
    ("datetime", None, "20260803 142000", "2026-08-03 14:20", None),
    ("datetime", None, "2026年8月3日 14時20分", "2026-08-03 14:20", None),
    ("datetime", datetime(2026, 8, 3, 0, 0), "", "2026-08-03", None),
    ("time", None, "142000", "14:20", None),
    ("time", None, "1420", "14:20", None),
    ("time", 0.5, "0.5", "12:00", None),
    ("time", timedelta(hours=1, minutes=5), "1:05", "01:05", None),
    ("time", dtime(9, 5), "09:05", "09:05", None),
])
def test_convert_dates_and_times(type_, value, text, expected, flag):
    col = ColumnSpec("c", "列", type=type_)
    cctx = ConvertContext(fiscal_year=2026, fiscal_start=4)
    out, error, got_flag = convert_cell(value, text, col, cctx)
    assert error is None
    assert out == expected
    assert got_flag == flag


@pytest.mark.parametrize("value, text, unit, header_unit, fmt, expected", [
    (None, "1,234", "", "", None, 1234),
    (None, "１，２３４", "", "", None, 1234),
    (None, "△500", "", "", None, -500),
    (None, "▲1,000", "円", "", None, -1000),
    (None, "300-", "", "", None, -300),
    (None, "98.5%", "%", "", None, 98.5),
    (0.985, "0.985", "%", "", "0.0%", 98.5),
    (None, "1.5h", "分", "", None, 90),
    (None, "2時間", "分", "", None, 120),
    (1.5, "1.5", "分", "h", None, 90),
    (None, "90分", "分", "h", None, 90),
    (timedelta(hours=2, minutes=30), "2:30", "分", "", None, 150),
    (timedelta(minutes=90), "1:30", "時間", "", None, 1.5),
    (None, "¥12,000", "円", "", None, 12000),
    (None, "1.23E+3", "", "", None, 1230),
    (3.0000000001, "3", "", "", None, 3),
])
def test_convert_numbers(value, text, unit, header_unit, fmt, expected):
    col = ColumnSpec("c", "列", type="number", unit=unit)
    out, error, _flag = convert_cell(value, text, col, ConvertContext(), fmt, header_unit)
    assert error is None
    assert out == expected


def test_convert_blank_na_errors_codes_and_maps():
    cctx = ConvertContext(na_tokens=frozenset({"-", "N/A"}))
    num = ColumnSpec("n", "数", type="number", unit="分")
    assert convert_cell(None, "－", num, cctx) == (None, None, None)
    assert convert_cell("#N/A", "#N/A", num, cctx) == (None, None, "error_value")
    out, error, _ = convert_cell(None, "約3", num, cctx)
    assert error and out == "約3"
    out, error, _ = convert_cell(None, "3件", num, cctx)
    assert error  # 分に換算できない単位
    date_col = ColumnSpec("d", "日付", type="date")
    out, error, flag = convert_cell(None, "8/5", date_col, ConvertContext())
    assert flag == "year_missing" and error
    code = ColumnSpec("k", "コード", type="code", normalize=["nfkc", "upper"])
    assert convert_cell(123, "123", code, cctx, "00000")[0] == "00123"
    assert convert_cell(None, "ｃｍｐ－１０１", code, cctx)[0] == "CMP-101"
    enum = ColumnSpec("e", "区分", type="enum", value_map={"電気系": "電気"})
    assert convert_cell(None, "電気系", enum, cctx)[0] == "電気"
    text = ColumnSpec("t", "文章", type="text")
    assert convert_cell(None, "  ＡＢＣ　 1\r\n  2行目  ", text, cctx)[0] == "ABC 1\n2行目"


def test_year_context():
    assert find_year_context([("タイトル行", "2026年度 故障一覧")])["fiscal_year"] == 2026
    assert find_year_context([("シート名", "FY25")])["fiscal_year"] == 2025
    assert find_year_context([("x", "令和8年度")])["fiscal_year"] == 2026
    ym = find_year_context([("ファイル名", "故障履歴_2026-08")])
    assert (ym["year"], ym["month"], ym["source"]) == (2026, 8, "ファイル名")
    assert find_year_context([("ファイル名", "T1_トラブル対応一覧_2023-2026")])["year"] is None
    assert find_year_context([("a", ""), ("b", "202608")])["month"] == 8


# ---- spec ------------------------------------------------------------------------------------

def test_spec_roundtrip_validate_and_resolve():
    spec = spec_from_dict(list_spec_dict())
    assert validate_spec(spec) == []
    again = spec_from_dict(json.loads(json.dumps(spec_to_dict(spec), ensure_ascii=False)))
    assert spec_hash(again) == spec_hash(spec)
    assert again.summaries()[0].id == "month"
    assert spec.date_key == "occurred_at" and spec.file_prefix == "トラブル対応一覧"

    bad = spec_from_dict({"name": "", "columns": [
        {"key": "1bad", "display": "x", "type": "weird", "role": "person"},
        {"key": "a", "display": "", "type": "number", "md": "nope"},
        {"key": "a", "display": "dup"},
    ], "record": {"key": ["missing"]}, "period": {"date_column": "a"},
        "markdown": {"group_by": "entity_month", "summaries": [{"id": "month", "metrics": ["sum:nothing", "median"]}]},
        "log_stage": {"column": "zzz"}})
    errors = validate_spec(bad)
    joined = "\n".join(errors)
    for fragment in ("設定名", "1bad", "重複", "表示名", "型「weird」", "mdでの扱い", "記録キーの列「missing」",
                     "日付", "設備×月", "sum:nothing", "median", "zzz"):
        assert fragment in joined, fragment

    res = resolve_columns(spec, ["管理No", "発生日", "設備番号", "設備名", "故障区分", "現象", "処置内容", "停止時間(h)",
                                 "費用(円)", "担当者", "新しい列"])
    assert res.positions["downtime"] == 7 and res.positions["cost"] == 8
    assert "action" not in res.positions and "処置" in res.missing
    assert res.unused_headers == ["処置内容", "新しい列"]
    assert res.missing_required == []


def test_spec_from_suggestions(tmp_path):
    path = make_list_book(tmp_path / "list.xlsx")
    source = open_source(path, path.name)
    layout = guess_layout(source, "2026年8月")
    suggestions = suggest_columns(layout.headers, sample_data_rows(source, "2026年8月", layout))
    spec = spec_from_suggestions("トラブル", layout, suggestions, {"description": "説明", "group_by": "entity_month"})
    assert validate_spec(spec) == []
    keys = [c.key for c in spec.columns]
    assert keys[:4] == ["record_no", "occurred_at", "equipment_id", "equipment_name"]
    assert spec.record["key"] == ["record_no"]
    assert spec.record["fallback_key"][0] == "occurred_at"
    assert spec.period == {"date_column": "occurred_at"}
    assert spec.log_stage is not None and spec.log_stage.column == "response_log"
    assert spec.markdown["group_by"] == "entity_month"


# ---- 一覧: 読み込み → チェック → md → zip ---------------------------------------------------------

def test_list_read_normalize_and_checks(tmp_path):
    path = make_list_book(tmp_path / "list.xlsx")
    spec = spec_from_dict(list_spec_dict())
    records, row_issues, stats = read_list(path, spec)
    by_key = {r.key: r for r in records}
    assert list(by_key) == ["TR-001", "TR-002", "TR-003", "TR-004", "TR-007"]
    r1, r2, r3, r4 = (by_key[k].values for k in ("TR-001", "TR-002", "TR-003", "TR-004"))
    assert r1["equipment_id"] == "CMP-101" and by_key["TR-001"].originals["equipment_id"] == "ｃｍｐ－101"
    assert r1["downtime"] == 90 and r1["cost"] == 1234 and r1["occurred_at"] == "2026-08-03"
    assert r2["occurred_at"] == "2026-08-10" and r2["equipment_name"] == "CMP研磨装置1号機"  # 縦結合を埋める
    assert r2["action"] == "- 再起動\n続き:ケーブル交換"  # 継続行を連結（NFKC）
    assert r2["downtime"] == 90 and r2["cost"] == -500
    assert r3["occurred_at"] == "2026-08-15" and "年を" in by_key["TR-003"].warnings[0]
    assert r3["cost"] == -300 and r3["equipment_name"] == "CVD 1号機"
    assert r4["occurred_at"] == "2026-08-20" and r4["equipment_name"] == "CVD 1号機"  # 空欄＝上と同じ
    assert stats.filled_down == {"equipment_name": 1}
    assert stats.continuation_merged == 1
    assert stats.excluded == {"非表示の行": 1, "取り消し線の行": 1, "小計": 1, "合計": 1}
    assert stats.year_inferred == 1 and stats.year_context["fiscal_year"] == 2026
    assert stats.records == 5 and stats.months == {"2026-07": 1, "2026-08": 4}
    assert len(stats.reconcile) == 2 and all(rc["expected"] == rc["actual"] == 390 for rc in stats.reconcile)
    assert row_issues == []

    issues = run_checks(records, spec, stats)
    codes = [i.code for i in issues]
    assert "year_inferred" in codes and "not_allowed" in codes
    assert codes.count("excluded_rows") == 2
    assert not has_blocking(issues)

    # 見出しが足りない・型エラーが多い → 確定を止める
    broken = spec_from_dict(list_spec_dict())
    broken.column("symptom").required = True
    broken.column("symptom").headers = ["不具合内容"]
    broken.column("symptom").display = "不具合"
    broken.column("failure_category").type = "number"
    _recs, _row_issues, st2 = read_list(path, broken)
    issues2 = run_checks(_recs, broken, st2)
    assert {"required_missing", "type_error_rate"} <= {i.code for i in issues2 if i.level == "error"}
    assert len(_row_issues) == 5 and _row_issues[0].column == "故障区分"


def test_list_markdown_and_determinism(tmp_path):
    path = make_list_book(tmp_path / "list.xlsx")
    spec = spec_from_dict(list_spec_dict())
    records, _issues, _stats = read_list(path, spec)
    aug = august(records)
    ai = {
        "TR-001": {"status": "ok", "result": {
            "entries": [{"id": "e1", "segs": ["s1"], "t": ["連絡", "初動"]}],
            "incident": {"root_cause": {"v": "フィルター目詰まり", "certainty": "確定"},
                         "permanent_actions": [{"v": "フィルター交換"}],
                         "parts": [{"name": "フィルター", "model": "FL-10", "qty_q": "×1"}]}}},
        "TR-002": {"status": "flagged", "result": {"incident": {"root_cause": {"v": "出してはいけない"}}}},
    }
    files = render_all(spec, aug, ai, {"coverage": {"start": "2026-08", "end": "2026-08"}})
    names = [f.name for f in files]
    assert names == [
        "トラブル対応一覧_00_データセット説明.md",
        "トラブル対応一覧_2026-08.md",
        "トラブル対応一覧_集計_月次_2026-08.md",
        "トラブル対応一覧_集計_設備別_CMP-101_2026年度.md",
        "トラブル対応一覧_集計_設備別_CVD-201_2026年度.md",
    ]
    rec = files[1].text
    assert rec.startswith("# トラブル対応一覧 2026年8月の記録\n\n- データ種別: トラブル対応一覧（1行＝1件）の記録\n"
                          "- 対象期間: 2026-08-01〜2026-08-31\n- このファイルの記録: 4件（2026年8月の全件）\n\n## ")
    block1 = rec.split("\n\n")[2]
    assert block1.split("\n")[0] == "## 【TR-001】CMP研磨装置1号機（CMP-101）スラリー流量低下｜2026-08-03"
    for line in ("- 管理No: TR-001", "- 発生日: 2026-08-03（2026年8月）", "- 設備: CMP研磨装置1号機（CMP-101）",
                 "- 停止時間: 90分", "- 費用: 1,234円", "- 対応の要点（AI抽出）:", "  原因: フィルター目詰まり（確定）",
                 "  恒久処置: フィルター交換", "  使用部品: フィルター FL-10 ×1", "- 対応の時系列:",
                 "- 出典: list.xlsx（管理No TR-001）"):
        assert line in block1.split("\n"), line
    timeline = [ln for ln in block1.split("\n") if ln.startswith("  1. ")]
    assert timeline and timeline[0].startswith("  1. 2026-08-03 10:00［連絡・初動］") and "CMP研磨装置1号機" in timeline[0]
    assert "  2. 2026-08-03 10:20" in block1
    assert "担当者" not in rec and "田中｜" in rec  # 担当者列は出さない（時系列の記入者は原文由来）
    assert "出してはいけない" not in rec
    assert "- 処置:\n  \\- 再起動\n  続き:ケーブル交換" in rec
    assert "- 現象: +停止(セル先頭が記号)" in rec
    assert "|" not in rec.replace("｜", "")  # パイプ表を使わない
    assert _no_blank_inside_records(rec)
    assert "\r" not in rec and rec.endswith("\n") and not rec.endswith("\n\n")

    month = files[2].text
    assert "- 2026年8月のトラブル対応一覧の記録は4件です。" in month
    assert "- 2026年8月の停止時間の合計は330分（5.5時間）です。" in month
    assert "- 1位: CMP研磨装置1号機（CMP-101） 停止時間 180分、2件、主な故障区分: 機械" in month
    assert "- 2位: CVD 1号機（CVD-201） 停止時間 150分、2件、主な故障区分: 機械" in month
    fy = files[4].text
    assert "- CVD-201の2026年度（2026年8月〜2026年8月）の記録は2件です。" in fy
    assert "- 停止時間の合計は150分（2.5時間）、1件あたり平均75分です。" in fy
    card = files[0].text
    assert "- 取り込み範囲: 2026-08-01〜2026-08-31（記録 4件）" in card and "- 担当者" not in card

    # 決定的: 入力の順番を変えても同じバイト列
    shuffled = list(aug)
    random.Random(3).shuffle(shuffled)
    again = render_all(spec, shuffled, ai, {"coverage": {"start": "2026-08", "end": "2026-08"}})
    assert [hashlib.sha256(f.data).hexdigest() for f in again] == [hashlib.sha256(f.data).hexdigest() for f in files]

    # 分割・ヒント・人名を出す設定、設備×月
    spec2 = spec_from_dict(list_spec_dict(max_records_per_file=3, lightrag_hint=True, omit_person=False))
    files2 = render_all(spec2, aug, {}, {})
    part1 = next(f for f in files2 if f.name == "トラブル対応一覧_2026-08" + HINT)
    assert "トラブル対応一覧_2026-08_part2" + HINT in [f.name for f in files2]
    assert "- このファイルの記録: 3件（2026年8月の全 4件のうち 1〜3件目）" in part1.text
    assert "- 担当者: 田中" in part1.text
    files3 = render_all(spec_from_dict(list_spec_dict(group_by="entity_month")), aug, {}, {})
    ent = next(f for f in files3 if f.name == "トラブル対応一覧_CMP-101_2026-08.md")
    assert "- このファイルの記録: 2件（CMP-101の2026年8月の全件）" in ent.text

    # 7月と8月の両方がある取り込みは月ごとのファイルになる
    all_names = [f.name for f in render_all(spec, [r.to_dict() for r in records], {}, {})]
    assert "トラブル対応一覧_2026-07.md" in all_names and "トラブル対応一覧_2026-08.md" in all_names


def test_zip_csv_and_formula_guard():
    assert outputs.guard_formula("=HYPERLINK(\"x\")") == "'=HYPERLINK(\"x\")"
    for s in ("+1", "-2", "@SUM", "\tx", "\rx"):
        assert outputs.guard_formula(s).startswith("'")
    assert outputs.guard_formula("通常の文字") == "通常の文字"
    assert outputs.guard_formula(-5) == "-5" and outputs.guard_formula(None) == ""
    data = outputs.issues_csv([Issue("warning", "x", "=1+1", row=3, column="-列")])
    assert data.startswith(b"\xef\xbb\xbf")
    rows = list(csv.reader(io.StringIO(data.decode("utf-8-sig"))))
    assert rows[1] == ["警告", "x", "3", "'-列", "'=1+1"]
    zipped = outputs.build_zip([("b.md", b"B\n"), ("a.md", b"A\n")], {"問題一覧.csv": data})
    assert zipped == outputs.build_zip([("a.md", b"A\n"), ("b.md", b"B\n")], {"問題一覧.csv": data})
    with zipfile.ZipFile(io.BytesIO(zipped)) as zf:
        assert zf.namelist() == ["RAG投入用/a.md", "RAG投入用/b.md", "管理用_RAGには入れない/問題一覧.csv"]


# ---- CSV ---------------------------------------------------------------------------------------

def test_system_csv_quirks(tmp_path):
    text = ("出力日時,2026/09/01 10:00:00\n抽出条件,期間=2026/08\n\n"
            "管理番号,発生日,発生時刻,設備コード,停止時間(分),費用(円),現象\n"
            '="00123",20260803,142000,cmp-101,95,"1,234",スラリー流量低下\n'
            '="00124",20260804,090500,CMP-101,30-,500,"複数行の\n現象"\n'
            "合計件数,2\n")
    path = tmp_path / "システム出力_2026-08.csv"
    path.write_bytes(text.encode("cp932"))
    spec = spec_from_dict({
        "name": "故障", "columns": [
            {"key": "record_no", "display": "管理番号", "headers": ["管理番号"], "type": "code", "role": "key"},
            {"key": "occurred_at", "display": "発生日", "headers": ["発生日"], "type": "date", "role": "date"},
            {"key": "occurred_time", "display": "発生時刻", "headers": ["発生時刻"], "type": "time"},
            {"key": "equipment_id", "display": "設備コード", "headers": ["設備コード"], "type": "code", "role": "entity",
             "normalize": ["nfkc", "upper"]},
            {"key": "downtime", "display": "停止時間", "headers": ["停止時間"], "type": "number", "role": "measure", "unit": "分"},
            {"key": "cost", "display": "費用", "headers": ["費用"], "type": "number", "role": "measure", "unit": "円"},
            {"key": "symptom", "display": "現象", "headers": ["現象"], "type": "text", "role": "text"},
        ],
        "record": {"key": ["record_no"]},
    })
    source = open_source(path, path.name)
    layout = guess_layout(source, path.name)
    records, row_issues, stats = read_records(source, {}, layout, spec)
    assert [r.values for r in records] == [
        {"record_no": "00123", "occurred_at": "2026-08-03", "occurred_time": "14:20", "equipment_id": "CMP-101",
         "downtime": 95, "cost": 1234, "symptom": "スラリー流量低下"},
        {"record_no": "00124", "occurred_at": "2026-08-04", "occurred_time": "09:05", "equipment_id": "CMP-101",
         "downtime": -30, "cost": 500, "symptom": "複数行の\n現象"},
    ]
    assert records[1].source == {"file": path.name, "sheet": "", "row": 6}
    assert stats.excluded == {"合計": 1} and row_issues == []
    files = render_all(spec, [r.to_dict() for r in records], {}, {})
    rec = next(f for f in files if f.name == "故障_2026-08.md").text
    assert "- 発生日: 2026-08-03 14:20（2026年8月）" in rec and "発生時刻" not in rec
    assert "## 【00124】CMP-101 複数行の｜2026-08-04" in rec
    assert "- 現象:\n  複数行の\n  現象\n" in rec


# ---- DB・ジョブを通した流れ ------------------------------------------------------------------------

class FakeCtx:
    def __init__(self):
        self.progress_calls = []

    def progress(self, **kw):
        self.progress_calls.append(kw)

    def check_cancel(self):
        pass


def _new_import(app, spec_dict=None, template_id=None, name="list.xlsx"):
    upload_dir = Path(app.config["UPLOAD_DIR"]) / "tables"
    upload_dir.mkdir(parents=True, exist_ok=True)
    path = make_list_book(upload_dir / name)
    if template_id is None:
        spec = spec_from_dict(spec_dict or list_spec_dict())
        template_id, version_id = store.create_template(spec.name, spec)
    else:
        version_id = store.get_template(template_id)["current_version_id"]
    file_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    import_id = store.create_import(name, file_hash, f"tables/{name}", {"sheet": "2026年8月"}, template_id, version_id)
    return template_id, import_id


def test_pipeline_read_render_download(app):
    with app.app_context():
        template_id, import_id = _new_import(app)
        ctx = FakeCtx()
        result = pipeline.run_read(ctx, import_id)
        assert result["records"] == 5 and result["error"] == 0
        assert any(c.get("phase") == "読み込み" for c in ctx.progress_calls)
        imp = store.get_import(import_id)
        assert imp["status"] == "preview" and imp["stats"]["records"] == 5
        assert Path(imp["rows_path"]).exists() and Path(imp["issues_path"]).exists()
        assert len(pipeline.load_rows(import_id)) == 5 and len(pipeline.load_rows(import_id, 1, 2)) == 2

        listing = pipeline.preview_files(import_id, imp, pipeline.spec_for_import(imp))
        assert "トラブル対応一覧_2026-07.md" in [f["name"] for f in listing]
        preview_dir = pipeline.import_files(import_id)["preview"]
        assert pipeline.md_text(preview_dir, "トラブル対応一覧_2026-08.md").startswith("# トラブル対応一覧 2026年8月")
        assert pipeline.md_text(preview_dir, "../x.md") is None

        done = pipeline.run_render(FakeCtx(), import_id)
        imp = store.get_import(import_id)
        assert imp["status"] == "confirmed" and imp["confirmed_at"] and done["records"] == 5
        assert len(pipeline.md_paths(import_id)) == done["files"]
        assert store.get_version(imp["template_version_id"])["used"] == 1
        data = pipeline.build_download(import_id, imp, pipeline.spec_for_import(imp))
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
            assert len([n for n in names if n.startswith("RAG投入用/")]) == done["files"]
            for extra in ("正規化データ.csv", "問題一覧.csv", "取込レポート.csv"):
                assert f"管理用_RAGには入れない/{extra}" in names
            normalized = list(csv.reader(io.StringIO(zf.read("管理用_RAGには入れない/正規化データ.csv").decode("utf-8-sig"))))
            assert normalized[0][:3] == ["記録キー", "管理No", "発生日"]
            assert "'+停止(セル先頭が記号)" in next(r for r in normalized if r[0] == "TR-004")

        # 確定に使った版は上書きせず、新しい版にする
        spec = store.get_template(template_id)["spec"]
        spec.description = "v2"
        v2 = store.save_template_version(template_id, spec)
        assert v2 != imp["template_version_id"] and store.save_template_version(template_id, spec) == v2


def test_pipeline_blocking_and_broken_file(app):
    with app.app_context():
        spec_dict = list_spec_dict()
        spec_dict["columns"][5].update(required=True, headers=["不具合内容"], display="不具合内容")
        template_id, import_id = _new_import(app, spec_dict)
        pipeline.run_read(FakeCtx(), import_id)
        assert store.get_import(import_id)["stats"]["issue_counts"]["error"] == 1
        assert has_blocking(pipeline.load_issues(import_id))

        bad_path = Path(app.config["UPLOAD_DIR"]) / "tables" / "broken.xlsx"
        bad_path.write_bytes(b"not a zip")
        bad = store.create_import("broken.xlsx", "x", "tables/broken.xlsx", {}, template_id,
                                  store.get_template(template_id)["current_version_id"])
        with pytest.raises(Exception):
            pipeline.run_read(FakeCtx(), bad)
        failed = store.get_import(bad)
        assert failed["status"] == "failed" and "Excel" in failed["stats"]["error"]


def test_read_job_runs_in_worker(app):
    from core import jobs

    with app.app_context():
        _tid, import_id = _new_import(app)
        job_id = pipeline.start_read_job(import_id)
        job = jobs.wait_job(job_id, timeout=60)
        assert job["status"] == "done", job
        assert job["result"]["records"] == 5
        imp = store.get_import(import_id)
        assert imp["status"] == "preview" and imp["job_id"] == job_id


# ---- samples（実ファイル） ------------------------------------------------------------------------

@pytest.mark.samples
def test_sample_t1_end_to_end():
    path = SAMPLES / "T1_トラブル対応一覧_2023-2026.xlsx"
    if not path.exists():
        pytest.skip(f"{path.name} がありません")
    source = open_source(path, path.name)
    layout = guess_layout(source, "トラブル一覧")
    spec = spec_from_suggestions("トラブル対応一覧", layout,
                                 suggest_columns(layout.headers, sample_data_rows(source, "トラブル一覧", layout)))
    assert validate_spec(spec) == []
    records, row_issues, stats = read_records(source, {"sheet": "トラブル一覧"}, layout, spec)
    assert stats.records == 8095 and len({r.key for r in records}) == 8095
    assert stats.type_errors == {} and row_issues == []
    assert Counter(str(r.values.get("occurred_at"))[:4] for r in records) == {"2023": 1534, "2024": 2195, "2025": 2553, "2026": 1813}
    assert not has_blocking(run_checks(records, spec, stats))
    files = render_all(spec, [r.to_dict() for r in records], {}, {})
    assert sum(f.text.count("\n## ") for f in files if f.kind == "records") == 8095
    first = next(f for f in files if f.name == "トラブル対応一覧_2023-04.md")
    assert "- 発生日: 2023-04-01 02:21（2023年4月）" in first.text and "担当者" not in first.text
    assert _no_blank_inside_records(first.text)
