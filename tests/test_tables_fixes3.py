"""一覧表の不具合修正（3巡目）の確認。"""
from __future__ import annotations

import pytest

from tables.checks import has_blocking, run_checks
from tables.detect import guess_layout, sample_data_rows
from tables.mapping import suggest_columns
from tables.normalize import read_records
from tables.source import open_source
from tables.spec import spec_from_suggestions


def _csv(tmp_path, name: str, lines: list[str], encoding: str = "utf-8"):
    p = tmp_path / name
    p.write_bytes(("\r\n".join(lines) + "\r\n").encode(encoding))
    return open_source(p, name)


def _auto(src, sheet, tweak=None):
    layout = guess_layout(src, sheet)
    suggestions = suggest_columns(layout.headers, sample_data_rows(src, sheet, layout))
    if tweak:
        tweak(suggestions)
    spec = spec_from_suggestions("テスト", layout, suggestions, {})
    records, issues, stats = read_records(src, {"sheet": sheet}, layout, spec)
    return layout, spec, records, stats


# ---- T3-1: 日付を書き省いた行 ------------------------------------------------------------------

def _date_once_lines():
    lines = ["日付,設備,現象,処置,停止時間(分)"]
    for i in range(30):
        date = f"2026/08/{i // 3 + 1:02d}" if i % 3 == 0 else ""
        lines.append(f"{date},設備{i + 1},ポンプ停止{i + 1},部品を交換した{i + 1},{(i + 1) * 5}")
    return lines


@pytest.mark.parametrize("fill_down", [False, True])
def test_row_with_blank_date_but_own_numbers_is_its_own_record(tmp_path, fill_down):
    """日付を1日1回だけ書いた表で、日付が空の行も数値（停止時間）を持つなら別の記録。前の記録に黙って混ぜない。"""
    src = _csv(tmp_path, "once.csv", _date_once_lines())

    def tweak(suggestions):
        if fill_down:
            for s in suggestions:
                if s.header == "日付":
                    s.fill_down_blank = True

    _layout, _spec, records, stats = _auto(src, "once.csv", tweak)
    assert len(records) == 30 and stats.continuation_merged == 0
    downtime = [v for r in records for k, v in r.values.items() if isinstance(v, (int, float)) and k != "record_no"]
    assert sum(downtime) == sum((i + 1) * 5 for i in range(30))
    if fill_down:
        assert all(any(str(v).startswith("2026-08") for v in r.values.values()) for r in records)


def test_text_only_row_with_blank_keys_is_still_a_continuation(tmp_path):
    lines = ["管理No,発生日,現象,停止時間(分)"]
    for i in range(10):
        lines.append(f"TR-{i:03d},2026/08/{i + 1:02d},停止{i},{i + 1}")
        if i == 4:
            lines.append(",,追記: 後日点検")
    src = _csv(tmp_path, "cont.csv", lines)
    _layout, _spec, records, stats = _auto(src, "cont.csv")
    assert len(records) == 10 and stats.continuation_merged == 1
    assert any("後日点検" in str(v) for v in records[4].values.values())


# ---- T3-2: 秒の小数・時差つきの日時 ------------------------------------------------------------

@pytest.mark.parametrize("tail", [".000", "+09:00", "Z"])
def test_datetime_with_fraction_or_offset_is_converted(tmp_path, tail):
    """SQL Server 等の「2026-08-02 10:01:00.000」、Webの「…T10:01:00+09:00」「…Z」も日時として読む（確定を止めない）。"""
    sep = " " if tail == ".000" else "T"
    lines = ["管理No,発生日時,現象"] + [f"TR-{i:03d},2026-08-{i + 1:02d}{sep}10:{i:02d}:00{tail},停止{i}" for i in range(20)]
    src = _csv(tmp_path, "sqlexp.csv", lines)
    _layout, spec, records, stats = _auto(src, "sqlexp.csv")
    col = next(c for c in spec.columns if "発生日時" in c.headers)
    assert col.type == "datetime"
    assert not stats.type_errors
    assert records[3].values[col.key] == "2026-08-04 10:03"
    assert not has_blocking(run_checks(records, spec, stats))


# ---- T3-3: 空行の後の書き添え ------------------------------------------------------------------

def _footer_lines(footer: str):
    lines = ["管理No,発生日,設備,現象"] + [f"TR-{i:03d},2026/08/{i + 1:02d},設備{i},停止{i + 1}" for i in range(20)]
    return lines + ["", footer]


@pytest.mark.parametrize("footer", ["以上", "作成者：保全課 田中", ",,,以上"])
def test_footer_after_blank_row_is_not_a_record(tmp_path, footer):
    """表の下の「以上」「作成者：…」は記録にも、最後の記録の文章にもしない。"""
    src = _csv(tmp_path, "foot.csv", _footer_lines(footer))
    layout, _spec, records, stats = _auto(src, "foot.csv")
    assert layout.data_end == 21
    assert len(records) == 20 and stats.continuation_merged == 0
    assert all("以上" not in str(v) and "田中" not in str(v) for r in records for v in r.values.values())
    assert any(rc.index == 23 and rc.kind == "note" for rc in layout.row_classes)


def test_text_only_row_after_blank_row_is_not_merged_silently(tmp_path):
    """空行を挟んだ文章だけの行は、前の記録に黙って混ぜない（1件として出し、警告で気づけるようにする）。"""
    lines = ["管理No,発生日,設備,現象"] + [f"TR-{i:03d},2026/08/{i + 1:02d},設備{i},停止{i + 1}" for i in range(5)]
    lines += ["", ",,,別件の停止", "TR-100,2026/08/20,設備9,停止9"]
    src = _csv(tmp_path, "gap.csv", lines)
    _layout, _spec, records, stats = _auto(src, "gap.csv")
    assert stats.continuation_merged == 0 and len(records) == 7
    assert "別件" not in str(records[4].values)


# ---- T3-4: 結合した2段見出しを書き出したCSV ----------------------------------------------------

def test_csv_from_merged_two_row_header_reads_both_rows_as_header(tmp_path):
    lines = ["管理No,発生,,設備,,停止時間(分)", ",日付,時刻,番号,名称,"]
    lines += [f"TR-{i:03d},2026/08/{i + 1:02d},10:{i:02d},M-{i:02d},プレス{i},{i + 5}" for i in range(20)]
    src = _csv(tmp_path, "two_nopre.csv", lines)
    layout, _spec, records, stats = _auto(src, "two_nopre.csv")
    assert layout.header_rows == [1, 2]
    assert layout.headers[:5] == ["管理No", "発生_日付", "発生_時刻", "設備_番号", "設備_名称"]
    assert len(records) == 20 and not stats.type_errors


def test_csv_single_header_row_is_not_joined_with_first_data_row(tmp_path):
    lines = ["管理No,発生日,設備,現象"] + [f"TR-{i:03d},2026/08/{i + 1:02d},設備{i},停止{i + 1}" for i in range(20)]
    layout = guess_layout(_csv(tmp_path, "one.csv", lines), "one.csv")
    assert layout.header_rows == [1]


# ---- T3-5: CP932 のファイルに UTF-8 の行が混ざる -----------------------------------------------

def test_utf8_lines_appended_to_cp932_csv_are_warned(tmp_path):
    from tables.csv_source import sniff_csv

    head = "管理No,発生日,現象\r\n" + "".join(f"TR-{i:03d},2026/08/{i % 28 + 1:02d},ポンプ停止{i}\r\n" for i in range(20))
    tail = "".join(f"TR-{i:03d},2026/08/{i % 28 + 1:02d},ポンプ停止{i}\r\n" for i in range(20, 28))
    p = tmp_path / "mixed.csv"
    p.write_bytes(head.encode("cp932") + tail.encode("utf-8"))
    sniff = sniff_csv(p)
    assert sniff.encoding == "cp932"
    assert any("UTF-8 の行が混ざっています" in w and "22・23" in w for w in sniff.warnings)


def test_plain_cp932_csv_has_no_utf8_warning(tmp_path):
    from tables.csv_source import sniff_csv

    text = "管理No,発生日,現象\r\n" + "".join(f"TR-{i:03d},2026/08/{i % 28 + 1:02d},ポンプ停止ｱｲｳ髙{i}\r\n" for i in range(30))
    p = tmp_path / "sjis.csv"
    p.write_bytes(text.encode("cp932"))
    assert not any("UTF-8" in w for w in sniff_csv(p).warnings)


# ---- F4: CSV の読み込みエラーに英語の例外文を出さない -------------------------------------------

def test_csv_error_message_has_no_english_exception_text(tmp_path):
    import csv
    import re

    from tables.source import UploadError

    p = tmp_path / "unclosed.csv"
    p.write_bytes(('管理No,発生日,現象\r\nTR-001,2026/08/01,"閉じていない\r\n' + "続き,x\r\n" * 50).encode("utf-8"))
    old = csv.field_size_limit(100)
    try:
        from tables.csv_source import CsvSource

        src = CsvSource(p, "unclosed.csv", {"encoding": "utf-8", "delimiter": ","})
        with pytest.raises(UploadError) as e:
            list(src.rows())
    finally:
        csv.field_size_limit(old)
    msg = str(e.value)
    assert "値が大きすぎます" in msg and not re.search(r"[A-Za-z]{4,}", msg.replace("CSV", ""))


# ---- T3-6: 「対応」「作業」の文章の列を人名の列にしない ------------------------------------------

def test_prose_columns_similar_to_person_headers_stay_in_the_body(tmp_path):
    lines = ["管理No,発生日,対応,作業,担当"]
    for i in range(20):
        lines.append(f"TR-{i:03d},2026/08/{i + 1:02d},"
                     f"モーターの異音を確認したためベアリングを交換し、試運転で異常がないことを確かめた{i},"
                     f"カバーを外して内部を清掃し、潤滑油を補充したうえで締め付けを点検した{i},石川 陸")
    src = _csv(tmp_path, "prose.csv", lines)
    layout = guess_layout(src, "prose.csv")
    by_header = {s.header: s for s in suggest_columns(layout.headers, sample_data_rows(src, "prose.csv", layout))}
    for h in ("対応", "作業"):
        assert by_header[h].role != "person" and by_header[h].md != "omit"
    assert by_header["担当"].role == "person"


# ---- C3-2: 同じ設定の別の取り込みの AI の結果で、下書きを作り直さない ------------------------------

def test_ai_results_of_another_import_leave_the_preview_signature_alone(app, client):
    from aiproc import items as ai_items
    from tables import pipeline, store
    from tests.test_tables_fixes2 import _confirmed_import

    import_id = _confirmed_import(app, client)
    with app.app_context():
        imp = store.get_import(import_id)
        spec = pipeline.spec_for_import(imp)
        before = pipeline._preview_signature(import_id, imp, spec)
        ai_items.upsert_item(imp["template_id"], "log", "row-x", status="ok", import_id=import_id + 1000)
        assert pipeline._preview_signature(import_id, imp, spec) == before
        ai_items.upsert_item(imp["template_id"], "log", "row-x", status="ok", import_id=import_id)
        assert pipeline._preview_signature(import_id, imp, spec) != before


# ---- F2: Excel で計算されていない数式 -------------------------------------------------------------

def test_uncached_formulas_are_warned(tmp_path):
    """プログラムが書いて Excel で保存していない数式は値がない。空欄として読むことを警告で知らせる。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["管理No", "発生日", "設備", "停止時間(分)", "対応内容"])
    for i in range(12):
        ws.append([f"TR-{i:03d}", "2026-08-01", f"設備{i}", f"={i}+1", f"内容{i}"])
    p = tmp_path / "formula.xlsx"
    wb.save(p)
    src = open_source(p, "formula.xlsx")
    _layout, spec, records, stats = _auto(src, ws.title)
    assert sum(stats.uncached_formulas.values()) == 12
    issues = run_checks(records, spec, stats)
    assert any(i.code == "uncached_formula" and "12個" in i.message and "停止時間" in i.message for i in issues)


def test_cached_formula_values_are_not_warned(tmp_path):
    """Excel が保存した数式（値あり、空文字列の結果は t="str" と空の v）は警告しない。"""
    import zipfile

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["管理No", "発生日", "停止時間(分)", "備考"])
    for i in range(12):
        ws.append([f"TR-{i:03d}", "2026-08-01", 5, "x"])
    p = tmp_path / "cached.xlsx"
    wb.save(p)
    # 「Excel で保存した」形に書き換える: C列は値つきの数式、D列は空文字列を返す数式
    with zipfile.ZipFile(p) as z:
        items = {n: z.read(n) for n in z.namelist()}
    xml = items["xl/worksheets/sheet1.xml"].decode("utf-8")
    import re

    xml = re.sub(r'<c r="(C\d+)" t="n"><v>5</v></c>', r'<c r="\1"><f>2+3</f><v>5</v></c>', xml)
    xml = re.sub(r'<c r="(D(?!1")\d+)" t="inlineStr"><is><t>x</t></is></c>', r'<c r="\1" t="str"><f>""</f><v></v></c>', xml)
    items["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(p, "w") as z:
        for n, data in items.items():
            z.writestr(n, data)
    src = open_source(p, "cached.xlsx")
    assert src.uncached_formulas(ws.title) == {}
    assert xml.count("<f>") == 24


# ---- F3: 打ち間違えた遠い年の日付 ----------------------------------------------------------------

def test_far_off_date_is_warned():
    from tables.markdown import render_all
    from tables.spec import spec_from_dict
    from tests.test_tables_pipe import _rec, list_spec_dict

    spec = spec_from_dict(list_spec_dict())
    dates = [f"2025-08-{d:02d}" for d in range(1, 11)] + ["2052-08-11"]
    records = [_rec(f"TR-{i:03d}", record_no=f"TR-{i:03d}", occurred_at=d, equipment_id="CMP-101", symptom="停止")
               for i, d in enumerate(dates)]
    for i, r in enumerate(records):
        r["source"]["row"] = i + 2
    issues = run_checks(records, spec, {})
    outlier = [i for i in issues if i.code == "date_outlier"]
    assert len(outlier) == 1 and "12行目（2052-08-11）" in outlier[0].message
    # 記録ファイルは記録のある月だけ（間の月は作らない）
    assert sorted(f.name for f in render_all(spec, records, {})) == ["トラブル対応一覧_2025-08.md",
                                                                     "トラブル対応一覧_2052-08.md"]


def test_short_gaps_between_months_are_not_warned():
    from tables.spec import spec_from_dict
    from tests.test_tables_pipe import _rec, list_spec_dict

    spec = spec_from_dict(list_spec_dict())
    records = [_rec("a", occurred_at="2025-01-10"), _rec("b", occurred_at="2025-06-10")]
    assert not [i for i in run_checks(records, spec, {}) if i.code == "date_outlier"]


def test_zero_serial_in_a_date_formatted_cell_is_a_type_error():
    from datetime import datetime

    from tables.normalize import ConvertContext, _convert_date

    out, error, _flag = _convert_date(datetime(1899, 12, 30), "1899-12-30", "date", ConvertContext())
    assert error and "変換できません" in error
    assert _convert_date(datetime(2026, 8, 1), "2026-08-01", "date", ConvertContext())[:2] == ("2026-08-01", None)


# ---- R3B-5: 「出さない」にした理由 -------------------------------------------------------------

def test_only_blank_columns_are_left_out_by_default(tmp_path):
    lines = ["管理No,発生日,起票者,予備,区分コード,現象"]
    for i in range(30):
        lines.append(f"TR-{i:03d},2026/08/{i % 28 + 1:02d},石川 陸,,A{i % 3},ポンプ停止{i}")
    src = _csv(tmp_path, "omit.csv", lines)
    layout = guess_layout(src, "omit.csv")
    by_header = {s.header: s for s in suggest_columns(layout.headers, sample_data_rows(src, "omit.csv", layout))}
    # 人名の列・コードの列も既定で出す（空欄だけの列だけ「出さない」にする）
    assert by_header["起票者"].md != "omit" and by_header["区分コード"].md != "omit"
    assert by_header["予備"].md == "omit" and by_header["予備"].omit_reason == "blank"
    assert by_header["現象"].omit_reason == ""


def test_preview_is_not_queued_behind_a_paused_ai_job(app, client):
    """AI整形が一時停止中のまま「内容の確認」の段を開いても、下書きのジョブを後ろに並べて黙って待たせない。"""
    import time

    from core import jobs
    from tests.tables_helpers import imported, panel

    import_id = imported(app, client, "トラブル一覧.csv", "停止中テスト", ai_role="log")

    def ai(ctx):
        while ctx.wait_if_paused():
            time.sleep(0.01)

    with app.app_context():
        ai_job = jobs.start_job("ai_format", "table_import", import_id, ai)
        deadline = time.time() + 5
        while jobs.get_job(ai_job)["status"] != "running" and time.time() < deadline:
            time.sleep(0.01)
        jobs.request_pause(ai_job)
        while jobs.get_job(ai_job)["status"] != "paused" and time.time() < deadline:
            time.sleep(0.01)
    try:
        preview = panel(client, import_id, "preview")
        assert "AI整形が動いています（一時停止中を含む）" in preview["locked"]
        with app.app_context():
            assert jobs.latest_job("table_import", import_id, kind="table_preview") is None
        assert "確認に進めません" in panel(client, import_id, "ai")["html"]
    finally:
        with app.app_context():
            jobs.request_cancel(ai_job)
            jobs.wait_job(ai_job, timeout=5)
