"""一覧表の不具合修正（4巡目）の確認。"""
from __future__ import annotations

from tables.detect import classify_rows, guess_layout
from tables.source import open_source


def _csv(tmp_path, name: str, lines: list[str], encoding: str = "utf-8"):
    p = tmp_path / name
    p.write_bytes(("\r\n".join(lines) + "\r\n").encode(encoding))
    return open_source(p, name)


# ---- T4-1: 計器名（pH計・温度計・電力量累計）の行は小計・合計ではない ----------------------------

_METERS = ["pH計", "温度計", "O2計", "流量計", "湿度計", "圧力計", "CO2計", "導電率計", "照度計", "粘度計"]


def test_meter_names_with_own_date_are_data_rows(tmp_path):
    lines = ["計器,点検日,指示値,備考"] + [f"{m},2026/08/{i + 1:02d},{i + 1}.5," for i, m in enumerate(_METERS)]
    src = _csv(tmp_path, "m1.csv", lines)
    layout = guess_layout(src, src.sheets()[0].name)
    assert layout.counts.get("subtotal", 0) == 0
    assert layout.counts.get("data") == 10


def test_cumulative_meter_rows_stay_data_with_auto_and_manual_range(tmp_path):
    lines = ["計器,点検日,指示値,備考"]
    names = ["pH計", "温度計", "電力量累計", "流量計", "稼働時間累計", "圧力計", "CO2計", "照度計", "粘度計", "湿度計"]
    for i, m in enumerate(names):
        lines.append(f"{m},2026/08/{i + 1:02d},{(i + 1) * 100},良好")
    lines.append("合計,,,1234")
    src = _csv(tmp_path, "m2.csv", lines)
    sheet = src.sheets()[0].name
    layout = guess_layout(src, sheet)
    assert layout.data_end >= 11
    kinds = {row.index: rc.kind for row, rc in classify_rows(src, sheet, layout)}
    assert kinds[4] == "data" and kinds[6] == "data"
    assert kinds[12] == "subtotal"  # 日付のない本当の合計行は今までどおり
    manual = guess_layout(src, sheet, header_rows=[1], data_end=12)
    kinds = {row.index: rc.kind for row, rc in classify_rows(src, sheet, manual)}
    assert kinds[4] == "data" and kinds[6] == "data" and kinds[12] == "subtotal"


def test_plain_subtotal_row_is_still_subtotal(tmp_path):
    lines = ["部署,日付,件数"] + [f"製造{i},2026/08/{i + 1:02d},{i}" for i in range(6)] + ["部署計,,123"]
    src = _csv(tmp_path, "s.csv", lines)
    sheet = src.sheets()[0].name
    layout = guess_layout(src, sheet, header_rows=[1], data_end=8)
    kinds = {row.index: rc.kind for row, rc in classify_rows(src, sheet, layout)}
    assert kinds[8] == "subtotal"


# ---- T4-2 / T4-3 / T4-5: 列の対応づけの保存で、値から決めた内容を壊さない ------------------------

def _editor_body(app, client, text=None):
    """取り込みを置いて「列の対応づけ」の段を開き、画面が送る形の表を返す。"""
    from tests.tables_helpers import SPEC_CSV, editor_body, spec_import

    import_id = spec_import(app, client, text=text or SPEC_CSV)
    return import_id, editor_body(client, import_id)


def _saved(app, client, import_id, body):
    from tables import pipeline, store
    from tests.tables_helpers import save_columns, wait_import_job

    res = save_columns(client, import_id, body)
    assert res.status_code == 200, res.get_json()
    wait_import_job(app, import_id)
    with app.app_context():
        return pipeline.spec_for_import(store.get_import(import_id))


def test_the_editor_only_asks_for_use_and_role(app, client):
    """画面に出すのは「使う・見出し・役割」だけ（キー・型・単位・出し方は値から決める）。"""
    from tests.tables_helpers import panel_html, spec_import

    import_id = spec_import(app, client)
    html = panel_html(client, import_id, "columns")
    for gone in ('data-field="key"', 'data-field="type"', 'data-field="unit"', 'data-field="md"',
                 'data-field="display"', 'data-field="description"', 'data-field="fill_down_blank"',
                 'data-field="ai"', 'data-setting="file_prefix"', 'data-setting="description"',
                 'data-setting="group_by"'):
        assert gone not in html, gone
    assert 'data-field="use"' in html and 'data-field="role"' in html and 'data-setting="name"' in html


def test_unticked_column_is_still_removed(app, client):
    import_id, body = _editor_body(app, client)
    for r in body["columns"]:
        if r["header"].startswith("停止時間"):
            r["use"] = False
    after = _saved(app, client, import_id, body)
    assert after.column("downtime") is None
    assert after.column("record_no") is not None


def test_the_record_key_follows_the_role_picked_on_screen(app, client):
    from tests.tables_helpers import set_role

    import_id, body = _editor_body(app, client)
    after = _saved(app, client, import_id, body)
    assert after.record["key"] == ["record_no"]

    set_role(body, "管理No", "attribute")
    set_role(body, "設備名", "key")
    after = _saved(app, client, import_id, body)
    assert after.record["key"] == [after.first_role("key").key]
    assert after.first_role("key").display == "設備名"


def test_the_table_name_decides_the_file_names(app, client):
    """「ファイル名の先頭」は無くなった。md の名前は「表の名前」を使う。"""
    import_id, body = _editor_body(app, client)
    body["name"] = "新しい名前"
    after = _saved(app, client, import_id, body)
    assert after.markdown["file_prefix"] == "" and after.file_prefix == "新しい名前"


# ---- T4-4 / S4-1: JSON で取り込む設定の AI整形の項目を確かめる ----------------------------------------

def _log_spec(**stage):
    from tables.spec import spec_from_dict

    return spec_from_dict({"name": "x", "columns": [{"key": "log", "display": "ログ", "type": "text", "role": "log"}],
                           "log_stage": {"column": "log", **stage}})


def test_mask_given_as_a_string_is_read_as_rules():
    from tables.markdown import parse_log_cell
    from tables.spec import validate_spec

    text = "8/1 10:00 田中: 連絡先 090-1234-5678 / taro@example.com に電話"
    for mask in ("email", "phone,email", "phone、email"):
        spec = _log_spec(mask=mask)
        assert validate_spec(spec) == []
        out = repr(parse_log_cell(spec, {"log": text}))
        assert "taro@example.com" not in out and "［メール］" in out
    assert "090-1234-5678" not in repr(parse_log_cell(_log_spec(mask="phone,email"), {"log": text}))


def test_malformed_log_stage_and_checks_are_refused():
    from tables.spec import spec_from_dict, validate_spec

    assert any("伏せ字" in e for e in validate_spec(_log_spec(mask=["phone", "住所"])))
    assert any("人名一覧" in e for e in validate_spec(_log_spec(people=["田中"])))
    assert any("用語集" in e for e in validate_spec(_log_spec(glossary=["a"])))
    assert any("区切り" in e for e in validate_spec(_log_spec(splitter="x")))
    assert validate_spec(_log_spec(people=[{"name": "田中 一郎", "aliases": ["田中"]}])) == []
    spec = spec_from_dict({"name": "x", "columns": [{"key": "a", "display": "A"}],
                           "checks": {"reconcile_tolerance": "なし"}})
    assert any("許容差" in e for e in validate_spec(spec))


def test_bad_or_slow_split_patterns_are_refused():
    from tables.spec import validate_spec

    assert any("正しくありません" in e for e in validate_spec(_log_spec(splitter={"extra_anchors": ["["]})))
    assert any("遅く" in e for e in validate_spec(_log_spec(splitter={"extra_anchors": ["(.+)+X"]})))
    assert any("遅く" in e for e in validate_spec(_log_spec(splitter={"not_date_patterns": ["(a*)*b"]})))
    assert validate_spec(_log_spec(splitter={"extra_anchors": [r"^【\d+】"]})) == []


# ---- R4-MD-1: 「〃」「同上」は直前の行の値で補う ----------------------------------------------------

def test_ditto_marks_are_filled_from_the_row_above(tmp_path):
    from tables.markdown import render_all
    from tests.test_tables_fixes3 import _auto

    lines = ["管理No,発生日,設備ID,設備名,現象,停止時間(分)",
             "TR-1,2025-03-01,CMP-101,CMP研磨機1号機,異音,30",
             "TR-2,〃,〃,〃,振動,20",
             "TR-3,2025-03-02,ETC-301,エッチャ1号機,停止,40",
             "TR-4,2025-03-02,同上,″,停止,50",
             "TR-5,2025-03-03,CVD-201,CVD1号機,警報,10",
             "TR-6,2025-03-04,CVD-202,CVD2号機,警報,10"]
    src = _csv(tmp_path, "d.csv", lines)
    layout, spec, records, stats = _auto(src, src.sheets()[0].name)
    by_key = {r.key: r.values for r in records}
    date_key = spec.date_key
    eq = spec.first_role("entity").key
    assert str(by_key["TR-2"][date_key]).startswith("2025-03-01") and by_key["TR-2"][eq] == "CMP-101"
    assert by_key["TR-4"][eq] == "ETC-301"
    assert stats.ditto_filled
    files = render_all(spec, [r.to_dict() for r in records], {})
    assert not any("〃" in f.name or "同上" in f.name for f in files)
    text = "\n".join(f.text for f in files)
    assert "〃" not in text and "日付なし" not in "".join(f.name for f in files)


def test_ditto_without_a_row_above_is_not_an_equipment(tmp_path):
    from tests.test_tables_fixes3 import _auto
    from tables.records import entity_value

    lines = ["管理No,発生日,設備ID,現象,停止時間(分)",
             "TR-1,2025-03-01,〃,異音,30",
             "TR-2,2025-03-02,CMP-101,振動,20",
             "TR-3,2025-03-02,ETC-301,停止,40",
             "TR-4,2025-03-03,CVD-201,警報,10"]
    src = _csv(tmp_path, "d2.csv", lines)
    _layout, spec, records, stats = _auto(src, src.sheets()[0].name)
    first = next(r for r in records if r.key == "TR-1")
    assert spec.first_role("entity") is not None
    assert entity_value(first.values, spec) == ("", "")
    assert not stats.ditto_filled


# ---- ux4-3: 変換できなかった日付は日付の範囲に入れない -------------------------------------------------

def test_date_range_ignores_unconverted_date_text(tmp_path):
    from tests.test_tables_fixes3 import _auto

    lines = ["管理No,発生日,現象"] + [f"TR-{i},{d},停止" for i, d in enumerate(
        ["2026-08-01", "2026/13/45", "不明", "2026-08-04", "2026-08-05", "2026-08-06"])]
    src = _csv(tmp_path, "u.csv", lines)
    _layout, _spec, _records, stats = _auto(src, src.sheets()[0].name)
    assert (stats.date_min, stats.date_max) == ("2026-08-01", "2026-08-06")


# ---- R4-MD-3: 時系列の記入者は、担当者の列から補った表記で出す ------------------------------------------

def _person_spec():
    from tables.spec import spec_from_dict

    return spec_from_dict({
        "name": "x", "columns": [
            {"key": "no", "display": "管理No", "type": "code", "role": "key"},
            {"key": "occurred_at", "display": "発生日", "type": "date", "role": "date"},
            {"key": "log", "display": "対応内容", "type": "text", "role": "log"},
            {"key": "worker", "display": "担当者", "type": "string", "role": "person"}],
        "log_stage": {"column": "log"}})


def _timeline(log: str) -> str:
    from tables.markdown import people_index_for, record_block

    spec = _person_spec()
    rec = {"key": "A-1", "values": {"no": "A-1", "occurred_at": "2026-07-01", "log": log, "worker": "井上 亮"},
           "originals": {}, "source": {"file": "a.csv", "row": 2}, "warnings": []}
    return "\n".join(record_block(rec, spec, None, people_index_for(spec, [rec])))


def test_the_person_column_is_written_and_fills_in_the_timeline_author():
    """人名の列は出す。時系列の記入者も担当者の列から補う（内容を隠さない）。"""
    log = "7/1 2:47 井上:連絡あり\n7/1 3:10 部品を交換した"
    text = _timeline(log)
    assert "- 担当者: 井上 亮" in text
    assert "02:47 井上 亮: 連絡あり" in text and "井上 亮（推定）" in text


# ---- R4-1: AIの試し実行中は取り込みを削除・ダウンロードできない -------------------------------------------

from tests.test_endpoints_ai import _read_import, ai_app, ai_client, fake  # noqa: E402,F401  (fixtures)


def test_delete_is_refused_while_an_ai_trial_is_running(ai_app, ai_client, monkeypatch):
    from aiproc import runner
    from tables import store

    import_id = _read_import(ai_app, ai_client)
    seen = {}

    def slow_trial(*_a, **_k):
        # 同じブラウザの別のタブから消しに来る（別のブラウザからは、そもそもこの取り込みが見えない）
        seen["delete"] = ai_client.post(f"/tables/imports/{import_id}/delete").get_json()
        raise runner.AIJobError("止めました")

    monkeypatch.setattr(runner, "trial_row", slow_trial)
    res = ai_client.post(f"/tables/imports/{import_id}/ai/trial", json={"row_key": "TR-001"})
    assert res.status_code == 400
    assert "AIの試し実行中は削除・ダウンロード・保存できません" in seen["delete"]["error"]
    with ai_app.app_context():
        assert store.get_import(import_id) is not None
    # 試し実行が終われば削除できる
    ai_client.post(f"/tables/imports/{import_id}/delete")
    with ai_app.app_context():
        assert store.get_import(import_id) is None


# ---- C4-1: AI整形の実行中・一時停止中は読み込み直さない ------------------------------------------------

def test_reread_is_refused_while_ai_format_is_running_or_paused(ai_app, ai_client, monkeypatch):
    from core import jobs
    from tables import store
    from views import tables as views_tables

    import_id = _read_import(ai_app, ai_client)
    with ai_app.app_context():
        before = store.get_import(import_id)
    monkeypatch.setattr(views_tables, "_ai_running", lambda _id: True)
    res = ai_client.post(f"/tables/imports/{import_id}/read")
    assert res.status_code == 409 and "処理中は変更できません" in res.get_json()["error"]
    with ai_app.app_context():
        after = store.get_import(import_id)
        assert after["status"] == before["status"] and after["job_id"] == before["job_id"]
        job = jobs.latest_job("table_import", import_id, kind="table_read")
        assert job is None or job["id"] == before["job_id"]


# ---- R4-FUZZ-1 / R4-FUZZ-2: 列数の上限・読めない値はアップロードの時点で断る ---------------------------------

def _no_import_left(app):
    from pathlib import Path

    from tables import store

    with app.app_context():
        assert store.list_imports(limit=10) == []
    assert [p for p in Path(app.config["UPLOAD_DIR"]).rglob("*") if p.is_file()] == []


def test_csv_with_a_huge_first_line_is_refused_at_upload(app, client):
    from tests.tables_helpers import upload

    header = ",".join(f"c{i}" for i in range(400_000))
    body = (header + "\r\n" + ",".join("1" for _ in range(3)) + "\r\n").encode("utf-8")
    res = upload(client, body, "wide.csv")
    assert res.status_code == 400 and "列数が上限" in res.get_json()["error"]
    _no_import_left(app)


def test_csv_rows_over_the_column_limit_are_refused_when_read(tmp_path):
    import pytest

    from core.files import UploadError
    from tables.csv_source import MAX_COLUMNS, CsvSource

    p = tmp_path / "w.csv"
    p.write_bytes(("a,b\r\n1,2\r\n" + ",".join("x" for _ in range(MAX_COLUMNS + 5)) + "\r\n").encode("utf-8"))
    src = CsvSource(p, "w.csv", {"encoding": "utf-8", "delimiter": ","})
    with pytest.raises(UploadError, match="3行目の列数が上限"):
        list(src.rows())


def test_xlsx_with_an_unreadable_value_far_down_is_refused_at_upload(app, client, tmp_path):
    import io
    import re
    import zipfile

    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["管理No", "発生日", "現象", "処置", "停止時間"])
    for i in range(300):
        ws.append([f"TR-{i}", "2026-08-01", "停止", "交換", i + 1000])
    wb.save(tmp_path / "ok.xlsx")
    src = zipfile.ZipFile(tmp_path / "ok.xlsx")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rb"<v>1250</v>", b"<v>NaN</v>", data, count=1)
            z.writestr(item, data)
    from tests.tables_helpers import upload

    res = upload(client, out.getvalue(), "nan.xlsx")
    assert res.status_code == 400 and "行目付近の値を読めません" in res.get_json()["error"]
    _no_import_left(app)


# ---- BR4-3: AI整形の対象を2列にして保存しない ------------------------------------------------------------

def test_two_ai_columns_are_refused_on_save(app, client):
    from tests.tables_helpers import save_columns, set_role

    import_id, body = _editor_body(app, client)
    set_role(body, "現象", "log")
    set_role(body, "対応内容", "log")
    res = save_columns(client, import_id, body)
    assert res.status_code == 400 and res.get_json()["error"] == "経過の記録の列は1つだけにしてください"


# ---- R4-2: 見出し行のないCSVは、見出しがデータのように見えると知らせる ------------------------------------

def test_headerless_csv_is_warned(tmp_path):
    lines = [f"TR-{i:03d},2026-08-{i:02d},EQ-0{i % 3},山田太郎: 主軸ベアリング損傷のため交換した。取引先に部品を手配し、翌日に復旧を確認した,{i}"
             for i in range(1, 9)]
    src = _csv(tmp_path, "nohead.csv", lines)
    layout = guess_layout(src, src.sheets()[0].name)
    assert any("見出し行がデータのように見えます" in w for w in layout.warnings)


def test_normal_header_is_not_warned(tmp_path):
    lines = ["管理No,発生日,設備,現象,停止時間(分)"] + [f"TR-{i},2026-08-{i:02d},EQ-1,停止,{i}" for i in range(1, 9)]
    src = _csv(tmp_path, "head.csv", lines)
    layout = guess_layout(src, src.sheets()[0].name)
    assert not any("見出し行がデータ" in w for w in layout.warnings)
