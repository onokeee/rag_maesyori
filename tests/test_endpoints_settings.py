"""画面の経路: 取り込み設定の編集保存（「列の対応づけ」の段）、帳票の種類の見本と状態、
帳票の元ファイルのダウンロード、Excel のシートの選択。"""
import io

import openpyxl

from models import database as db
from tables import store
from tables.spec import spec_to_dict
from tests.tables_helpers import (CSV_TEXT, columns_payload, editor_body, json_only_spec, panel, panel_html,
                                  save_columns, save_layout, save_source, template_import, upload_bytes, upload_csv,
                                  wait_import_job)


# ---- 取り込み設定の編集（「列の対応づけ」の段。設定だけを開く画面は無い） ----------------------------

def test_saving_the_template_editor_unchanged_keeps_json_only_settings(app, client):
    """画面に出ない設定（見出しの別名・単位換算・値の置き換え・集計）は、そのまま保存しても消えない。"""
    import_id, template_id = template_import(app, client, json_only_spec())
    with app.app_context():
        before = spec_to_dict(store.get_template(template_id)["spec"])

    body = editor_body(client, import_id)
    res = save_columns(client, import_id, body)
    assert res.status_code == 200, res.get_json()
    wait_import_job(app, import_id)
    with app.app_context():
        after = spec_to_dict(store.get_template(template_id)["spec"])

    assert after["columns"] == before["columns"]
    assert after["markdown"]["summaries"] == before["markdown"]["summaries"]
    assert after["markdown"]["dataset_card"] is False
    assert after["markdown"]["title_columns"] == ["equipment_name", "symptom"]
    # search_rows は判定に使っていないので読み捨てる（R6-T-04）。anchors は残る
    assert "search_rows" not in after["header"] and after["header"]["anchors"] == ["管理No"]
    assert after["custom_stages"] == before["custom_stages"]
    assert after["log_stage"] == before["log_stage"]
    assert after["period"]["grain"] == "month"


def test_template_editor_changes_are_saved_and_stale_metrics_dropped(app, client):
    import_id, template_id = template_import(app, client, json_only_spec())
    body = editor_body(client, import_id)
    for row in body["columns"]:
        if row["key"] == "downtime":
            row["use"] = False         # 集計に使っていた列を外す
        if row["key"] == "equipment_name":
            row["type"], row["role"] = "code", "entity"    # 型と役割を変えた列は画面の役割に合わせる
    body["file_prefix"] = "新しい接頭辞"
    res = save_columns(client, import_id, body)
    assert res.status_code == 200, res.get_json()
    wait_import_job(app, import_id)
    with app.app_context():
        spec = store.get_template(template_id)["spec"]
    assert spec.column("downtime") is None
    assert [s.metrics for s in spec.summaries()] == [["count"], ["count"]]
    assert spec.markdown["file_prefix"] == "新しい接頭辞"
    eq = spec.column("equipment_name")
    assert eq.role == "entity" and eq.value_map == {"ロボ2": "搬送ロボット2号機"}   # 値の置き換えは引き継ぐ


def test_renamed_header_keeps_the_old_header_in_the_template(app, client):
    """見出しの名前が変わったファイルを読み直しても、前の見出しのファイルは対応づけをやり直さずに読める。"""
    first = upload_csv(client, "a.csv")
    assert save_source(client, first, encoding="cp932", delimiter=",", template="new",
                       new_template_name="T").status_code == 200
    assert save_layout(client, first).status_code == 200
    assert save_columns(client, first, columns_payload("T")).status_code == 200
    wait_import_job(app, first)
    with app.app_context():
        template_id = store.get_import(first)["template_id"]

    # 新しいファイルで見出し「設備番号」が「機番」に変わった
    second = upload_csv(client, "b.csv", CSV_TEXT.replace("設備番号", "機番", 1))
    assert save_source(client, second, encoding="cp932", delimiter=",",
                       template=str(template_id)).status_code == 200
    res = save_layout(client, second)
    assert res.status_code == 200 and res.get_json()["next"] == "columns"   # 見出しが合わないので対応づけへ
    headers = {"equipment_id": "機番"}
    payload = columns_payload("T")
    for row in payload["columns"]:
        if row["key"] in headers:
            row["header"] = row["display"] = headers[row["key"]]
    assert save_columns(client, second, payload).status_code == 200
    wait_import_job(app, second)
    with app.app_context():
        headers = store.get_template(template_id)["spec"].column("equipment_id").headers
    assert headers[0] == "機番" and "設備番号" in headers

    # 前の見出しのファイルも、列の対応づけをやり直さずに読める
    third = upload_csv(client, "c.csv")
    assert save_source(client, third, encoding="cp932", delimiter=",",
                       template=str(template_id)).status_code == 200
    res = save_layout(client, third)
    assert res.status_code == 200 and res.get_json()["next"] != "columns"
    wait_import_job(app, third)


# ---- 帳票の種類: 見本・状態・削除、帳票の JSON と元ファイル ------------------------------------------

def _upload(client, url, path, field, extra=None):
    data = {**(extra or {}), field: (io.BytesIO(path.read_bytes()), path.name)}
    return client.post(url, data=data, content_type="multipart/form-data")


def _make_form_type(app, client, sample_dir) -> int:
    from tests.test_forms_flow import add_field, create_type

    pattern_id = create_type(client, sample_dir / "standard.xlsx", "設備修理報告書")
    add_field(client, pattern_id, "修理報告書", "A3", "B3")
    add_field(client, pattern_id, "修理報告書", "A4", "B4")
    return pattern_id


def _upload_files(tmp_path) -> set[str]:
    return {p.name for p in (tmp_path / "uploads").rglob("*") if p.is_file()}


def test_form_type_samples_status_and_delete(app, client, sample_dir, tmp_path):
    pattern_id = _make_form_type(app, client, sample_dir)
    other_id = _make_form_type(app, client, sample_dir)

    res = client.post(f"/form-types/{pattern_id}/status", json={"status": "active"})
    assert res.status_code == 200 and "使用を開始しました" in res.get_json()["message"]
    with app.app_context():
        assert db.load_pattern(pattern_id).status == "active"
    assert client.post(f"/form-types/{pattern_id}/status", json={"status": "zzz"}).status_code == 400
    assert client.post("/form-types/9999/status", json={"status": "active"}).status_code == 404

    # 見本の追加（Excel でないファイルはエラー）
    res = _upload(client, f"/form-types/{pattern_id}/samples", sample_dir / "shifted.xlsx", "samples")
    assert res.status_code == 200 and "見本ファイルを1件追加しました" in res.get_json()["message"]
    before = _upload_files(tmp_path)
    res = client.post(f"/form-types/{pattern_id}/samples",
                      data={"samples": (io.BytesIO(b"not excel"), "x.xlsx")},
                      content_type="multipart/form-data")
    assert res.status_code == 400 and "1件目のファイル: Excelファイル" in res.get_json()["error"]
    assert _upload_files(tmp_path) == before          # 読めなかったファイルは残さない
    with app.app_context():
        samples = db.list_samples(pattern_id)
        other_samples = db.list_samples(other_id)
    assert len(samples) == 2

    # 他の種類の見本は消せない
    assert client.post(f"/form-types/{pattern_id}/samples/{other_samples[0]['id']}/delete").status_code == 404
    assert client.post(f"/form-types/{pattern_id}/samples/9999/delete").status_code == 404
    with app.app_context():
        assert len(db.list_samples(other_id)) == 1

    stored = samples[-1]["stored_path"].split("/")[-1]
    assert stored in _upload_files(tmp_path)
    res = client.post(f"/form-types/{pattern_id}/samples/{samples[-1]['id']}/delete")
    assert res.get_json()["message"] == "見本ファイルを削除しました"
    assert stored not in _upload_files(tmp_path)
    with app.app_context():
        assert [s["id"] for s in db.list_samples(pattern_id)] == [samples[0]["id"]]

    res = client.post(f"/form-types/{pattern_id}/status", json={"status": "inactive"})
    assert "使用を停止しました" in res.get_json()["message"]
    with app.app_context():
        assert db.load_pattern(pattern_id).status == "inactive"

    # 種類を削除すると見本ファイルも消える
    remaining = samples[0]["stored_path"].split("/")[-1]
    res = client.post(f"/form-types/{pattern_id}/delete")
    assert "を削除しました" in res.get_json()["message"]
    assert remaining not in _upload_files(tmp_path)
    with app.app_context():
        assert db.list_samples(pattern_id) == []
        assert len(db.list_samples(other_id)) == 1
    assert client.post(f"/form-types/{pattern_id}/delete").status_code == 404


def test_the_original_file_can_be_downloaded_without_losing_the_form(app, client, sample_dir):
    """元のファイルのダウンロードでは何も消さない（消えるのは Markdown を渡したときだけ）。"""
    from tests.test_forms_flow import activate, read_form, upload_forms

    pattern_id = _make_form_type(app, client, sample_dir)
    activate(client, pattern_id)
    doc_id, = upload_forms(client, sample_dir / "standard.xlsx")
    read_form(client, doc_id, pattern_id, ["修理報告書"])

    res = client.get(f"/forms/{doc_id}/original")
    assert res.status_code == 200 and res.data == (sample_dir / "standard.xlsx").read_bytes()
    assert "standard.xlsx" in res.headers["Content-Disposition"]
    with app.app_context():
        assert db.get_document(doc_id) is not None
    assert client.get("/forms/9999/original").status_code == 404
    # 確定していない帳票の Markdown は渡さない
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 404


# ---- 一覧表: Excel のシートの選択 ---------------------------------------------------------------

def test_save_source_uses_the_chosen_excel_sheet(app, client, tmp_path):
    book = openpyxl.Workbook()
    memo = book.active
    memo.title = "メモ"
    memo["A1"] = "このブックの説明"
    sheet = book.create_sheet("一覧")
    sheet.append(["管理No", "発生日", "現象"])
    for i in range(1, 6):
        sheet.append([f"TR-{i:03d}", f"2026-08-{i:02d}", f"アラーム停止{i}"])
    path = tmp_path / "二つのシート.xlsx"
    book.save(path)

    import_id = upload_bytes(client, path.read_bytes(), path.name)
    page = panel_html(client, import_id, "source")
    assert "メモ" in page and "一覧" in page

    res = save_source(client, import_id, sheet="一覧", template="new", new_template_name="シート選択")
    assert res.status_code == 200 and res.get_json()["next"] == "layout"
    with app.app_context():
        assert store.get_import(import_id)["source"]["sheet"] == "一覧"
    page = panel_html(client, import_id, "layout")
    assert "TR-001" in page and "このブックの説明" not in page
    detect = client.post(f"/tables/imports/{import_id}/layout/detect", json={"header_rows": [1]}).get_json()
    assert detect["data_end"] == 6
    res = save_layout(client, import_id)
    assert res.status_code == 200 and res.get_json()["next"] == "columns"
    assert "発生日" in panel_html(client, import_id, "columns")

    # シートを選び直すと、前のシートで決めた見出し行・範囲は使わない
    save_source(client, import_id, sheet="メモ", template="new", new_template_name="シート選択")
    with app.app_context():
        src = store.get_import(import_id)["source"]
    assert src["sheet"] == "メモ" and "header_rows" not in src
    assert "このブックの説明" in panel_html(client, import_id, "layout")


def test_the_section_of_a_clicked_cell_is_saved_with_the_field(app, client, tmp_path):
    """回答欄の中のセルをクリックした項目は「探す区画」を持つ（消えると回答側の欄を読めない）。"""
    from openpyxl.styles import Font, PatternFill

    from tests.test_forms_flow import add_field, create_type

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "報告書"
    for coord, text in (("A1", "■ 発行側"), ("A4", "▼ 回答欄")):
        ws[coord] = text
        ws[coord].font = Font(bold=True)
        ws[coord].fill = PatternFill("solid", fgColor="DDDDDD")
    for coord, text in (("A2", "処置内容"), ("A5", "処置内容")):
        ws[coord] = text
        ws[coord].fill = PatternFill("solid", fgColor="DDDDDD")
    ws["B2"], ws["B5"] = "発行側の処置", "回答側の処置"
    path = tmp_path / "回答欄.xlsx"
    wb.save(path)

    pattern_id = create_type(client, path, "工程異常連絡書")
    add_field(client, pattern_id, "報告書", "A5", "B5")
    with app.app_context():
        field = db.load_pattern(pattern_id).fields[0]
    assert field.section == "回答"
    # 画面を開き直しても区画は残り、読み取りテストは回答側の値を出す
    # （見本のシートには両方の値が写っているので、Markdown の中身で確かめる）
    panel = client.get(f"/form-types/{pattern_id}/panel").get_json()["html"]
    markdown = panel[panel.index("md-preview"):]
    assert "## 処置内容\n回答側の処置" in markdown and "発行側の処置" not in markdown
    with app.app_context():
        assert db.load_pattern(pattern_id).fields[0].section == "回答"
