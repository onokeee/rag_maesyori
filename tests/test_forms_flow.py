"""帳票フローの画面の通しテスト（帳票の種類の作成 → 使用開始 → 取り込み → 確認 → 確定 → ダウンロード）。"""
import io
import json
import re

from models import database as db


def _upload(client, url, path, field, extra=None):
    data = {**(extra or {}), field: (io.BytesIO(path.read_bytes()), path.name)}
    return client.post(url, data=data, content_type="multipart/form-data")


def test_form_flow_happy_path(app, client, sample_dir):
    # 帳票の種類を作る（見本ファイル）→ 候補の確認 → 保存しても作成中のまま
    res = _upload(client, "/settings/form-types/new", sample_dir / "standard.xlsx", "samples", {"name": "設備修理報告書"})
    assert res.status_code == 302
    pattern_id = int(re.search(r"/settings/form-types/(\d+)/review", res.headers["Location"]).group(1))
    page = client.get(f"/settings/form-types/{pattern_id}/review").get_data(as_text=True)
    assert "読み取る項目の確認" in page and "RAGに出す" in page

    with app.app_context():
        from pattern.builder import suggest_rows
        from excel.workbook import load_workbook_info
        _, rows = suggest_rows([load_workbook_info(sample_dir / "standard.xlsx")])
    form = {"name": "設備修理報告書", "version": "v1", "mode": "review", "md_options_form": "1",
            "md_omit_person_fields": "on", "image_processing": "none"}
    for i, r in enumerate(rows):
        for key in ("field_name", "display_name", "candidates", "data_type", "direction", "unit", "rag_output"):
            form[f"fields-{i}-{key}"] = r[key]
        if r["use"]:
            form[f"fields-{i}-use"] = "on"
        if r["field_name"] == "report_id":
            form[f"fields-{i}-required"] = "on"
            form[f"fields-{i}-title_order"] = "1"
    res = client.post(f"/settings/form-types/{pattern_id}/save", data=form)
    assert res.status_code == 302 and res.headers["Location"].endswith(f"/{pattern_id}/test")
    with app.app_context():
        assert db.load_pattern(pattern_id).status == "draft"
    page = client.get(f"/settings/form-types/{pattern_id}/test").get_data(as_text=True)
    assert "使用開始" in page and "R2026-00123" in page
    assert client.post(f"/settings/form-types/{pattern_id}/status", data={"status": "active"}).status_code == 302

    # 帳票を取り込む
    assert client.get("/forms/new").status_code == 200
    res = _upload(client, "/forms/upload", sample_dir / "standard.xlsx", "file")
    assert res.status_code == 302
    doc_id = int(re.search(r"/forms/(\d+)/type", res.headers["Location"]).group(1))
    page = client.get(f"/forms/{doc_id}/type").get_data(as_text=True)
    assert "項目中" in page and "設備修理報告書" in page

    res = client.post(f"/forms/{doc_id}/read", data={"pattern_id": pattern_id, "sheets": ["修理報告書"]})
    assert res.headers["Location"].endswith(f"/forms/{doc_id}/review")
    page = client.get(f"/forms/{doc_id}/review").get_data(as_text=True)
    assert "確定してMarkdownを作成" in page and "作業内容は保存されています" in page
    assert 'data-cell="B3"' in page and "読み取り前" not in page

    # 途中保存（204）とプレビュー
    assert client.post(f"/forms/{doc_id}/draft", json={"values": {"equipment_name": "CMP研磨装置"}}).status_code == 204
    summary = client.post(f"/forms/{doc_id}/preview", json={"values": {}}).get_json()
    assert "CMP研磨装置" in summary["markdown"] and summary["counts"]["manual"] == 1

    # 必須が空なら止まる → チェックで確定
    res = client.post(f"/forms/{doc_id}/confirm", data={"value-report_id": ""})
    assert "/review" in res.headers["Location"]
    with app.app_context():
        assert db.get_document(doc_id)["state"] == "reviewing"
    res = client.post(f"/forms/{doc_id}/confirm", data={"value-report_id": "R2026-00123"})
    assert res.headers["Location"].endswith(f"/forms/{doc_id}/done")
    assert "Markdownをダウンロード" in client.get(f"/forms/{doc_id}/done").get_data(as_text=True)

    # 修正中 → 変更を破棄 → 確定済み。詳細・削除
    client.post(f"/forms/{doc_id}/draft", json={"values": {"equipment_name": "別の名前"}})
    with app.app_context():
        assert db.get_document(doc_id)["state"] == "modified"
    # 開いたままの確認・修正画面の状態タグを書き替えられるよう、プレビューは文書の状態も返す
    assert client.post(f"/forms/{doc_id}/preview", json={"values": {}}).get_json()["state"] == "modified"
    client.post(f"/forms/{doc_id}/discard-changes")
    with app.app_context():
        assert db.get_document(doc_id)["state"] == "confirmed"
    detail = client.get(f"/forms/{doc_id}").get_data(as_text=True)
    assert "確定済み" in detail and "元に戻せません" in detail
    assert client.get("/settings/form-types/").status_code == 200
    assert client.get(f"/settings/form-types/{pattern_id}/edit").status_code == 200

    # 確定済みの読み取り直しは確認のチェックが必要
    res = client.post(f"/forms/{doc_id}/reread", data={"pattern_id": pattern_id, "sheets": ["修理報告書"]})
    assert "/type" in res.headers["Location"]

    # ダウンロードが最後の手順（ダウンロードするとこのPCからデータが消える）
    md = client.get(f"/forms/{doc_id}/download.md")
    assert md.status_code == 200
    assert md.headers["Content-Type"] == "text/markdown; charset=utf-8"  # charset は1つだけ
    text = md.get_data(as_text=True)
    assert text.startswith("# 設備修理報告書 R2026-00123") and "CMP研磨装置" in text
    assert "別の名前" not in text          # 確定済みの版で作る
    with app.app_context():
        assert db.get_document(doc_id) is None
    assert client.get(f"/forms/{doc_id}").status_code == 404


def _parts_report(path):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "点検報告書"
    fill = PatternFill("solid", fgColor="FFD9E1F2")
    for coord, value in {"A1": "報告番号", "B1": "IN-2026-001", "A2": "設備番号", "B2": "CMP-101",
                         "A4": "■ 交換部品", "A5": "No.", "B5": "品番", "C5": "品名", "D5": "数量",
                         "A6": 1, "B6": "PW48-1591", "C6": "ベアリング", "D6": 2,
                         "A7": 2, "B7": "PW35-1577", "C7": "スピンモータ", "D7": 1,
                         "A9": "所見", "B9": "異常なし"}.items():
        ws[coord] = value
    for coord in ("A1", "A2", "A5", "B5", "C5", "D5", "A9"):
        ws[coord].fill = fill
    wb.save(path)
    return path


def test_form_flow_with_table_field(app, client, tmp_path):
    """明細表の項目: 見本から候補 → 保存（列見出しも保存）→ 取り込み → 確認画面で行を修正 → Markdown。"""
    path = _parts_report(tmp_path / "点検報告書_CMP-101.xlsx")
    res = _upload(client, "/settings/form-types/new", path, "samples", {"name": "点検報告書"})
    pattern_id = int(re.search(r"/settings/form-types/(\d+)/review", res.headers["Location"]).group(1))
    page = client.get(f"/settings/form-types/{pattern_id}/review").get_data(as_text=True)
    assert "明細表（列見出しと行）" in page and "品番: PW48-1591／品名: ベアリング" in page

    with app.app_context():
        from excel.workbook import load_workbook_info
        from pattern.builder import suggest_rows
        _, rows = suggest_rows([load_workbook_info(path)])
    form = {"name": "点検報告書", "version": "v1", "mode": "review", "md_options_form": "1", "image_processing": "none"}
    for i, r in enumerate(rows):
        for key in ("field_name", "display_name", "candidates", "data_type", "direction", "unit", "rag_output",
                    "table_columns"):
            form[f"fields-{i}-{key}"] = r[key]
        if r["use"]:
            form[f"fields-{i}-use"] = "on"
    assert client.post(f"/settings/form-types/{pattern_id}/save", data=form).status_code == 302
    with app.app_context():
        parts = next(f for f in db.load_pattern(pattern_id).fields if f.field_name == "parts")
        assert parts.data_type == "table" and parts.table_columns == ["No.", "品番", "品名", "数量"]
    page = client.get(f"/settings/form-types/{pattern_id}/test").get_data(as_text=True)
    assert "品番: PW35-1577／品名: スピンモータ" in page
    assert "品番" in client.get(f"/settings/form-types/{pattern_id}/edit").get_data(as_text=True)
    client.post(f"/settings/form-types/{pattern_id}/status", data={"status": "active"})

    res = _upload(client, "/forms/upload", path, "file")
    doc_id = int(re.search(r"/forms/(\d+)/type", res.headers["Location"]).group(1))
    client.post(f"/forms/{doc_id}/read", data={"pattern_id": pattern_id, "sheets": ["点検報告書"]})
    page = client.get(f"/forms/{doc_id}/review").get_data(as_text=True)
    assert "data-table-editor" in page and "スピンモータ</textarea>" in page and 'name="value-parts"' in page

    # 画面の値（hidden の JSON）が表示中と同じなら「手で修正」にしない
    summary = client.post(f"/forms/{doc_id}/preview", json={"values": {}}).get_json()
    assert summary["counts"]["manual"] == 0
    assert "## 交換部品\n- 品番: PW48-1591／品名: ベアリング／数量: 2\n- 品番: PW35-1577／品名: スピンモータ／数量: 1\n" \
        in summary["markdown"]
    with app.app_context():
        current = json.loads(db.get_document(doc_id)["data_json"])
    parts_value = next(f for f in current["fields"] if f["field_name"] == "parts")["value"]
    assert client.post(f"/forms/{doc_id}/draft", json={"values": {"parts": json.dumps(parts_value, ensure_ascii=False)}}
                       ).status_code == 204
    assert client.post(f"/forms/{doc_id}/preview", json={"values": {}}).get_json()["counts"]["manual"] == 0

    edited = {"columns": parts_value["columns"], "rows": [["PW48-1591", "ベアリング", "4"]]}
    res = client.post(f"/forms/{doc_id}/confirm", data={"value-parts": json.dumps(edited, ensure_ascii=False)})
    assert res.headers["Location"].endswith(f"/forms/{doc_id}/done")
    assert "品番: PW48-1591／品名: ベアリング／数量: 4" in client.get(f"/forms/{doc_id}").get_data(as_text=True)

    # 行を全部消しても列見出しと表の入力欄は残る（人が入れ直せる）。Markdown には行の無い明細表を出さない
    empty = {"columns": parts_value["columns"], "rows": []}
    client.post(f"/forms/{doc_id}/draft", json={"values": {"parts": json.dumps(empty, ensure_ascii=False)}})
    page = client.get(f"/forms/{doc_id}/review").get_data(as_text=True)
    assert "data-table-editor" in page and "行を追加" in page and "明細表が見つかりませんでした" not in page
    assert "## 交換部品" not in client.post(f"/forms/{doc_id}/preview", json={"values": {}}).get_json()["markdown"]

    # ダウンロードは確定済みの版から作る（最後の手順。ここでデータは消える）
    md = client.get(f"/forms/{doc_id}/download.md").get_data(as_text=True)
    assert "- 品番: PW48-1591／品名: ベアリング／数量: 4\n" in md and "スピンモータ" not in md
