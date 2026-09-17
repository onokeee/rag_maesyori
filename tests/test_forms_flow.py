"""帳票フローの画面の通しテスト（帳票の種類の作成 → 使用開始 → 取り込み → 確認 → 確定 → ダウンロード）。"""
import io
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

    md = client.get(f"/forms/{doc_id}/download.md")
    assert md.status_code == 200
    text = md.get_data(as_text=True)
    assert text.startswith("# 設備修理報告書 R2026-00123") and "CMP研磨装置" in text

    # 修正中 → 変更を破棄 → 確定済み。詳細・履歴・削除
    client.post(f"/forms/{doc_id}/draft", json={"values": {"equipment_name": "別の名前"}})
    with app.app_context():
        assert db.get_document(doc_id)["state"] == "modified"
    assert "別の名前" not in client.get(f"/forms/{doc_id}/download.md").get_data(as_text=True)
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

    res = client.post(f"/forms/{doc_id}/delete", data={"next": "/history/"})
    assert res.status_code == 302
    with app.app_context():
        assert db.get_document(doc_id) is None
