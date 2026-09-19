import json
import re

import pytest
import yaml

from app import create_app
from excel.workbook import load_workbook_info
from models import database as db
from pattern.builder import suggest_rows
from pattern.forms import rows_to_pattern
from services import llm
from tests.conftest import make_config
from tests.fake_servers import OPENAI_KEY, FakeServer


@pytest.fixture
def fake():
    with FakeServer() as server:
        yield server


@pytest.fixture
def ai_app(tmp_path, fake):
    return create_app(make_config(tmp_path, OPENAI_BASE_URL=f"{fake.url}/v1", OPENAI_API_KEY="",
                                  OPENAI_MODELS=[], OPENAI_MODEL="gpt-test"))


@pytest.fixture
def ai_client(ai_app):
    return ai_app.test_client()


def _save_ai_settings(client, **overrides):
    form = {"models": ["gpt-test", "picky-model"], "default": "gpt-test", "api_key": OPENAI_KEY,
            "chat_url": "", "models_url": "", "add_models": ""}
    form.update(overrides)
    return client.post("/settings/ai", data=form, follow_redirects=True)


def test_unsupported_parameter_is_dropped_and_retried(ai_client, ai_app, fake):
    _save_ai_settings(ai_client, default="picky-model")
    fake.chat_replies = ['```json\n{"answer": 42}\n```']
    with ai_app.app_context():
        assert llm.ask_json("system", "user") == {"answer": 42}
    chats = [r for r in fake.requests if r["path"] == "/v1/chat/completions"]
    assert len(chats) == 2
    assert "temperature" in chats[0]["body"] and "temperature" not in chats[1]["body"]


def _create_pattern(app, sample_dir, extra_field=True):
    info = load_workbook_info(sample_dir / "standard.xlsx")
    sheet_rows, field_rows = suggest_rows([info])
    if extra_field:
        field_rows.append({"use": True, "field_name": "checker", "display_name": "確認者", "candidates": "確認者",
                           "data_type": "string", "required": False, "direction": "auto"})
    with app.app_context():
        pattern_id = db.create_pattern("設備修理報告書", "v1", "")
        pattern = rows_to_pattern(pattern_id, {"name": "設備修理報告書", "version": "v1", "description": "",
                                               "image_processing": "none"}, sheet_rows, field_rows)
        db.save_pattern(pattern, "active")
    return pattern_id


def _upload(client, sample_dir, name="standard.xlsx"):
    with open(sample_dir / name, "rb") as f:
        res = client.post("/forms/upload", data={"file": (f, "修理報告書_標準.xlsx")},
                          content_type="multipart/form-data")
    return int(re.search(r"/forms/(\d+)/type", res.headers["Location"])[1])


def test_ai_classify_and_fill(ai_client, ai_app, fake, sample_dir):
    _save_ai_settings(ai_client)
    pattern_id = _create_pattern(ai_app, sample_dir)
    doc_id = _upload(ai_client, sample_dir)

    fake.chat_replies = [json.dumps({"pattern_id": pattern_id, "sheets": ["修理報告書", "存在しないシート"],
                                     "confidence": 88, "reason": "項目構成が一致"})]
    page = ai_client.post(f"/forms/{doc_id}/ai-classify").get_data(as_text=True)
    assert "AIの推定（gpt-test）" in page and "項目構成が一致" in page
    ai_request = fake.requests[-1]["body"]["messages"][1]["content"]
    assert "修理報告書!A4: 設備番号" in ai_request  # セルの内容をAIに渡している

    ai_client.post(f"/forms/{doc_id}/read", data={"pattern_id": pattern_id, "sheets": ["修理報告書"]})
    page = ai_client.get(f"/forms/{doc_id}/review").get_data(as_text=True)
    assert "確定してMarkdownを作成" in page

    fake.chat_replies = [json.dumps({"values": {"checker": {"value": "田中", "sheet": "修理報告書", "cell": "C5"}}})]
    page = ai_client.post(f"/forms/{doc_id}/ai-fill", data={"value-reporter": "山田 花子"},
                          follow_redirects=True).get_data(as_text=True)
    assert "AIが1項目を入力しました" in page and "AIが入力（要確認）" in page
    with ai_app.app_context():
        extraction = json.loads(db.get_document(doc_id)["data_json"])
    assert extraction["values"]["checker"] == "田中"
    assert extraction["values"]["reporter"] == "山田 花子"  # 画面で入力中の修正は保持




def test_ai_filled_numbers_use_the_same_unit_rule_as_reading(ai_app, monkeypatch):
    """AIが入れた数値の単位も、読み取りと同じ決め方（種類の設定→書かれた値。勝手に補わない）にする。"""
    from services import ai_assist

    class _Info:
        date1904 = False
        grids: dict = {}

    def fields():
        return [{"field_name": "downtime", "display_name": "停止時間", "data_type": "number",
                 "value": None, "unit": ""},
                {"field_name": "repair_cost", "display_name": "修理費用", "data_type": "number",
                 "value": None, "unit": ""},
                {"field_name": "part_count", "display_name": "交換部品数", "data_type": "number",
                 "value": None, "unit": ""}]

    monkeypatch.setattr(llm, "ask_json", lambda *a, **k: {"values": {
        "downtime": {"value": "390分"}, "repair_cost": {"value": "12000"}, "part_count": {"value": "3"}}})
    extraction = {"sheets": [], "fields": fields()}
    with ai_app.app_context():
        assert ai_assist.fill_missing(_Info(), extraction) == ["停止時間", "修理費用", "交換部品数"]
    by_name = {f["field_name"]: f for f in extraction["fields"]}
    # 書かれた単位を取り込む（md は「390分」になる）。単位だけの読み落としの警告は出さない
    assert by_name["downtime"]["value"] == 390 and by_name["downtime"]["unit"] == "分"
    assert "数値の部分だけ" not in (by_name["downtime"]["warning"] or "")
    # どこにも単位が無く、単位で意味が変わる項目は要確認
    assert by_name["repair_cost"]["unit"] == "" and "単位が書かれていません" in by_name["repair_cost"]["warning"]
    # 件数・個数は単位不明で警告しない（AIが入力したことの案内だけ）
    assert by_name["part_count"]["unit"] == "" and "単位" not in by_name["part_count"]["warning"]


def test_fill_missing_does_not_send_table_fields(monkeypatch):
    """明細表の項目は AI の補完の対象にしない（行と列の形があるため、確認画面で人が入力する）。"""
    from services import ai_assist

    def fail(*args, **kwargs):
        raise AssertionError("AI を呼ばない")

    monkeypatch.setattr(llm, "ask_json", fail)
    extraction = {"sheets": [], "fields": [{"field_name": "parts", "display_name": "交換部品", "data_type": "table",
                                            "value": None, "unit": ""}]}
    assert ai_assist.fill_missing(None, extraction) == []


def test_ai_filled_number_is_judged_against_the_unit_set_on_the_form_type(ai_app, monkeypatch):
    """読み取りで unit が書かれた単位（時間）に置き換わっていても、AIの値は種類で決めた単位（分）で判定する。"""
    from services import ai_assist

    class _Info:
        date1904 = False
        grids: dict = {}

    monkeypatch.setattr(llm, "ask_json", lambda *a, **k: {"values": {"downtime": {"value": "894"}}})
    extraction = {"sheets": [], "fields": [{"field_name": "downtime", "display_name": "停止時間",
                                            "data_type": "number", "value": None, "unit": "時間",
                                            "spec_unit": "分"}]}
    with ai_app.app_context():
        assert ai_assist.fill_missing(_Info(), extraction) == ["停止時間"]
    f = extraction["fields"][0]
    assert f["value"] == 894 and f["unit"] == "分"
