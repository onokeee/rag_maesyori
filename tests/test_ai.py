
import pytest

from app import create_app
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
    """AI接続の設定（表の取り込みの画面の中にある欄）を保存する。"""
    payload = {"models": ["gpt-test", "picky-model"], "default": "gpt-test", "api_key": OPENAI_KEY,
               "chat_url": "", "models_url": "", "add_models": ""}
    payload.update(overrides)
    res = client.post("/tables/ai-connection", json=payload)
    assert res.status_code == 200 and res.get_json()["ok"] is True
    return res


def test_unsupported_parameter_is_dropped_and_retried(ai_client, ai_app, fake):
    _save_ai_settings(ai_client, default="picky-model")
    fake.chat_replies = ['```json\n{"answer": 42}\n```']
    with ai_app.app_context():
        assert llm.ask_json("system", "user") == {"answer": 42}
    chats = [r for r in fake.requests if r["path"] == "/v1/chat/completions"]
    assert len(chats) == 2
    assert "temperature" in chats[0]["body"] and "temperature" not in chats[1]["body"]
