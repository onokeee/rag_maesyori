"""AI まわり（5巡目）：レート制限が続くときの一時停止、方式判定の出力上限、帳票AIのクライアント・応答の読み取り、
一時停止を頼んだ直後の画面。"""
import pytest

from aiproc import items, runner
from core import jobs
from models import database
from services import ai_assist, llm
from tests.fake_servers import Reply, keep_scenario, segments_of
from tests.test_aiproc import ROWS, _make_import, _wait, ai_app, fake  # noqa: F401  (fixture)


# ---- R5-AI-1 レート制限が続いたら、行を次々エラーにせず一時停止する ------------------------------------

def test_persistent_rate_limit_pauses_job_instead_of_failing_rows(ai_app, fake, monkeypatch):
    monkeypatch.setattr(llm, "LOCAL_TIMEOUT", 0.5)       # 行の上限は1秒
    limited = {"on": True}

    def responder(body, srv):
        if segments_of(body) and limited["on"]:
            return Reply(status=429, body={"error": {"message": "Rate limit reached"}}, headers={"retry-after": "0.3"})
        return keep_scenario(body, srv)

    fake.responder = responder
    rows = {k: ROWS[k] for k in ("R1", "R2", "R3", "R4", "R5")}
    iid = _make_import(ai_app, rows=rows)
    with ai_app.app_context():
        job_id = runner.start_ai_job(iid, concurrency=1, stage_ids=["log"])
        job = _wait(lambda: (j := jobs.get_job(job_id))["status"] == "paused" and j, timeout=30)
        assert "混み合っています" in job["message"] and "再開" in job["message"]
        # 打ち切りになった行はエラーとして保存せず、未処理に戻している
        assert not [v for v in items.items_by_key(1, "log", import_id=iid).values() if v["status"] == "error"]
        sent = sum(1 for r in fake.chat_requests() if segments_of(r["body"]))
        assert sent <= runner.RATE_LIMIT_PAUSE_ROWS * 5      # 残りの行に送り続けない

        limited["on"] = False                                # 混雑が解けてから再開する
        jobs.request_resume(job_id)
        done = jobs.wait_job(job_id, timeout=60)
        assert done["status"] == "done", done["message"]
        assert done["message"] == ""                         # 案内は再開で消える
        got = items.items_by_key(1, "log", import_id=iid)
        assert len(got) == len(rows)
        assert not [v for v in got.values() if "混み合っています" in (v["error"] or "")]


def test_rate_limited_rows_are_errors_when_other_rows_succeed(ai_app, fake, monkeypatch):
    """一時的な混雑（他の行は通る）なら、打ち切りの行はこれまでどおりエラーにして最後まで進む。"""
    monkeypatch.setattr(llm, "LOCAL_TIMEOUT", 0.5)

    def responder(body, srv):
        if segments_of(body) and "ID抜け" in str(body.get("messages")):
            return Reply(status=429, body={"error": {"message": "Rate limit reached"}}, headers={"retry-after": "0.3"})
        return keep_scenario(body, srv)

    fake.responder = responder
    rows = {k: ROWS[k] for k in ("R1", "R5", "R2")}
    iid = _make_import(ai_app, rows=rows)
    with ai_app.app_context():
        job = jobs.wait_job(runner.start_ai_job(iid, concurrency=1, stage_ids=["log"]), timeout=60)
        assert job["status"] == "done"
        got = items.items_by_key(1, "log", import_id=iid)
        assert got["R5"]["status"] == "error" and "混み合っています" in got["R5"]["error"]
        assert got["R1"]["status"] != "error" and got["R2"]["status"] != "error"


# ---- R5-AI-2 方式判定は推論モデルでも使い切らない上限で試し、打ち切りは覚えない ------------------------

def test_detect_mode_with_reasoning_model_that_truncates_small_budgets(ai_app, fake):
    def responder(body, srv):
        budget = body.get("max_tokens") or body.get("max_completion_tokens") or 10 ** 6
        if budget <= 200:                                    # 小さい上限は考える途中で使い切る
            return Reply(content="", finish_reason="length")
        return keep_scenario(body, srv)

    fake.responder = responder
    with ai_app.app_context():
        s = llm.job_client_settings()
        assert llm.detect_structured_mode(s) == "json_schema"
        assert llm._MODES[(s["chat_url"], "gpt-test")] == "json_schema"


def test_detect_mode_truncated_every_time_is_not_remembered(ai_app, fake):
    fake.responder = lambda body, srv: Reply(content="<think>考え中", finish_reason="length")
    with ai_app.app_context():
        s = llm.job_client_settings()
        before = len(fake.chat_requests())
        assert llm.detect_structured_mode(s) == "json_schema"    # 指定は受け付けた（400 ではない）
        sent = fake.chat_requests()[before:]
        assert [r["body"].get("max_tokens") for r in sent] == list(llm._PROBE_BUDGETS)
        assert (s["chat_url"], "gpt-test") not in llm._MODES     # 次回また判定する


# ---- R5-AI-3 帳票AIのクライアントは明示タイムアウト・SDK の再試行なし ------------------------------------

def test_ask_json_client_has_finite_timeout_and_no_sdk_retries(ai_app, fake):
    with ai_app.app_context():
        cli = llm.client()
        assert cli.max_retries == 0 and cli.timeout == llm.LOCAL_TIMEOUT     # 偽サーバーは 127.0.0.1
        assert llm.models_client().max_retries == 0 and llm.models_client().timeout == llm.MODELS_TIMEOUT
        fake.responder = lambda body, srv: Reply(status=500, body={"error": {"message": "boom"}})
        with pytest.raises(Exception):
            llm.ask_json("system", "user")
        assert len(fake.chat_requests()) == 1                  # SDK が投げ直さない


# ---- R5-AI-5 ask_json は <think> を除いてから読む。読めなければ日本語 ------------------------------------

def test_ask_json_strips_think_blocks_and_errors_in_japanese(ai_app, fake):
    with ai_app.app_context():
        fake.chat_replies = ['<think>候補は {"values": ...} かな</think>\n{"values": {"a": null}}']
        assert llm.ask_json("s", "u") == {"values": {"a": None}}
        fake.chat_replies = ['{"values": {"a": ']
        with pytest.raises(ValueError) as e:
            llm.ask_json("s", "u", what="項目の補完")
        assert str(e.value).startswith("項目の補完をJSONとして解析できませんでした") and "Expecting" not in str(e.value)
        assert llm.friendly_error(e.value) == str(e.value)


# ---- R5-AI-6 fill_missing は想定外の型を捨て、途中で失敗しても一部だけ入力済みにしない -----------------------

class _Info:
    grids = {"Sheet1": object()}
    date1904 = False


def _field(name):
    return {"field_name": name, "display_name": name, "data_type": "text", "value": None}


def test_fill_missing_ignores_odd_types_from_ai(monkeypatch, app):
    reply = {"values": {"a": {"value": "山田", "sheet": ["Sheet1"], "cell": {"x": 1}},
                        "b": {"value": ["x"], "sheet": "Sheet1", "cell": "B2"},
                        "c": {"value": "3号機", "sheet": "Sheet1", "cell": " C3 "}}}
    monkeypatch.setattr(llm, "ask_json", lambda *a, **k: reply)
    monkeypatch.setattr(llm, "current_model", lambda: "gpt-test")
    monkeypatch.setattr(ai_assist, "_dump_cells", lambda info, sheets: [])
    extraction = {"fields": [_field("a"), _field("b"), _field("c")], "sheets": ["Sheet1"]}
    with app.app_context():
        assert ai_assist.fill_missing(_Info(), extraction) == ["a", "c"]
    a, b, c = extraction["fields"]
    assert a["value"] == "山田" and a["sheet"] is None and a["value_cell"] is None and a["ai_filled"]
    assert b["value"] is None and "ai_filled" not in b
    assert c["sheet"] == "Sheet1" and c["value_cell"] == "C3"


def test_fill_missing_leaves_no_partial_fill_when_it_fails(monkeypatch, app):
    reply = {"values": {"a": {"value": "山田"}, "b": {"value": "2026/13/45"}}}
    monkeypatch.setattr(llm, "ask_json", lambda *a, **k: reply)
    monkeypatch.setattr(llm, "current_model", lambda: "gpt-test")
    monkeypatch.setattr(ai_assist, "_dump_cells", lambda info, sheets: [])

    def boom(*a, **k):
        raise RuntimeError("想定外")
    monkeypatch.setattr(ai_assist, "to_date", boom)
    extraction = {"fields": [_field("a"), dict(_field("b"), data_type="date")], "sheets": ["Sheet1"]}
    with app.app_context(), pytest.raises(RuntimeError):
        ai_assist.fill_missing(_Info(), extraction)
    assert all(f["value"] is None and "ai_filled" not in f for f in extraction["fields"])


# ---- R5-AI-4 一時停止を頼んだ直後（まだ実行中）でも［再開］を出す ------------------------------------------

def test_ai_page_shows_resume_while_pause_is_pending(ai_app):
    iid = _make_import(ai_app, rows={"R1": ROWS["R1"]})
    with ai_app.app_context():
        conn = database.connect()
        now = database.now()
        conn.execute("INSERT INTO jobs (kind, ref_type, ref_id, status, pause_requested, heartbeat_at, created_at, "
                     "updated_at) VALUES ('ai_format', 'table_import', ?, 'running', 1, ?, ?, ?)", (iid, now, now, now))
        conn.commit()
        conn.close()
    page = ai_app.test_client().get(f"/tables/imports/{iid}/ai").get_data(as_text=True)
    assert "/ai/resume" in page and "/ai/pause" not in page and "一時停止中…" in page
