"""AI まわり（6巡目の修正）:
- max_tokens を受け付けないモデルでは「名前の付け替え」を覚える（1回目の値で固定しない）
- 見積もりは DB を開き直さない（行数が増えても接続回数は変わらない）
- 基準日の探し方が AI整形と Markdown で同じ
- AI接続が外れても、動いているジョブの［中止］は画面に残る
"""
import json

import pytest

from aiproc import estimate, runner
from app import create_app
from core import jobs
from models import database
from services import llm
from tables import markdown as tmd
from tables.spec import spec_from_dict
from tests.conftest import BufferedClient, make_config
from tests.fake_servers import Reply
from tests.test_aiproc import SPEC, ai_app, fake, _write_rows  # noqa: F401  (fixture)


# ---- R6-AI-1 max_tokens → max_completion_tokens は「名前の付け替え」として覚える ----------------

_MAX_TOKENS_400 = {"error": {
    "message": "Unsupported parameter: 'max_tokens' is not supported with this model. "
               "Use 'max_completion_tokens' instead.",
    "type": "invalid_request_error", "param": "max_tokens", "code": "unsupported_parameter"}}


def _reject_max_tokens(body, server):
    if "max_tokens" in body:
        return Reply(status=400, body=_MAX_TOKENS_400)
    return Reply(content='{"ok": true}')


def test_max_tokens_quirk_keeps_each_call_budget(ai_app, fake):
    fake.responder = _reject_max_tokens
    with ai_app.app_context():
        s = llm.job_client_settings()
        llm.chat_raw(s, [{"role": "user", "content": "x"}], max_tokens=111)   # 接続テストのような小さい上限
        llm.chat_raw(s, [{"role": "user", "content": "x"}], max_tokens=2000)  # 本番の行（大きい上限）
        sent = [r["body"] for r in fake.chat_requests()]
        # 1回目は max_tokens で拒否 → 付け替えて投げ直し。2回目以降はその呼び出し自身の値を送る
        assert [b.get("max_tokens") for b in sent] == [111, None, None]
        assert [b.get("max_completion_tokens") for b in sent] == [None, 111, 2000]


def test_reset_llm_client_forgets_quirks(ai_app, fake):
    fake.responder = _reject_max_tokens
    with ai_app.app_context():
        llm.chat_raw(llm.job_client_settings(), [{"role": "user", "content": "x"}], max_tokens=50)
        assert llm._QUIRKS                       # 覚えている
        assert list(llm._QUIRKS)[0][0].endswith("/chat/completions")   # 接続先×モデルで覚える
        llm.reset_llm_client()                   # 接続先やキーを保存し直したとき
        assert llm._QUIRKS == {}                 # 再起動しなくても忘れる


# ---- R6-AI-2 見積もりは行数に関係なく DB を開く回数が変わらない --------------------------------

def _log_rows(n: int) -> dict:
    return {f"R{i}": (f"4/{i % 27 + 1} 田中：点検{i}の連絡あり。停止を確認。\n"
                      f"4/{i % 27 + 2} 佐藤：部品を交換して復旧{i}を確認。", "") for i in range(1, n + 1)}


def _count_connects(app, monkeypatch, iid) -> tuple[int, dict]:
    real = database.connect
    calls = []

    def counted(*a, **kw):
        calls.append(1)
        return real(*a, **kw)

    monkeypatch.setattr(database, "connect", counted)
    try:
        with app.app_context():
            result = estimate.estimate(iid, [], scope="pending", settings=llm.job_client_settings(),
                                       stage_ids=["log"])
    finally:
        monkeypatch.setattr(database, "connect", real)
    return len(calls), result


def _import_named(app, name: str, rows: dict) -> int:
    """取り込み設定の名前は一意なので、2つの取り込みを作るために名前を変えて登録する。"""
    spec = dict(SPEC, name=name)
    with app.app_context():
        conn = database.connect()
        now = database.now()
        try:
            tid = conn.execute("INSERT INTO table_templates (name, created_at, updated_at) VALUES (?, ?, ?)",
                               (name, now, now)).lastrowid
            vid = conn.execute("INSERT INTO table_template_versions (template_id, version, spec_json, spec_hash,"
                               " created_at) VALUES (?, 1, ?, 'h', ?)",
                               (tid, json.dumps(spec, ensure_ascii=False), now)).lastrowid
            iid = conn.execute("INSERT INTO table_imports (template_id, template_version_id, file_name, file_hash,"
                               " stored_path, status, created_at, updated_at)"
                               " VALUES (?, ?, 'T1.xlsx', 'x', 'x', 'preview', ?, ?)", (tid, vid, now, now)).lastrowid
            conn.commit()
        finally:
            conn.close()
        _write_rows(app, iid, rows)
    return iid


def test_estimate_opens_db_a_constant_number_of_times(ai_app, fake, monkeypatch):
    small = _import_named(ai_app, "少ない方", _log_rows(2))
    big = _import_named(ai_app, "多い方", _log_rows(12))
    n_small, est_small = _count_connects(ai_app, monkeypatch, small)
    n_big, est_big = _count_connects(ai_app, monkeypatch, big)
    assert est_small["ai_rows"] == 2 and est_big["ai_rows"] == 12   # キャッシュ照会が行数分あること
    assert n_small == n_big and n_big <= 5


# ---- R6-AI-3 基準日の探し方を AI整形と Markdown でそろえる ------------------------------------

def _spec_two_dates():
    data = json.loads(json.dumps(SPEC))
    data["columns"].insert(2, {"key": "repaired_at", "display": "復旧日", "type": "date", "role": "date"})
    return spec_from_dict(data)


def _row(occurred_at: str, repaired_at: str) -> dict:
    return {"key": "R1", "values": {
        "record_no": "R1", "occurred_at": occurred_at, "repaired_at": repaired_at,
        "equipment_name": "搬送ロボット2号機", "symptom": "停止", "cause": "",
        "response_log": "4/1 10:00 田中：ライン停止。\n翌週 佐藤：ケーブル交換済。\n4/19 佐藤：再発なし→クローズ",
        "worker": "佐藤"}, "originals": {}, "source": {}}


def _whens(parse) -> list:
    return [(s.id, s.when.date if s.when else None) for s in parse.segments]


@pytest.mark.parametrize("occurred_at, repaired_at", [("", "2024-04-05"), ("2024/04/05", "")])
def test_base_date_same_for_ai_and_markdown(occurred_at, repaired_at):
    spec = _spec_two_dates()
    row = _row(occurred_at, repaired_at)
    data = runner.ImportData(1, 1, 1, spec, [row])
    people = runner.people_index(data, spec.log_stage)
    work = runner._prepare_log(row, data, spec.log_stage, people)
    md_parse = tmd.parse_log_cell(spec, row["values"])
    assert _whens(work.parse) == _whens(md_parse)
    assert [w for _, w in _whens(md_parse) if w] and all(
        w.startswith("2024-") for _, w in _whens(md_parse) if w)


# ---- ux6-1 AI接続が外れても、動いているジョブの［中止］は残る ----------------------------------

@pytest.fixture
def plain_app(tmp_path):
    """AI接続が未設定（APIキーなし）のアプリ。"""
    llm.reset_llm_client()
    app = create_app(make_config(tmp_path))
    app.test_client_class = BufferedClient
    yield app
    llm.reset_llm_client()


def _read_csv_import(app, client) -> int:
    from tests.tables_helpers import imported

    return imported(app, client, "a.csv", "T", ai_role="log")


def test_paused_ai_job_can_still_be_cancelled_without_ai_settings(plain_app):
    client = plain_app.test_client()
    import_id = _read_csv_import(plain_app, client)
    with plain_app.app_context():
        assert not llm.is_configured()
        conn = database.connect()
        try:
            conn.execute("INSERT INTO jobs (kind, ref_type, ref_id, status, params_json, progress_json,"
                         " created_at, updated_at) VALUES ('ai_format', 'table_import', ?, 'paused', '{}', '{}', ?, ?)",
                         (import_id, database.now(), database.now()))
            conn.commit()
        finally:
            conn.close()
    from tests.tables_helpers import panel_html

    html = panel_html(client, import_id, "ai")
    assert "APIキーが設定されていません" in html            # AI接続のパネルは「未設定」と出る
    assert 'data-ai-control="cancel"' in html and "中止" in html
    # 接続が無いので実行・再開はできない（中止だけ残る）
    assert 'data-ai-control="resume"' not in html and 'data-ai-control="pause"' not in html
