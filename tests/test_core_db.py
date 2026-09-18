"""models/database.py: マイグレーション・PRAGMA・帳票のデータアクセス。"""
import json
import sqlite3

import pytest
from flask import Flask

from models import database as db
from pattern.model import FieldDef, PatternDef, SheetDef


def make_app(tmp_path) -> Flask:
    """DB だけを初期化した最小のアプリ（app.py の画面構成に依存しない）。"""
    app = Flask(__name__)
    app.config.update(TESTING=True, DATABASE=tmp_path / "app.db", UPLOAD_DIR=tmp_path / "uploads")
    db.init_app(app)
    return app


@pytest.fixture
def core_app(tmp_path):
    return make_app(tmp_path)


# 作り直し前（user_version=0）のスキーマ
OLD_SCHEMA = """
CREATE TABLE patterns (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, version TEXT NOT NULL DEFAULT 'v1',
    description TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'draft', image_processing TEXT NOT NULL DEFAULT 'none',
    created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE pattern_sheets (id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_id INTEGER NOT NULL REFERENCES patterns(id) ON DELETE CASCADE, sheet_name TEXT NOT NULL,
    required INTEGER NOT NULL DEFAULT 1);
CREATE TABLE pattern_fields (id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_id INTEGER NOT NULL REFERENCES patterns(id) ON DELETE CASCADE, sort_order INTEGER NOT NULL DEFAULT 0,
    field_name TEXT NOT NULL, display_name TEXT NOT NULL, candidates TEXT NOT NULL DEFAULT '[]',
    required INTEGER NOT NULL DEFAULT 0, data_type TEXT NOT NULL DEFAULT 'string',
    extraction_rule TEXT NOT NULL DEFAULT '{}');
CREATE TABLE pattern_samples (id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_id INTEGER NOT NULL REFERENCES patterns(id) ON DELETE CASCADE, file_name TEXT NOT NULL,
    file_hash TEXT NOT NULL, stored_path TEXT NOT NULL, created_at TEXT NOT NULL);
CREATE TABLE documents (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT NOT NULL, file_hash TEXT NOT NULL,
    stored_path TEXT NOT NULL, pattern_id INTEGER REFERENCES patterns(id) ON DELETE SET NULL,
    status TEXT NOT NULL DEFAULT 'uploaded', data_json TEXT, markdown TEXT, created_at TEXT NOT NULL,
    registered_at TEXT);
"""

DATA = json.dumps({"values": {"report_id": "R-1"}}, ensure_ascii=False)


def _make_old_db(path):
    conn = sqlite3.connect(path)
    conn.executescript(OLD_SCHEMA)
    ts = "2026-01-01T00:00:00"
    conn.execute("INSERT INTO patterns (name, status, created_at, updated_at) VALUES (?, 'active', ?, ?)",
                 ("設備修理報告書", ts, ts))
    conn.execute("INSERT INTO pattern_fields (pattern_id, field_name, display_name, candidates, extraction_rule) "
                 "VALUES (1, 'report_id', '報告番号', ?, ?)", ('["報告番号"]', '{"direction": "right"}'))
    conn.execute("INSERT INTO documents (file_name, file_hash, stored_path, pattern_id, status, data_json, markdown, "
                 "created_at, registered_at) VALUES (?, 'h1', 'documents/a.xlsx', 1, 'registered', ?, '# md', ?, ?)",
                 ("登録済み.xlsx", DATA, "2026-02-01T10:00:00", "2026-02-02T11:00:00"))
    conn.execute("INSERT INTO documents (file_name, file_hash, stored_path, pattern_id, status, data_json, created_at) "
                 "VALUES (?, 'h2', 'documents/b.xlsx', 1, 'extracted', ?, ?)", ("確認中.xlsx", DATA, "2026-03-01T10:00:00"))
    conn.execute("INSERT INTO documents (file_name, file_hash, stored_path, status, created_at) "
                 "VALUES (?, 'h3', 'documents/c.xlsx', 'uploaded', ?)", ("読み取り前.xlsx", "2026-04-01T10:00:00"))
    conn.commit()
    conn.close()


def test_migrates_old_db_with_registered_document(tmp_path):
    _make_old_db(tmp_path / "app.db")
    app = make_app(tmp_path)
    with app.app_context():
        conn = db.get_db()
        assert conn.execute("PRAGMA user_version").fetchone()[0] == len(db.MIGRATIONS)
        assert conn.execute("PRAGMA journal_mode").fetchone()[0].lower() == "wal"
        assert conn.execute("PRAGMA foreign_keys").fetchone()[0] == 1
        assert conn.execute("PRAGMA busy_timeout").fetchone()[0] == db.BUSY_TIMEOUT_MS
        tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        assert {"table_templates", "table_template_versions", "table_imports", "jobs", "llm_calls",
                "ai_items"} <= tables
        # 使っていない表は残さない（取り込みを指す列が無く、行が入ると purge の探索から漏れる。design.md 3.3）
        assert tables.isdisjoint({"table_outputs", "table_downloads", "table_template_samples", "alias_entries"})
        # AI整形の控えは取り込み単位で消せる（同じ設定の別の取り込みを巻き添えにしない）
        assert "import_id" in {r[1] for r in conn.execute("PRAGMA table_info(ai_items)")}

        doc = db.get_document(1)
        assert doc["state"] == "confirmed"
        assert doc["confirmed_json"] == DATA
        assert doc["confirmed_at"] == "2026-02-02T11:00:00"
        assert doc["title"] == ""
        assert db.get_document(2)["state"] == "reviewing"
        assert db.get_document(3)["state"] == "unread"

        pattern = db.load_pattern(1)
        assert pattern.title_fields == [] and pattern.md_options == {} and pattern.version_no == 1
        assert pattern.fields[0].unit == "" and pattern.fields[0].rag_output == "show"
        assert pattern.fields[0].direction == "right"
        assert db.list_patterns()[0]["document_count"] == 1
        assert db.find_confirmed_by_hash("h1")["id"] == 1

    # 2回目の起動でも壊れない（冪等）
    app2 = make_app(tmp_path)
    with app2.app_context():
        assert db.get_document(1)["state"] == "confirmed"
        assert db.get_db().execute("SELECT COUNT(*) FROM documents").fetchone()[0] == 3


def test_new_db_document_states_and_filters(core_app):
    with core_app.app_context():
        pid = db.create_pattern("設備修理報告書", "v1", "")
        d1 = db.create_document("修理報告書_標準.xlsx", "hash1", "forms/x.xlsx", pattern_id=pid)
        d2 = db.create_document("点検記録表.xlsx", "hash2", "forms/y.xlsx")
        assert db.get_document(d1)["state"] == "unread"

        db.save_draft(d1, {"values": {"report_id": "R2026-00123"}}, title="R2026-00123 CMP-101")
        assert db.get_document(d1)["state"] == "reviewing"
        assert db.find_confirmed_by_hash("hash1") is None
        assert db.confirm_document(d2) is False  # 読み取り前は確定できない

        assert db.confirm_document(d1, title="R2026-00123 CMP研磨装置1号機")
        doc = db.get_document(d1)
        assert doc["state"] == "confirmed" and doc["confirmed_at"]
        assert db.find_confirmed_by_hash("hash1")["id"] == d1
        assert db.find_confirmed_by_hash("hash1", exclude_id=d1) is None

        db.save_draft(d1, {"values": {"report_id": "R2026-00124"}})
        assert db.get_document(d1)["state"] == "modified"
        assert db.discard_changes(d1)
        doc = db.get_document(d1)
        assert doc["state"] == "confirmed"
        assert json.loads(doc["data_json"])["values"]["report_id"] == "R2026-00123"
        assert db.discard_changes(d2) is False

        # ホームの一覧（状態で絞る。取り込み履歴は無いので検索・絞り込みは持たない）
        assert [r["id"] for r in db.list_documents()] == [d2, d1]
        rows = db.list_documents(state="confirmed")
        assert len(rows) == 1 and rows[0]["pattern_name"] == "設備修理報告書" and rows[0]["state"] == "confirmed"
        assert len(db.list_documents(state=["unread", "reviewing"])) == 1
        assert [r["id"] for r in db.list_documents(limit=1)] == [d2]
        assert [d["id"] for d in db.list_confirmed_documents()] == [d1]
        assert db.list_confirmed_documents([]) == []

        # まとめ取り込み（同じ batch_id を選んだ順に返す）
        b1 = db.create_document("1.xlsx", "h1", "documents/1.xlsx", batch_id="B", batch_order=0)
        b2 = db.create_document("2.xlsx", "h2", "documents/2.xlsx", batch_id="B", batch_order=1)
        assert [d["id"] for d in db.list_batch_documents("B")] == [b1, b2]
        assert db.list_batch_documents("") == [] and db.list_batch_documents("ない") == []
        assert db.get_document(b1)["batch_id"] == "B"


def test_save_pattern_new_fields_and_version_no(core_app):
    with core_app.app_context():
        pid = db.create_pattern("設備修理報告書")
        fd = FieldDef("work_hours", "作業時間", ["作業時間"], data_type="number")
        fd.unit = "時間"
        fd.rag_output = "omit"
        pattern = PatternDef(name="設備修理報告書", id=pid, sheets=[SheetDef("修理報告書")], fields=[fd])
        pattern.title_fields = ["report_id", "equipment"]
        pattern.md_options = {"omit_person_fields": True}
        db.save_pattern(pattern, "active")
        assert pattern.version_no == 2

        loaded = db.load_pattern(pid)
        assert loaded.title_fields == ["report_id", "equipment"]
        assert loaded.md_options == {"omit_person_fields": True}
        assert loaded.version_no == 2 and loaded.status == "active"
        assert loaded.fields[0].unit == "時間" and loaded.fields[0].rag_output == "omit"

        # 新しい属性を持たない PatternDef / FieldDef でも保存できる
        db.save_pattern(PatternDef(name="別名", id=pid, fields=[FieldDef("a", "A")]), "draft")
        loaded = db.load_pattern(pid)
        assert loaded.version_no == 3 and loaded.title_fields == [] and loaded.fields[0].rag_output == "show"


def test_foreign_keys_cascade_with_background_connection(core_app):
    with core_app.app_context():
        conn = db.connect()
        try:
            ts = db.now()
            tid = conn.execute("INSERT INTO table_templates (name, created_at, updated_at) VALUES (?, ?, ?)",
                               ("故障履歴", ts, ts)).lastrowid
            conn.execute("INSERT INTO table_template_versions (template_id, version, spec_json, spec_hash, created_at) "
                         "VALUES (?, 1, '{}', 'x', ?)", (tid, ts))
            iid = conn.execute("INSERT INTO table_imports (template_id, file_name, file_hash, stored_path, created_at, "
                               "updated_at) VALUES (?, 'a.csv', 'h', 'tables/a.csv', ?, ?)", (tid, ts, ts)).lastrowid
            conn.commit()
            conn.execute("DELETE FROM table_templates WHERE id = ?", (tid,))
            conn.commit()
            assert conn.execute("SELECT COUNT(*) FROM table_template_versions").fetchone()[0] == 0
            assert conn.execute("SELECT template_id FROM table_imports WHERE id = ?", (iid,)).fetchone()[0] is None
        finally:
            conn.close()
