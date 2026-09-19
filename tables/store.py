"""一覧表の DB アクセス（取り込み設定と版・取り込み）。

接続は models.database.get_db()（リクエスト中もジョブの app_context 中も使える）。conn を渡せばそれを使う。
JSON の列は読み出し時に dict/list に直した値を別名（source, period, stats, spec）で付ける。
"""
from __future__ import annotations

import json
import sqlite3

from core import purge
from models import database
from tables.spec import TableSpec, spec_from_dict, spec_hash, spec_json

IMPORT_JSON_COLUMNS = {"source_json": "source", "period_json": "period", "stats_json": "stats"}


def _db(conn=None) -> sqlite3.Connection:
    return conn if conn is not None else database.get_db()


def _loads(text, default):
    try:
        value = json.loads(text) if text else default
    except (TypeError, ValueError):
        return default
    return value if isinstance(value, type(default)) else default


def _dumps(value) -> str:
    return value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)


def _spec_or_none(spec_text):
    try:
        return spec_from_dict(_loads(spec_text, {})) if spec_text else None
    except ValueError:
        return None


# ---- 取り込み設定と版 -------------------------------------------------------------------
# 版の履歴は画面で扱わない。保存は今の版を上書きする（確定に使われた版だけ新しい版にする）。

def create_template(name: str, spec: TableSpec, description: str = "", note: str = "", conn=None) -> tuple[int, int]:
    """設定と版1を作る。戻り値: (template_id, version_id)"""
    db = _db(conn)
    ts = database.now()
    cur = db.execute("INSERT INTO table_templates (name, description, created_at, updated_at) VALUES (?, ?, ?, ?)",
                     (name, description or spec.description or "", ts, ts))
    template_id = cur.lastrowid
    cur = db.execute("INSERT INTO table_template_versions (template_id, version, spec_json, spec_hash, note, created_at) "
                     "VALUES (?, 1, ?, ?, ?, ?)", (template_id, spec_json(spec), spec_hash(spec), note, ts))
    version_id = cur.lastrowid
    db.execute("UPDATE table_templates SET current_version_id = ? WHERE id = ?", (version_id, template_id))
    db.commit()
    return template_id, version_id


def save_template_version(template_id: int, spec: TableSpec, note: str = "", conn=None) -> int:
    """設定を保存する。今の版が確定に使われていなければ上書き、使われていれば新しい版を作る。戻り値: version_id"""
    db = _db(conn)
    ts = database.now()
    current = db.execute("SELECT v.* FROM table_templates t JOIN table_template_versions v ON v.id = t.current_version_id "
                         "WHERE t.id = ?", (template_id,)).fetchone()
    if current is not None and not current["used"]:
        if current["spec_hash"] != spec_hash(spec):
            db.execute("UPDATE table_template_versions SET spec_json = ?, spec_hash = ?, note = ? WHERE id = ?",
                       (spec_json(spec), spec_hash(spec), note or current["note"], current["id"]))
        version_id = current["id"]
    else:
        last = db.execute("SELECT COALESCE(MAX(version), 0) FROM table_template_versions WHERE template_id = ?",
                          (template_id,)).fetchone()[0]
        cur = db.execute("INSERT INTO table_template_versions (template_id, version, spec_json, spec_hash, note, created_at) "
                         "VALUES (?, ?, ?, ?, ?, ?)", (template_id, last + 1, spec_json(spec), spec_hash(spec), note, ts))
        version_id = cur.lastrowid
    db.execute("UPDATE table_templates SET current_version_id = ?, name = ?, description = ?, updated_at = ? WHERE id = ?",
               (version_id, spec.name, spec.description or "", ts, template_id))
    db.commit()
    return version_id


def get_template(template_id: int, conn=None) -> dict | None:
    row = _db(conn).execute("""
        SELECT t.*, v.version, v.spec_json, v.spec_hash, v.used
        FROM table_templates t LEFT JOIN table_template_versions v ON v.id = t.current_version_id
        WHERE t.id = ?""", (template_id,)).fetchone()
    if row is None:
        return None
    d = dict(row)
    d["spec"] = _spec_or_none(d.get("spec_json"))
    return d


def list_templates(conn=None) -> list[dict]:
    rows = _db(conn).execute("""
        SELECT t.*, v.version, v.spec_json, v.spec_hash,
               (SELECT COUNT(*) FROM table_imports i WHERE i.template_id = t.id) AS import_count,  -- まだダウンロードしていない取り込み（ダウンロードで消える）
               (SELECT MAX(i.confirmed_at) FROM table_imports i WHERE i.template_id = t.id) AS last_confirmed_at
        FROM table_templates t LEFT JOIN table_template_versions v ON v.id = t.current_version_id
        ORDER BY t.name""").fetchall()
    out = []
    for row in rows:
        d = dict(row)
        d["spec"] = _spec_or_none(d.pop("spec_json", None))
        out.append(d)
    return out


def get_version(version_id: int, conn=None) -> dict | None:
    row = _db(conn).execute("SELECT * FROM table_template_versions WHERE id = ?", (version_id,)).fetchone()
    if row is None:
        return None
    d = dict(row)
    d["spec"] = spec_from_dict(_loads(d.get("spec_json"), {}))
    return d


def mark_version_used(version_id: int, conn=None) -> None:
    """確定に使った版は、以後の保存で上書きしない（その取り込みの結果を作った設定を残す）。"""
    db = _db(conn)
    db.execute("UPDATE table_template_versions SET used = 1 WHERE id = ?", (version_id,))
    db.commit()


def delete_template(template_id: int, conn=None) -> None:
    db = _db(conn)
    db.execute("DELETE FROM table_templates WHERE id = ?", (template_id,))
    db.execute("DELETE FROM ai_items WHERE template_id = ?", (template_id,))
    # AIの応答の控え（llm_calls）もどこからも使われなくなった分をすぐ消す（起動時まで残さない。design.md 3.3）。
    # sweep_orphan_ai が commit と DB の縮小までする
    purge.sweep_orphan_ai(db)


# ---- 取り込み ------------------------------------------------------------------------------

def _decode_import(row) -> dict | None:
    if row is None:
        return None
    d = dict(row)
    for column, alias in IMPORT_JSON_COLUMNS.items():
        d[alias] = _loads(d.get(column), {})
    return d


def create_import(file_name: str, file_hash: str, stored_path: str, source: dict | None = None,
                  template_id: int | None = None, template_version_id: int | None = None, conn=None) -> int:
    db = _db(conn)
    ts = database.now()
    cur = db.execute("""INSERT INTO table_imports (template_id, template_version_id, file_name, file_hash, stored_path,
                        source_json, status, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, 'uploaded', ?, ?)""",
                     (template_id, template_version_id, file_name, file_hash, stored_path, _dumps(source or {}), ts, ts))
    db.commit()
    return cur.lastrowid


def get_import(import_id: int, conn=None) -> dict | None:
    return _decode_import(_db(conn).execute("""
        SELECT i.*, t.name AS template_name FROM table_imports i LEFT JOIN table_templates t ON t.id = i.template_id
        WHERE i.id = ?""", (import_id,)).fetchone())


def update_import(import_id: int, conn=None, commit: bool = True, **columns) -> None:
    """列を更新する。source/period/stats は dict のまま渡してよい（*_json に保存）。"""
    if not columns:
        return
    sets, args = [], []
    for name, value in columns.items():
        column = f"{name}_json" if name in ("source", "period", "stats") else name
        if column.endswith("_json") and not isinstance(value, str):
            value = _dumps(value)
        sets.append(f"{column} = ?")
        args.append(value)
    sets.append("updated_at = ?")
    args.append(database.now())
    db = _db(conn)
    db.execute(f"UPDATE table_imports SET {', '.join(sets)} WHERE id = ?", (*args, import_id))
    if commit:
        db.commit()


# 取り込み1件を消すのは core/purge.py の purge_table_import（ファイルと DB の行をまとめて消す。design.md 3.3）


def list_imports(template_id: int | None = None, status: str | list | None = None, limit: int = 100, conn=None) -> list[dict]:
    where, args = [], []
    if template_id is not None:
        where.append("i.template_id = ?")
        args.append(template_id)
    if status:
        statuses = [status] if isinstance(status, str) else list(status)
        where.append(f"i.status IN ({','.join('?' * len(statuses))})")
        args += statuses
    sql = ("SELECT i.*, t.name AS template_name FROM table_imports i LEFT JOIN table_templates t ON t.id = i.template_id"
           + (f" WHERE {' AND '.join(where)}" if where else "") + " ORDER BY i.id DESC LIMIT ?")
    return [_decode_import(r) for r in _db(conn).execute(sql, (*args, limit)).fetchall()]
