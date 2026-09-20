"""一覧表の DB アクセス（取り込み1件と、その取り込みが使う取り込み設定）。

取り込み設定は保存しない（利用者の指示 2026-09-20:「表の方には、取り込み設定を保持しておく機能はいらない」）。
取り込みごとに列の対応づけを決め、その設定（TableSpec）を取り込みの行（spec_json）に持つ。
接続は models.database.get_db()（リクエスト中もジョブの app_context 中も使える）。conn を渡せばそれを使う。
JSON の列は読み出し時に dict/list に直した値を別名（source, stats）で付け、spec_json は TableSpec にする。
"""
from __future__ import annotations

import json
import sqlite3

from models import database
from tables.spec import TableSpec, spec_from_dict, spec_hash, spec_json

IMPORT_JSON_COLUMNS = {"source_json": "source", "stats_json": "stats"}


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


# ---- 取り込み ------------------------------------------------------------------------------

def _decode_import(row) -> dict | None:
    if row is None:
        return None
    d = dict(row)
    for column, alias in IMPORT_JSON_COLUMNS.items():
        d[alias] = _loads(d.get(column), {})
    d["spec"] = _spec_or_none(d.get("spec_json"))
    return d


def create_import(file_name: str, file_hash: str, stored_path: str, source: dict | None = None, conn=None,
                  session_id: str | None = None) -> int:
    """取り込みを1件作る。session_id は置いたブラウザ（views.current_session_id）。

    template_id / template_version_id は取り込み自身の番号にそろえる（設定はもう無いが、AI整形の控え
    （ai_items）がこの番号で取り込みを束ねている。design.md 3.2）。
    """
    db = _db(conn)
    ts = database.now()
    cur = db.execute("""INSERT INTO table_imports (file_name, file_hash, stored_path, source_json, session_id,
                        status, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, 'uploaded', ?, ?)""",
                     (file_name, file_hash, stored_path, _dumps(source or {}), session_id, ts, ts))
    import_id = cur.lastrowid
    db.execute("UPDATE table_imports SET template_id = id, template_version_id = id WHERE id = ?", (import_id,))
    db.commit()
    return import_id


def get_import(import_id: int, conn=None) -> dict | None:
    return _decode_import(_db(conn).execute("SELECT * FROM table_imports WHERE id = ?", (import_id,)).fetchone())


def save_spec(import_id: int, spec: TableSpec, conn=None) -> None:
    """この取り込みが使う取り込み設定を保存する（列の対応づけを保存するたびに上書きする）。"""
    update_import(import_id, conn=conn, spec_json=spec_json(spec), spec_hash=spec_hash(spec))


def update_import(import_id: int, conn=None, commit: bool = True, **columns) -> None:
    """列を更新する。source/stats は dict のまま渡してよい（*_json に保存）。"""
    if not columns:
        return
    sets, args = [], []
    for name, value in columns.items():
        column = f"{name}_json" if name in ("source", "stats") else name
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


def list_imports(status: str | list | None = None, limit: int = 100, conn=None,
                 session_id: str | None = None) -> list[dict]:
    """取り込みの一覧。session_id を渡すとそのブラウザの分だけ（持ち主の分からない古い行は含む）。"""
    where, args = [], []
    if session_id:
        where.append("(session_id IS NULL OR session_id = ?)")
        args.append(session_id)
    if status:
        statuses = [status] if isinstance(status, str) else list(status)
        where.append(f"status IN ({','.join('?' * len(statuses))})")
        args += statuses
    sql = ("SELECT * FROM table_imports" + (f" WHERE {' AND '.join(where)}" if where else "")
           + " ORDER BY id DESC LIMIT ?")
    return [_decode_import(r) for r in _db(conn).execute(sql, (*args, limit)).fetchall()]
