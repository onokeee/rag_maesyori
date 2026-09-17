"""画面（blueprint）共通の小物。"""
from __future__ import annotations

import sqlite3

from flask import request, url_for

from models import database


def safe_next(default: str) -> str:
    """フォームの next パラメータ（同一サイト内の相対パスのみ許可）"""
    target = request.form.get("next", "") or request.args.get("next", "")
    return target if target.startswith("/") and not target.startswith("//") else default


# ---- DB の読み取り（ホーム・履歴の一覧用の直接照会。表が無ければ空） ------------

def query_all(sql: str, args=()) -> list[dict]:
    try:
        return [dict(r) for r in database.get_db().execute(sql, args).fetchall()]
    except sqlite3.OperationalError:
        return []


def query_value(sql: str, args=(), default=0):
    try:
        row = database.get_db().execute(sql, args).fetchone()
    except sqlite3.OperationalError:
        return default
    return row[0] if row else default


# ---- 帳票の状態 --------------------------------------------------------------------
# models/database.py の導出値: unread / reviewing / modified / confirmed

FORM_STATES = database.DOCUMENT_STATES


def form_link(doc: dict) -> str:
    """帳票の状態に応じた「続き」の行き先（design.md 2.3 のルート）。"""
    state, doc_id = doc.get("state"), doc["id"]
    if state == "unread":
        return url_for("forms.type_select", doc_id=doc_id)
    if state in ("reviewing", "modified"):
        return url_for("forms.review", doc_id=doc_id)
    return url_for("forms.detail", doc_id=doc_id)


# ---- 一覧表の取り込みの状態 ------------------------------------------------------------

TABLE_IMPORT_STATES = {"uploaded": "読み込み前", "reading": "読み込み中", "preview": "確認中", "confirming": "確定処理中",
                       "confirmed": "確定済み", "failed": "失敗"}
TABLE_IMPORT_ACTIVE = ("uploaded", "reading", "preview", "confirming")


def table_import_link(item: dict) -> str:
    status, import_id = item.get("status"), item["id"]
    step = {"uploaded": "source", "reading": "preview", "preview": "preview",
            "confirming": "done", "confirmed": "done"}.get(status, "source")
    return url_for(f"tables.{step}", import_id=import_id)
