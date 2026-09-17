"""ホーム: 2つの入口、作業中の一覧、最近の確定、はじめての案内。"""
from __future__ import annotations

from flask import Blueprint, render_template

from models import database
from views import TABLE_IMPORT_ACTIVE, form_link, query_all, query_value, table_import_link

bp = Blueprint("home", __name__)

RECENT_LIMIT = 5
WORKING_LIMIT = 10


def _forms_working() -> list[dict]:
    rows, _ = database.list_documents(state=("unread", "reviewing", "modified"), limit=WORKING_LIMIT)
    for row in rows:
        row["href"] = form_link(row)
    return rows


def _forms_recent() -> list[dict]:
    rows, _ = database.list_documents(state="confirmed", limit=RECENT_LIMIT)
    for row in rows:
        row["href"] = form_link(row)
    return rows


def _tables(statuses: tuple[str, ...], limit: int, order: str) -> list[dict]:
    marks = ",".join("?" * len(statuses))
    rows = query_all(f"""
        SELECT i.id, i.file_name, i.status, i.created_at, i.updated_at, i.confirmed_at, i.template_id,
               t.name AS template_name
        FROM table_imports i LEFT JOIN table_templates t ON t.id = i.template_id
        WHERE i.status IN ({marks}) ORDER BY {order} LIMIT ?
    """, (*statuses, limit))
    for row in rows:
        row["href"] = table_import_link(row)
    return rows


@bp.get("/")
def index():
    pattern_count = query_value("SELECT COUNT(*) FROM patterns")
    template_count = query_value("SELECT COUNT(*) FROM table_templates")
    return render_template(
        "home.html",
        first_run=pattern_count == 0 and template_count == 0,
        pattern_count=pattern_count,
        active_pattern_count=query_value("SELECT COUNT(*) FROM patterns WHERE status = 'active'"),
        template_count=template_count,
        forms_working=_forms_working(),
        tables_working=_tables(TABLE_IMPORT_ACTIVE, WORKING_LIMIT, "i.id DESC"),
        forms_recent=_forms_recent(),
        tables_recent=_tables(("confirmed",), RECENT_LIMIT, "i.confirmed_at DESC, i.id DESC"),
    )
