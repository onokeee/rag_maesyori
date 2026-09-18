"""ホーム: 2つの入口、作業中の一覧、ダウンロード待ちの一覧、はじめての案内。

データを残さない方針（design.md 3.3）なので、取り込み履歴の画面は無い。ここに出るのは
「まだ作業中のもの」と「確定したがまだダウンロードしていないもの」だけで、ダウンロードすると消える。
"""
from __future__ import annotations

from flask import Blueprint, render_template

from models import database
from views import TABLE_IMPORT_ACTIVE, form_link, query_all, query_value, table_import_link

bp = Blueprint("home", __name__)

LIST_LIMIT = 50


def _forms_working() -> list[dict]:
    rows = database.list_documents(state=("unread", "reviewing"), limit=LIST_LIMIT)
    for row in rows:
        row["href"] = form_link(row)
    return rows


READY_STATES = ("confirmed", "modified")


def _forms_ready() -> list[dict]:
    """ダウンロード待ちの帳票。まとめ取り込みは「まとまり1行」にする。

    1件ずつ .md を押すと、その帳票だけがまとまりから消えてしまう（zip は残りの分だけになる）。
    ホームからそれが黙って起きないよう、まとまりは zip への導線と一緒に1行で出す（design.md 2.2）。
    修正中でも確定済みの版は残っている（ダウンロードもその版）ので、ダウンロード待ちに入れる。
    """
    rows = database.list_documents(state=READY_STATES, limit=LIST_LIMIT)
    items: list[dict] = []
    seen: set[str] = set()
    for row in rows:
        row["href"] = form_link(row)
        batch_id = row.get("batch_id") or ""
        if not batch_id:
            items.append({"doc": row})
            continue
        if batch_id in seen:
            continue
        seen.add(batch_id)
        docs = database.list_batch_documents(batch_id)
        ready = [d for d in docs if d["state"] in READY_STATES]
        for d in ready:
            d["href"] = form_link(d)
        items.append({"batch": {"id": batch_id, "total": len(docs), "confirmed": len(ready),
                                "all_confirmed": len(ready) == len(docs), "docs": ready}})
    return items


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
        tables_working=_tables(TABLE_IMPORT_ACTIVE, LIST_LIMIT, "i.id DESC"),
        forms_ready=_forms_ready(),
        tables_ready=_tables(("confirmed",), LIST_LIMIT, "i.confirmed_at DESC, i.id DESC"),
    )
