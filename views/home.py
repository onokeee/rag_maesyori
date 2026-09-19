"""ホーム: 2つの入口、作業中の一覧、ダウンロード待ちの一覧、はじめての案内。

データを残さない方針（design.md 3.3）なので、取り込み履歴の画面は無い。ここに出るのは
「まだ作業中のもの」と「確定したがまだダウンロードしていないもの」だけで、ダウンロードすると消える。
"""
from __future__ import annotations

from flask import Blueprint, render_template, request

from models import database
from views.forms import batch_save_confirm, batch_zip_confirm, delete_confirm, save_confirm
from views.tables import SAVE_TO_FOLDER_CONFIRM as TABLE_SAVE_CONFIRM
from views import TABLE_IMPORT_ACTIVE, form_link, query_all, query_value, table_import_link

bp = Blueprint("home", __name__)

LIST_LIMIT = 50
# 「すべて表示」のときの上限（ホームはこのPCに残っているデータの唯一の一覧なので、実質すべて出す）
ALL_LIMIT = 1_000_000


def _forms_working(limit: int) -> list[dict]:
    rows = database.list_documents(state=("unread", "reviewing"), limit=limit)
    for row in rows:
        row["href"] = form_link(row)
    return rows


READY_STATES = ("confirmed", "modified")


def _forms_ready(limit: int) -> list[dict]:
    """ダウンロード待ちの帳票。まとめ取り込みは「まとまり1行」にする。

    1件ずつ .md を押すと、その帳票だけがまとまりから消えてしまう（zip は残りの分だけになる）。
    ホームからそれが黙って起きないよう、まとまりは zip への導線と一緒に1行で出す（design.md 2.2）。
    修正中でも確定済みの版は残っている（ダウンロードもその版）ので、ダウンロード待ちに入れる。
    """
    rows = database.list_documents(state=READY_STATES, limit=limit)
    items: list[dict] = []
    seen: set[str] = set()
    for row in rows:
        row["href"] = form_link(row)
        # 修正中なら「確定し直していない変更は入らずに消える」ことを確認文で先に伝える
        row["download_confirm"] = delete_confirm(row)
        row["save_confirm"] = save_confirm(row)
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
            d["modified_note"] = ("この帳票は修正中です。確定し直していない変更は Markdown に入らず、消えます。"
                                  if d["state"] == "modified" else "")
        items.append({"batch": {"id": batch_id, "total": len(docs), "confirmed": len(ready),
                                "all_confirmed": len(ready) == len(docs), "docs": ready,
                                "zip_confirm": batch_zip_confirm(docs),
                                "save_confirm": batch_save_confirm(docs)}})
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


def _forms_ready_shown(items: list[dict]) -> int:
    """ダウンロード待ちの一覧に出た帳票の件数（まとまりは中の確定済みの件数で数える）。"""
    return sum(len(item["batch"]["docs"]) if "batch" in item else 1 for item in items)


def _count_tables(statuses: tuple[str, ...]) -> int:
    marks = ",".join("?" * len(statuses))
    return query_value(f"SELECT COUNT(*) FROM table_imports WHERE status IN ({marks})", statuses)


@bp.get("/")
def index():
    pattern_count = query_value("SELECT COUNT(*) FROM patterns")
    template_count = query_value("SELECT COUNT(*) FROM table_templates")
    # 一覧は新しい順に LIST_LIMIT 件まで。超えた分は「ほかにN件」と出し、?all=1 ですべて出す
    # （履歴画面は無いので、ここに出ないとダウンロードも削除もできないまま残ってしまう）。
    show_all = request.args.get("all") == "1"
    limit = ALL_LIMIT if show_all else LIST_LIMIT
    forms_working = _forms_working(limit)
    tables_working = _tables(TABLE_IMPORT_ACTIVE, limit, "i.id DESC")
    forms_ready = _forms_ready(limit)
    tables_ready = _tables(("confirmed",), limit, "i.confirmed_at DESC, i.id DESC")
    # 作業中＝確定済みの版が無い（unread/reviewing）、ダウンロード待ち＝確定済みの版がある（confirmed/modified）
    hidden = {
        "forms_working": query_value("SELECT COUNT(*) FROM documents WHERE confirmed_json IS NULL")
                         - len(forms_working),
        "tables_working": _count_tables(TABLE_IMPORT_ACTIVE) - len(tables_working),
        "forms_ready": query_value("SELECT COUNT(*) FROM documents WHERE confirmed_json IS NOT NULL")
                       - _forms_ready_shown(forms_ready),
        "tables_ready": _count_tables(("confirmed",)) - len(tables_ready),
    }
    return render_template(
        "home.html",
        first_run=pattern_count == 0 and template_count == 0,
        pattern_count=pattern_count,
        active_pattern_count=query_value("SELECT COUNT(*) FROM patterns WHERE status = 'active'"),
        template_count=template_count,
        forms_working=forms_working,
        tables_working=tables_working,
        forms_ready=forms_ready,
        tables_ready=tables_ready,
        hidden={key: max(0, n) for key, n in hidden.items()},
        list_limit=LIST_LIMIT,
        table_save_confirm=TABLE_SAVE_CONFIRM,
    )
