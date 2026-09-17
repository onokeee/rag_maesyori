"""取り込み履歴（帳票 / 一覧表）と、確定済み帳票の Markdown 一括ダウンロード。"""
from __future__ import annotations

import io
import json
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Blueprint, flash, redirect, render_template, request, send_file, url_for

from export import formats
from models import database
from views import FORM_STATES, TABLE_IMPORT_STATES, form_link, query_all, query_value, table_import_link

bp = Blueprint("history", __name__, url_prefix="/history")

PAGE_SIZE = 50
MAX_ZIP_DOCS = 1000


def _page() -> int:
    try:
        return max(1, int(request.args.get("page", "1")))
    except ValueError:
        return 1


def _int_arg(name: str) -> int | None:
    try:
        return int(request.args.get(name, ""))
    except ValueError:
        return None


def _like(text: str) -> str:
    escaped = text.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return f"%{escaped}%"


def _date_arg(name: str) -> str:
    raw = request.args.get(name, "").strip()
    try:
        return datetime.strptime(raw, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return ""


# ---- 帳票 ------------------------------------------------------------------------

def _form_filters() -> dict:
    state = request.args.get("state", "")
    return {
        "pattern_id": _int_arg("pattern_id"),
        "state": state if state in FORM_STATES else "",
        "date_from": _date_arg("date_from"),
        "date_to": _date_arg("date_to"),
        "q": request.args.get("q", "").strip()[:100],
    }


def _forms_page(filters: dict, page: int) -> tuple[list[dict], int]:
    rows, total = database.list_documents(state=filters["state"] or None, pattern_id=filters["pattern_id"],
                                          q=filters["q"] or None, date_from=filters["date_from"] or None,
                                          date_to=filters["date_to"] or None,
                                          limit=PAGE_SIZE, offset=(page - 1) * PAGE_SIZE)
    for row in rows:
        row["href"] = form_link(row)
    return rows, total


# ---- 一覧表 -----------------------------------------------------------------------

def _table_filters() -> dict:
    status = request.args.get("status", "")
    return {
        "template_id": _int_arg("template_id"),
        "status": status if status in TABLE_IMPORT_STATES else "",
        "q": request.args.get("q", "").strip()[:100],
    }


def _tables_page(filters: dict, page: int) -> tuple[list[dict], int]:
    where, args = [], []
    if filters["template_id"]:
        where.append("i.template_id = ?")
        args.append(filters["template_id"])
    if filters["status"]:
        where.append("i.status = ?")
        args.append(filters["status"])
    if filters["q"]:
        where.append("(i.file_name LIKE ? ESCAPE '\\' OR t.name LIKE ? ESCAPE '\\')")
        args += [_like(filters["q"])] * 2
    cond = f"WHERE {' AND '.join(where)}" if where else ""
    base = "FROM table_imports i LEFT JOIN table_templates t ON t.id = i.template_id"
    total = query_value(f"SELECT COUNT(*) {base} {cond}", args)
    rows = query_all(f"""
        SELECT i.id, i.file_name, i.status, i.stats_json, i.created_at, i.updated_at, i.confirmed_at, i.template_id,
               t.name AS template_name
        {base} {cond} ORDER BY i.id DESC LIMIT ? OFFSET ?
    """, (*args, PAGE_SIZE, (page - 1) * PAGE_SIZE))
    for row in rows:
        row["href"] = table_import_link(row)
        row["record_count"] = _record_count(row.get("stats_json"))
    return rows, total


def _record_count(stats_json: str | None) -> int | None:
    try:
        stats = json.loads(stats_json or "{}")
    except ValueError:
        return None
    if not isinstance(stats, dict):
        return None
    return stats["records"] if isinstance(stats.get("records"), int) else None


@bp.get("/")
def index():
    tab = "tables" if request.args.get("tab") == "tables" else "forms"
    page = _page()
    form_count = query_value("SELECT COUNT(*) FROM documents")
    table_count = query_value("SELECT COUNT(*) FROM table_imports")
    context = {"tab": tab, "page": page, "page_size": PAGE_SIZE, "form_count": form_count, "table_count": table_count}
    if tab == "forms":
        filters = _form_filters()
        rows, total = _forms_page(filters, page)
        context.update(filters=filters, rows=rows, total=total, form_states=FORM_STATES,
                       patterns=query_all("SELECT id, name FROM patterns ORDER BY name"))
    else:
        filters = _table_filters()
        rows, total = _tables_page(filters, page)
        context.update(filters=filters, rows=rows, total=total, import_states=TABLE_IMPORT_STATES,
                       templates=query_all("SELECT id, name FROM table_templates ORDER BY name"))
    context["pages"] = max(1, -(-context["total"] // PAGE_SIZE))
    context["query"] = {k: v for k, v in filters.items() if v}
    return render_template("history.html", **context)


# ---- 選択した確定済み帳票を zip で ----------------------------------------------------------

def _unique(name: str, used: set[str]) -> str:
    stem, suffix = Path(name).stem, Path(name).suffix or ".md"
    candidate, n = name, 2
    while candidate in used:
        candidate = f"{stem}_{n}{suffix}"
        n += 1
    used.add(candidate)
    return candidate


@bp.post("/forms/download")
def download_forms():
    ids = []
    for raw in request.form.getlist("ids"):
        try:
            ids.append(int(raw))
        except ValueError:
            continue
    ids = list(dict.fromkeys(ids))[:MAX_ZIP_DOCS]
    if not ids:
        flash("ダウンロードする帳票を選んでください", "error")
        return redirect(url_for(".index", tab="forms"))

    buffer, used = io.BytesIO(), set()
    docs = database.list_confirmed_documents(ids)
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for doc in docs:
            # 修正中の帳票も、確定済みの版で作る（Markdown は確定済みデータから毎回生成）
            try:
                extraction = json.loads(doc["confirmed_json"])
            except (TypeError, ValueError):
                continue
            name = _unique(formats.markdown_filename(doc, extraction), used)
            zf.writestr(name, formats.build_markdown(doc, extraction).encode("utf-8"))
    if not used:
        flash("選んだ帳票に確定済みのものがありません（まだ確定していない帳票はダウンロードできません）", "error")
        return redirect(url_for(".index", tab="forms"))
    buffer.seek(0)
    name = f"帳票Markdown_{datetime.now():%Y%m%d_%H%M%S}.zip"
    return send_file(buffer, mimetype="application/zip", as_attachment=True, download_name=name)

