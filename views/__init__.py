"""画面（blueprint）共通の小物。"""
from __future__ import annotations

import re
import sqlite3
from pathlib import Path
from urllib.parse import quote, urlsplit

from flask import request, url_for

from models import database


def safe_next(default: str) -> str:
    """フォームの next パラメータ（同一サイト内の相対パスのみ許可）"""
    target = request.form.get("next", "") or request.args.get("next", "")
    return target if target.startswith("/") and not target.startswith("//") else default


# ---- ダウンロードのファイル名 ---------------------------------------------------------
# Content-Disposition には UTF-8 の名前（filename*）と、それを読めない相手向けの ASCII の名前
# （filename）が並ぶ。日本語だけの名前だと ASCII 側が「_2026-09-19.zip」のように意味を失い、
# ログや古い環境では何のファイルか分からなくなるので、英字の既定名にそろえる。

_ASCII_UNSAFE = re.compile(r"[^A-Za-z0-9._-]+")


def ascii_download_name(name: str, default_stem: str) -> str:
    """name から作る ASCII のファイル名。英字が残らなければ default_stem を使う（拡張子は保つ）。"""
    stem = _ASCII_UNSAFE.sub("_", Path(name).stem).strip("._-")
    if not re.search(r"[A-Za-z]", stem):
        stem = f"{default_stem}_{stem}".strip("._-") if stem else default_stem
    return f"{stem}{Path(name).suffix}"


def set_download_name(response, name: str, default_stem: str):
    """添付ファイル名を「意味のある ASCII 名 ＋ 日本語の名前」にそろえる。"""
    fallback = ascii_download_name(name, default_stem)
    encoded = quote(name, safe="")
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{encoded}")
    return response


# ---- 書き込みの発火元を確かめる（他サイトからの操作を断る） ----------------------------
# 127.0.0.1 だけで待ち受けても、「利用者が開いた別のサイトのページが、そのブラウザから
# このアプリへ POST する」ことは防げない（AI接続先の書き換え・帳票の削除などができてしまう）。
# ブラウザは POST に必ず Origin を付け、最近のブラウザは Sec-Fetch-Site も付けるので、
# 他サイト発と分かる書き込みだけを断る（curl などヘッダの無い要求はこのPCからの操作として許す）。

SAFE_METHODS = ("GET", "HEAD", "OPTIONS")
SAME_SITE_FETCH = ("same-origin", "none")

# GET でもデータを消すルート（Markdown をダウンロードすると、その帳票・取り込みを消す。design.md 3.3）。
# 他サイトのページの <img>・リンク・window.open からでも GET は出せるので、書き込みと同じく発火元を確かめる。
# アドレス欄に打った URL（Sec-Fetch-Site: none）とアプリ内のクリック（same-origin）は通す。
PURGING_ENDPOINTS = frozenset({"forms.download_md", "forms.download_batch", "tables.download_zip"})


def _origin_of(url: str) -> str:
    parts = urlsplit(url)
    return f"{parts.scheme}://{parts.netloc}" if parts.scheme and parts.netloc else ""


def is_cross_site_write(req=None) -> bool:
    """他サイトのページから出された書き込み要求（データを消すダウンロードを含む）なら True。"""
    req = req if req is not None else request
    purging = req.endpoint in PURGING_ENDPOINTS
    if req.method in SAFE_METHODS and not purging:
        return False
    site = (req.headers.get("Sec-Fetch-Site") or "").strip().lower()
    if site and site not in SAME_SITE_FETCH:
        return True
    host = req.host_url.rstrip("/")
    origin = (req.headers.get("Origin") or "").strip()
    if origin and origin.rstrip("/") != host:
        return True
    if purging and not site:
        # Sec-Fetch-Site を付けない古いブラウザ向け。Referer が別サイトなら断る（無ければ手入力として通す）
        referer = (req.headers.get("Referer") or "").strip()
        return bool(referer) and _origin_of(referer) != host
    return False


# ---- DB の読み取り（ホームの一覧用の直接照会。表が無ければ空） ------------

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

def form_link(doc: dict) -> str:
    """帳票の状態に応じた「続き」の行き先（design.md 2.3 のルート）。"""
    state, doc_id = doc.get("state"), doc["id"]
    if state == "unread":
        return url_for("forms.type_select", doc_id=doc_id)
    if state in ("reviewing", "modified"):
        return url_for("forms.review", doc_id=doc_id)
    return url_for("forms.detail", doc_id=doc_id)


# ---- 一覧表の取り込みの状態 ------------------------------------------------------------

# 状態: uploaded=読み込み前 / reading=読み込み中 / preview=確認中 / confirming=確定処理中 / confirmed=確定済み / failed=失敗
TABLE_IMPORT_ACTIVE = ("uploaded", "reading", "preview", "confirming", "failed")


def table_import_link(item: dict) -> str:
    status, import_id = item.get("status"), item["id"]
    step = {"uploaded": "source", "reading": "preview", "preview": "preview", "failed": "preview",
            "confirming": "done", "confirmed": "done"}.get(status, "source")
    return url_for(f"tables.{step}", import_id=import_id)
