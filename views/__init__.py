"""画面（blueprint）共通の小物。"""
from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import quote, urlsplit
from uuid import uuid4

from flask import request, session


# ---- 使っている人ごとの作業場所 ---------------------------------------------------------
# 社内LANのサーバーで動かし、数人が同時に別々のPCから使う（2026-09-20 の利用者の指示）。
# ログインは無いので「誰か」は分からないが、「どのブラウザか」はセッションのクッキーで分かる。
# 取り込んだ帳票・一覧表はそのブラウザのものとして持ち主を記録し、ほかのブラウザからは
# 見えない・触れないようにする（views/forms.py・views/tables.py の 404）。
# 「設定」（帳票の種類・取り込み設定・AI接続）はみんなで使うものなので分けない。

SESSION_ID_KEY = "sid"


def current_session_id() -> str:
    """このブラウザの作業場所の id（クッキーに無ければ作る）。"""
    sid = session.get(SESSION_ID_KEY)
    if not isinstance(sid, str) or len(sid) != 32:
        sid = uuid4().hex
        session[SESSION_ID_KEY] = sid
    return sid


def owns(row) -> bool:
    """帳票・取り込みの行がこのブラウザのものか。

    session_id が空の行は持ち主が分からない（この仕組みを入れる前のDBの行・テストで直接作った行）ので、
    これまでどおり誰からでも扱えるものとして扱う。起動時の片付けで消えるので、普段は残らない。
    """
    owner = (row or {}).get("session_id") if hasattr(row, "get") else None
    return not owner or owner == current_session_id()


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
# 他サイト発と分かる書き込みだけを断る（curl などヘッダの無い要求はアプリを直接たたく操作として許す）。

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
