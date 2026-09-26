"""Flask アプリ本体。設定・土台（DBを含む）・画面を1ファイルにまとめてある。
起動は flask --app app run --host=0.0.0.0 --port=5000（--app app はこのファイル app.py のこと。
waitress で動かすときだけ flask --app app serve）。

- 設定（旧 config.py）: env ファイルの読み込み、待ち受け先（HOST・PORT）、受け付ける宛先の名前。
- 土台（旧 core.py / models/database.py）: SQLite のスキーマ・移行・データアクセス、安全なファイル名、
  Markdown テキスト処理、アップロードの保存と事前チェック、ジョブ実行、取り込んだデータの削除。
- 画面（旧 views.py）: 2つの Blueprint（home＝「/」の転送と解説、tables＝表の取り込み）と段の断片描画。
- create_app とエラー画面、待ち受け先の決め方、起動コマンド serve（waitress 用）。

読み取りは extract.py（一覧表）、AI は ai.py（起動時には読み込まない）。
同じ名前で中身の違うものは分けてある（JOB_KIND_LABELS / ROW_KIND_LABELS、
save_ai_connection_row と画面側の save_ai_connection ルート）。
"""
from __future__ import annotations

import hashlib
import io
import json
import logging
import os
import posixpath
import queue
import re
import secrets
import shutil
import socket
import sqlite3
import sys
import threading
import time
import unicodedata
import zipfile
import zlib
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Callable
from urllib.parse import quote, urlsplit
from uuid import uuid4

from dotenv import load_dotenv
from flask import abort, Blueprint, current_app, flash, Flask, g, has_app_context, jsonify, redirect, render_template, request, send_file, session, url_for


####################################################################################################
# 設定（旧 config.py）
####################################################################################################

# ====================================================================================================
# 設定（元 config.py）
# ====================================================================================================

BASE_DIR = Path(__file__).resolve().parent   # プロジェクトの根（env・instance/・data/・uploads/ の置き場所）

# 接続設定は aiagent_minimal_rag_tougou と同じく、ドット無しの "env" ファイル（export KEY="..." 形式も可）から読む
load_dotenv(BASE_DIR / "env")


def _split(value: str) -> list[str]:
    return [v.strip() for v in value.replace(",", ";").split(";") if v.strip()]


def _optional_float(name: str) -> float | None:
    raw = os.environ.get(name, "").strip()
    return float(raw) if raw else None


class Config:
    # セッション署名鍵は create_app が FLASK_SECRET_KEY または .flask_secret ファイルから設定する
    SECRET_KEY = None
    DATABASE = Path(os.environ.get("DATABASE", BASE_DIR / "instance" / "app.db"))
    UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", BASE_DIR / "uploads"))
    # 行データ・状態ファイルの置き場所（下の TABLES_DIR。前の版の data/model_settings.yaml は読むだけ）
    DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data"))
    # 一覧表の行データ・状態ファイルの置き場所（DATA_DIR/tables。create_app で DATA_DIR に合わせて決め直す）
    TABLES_DIR = Path(os.environ.get("TABLES_DIR", DATA_DIR / "tables"))
    # 一覧表（Excel/CSV・1行＝1件）
    TABLE_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".tsv", ".txt"}
    TABLE_MAX_UPLOAD_BYTES = int(os.environ.get("TABLE_MAX_UPLOAD_BYTES", str(200 * 1024 * 1024)))
    # 1回のリクエストの上限（一覧表の大きいCSV/Excelを想定）。受け口はここ。
    # 200MB 固定にしていたころは、TABLE_MAX_UPLOAD_BYTES を広げても Flask が先に 413 で断り、
    # 設定が効かなかった（2026-09-23 のレビューで実測）
    MAX_CONTENT_LENGTH = max(200 * 1024 * 1024, TABLE_MAX_UPLOAD_BYTES)
    # Excel のセル数の上限（これを超えるシートは読み込まない）
    EXCEL_MAX_CELLS = int(os.environ.get("EXCEL_MAX_CELLS", "500000"))
    # 読み込み・下書き・AI整形を同時に動かす本数（列ごと。core/jobs.py）。
    # 社内LANのサーバーで数人が同時に使うので、1本だと誰かの長い読み込みでほかの人が待たされる。
    # 1〜8 に丸められる。列ごとの本数はプロセス内でその列を最初に使ったときに決まり、あとから減らない
    JOB_WORKERS = int(os.environ.get("JOB_WORKERS", "3"))

    # ---- LLM（OpenAI互換API） ----
    OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1").strip().rstrip("/")
    OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
    OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.6-sol").strip()
    OPENAI_MODELS = _split(os.environ.get("OPENAI_MODELS", ""))
    OPENAI_TEMPERATURE = float(os.environ.get("OPENAI_TEMPERATURE", "0"))
    OPENAI_TOP_P = _optional_float("OPENAI_TOP_P")
    # OPENAI_MAX_TOKENS・LLM_RATE_LIMIT_RETRIES・LLM_RATE_LIMIT_MAX_WAIT は持たない。
    # 読み込むだけでどこからも使っておらず、書かれた既定値（3回・20秒）も実際の値と違っていた
    # （再試行は ai.py の MAX_RETRIES=5・MAX_WAIT=120.0。Retry-After の優先や、429 が3行続いたら
    #  一時停止する決まりと噛み合っているので、外から2つだけ変えられる形にはしない。2026-09-23）


# ====================================================================================================
# アプリ（元 app.py）
# ====================================================================================================

_SECRET_FILE = BASE_DIR / ".flask_secret"

# --- 「/」は表の取り込みへ --------------------------------------------------------------
# ホーム画面は無い。画面は「表の取り込み」と「解説」の2つだけで、作業中の一覧も
# ダウンロード待ちの一覧も持たない（利用者の指示 2026-09-20）。
# ここに残るのは「/」を開いたときの転送だけ（古いブックマーク・開いたままの画面の戻り先も拾う）。
home_bp = Blueprint("home", __name__)


@home_bp.get("/", endpoint="index")
def to_tables():
    return redirect(url_for("tables.new"))


@home_bp.get("/guide", endpoint="guide")
def guide():
    """解説（Markdown がどう作られるか）。上部タブの2つ目。読むだけの画面で、データには触れない
    （利用者の求め 2026-09-22）。中身は templates/base.html の screen == "guide" の節。"""
    from flask import render_template

    return render_template("base.html", screen="guide")


def _secret_key() -> str:
    """セッション署名鍵（画面のメッセージ表示に使う）。再起動しても変わらないようファイルに保持する。"""
    env = os.getenv("FLASK_SECRET_KEY", "").strip()
    if env:
        return env
    if _SECRET_FILE.exists():
        return _SECRET_FILE.read_text(encoding="utf-8").strip()
    key = secrets.token_urlsafe(48)
    _SECRET_FILE.write_text(key, encoding="utf-8")
    try:
        os.chmod(_SECRET_FILE, 0o600)
    except OSError:
        pass
    return key


# --- デバッガは開かせない -----------------------------------------------------------
# デバッガが開いていると、例外が出たときにブラウザからこのサーバの Python を実行できる。
# PIN はユーザー名・MACアドレス・マシンIDから作る固定値で、守りにならない。

_TRUE = ("1", "true", "yes", "on", "t", "y")
_DEBUG_ARGS = ("--debugger", "--debug")
_NO_DEBUG_MSG = (
    "\n[app] デバッグモードでは起動しません。\n"
    "  理由: デバッガが開くと、例外が出たときにブラウザからこのサーバの Python を実行できます。\n"
    "  断ったもの: flask run の --debug / --debugger と、環境変数 FLASK_DEBUG=1。\n"
    "  通常の起動: flask --app app run --host=0.0.0.0 --port=5000\n"
    "  詳しいエラーを見たいとき: app.py の DEBUG を True にして flask --app app serve\n"
)


def _refuse_debugger() -> None:
    if str(os.environ.get("FLASK_DEBUG") or "").strip().lower() in _TRUE:
        raise SystemExit(_NO_DEBUG_MSG)
    if any(a in _DEBUG_ARGS for a in sys.argv[1:]):
        raise SystemExit(_NO_DEBUG_MSG)


# --- 待ち受け先と、受け付ける宛先の名前 -------------------------------------------------
# サーバ（JupyterLab のターミナルなど）で起動し、社内LANの他のPCから数人で開いて使う（design.md 0）。
# 既定はこのサーバの中からだけ開ける 127.0.0.1。LAN に出すときは起動のコマンドで渡す:
#   flask --app app run --host=0.0.0.0 --port=5000
# 待ち受け先は「コマンド行（--host/--port）→ FLASK_RUN_HOST/FLASK_RUN_PORT → HOST/PORT → 既定」の順に決める。
# コマンド行を最優先にするのは、実際に待ち受けるのがその値だから。ここがずれると
# 受け付ける宛先（allowed_hosts）がループバックだけのままになり、他のPCから開いて断られる。
_ALL_ADDRESSES = ("", "0.0.0.0", "::")
_SERVER_COMMANDS = ("run", "serve")


def _start_command() -> str:
    """flask のコマンド行のサブコマンド（"run" / "serve"。それ以外・テストからの呼び出しは ""）。"""
    for arg in sys.argv[1:]:
        if arg in _SERVER_COMMANDS:
            return arg
    return ""


def _cli_option(*names: str) -> str:
    """コマンド行から `--host 0.0.0.0` `--host=0.0.0.0` `-h 0.0.0.0` の値を取り出す（無ければ ""）。"""
    if not _start_command():
        return ""
    args = sys.argv[1:]
    for i, arg in enumerate(args):
        for name in names:
            if arg == name:
                return args[i + 1].strip() if i + 1 < len(args) else ""
            if arg.startswith(f"{name}="):
                return arg.split("=", 1)[1].strip()
    return ""


def _env_host(default: str = "127.0.0.1") -> str:
    for value in (_cli_option("--host", "-h"), os.environ.get("FLASK_RUN_HOST"), os.environ.get("HOST")):
        text = str(value or "").strip()
        if text:
            return text
    return default


def _env_port(default: int = 5000) -> int:
    for where, value in (("--port", _cli_option("--port", "-p")),
                         ("FLASK_RUN_PORT", os.environ.get("FLASK_RUN_PORT")),
                         ("PORT", os.environ.get("PORT"))):
        raw = str(value or "").strip()
        if not raw:
            continue
        try:
            port = int(raw)
        except ValueError:
            raise SystemExit(f"\n[app] {where} が数字ではありません: {raw!r}\n"
                             f"  例: flask --app app run --host=0.0.0.0 --port=5000\n") from None
        if not 1 <= port <= 65535:
            raise SystemExit(f"\n[app] {where} が範囲外です: {port}（1〜65535）\n")
        return port
    return default


HOST = _env_host()
PORT = _env_port()


def _is_loopback(host) -> bool:
    # 空文字は 0.0.0.0 と同じで全てのネットワークから届くので loopback に含めない
    return str(host or "").strip().lower() in ("127.0.0.1", "localhost", "::1")


def _hostname_only(value) -> str:
    """"mypc:5000" や "[::1]:5000" から、宛先の名前だけを取り出す（小文字・末尾のドットは落とす）。"""
    try:
        name = urlsplit(f"//{str(value or '').strip()}").hostname or ""
    except ValueError:
        return ""
    return name.rstrip(".").lower()


def _lan_address() -> str:
    """LAN の他のPCから届くこのサーバのアドレス。取れなければ空（通信はしない）。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("192.0.2.1", 9))   # TEST-NET-1。UDP なので何も送らず、どの口から出るかだけ決めさせる
        return sock.getsockname()[0]
    except OSError:
        return ""
    finally:
        sock.close()


@lru_cache(maxsize=1)
def _machine_names() -> tuple[str, ...]:
    """このサーバ自身の名前とアドレス（LAN から届く宛先）。起動中は変わらない前提で1回だけ調べる。"""
    names: set[str] = set()
    try:
        hostname = socket.gethostname()
    except OSError:
        hostname = ""
    if hostname:
        names.update({hostname.lower(), hostname.split(".")[0].lower()})
        try:
            canonical, aliases, addresses = socket.gethostbyname_ex(hostname)
        except OSError:
            pass
        else:
            names.update(n.lower() for n in (canonical, *aliases) if n)
            names.update(addresses)
    lan = _lan_address()
    if lan:
        names.add(lan)
    # 日本語のPC名のときブラウザは punycode（xn--…）で送ってくるので、その形も受け付ける
    for name in list(names):
        try:
            names.add(name.encode("idna").decode("ascii").lower())
        except (UnicodeError, ValueError):
            pass
    return tuple(sorted(n for n in names if n))


def allowed_hosts(host: str | None = None) -> set[str]:
    """受け付ける宛先の名前（Host ヘッダー）。

    ログインが無いので、宛先の名前まで確かめる。攻撃者のドメインがこのサーバのアドレスを指していると、
    そのページとこのアプリが同一オリジンになり、画面の中身を読み取られてしまう（DNSリバインディング）。
    受け付けるのは、ループバックと、LAN に出しているときはこのサーバ自身の名前・アドレス。
    ロードバランサや別名で開くときは ALLOWED_HOSTS（`;` か `,` 区切り）で足す。`*` ですべて受け付ける。
    """
    bound = (HOST if host is None else host).strip().lower()
    names = {"127.0.0.1", "localhost", "::1"}
    for entry in (os.environ.get("ALLOWED_HOSTS") or "").replace(",", ";").split(";"):
        entry = entry.strip()
        if entry == "*":
            return {"*"}
        if entry:
            names.add(_hostname_only(entry) or entry.lower())
    if bound not in _ALL_ADDRESSES:
        names.add(_hostname_only(bound) or bound)
    if bound in _ALL_ADDRESSES or not _is_loopback(bound):
        names.update(_machine_names())   # LAN に出しているときだけ、PC名・LANのアドレスでも開ける
    return names


def _home_host() -> str:
    """画面や起動時の案内に出す「開けるアドレス」の宛先。"""
    bound = HOST.strip().lower()
    if _is_loopback(bound):
        return "127.0.0.1"
    if bound in _ALL_ADDRESSES:
        return _lan_address() or (_machine_names()[0] if _machine_names() else "127.0.0.1")
    return bound


def startup_notice(port: int | None = None) -> str:
    """起動したときに出す案内。LAN に出しているときは、ログインが無いことを必ず知らせる。"""
    url = f"http://{_home_host()}:{port or PORT}/"
    if _is_loopback(HOST):
        return f"[app] {url} で起動しました（このサーバの中からだけ開けます）"
    line = "=" * 68
    return (
        f"\n{line}\n"
        f"[app] 社内LANに公開して起動しました。ほかのPCからは次のアドレスを開いてください。\n"
        f"\n      {url}\n\n"
        f"  ・ログインはありません。このアドレスを知っている人は誰でも使えます。\n"
        f"    いま取り込んでいる表の中身も、開かれれば見えます。\n"
        f"    信頼できる社内LANの中だけで使ってください。\n"
        f"  ・ダウンロードするとサーバーからデータが消えます。ダウンロードしていないものも\n"
        f"    しばらくすると捨てます。要るものはその場でダウンロードしてください。\n"
        f"  ・止めるとき: Ctrl+C（nohup で動かしているときは kill <PID>）\n"
        f"{line}\n"
    )


####################################################################################################
# 土台（旧 core.py・models/database.py）
####################################################################################################

# ====================================================================================================
# 元 models/database.py
# SQLite のスキーマ・マイグレーションとデータアクセス。
#
# 一覧表・ジョブ・AI のテーブルをここで作る。一覧表のデータアクセスは tables/store.py、
# ジョブは core/jobs.py が持つ。
# ====================================================================================================

BUSY_TIMEOUT_MS = 5000


# ---- スキーマ -----------------------------------------------------------------

# 第1版（作り直し前）のスキーマ。古いDBにも新規DBにも同じ形で適用できるよう IF NOT EXISTS
_SCHEMA_TABLES = """
CREATE TABLE IF NOT EXISTS table_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL DEFAULT '',
    current_version_id INTEGER,                    -- table_template_versions.id（循環するので外部キーにしない）
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS table_template_versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id INTEGER NOT NULL REFERENCES table_templates(id) ON DELETE CASCADE,
    version INTEGER NOT NULL,
    spec_json TEXT NOT NULL,
    spec_hash TEXT NOT NULL,
    used INTEGER NOT NULL DEFAULT 0,
    note TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    UNIQUE (template_id, version)
);

CREATE TABLE IF NOT EXISTS table_imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id INTEGER REFERENCES table_templates(id) ON DELETE SET NULL,
    template_version_id INTEGER REFERENCES table_template_versions(id) ON DELETE SET NULL,
    file_name TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    source_json TEXT NOT NULL DEFAULT '{}',        -- {"kind","encoding","delimiter","preamble_rows","sheet","header_row","header_rows","data_end_row"}
    status TEXT NOT NULL DEFAULT 'uploaded',       -- uploaded / reading / preview / confirming / confirmed / discarded / failed
    stats_json TEXT NOT NULL DEFAULT '{}',
    issues_path TEXT,
    rows_path TEXT,
    job_id INTEGER,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    confirmed_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_table_imports_template ON table_imports(template_id);

CREATE TABLE IF NOT EXISTS jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL,
    ref_type TEXT NOT NULL DEFAULT '',
    ref_id INTEGER,
    status TEXT NOT NULL DEFAULT 'queued',         -- queued / running / paused / done / failed / cancelled / interrupted
    params_json TEXT NOT NULL DEFAULT '{}',
    progress_json TEXT NOT NULL DEFAULT '{}',
    message TEXT NOT NULL DEFAULT '',
    cancel_requested INTEGER NOT NULL DEFAULT 0,
    pause_requested INTEGER NOT NULL DEFAULT 0,
    heartbeat_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status);
CREATE INDEX IF NOT EXISTS idx_jobs_ref ON jobs(ref_type, ref_id);

CREATE TABLE IF NOT EXISTS llm_calls (
    cache_key TEXT PRIMARY KEY,
    raw_text TEXT,
    parsed_json TEXT,
    model TEXT,
    params_json TEXT,
    structured_mode TEXT,
    finish_reason TEXT,
    tokens_in INTEGER,
    tokens_out INTEGER,
    latency_ms INTEGER,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS ai_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id INTEGER NOT NULL,
    stage_id TEXT NOT NULL,
    row_key TEXT NOT NULL,
    template_version_id INTEGER,
    source_hash TEXT,
    context_hash TEXT,
    segments_hash TEXT,
    cache_key TEXT,
    status TEXT NOT NULL DEFAULT 'pending',        -- pending / ok / flagged / rule_only / error / skipped / outdated / excluded
    result_json TEXT,
    checks_json TEXT,
    attempts INTEGER NOT NULL DEFAULT 0,
    error TEXT,
    job_id INTEGER,
    updated_at TEXT NOT NULL,
    UNIQUE (template_id, stage_id, row_key)
);
CREATE INDEX IF NOT EXISTS idx_ai_items_status ON ai_items(template_id, stage_id, status);
"""


def _column_names(conn: sqlite3.Connection, table: str) -> set[str]:
    return {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def _add_column(conn: sqlite3.Connection, table: str, name: str, definition: str) -> None:
    """列が無ければ追加する（古いDBに途中まで手で足した場合も通るように）。"""
    if name not in _column_names(conn, table):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")


def _run_script(conn: sqlite3.Connection, script: str) -> None:
    # executescript は途中で COMMIT するので、トランザクション内では1文ずつ実行する
    for statement in script.split(";"):
        lines = [ln for ln in statement.splitlines() if ln.strip()]
        if lines:
            conn.execute("\n".join(lines))


def _retired_forms(conn: sqlite3.Connection) -> None:
    """もう無い機能（1ファイル＝1件の取り込み）の表を作っていた版。機能ごと外したので何もしない。

    版の番号は「適用済みの件数」なので、消して番号をずらすと古いDBが壊れる。並びだけ残す。
    古いDBに残っているその表は、最後の版（_m15_drop_forms）で落とす。
    """


def _m1_base(conn: sqlite3.Connection) -> None:
    _retired_forms(conn)


def _m2_forms(conn: sqlite3.Connection) -> None:
    _retired_forms(conn)


def _m3_tables(conn: sqlite3.Connection) -> None:
    _run_script(conn, _SCHEMA_TABLES)


def _m4_form_batches(conn: sqlite3.Connection) -> None:
    _retired_forms(conn)


# 使っていない表（8.1 で外した機能のもの）。取り込みを指す列が無く purge の探索に乗らないので、
# 行が入れば永久に残ってしまう。作らない機能の表は置かない（design.md 3.3）
_UNUSED_TABLES = ("table_outputs", "table_downloads", "table_template_samples", "alias_entries")


def _m5_purge_scope(conn: sqlite3.Connection) -> None:
    """AI整形の控えを「取り込み単位」で消せるようにし、使っていない表を落とす。

    ai_items は (設定, 段, 行) で一意なので、設定単位で消すと同じ設定で作業中の別の取り込みの結果まで
    消えていた（課金済みの応答キャッシュも道連れ）。import_id は次の _m6 で表を作り直すときに持たせる。
    """
    for table in _UNUSED_TABLES:
        conn.execute(f"DROP TABLE IF EXISTS {table}")


_AI_ITEMS_COLUMNS = ("id, template_id, stage_id, row_key, template_version_id, source_hash, context_hash, "
                     "segments_hash, cache_key, status, result_json, checks_json, attempts, error, job_id, "
                     "updated_at, import_id")


def _m6_ai_items_per_import(conn: sqlite3.Connection) -> None:
    """ai_items を「取り込み × 段 × 行」で一意にする（表を作り直す。SQLite は UNIQUE を外せないため）。

    (設定, 段, 行) で一意だと、同じ設定・同じ行の2つ目の取り込みは1つ目の結果を「処理済み」とみなして
    自分の行を持たず、1つ目をダウンロード（削除）すると2つ目の AI 結果まで消えていた（design.md 3.3）。
    """
    _add_column(conn, "ai_items", "import_id", "INTEGER")
    conn.execute("""CREATE TABLE ai_items_new (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER NOT NULL,
        stage_id TEXT NOT NULL,
        row_key TEXT NOT NULL,
        template_version_id INTEGER,
        source_hash TEXT,
        context_hash TEXT,
        segments_hash TEXT,
        cache_key TEXT,
        status TEXT NOT NULL DEFAULT 'pending',
        result_json TEXT,
        checks_json TEXT,
        attempts INTEGER NOT NULL DEFAULT 0,
        error TEXT,
        job_id INTEGER,
        updated_at TEXT NOT NULL,
        import_id INTEGER,
        UNIQUE (import_id, template_id, stage_id, row_key)
    )""")
    conn.execute(f"INSERT INTO ai_items_new ({_AI_ITEMS_COLUMNS}) SELECT {_AI_ITEMS_COLUMNS} FROM ai_items")
    conn.execute("DROP TABLE ai_items")
    conn.execute("ALTER TABLE ai_items_new RENAME TO ai_items")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ai_items_status ON ai_items(template_id, stage_id, status)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_ai_items_import ON ai_items(import_id)")


def _m7_llm_calls_owner(conn: sqlite3.Connection) -> None:
    """AI の生の応答に「どの取り込みが払ったか」を持たせる（design.md 3.3）。

    ai_items が覚えているのは最後に使った応答のキーだけなので、再依頼で直した行の1回目の応答や、
    一時停止の直前に受け取った応答は、どの ai_items からも参照されない。持ち主が分からないと、
    その取り込みを消しても（ほかの取り込みが作業中の間は）生の応答が残ってしまう。
    古いDBの行は NULL（持ち主が分からない）のまま、これまでどおり扱う。
    """
    _add_column(conn, "llm_calls", "import_id", "INTEGER")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_llm_calls_import ON llm_calls(import_id)")


def _m8_session_scope(conn: sqlite3.Connection) -> None:
    """取り込んだものに「どのブラウザのものか」を持たせる（社内LANで数人が同時に使うため）。

    ログインは無いので利用者は分からないが、セッションのクッキー（current_session_id）で
    ブラウザごとの作業場所は分けられる。ほかのブラウザの取り込みは見えない（404）。
    古いDBの行は NULL（持ち主が分からない）のままで、これまでどおり扱う（起動時の片付けで消える）。
    """
    _add_column(conn, "table_imports", "session_id", "TEXT")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_table_imports_session ON table_imports(session_id)")


_TABLE_IMPORTS_COLUMNS = ("id, template_id, template_version_id, file_name, file_hash, stored_path, source_json, "
                          "status, stats_json, issues_path, rows_path, job_id, created_at, updated_at, "
                          "confirmed_at, session_id, spec_json, spec_hash")

def _m9_import_spec(conn: sqlite3.Connection) -> None:
    """取り込み設定の保存をやめ、その取り込みが使う設定（spec）を取り込みの行に持たせる。

    利用者の指示（2026-09-20）:「表の方には、取り込み設定を保持しておく機能はいらない」。
    取り込みごとに列の対応づけを決めるので、設定を名前で残して選び直す仕組み（table_templates と
    その版）は要らない。表を作り直すのは table_templates への外部キーを外すため（参照先の表を
    落とすと、外部キーを有効にした接続では table_imports への書き込みがすべて落ちる）。
    template_id / template_version_id は取り込み自身の id にそろえて残す（ai.py・purge_table_import が
    AI整形の控えを束ねる鍵に使っている）。
    """
    _add_column(conn, "table_imports", "spec_json", "TEXT NOT NULL DEFAULT ''")
    _add_column(conn, "table_imports", "spec_hash", "TEXT NOT NULL DEFAULT ''")
    if conn.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'table_template_versions'").fetchone():
        # まだダウンロードしていない取り込みが使っている版を、取り込みの行に写す
        conn.execute("""UPDATE table_imports SET
            spec_json = COALESCE((SELECT v.spec_json FROM table_template_versions v
                                  WHERE v.id = table_imports.template_version_id), ''),
            spec_hash = COALESCE((SELECT v.spec_hash FROM table_template_versions v
                                  WHERE v.id = table_imports.template_version_id), '')""")
    conn.execute("""CREATE TABLE table_imports_new (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER,                           -- = id（AI整形の控えを束ねる鍵。設定はもう無い）
        template_version_id INTEGER,                   -- = id（AI整形の控えの「作り直し判定」に使う）
        file_name TEXT NOT NULL,
        file_hash TEXT NOT NULL,
        stored_path TEXT NOT NULL,
        source_json TEXT NOT NULL DEFAULT '{}',        -- {"kind","encoding","delimiter","preamble_rows","sheet","header_rows","data_end_row"}
        status TEXT NOT NULL DEFAULT 'uploaded',       -- uploaded / reading / preview / confirming / confirmed / failed
        stats_json TEXT NOT NULL DEFAULT '{}',
        issues_path TEXT,
        rows_path TEXT,
        job_id INTEGER,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        confirmed_at TEXT,
        session_id TEXT,
        spec_json TEXT NOT NULL DEFAULT '',            -- この取り込みの取り込み設定（tables.spec.TableSpec の JSON）
        spec_hash TEXT NOT NULL DEFAULT ''
    )""")
    conn.execute(f"INSERT INTO table_imports_new ({_TABLE_IMPORTS_COLUMNS}) "
                 f"SELECT {_TABLE_IMPORTS_COLUMNS} FROM table_imports")
    conn.execute("DROP TABLE table_imports")
    conn.execute("ALTER TABLE table_imports_new RENAME TO table_imports")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_table_imports_session ON table_imports(session_id)")
    conn.execute("DROP TABLE IF EXISTS table_template_versions")
    conn.execute("DROP TABLE IF EXISTS table_templates")
    conn.execute("UPDATE table_imports SET template_id = id, template_version_id = id")
    # AI整形の控えは「設定」ではなく取り込みで束ねる。持ち主の分からない古い行（import_id なし）は消す
    conn.execute("DELETE FROM ai_items WHERE import_id IS NULL")
    conn.execute("UPDATE ai_items SET template_id = import_id, template_version_id = import_id")


def _m10_drop_unused(conn: sqlite3.Connection) -> None:
    """使わなくなったものを落とす。

    - 互換ビュー table_template_versions: _m9 が一時的に置いたもの。ai.load_spec が
      table_imports.spec_json を直接読むようになったので要らない。
    - table_imports.period_json: 書く側がもう無く、常に '{}' のまま（期間は取り込み設定が持つ）。
    """
    conn.execute("DROP VIEW IF EXISTS table_template_versions")
    if "period_json" in _column_names(conn, "table_imports"):
        conn.execute("ALTER TABLE table_imports DROP COLUMN period_json")


def _m11_document_touch(conn: sqlite3.Connection) -> None:
    _retired_forms(conn)


def _m12_drop_pattern_samples(conn: sqlite3.Connection) -> None:
    _retired_forms(conn)


def _m13_ai_connections(conn: sqlite3.Connection) -> None:
    """AI接続（APIキー・接続先・モデル）を「ブラウザごと」に持つ（利用者の指示 2026-09-21）。

    利用者の指示:「AI接続の設定は、もともとの位置ヘッダーの画面右上『AI接続』に移動させる。全部空欄にしておいて
    cookieでユーザー毎に登録内容をずっと保持させるようにしてほしい」。
    これまでは data/model_settings.yaml 1つを全員で使っていたので、社内LANで数人が使うと互いのキーを
    上書きし合い、1つのキーを共有していた。持ち主は取り込みと同じ session_id（current_session_id）。
    この表は「設定」なので、取り込みを捨てる片付け（purge_session / sweep_stale / purge_all_pending）の
    対象にしない（SETTINGS_TABLES）。最後の接続確認の結果もここに持ち、ヘッダーはそれを表示するだけ
    （画面を開くたびに確認しに行かない）。
    """
    conn.execute("""CREATE TABLE IF NOT EXISTS ai_connections (
        session_id TEXT PRIMARY KEY,                   -- ブラウザの作業場所（current_session_id）
        api_key TEXT NOT NULL DEFAULT '',              -- 平文（暗号化はしていない。画面にそう書く）
        chat_url TEXT NOT NULL DEFAULT '',             -- …/chat/completions までのフルパス（空欄＝サーバー共通の値）
        models_url TEXT NOT NULL DEFAULT '',           -- …/models までのフルパス（空欄＝サーバー共通の値）
        model TEXT NOT NULL DEFAULT '',                -- 使うモデル（空欄＝サーバー共通の既定）
        models_json TEXT NOT NULL DEFAULT '[]',        -- APIから取得した候補（datalist に出すだけ）
        last_check_ok INTEGER,                         -- NULL=未確認 / 1=つながった / 0=つながらなかった
        last_check_at TEXT,
        last_check_detail TEXT NOT NULL DEFAULT '',    -- 最後の確認の結果（つながらなかった理由）
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    )""")


def _m14_pattern_books(conn: sqlite3.Connection) -> None:
    _retired_forms(conn)


def _m15_drop_forms(conn: sqlite3.Connection) -> None:
    """もう使わない表（1ファイル＝1件の取り込みと、その種類の登録）を落とす。機能ごと外した（利用者の指示 2026-09-26）。

    残っていても使わないが、置いたままだと「消えないデータ」になる（design.md 3.3）。
    新しいDBでは作っていないので、古いDBのときだけ効く。
    """
    for table in ("pattern_books", "pattern_samples", "pattern_fields", "pattern_sheets",
                  "documents", "patterns"):
        conn.execute(f"DROP TABLE IF EXISTS {table}")


# PRAGMA user_version = 適用済みの件数。追加は末尾にだけ行う
MIGRATIONS: list[Callable[[sqlite3.Connection], None]] = [_m1_base, _m2_forms, _m3_tables, _m4_form_batches,
                                                          _m5_purge_scope, _m6_ai_items_per_import,
                                                          _m7_llm_calls_owner, _m8_session_scope, _m9_import_spec,
                                                          _m10_drop_unused, _m11_document_touch,
                                                          _m12_drop_pattern_samples, _m13_ai_connections,
                                                          _m14_pattern_books, _m15_drop_forms]


def migrate(conn: sqlite3.Connection) -> int:
    """未適用のマイグレーションを1つずつトランザクションで適用し、適用後の版を返す。"""
    current = conn.execute("PRAGMA user_version").fetchone()[0]
    for number, migration in enumerate(MIGRATIONS[current:], start=current + 1):
        conn.execute("BEGIN IMMEDIATE")
        try:
            migration(conn)
            conn.execute(f"PRAGMA user_version = {number}")
            conn.commit()
        except Exception:
            conn.rollback()
            raise
    return max(current, len(MIGRATIONS))


# ---- 接続 -----------------------------------------------------------------------

def connect(path: str | Path | None = None) -> sqlite3.Connection:
    """g を使わない接続（バックグラウンドスレッド用）。呼び出し側で close する。

    path 省略時はアプリ設定の DATABASE（app_context が必要）。
    """
    if path is None:
        path = current_app.config["DATABASE"]
    conn = sqlite3.connect(path, timeout=BUSY_TIMEOUT_MS / 1000, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute(f"PRAGMA busy_timeout = {BUSY_TIMEOUT_MS}")
    conn.execute("PRAGMA journal_mode = WAL")
    # 消した行の中身をその場でゼロ埋めする（design.md 3.3）。これが無いと、行を消しても解放された
    # ページに取り込んだ値・対象の名前・人名が平文で残り、DBファイル（と app.db-wal）から読めてしまう。
    # VACUUM は他の接続があると失敗することがあるので、消し方そのものを安全側にしておく。
    conn.execute("PRAGMA secure_delete = ON")
    return conn


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = connect()
    return g.db


def close_db(_exc=None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_app(app) -> None:
    Path(app.config["DATABASE"]).parent.mkdir(parents=True, exist_ok=True)
    conn = connect(app.config["DATABASE"])
    try:
        migrate(conn)
    finally:
        conn.close()
    app.teardown_appcontext(close_db)


def now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _one(sql: str, args=()) -> dict | None:
    row = get_db().execute(sql, args).fetchone()
    return dict(row) if row else None


# ---- ai_connections（ブラウザごとの AI接続） ---------------------------------------------------
# 持ち主は session_id（current_session_id）。行は「保存」か「接続の確認」で初めてできる（読むだけでは作らない）。
# 取り込みの片付け（purge_session など）では消えない「設定」。長く使われていない行だけ sweep_stale_ai_connections が消す。


def get_ai_connection(session_id: str | None) -> dict | None:
    """そのブラウザの AI接続の行（無ければ None）。models は list にして返す。"""
    if not session_id:
        return None
    row = _one("SELECT * FROM ai_connections WHERE session_id = ?", (session_id,))
    if row is None:
        return None
    try:
        row["models"] = [str(m) for m in json.loads(row.get("models_json") or "[]") if str(m).strip()]
    except ValueError:
        row["models"] = []
    return row


def save_ai_connection_row(session_id: str, *, api_key: str | None = None, chat_url: str | None = None,
                       models_url: str | None = None, model: str | None = None,
                       models: list[str] | None = None) -> dict:
    """そのブラウザの AI接続を保存する（None の項目は変えない）。接続先・キーを変えたら確認の結果は「未確認」に戻す。"""
    ts = now()
    db = get_db()
    with db:
        row = db.execute("SELECT * FROM ai_connections WHERE session_id = ?", (session_id,)).fetchone()
        if row is None:
            db.execute("INSERT INTO ai_connections (session_id, created_at, updated_at) VALUES (?, ?, ?)",
                       (session_id, ts, ts))
        sets, args = ["updated_at = ?"], [ts]
        for name, value in (("api_key", api_key), ("chat_url", chat_url), ("models_url", models_url), ("model", model)):
            if value is not None:
                sets.append(f"{name} = ?")
                args.append(str(value))
        if models is not None:
            sets.append("models_json = ?")
            args.append(json.dumps([str(m) for m in models], ensure_ascii=False))
        # つながるかどうかは接続先・キー・モデルで決まる。どれかが変わったら前の結果は当てにならない
        changed = row is None or any(value is not None and str(value) != (row[name] or "")
                                     for name, value in (("api_key", api_key), ("chat_url", chat_url),
                                                         ("models_url", models_url), ("model", model)))
        if changed:
            sets.append("last_check_ok = NULL")
            sets.append("last_check_at = NULL")
            sets.append("last_check_detail = ''")
        db.execute(f"UPDATE ai_connections SET {', '.join(sets)} WHERE session_id = ?", (*args, session_id))
    return get_ai_connection(session_id) or {}


def set_ai_connection_check(session_id: str, ok: bool, detail: str = "") -> None:
    """接続の確認の結果（ヘッダーの「接続中／つながりません」と「最終確認」）を覚える。行が無ければ作る。"""
    ts = now()
    db = get_db()
    with db:
        db.execute("INSERT OR IGNORE INTO ai_connections (session_id, created_at, updated_at) VALUES (?, ?, ?)",
                   (session_id, ts, ts))
        db.execute("UPDATE ai_connections SET last_check_ok = ?, last_check_at = ?, last_check_detail = ?, "
                   "updated_at = ? WHERE session_id = ?",
                   (1 if ok else 0, ts, str(detail or "")[:500], ts, session_id))


def clear_ai_connection_key(session_id: str) -> bool:
    """［キーを消す］（共有PC）。接続先・モデルは残し、キーと確認の結果だけ消す。戻り値: 行があったか。"""
    db = get_db()
    with db:
        cur = db.execute("UPDATE ai_connections SET api_key = '', last_check_ok = NULL, last_check_at = NULL, "
                         "last_check_detail = '', updated_at = ? WHERE session_id = ?", (now(), session_id))
    return cur.rowcount > 0


# ====================================================================================================
# 元 core/naming.py
# 出力ファイル名（Windows と LightRAG の両方で安全な名前）。
#
# LightRAG のファイル名ヒント（`.[legacy-R(...)]`）は付けない。ヒントはサーバー側の取り込み設定より
# 優先されてしまう上に、そのサーバーが知らない書き方だと取り込みが HTTP 400 で断られる。
# 名前は「安定・意味が分かる・重複しない」だけを満たし、チャンクへの耐性は本文の作り方（記録の分割）で確保する。
# ====================================================================================================

_UNSAFE = re.compile(r'[\\/:*?"<>|\[\]\s\x00-\x1f\x7f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def safe_filename_part(text, max_len: int = 60) -> str:
    """ファイル名の1部品。NFKC、禁止文字・空白・角かっこを _ に、'.[' を除去、前後の . _ を除去。

    ここは本文と違って NFKC をそのままかける（囲み文字は残さない）。ファイル名は同じ文書に対して
    安定・一意であればよく、①などを残すと OS や LightRAG 側での扱いが揺れるため。
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = s.replace(".[", "")  # LightRAG のヒント記法と誤認されないように
    s = _UNSAFE.sub("_", s)
    s = re.sub(r"_+", "_", s).strip("._")
    s = s[:max_len].strip("._")
    if s.split(".")[0].upper() in _WINDOWS_RESERVED:
        s = f"{s}_"
    return s


def md_filename(parts: list[str]) -> str:
    """部品を _ で連結した .md ファイル名。空の部品は飛ばす。"""
    safe = [p for p in (safe_filename_part(x) for x in parts) if p]
    return f"{'_'.join(safe) or '無題'}.md"


# ====================================================================================================
# 元 core/mdtext.py
# Markdown 出力のテキスト処理。
#
# 決まり（docs/design.md 6章）: UTF-8・LF、レコード内に空行を入れない、パイプ表を使わない（`- 項目: 値`）、
# 複数行の値は2文字下げの連続行。
# ====================================================================================================

_SPACES = re.compile(r"[^\S\n]+")          # 改行以外の空白の連続
_HEADING = re.compile(r"^(#{1,6})(\s|$)")
_LIST = re.compile(r"^([-*+])(\s|$)")
_ORDERED = re.compile(r"^(\d{1,9})([.)])(\s|$)")
_RULE = re.compile(r"^(?:-[ \t]*){3,}$|^(?:=[ \t]*){3,}$|^(?:\*[ \t]*){3,}$|^(?:_[ \t]*){3,}$")
_FENCE = re.compile(r"^(```|~~~)")


def _enclosed_marks() -> str:
    """NFKC で囲みが外れてしまう囲み文字（丸数字 ①、丸英字 Ⓐ、丸カナ ㋐、丸漢字 ㊤ など）を集める。"""
    marks = []
    for start, end in ((0x2460, 0x24FF), (0x3240, 0x32FF), (0x1F100, 0x1F1FF)):
        for cp in range(start, end + 1):
            ch = chr(cp)
            if "CIRCLED" in unicodedata.name(ch, "") and unicodedata.normalize("NFKC", ch) != ch:
                marks.append(ch)
    return "".join(marks)


# ①→1 のように囲みが外れると「①破損…」が「1破損…」になり、番号と本文の区切りが消えて読めなくなる。
# 丸数字は手順や項目番号でごく普通に使われるので、囲み文字だけは NFKC をかけずに残す。
# （㈱→(株)、⑴→(1) のように区切りが残る表記はそのまま NFKC で正規化する）
_ENCLOSED = re.compile(f"([{re.escape(_enclosed_marks())}])")


def nfkc_keep_enclosed(text: str) -> str:
    """NFKC 正規化。ただし囲み文字（①Ⓐ㋐㊤…）はそのまま残す。

    囲み文字は前後の文字と結合しないので、そこで区切って正規化しても結果は変わらない。
    """
    if not _ENCLOSED.search(text):
        return unicodedata.normalize("NFKC", text)
    return "".join(part if _ENCLOSED.fullmatch(part) else unicodedata.normalize("NFKC", part)
                   for part in _ENCLOSED.split(text))


def escape_md_line(line: str) -> str:
    """行頭の見出し・箇条書き・引用・番号付きリスト、区切り線、コードフェンスとして解釈されないようにする。"""
    body = line.lstrip(" \t")
    indent = line[: len(line) - len(body)]
    if not body:
        return line
    if _FENCE.match(body):
        body = "\\" + body[0] + "\\" + body[1] + "\\" + body[2:]
    elif _RULE.match(body):
        body = "\\" + body
    elif _HEADING.match(body) or _LIST.match(body) or body.startswith(">"):
        body = "\\" + body
    else:
        m = _ORDERED.match(body)
        if m:
            body = m.group(1) + "\\" + body[len(m.group(1)):]
    return indent + body


def _value_lines(value) -> list[str]:
    if value is None:
        return []
    items = value if isinstance(value, (list, tuple)) else [value]
    lines: list[str] = []
    for item in items:
        if item is None:
            continue
        text = str(item).replace("\r\n", "\n").replace("\r", "\n")
        lines += [ln.strip() for ln in text.split("\n") if ln.strip()]
    return lines


def md_bullet(label: str, value) -> list[str]:
    """`- label: value`。複数行は `- label:` の後に2文字下げで続ける。値が空なら出さない（[]）。"""
    lines = _value_lines(value)
    if not lines:
        return []
    label = " ".join(str(label).split())
    if len(lines) == 1:
        return [f"- {label}: {lines[0]}"]
    return [f"- {label}:"] + [f"  {escape_md_line(ln)}" for ln in lines]


# estimate_tokens 用：UTF-8 のバイトを種類の印に置き換える表（数字 → "0"、ASCII の記号 → "."、それ以外 → "x"）
_TOKEN_CLASS = bytes(
    ord("0") if 0x30 <= b <= 0x39 else ord(".") if (0x21 <= b <= 0x2F or 0x3A <= b <= 0x40 or 0x5B <= b <= 0x60
                                                     or 0x7B <= b <= 0x7E) else ord("x")
    for b in range(256))


def estimate_tokens(text: str) -> int:
    """推定トークン数（実トークン以上になる見積もり）。

    非ASCII 1文字=1.1、ASCII の記号 1文字=1、数字は「連続する数字のまとまり1つ=1 ＋ 3桁ごとに1」、
    それ以外の ASCII（英字・空白・改行）2文字=1。
    LightRAG の o200k_base は数字を3桁ずつに区切り、記号（- : / . = など）もほぼ1文字ずつ別のトークンにする。
    日時・品番・計測値の多い記録（「2023-09-01 09:44」は実10トークン）を英字と同じ2文字=1で数えると、
    実トークンより3割ほど少なく見積もり、記録の上限（チャンク 1,500）を超えることがあった。
    実出力 142,952 ブロックとの実測で、この式は実/推定の最大 0.99（旧式は 1.32）、全体では 23% 多めに見積もる。
    例外：まれな漢字（髙・﨑 など）は1文字が2〜3トークンになる。記録全体では他の文字の余裕に吸収される。
    """
    if not text:
        return 0
    ascii_count = len(text.encode("ascii", "ignore"))  # ASCII の文字数（1文字ずつ数えるより速い）
    marks = text.encode("utf-8", "surrogatepass").translate(_TOKEN_CLASS)
    digits = marks.count(b"0")
    punct = marks.count(b".")
    digit_runs = marks.count(b"x0") + marks.count(b".0") + (marks[:1] == b"0")   # 数字のまとまりの数
    other = ascii_count - digits - punct
    # 30分の1トークン単位の整数で数えて ceil する（浮動小数の誤差を避ける）
    total = 33 * (len(text) - ascii_count) + 15 * other + 30 * punct + 10 * digits + 30 * digit_runs
    return -(-total // 30)


def join_blocks(blocks: list[list[str]]) -> str:
    """ブロック間に空行1つ、ブロック内の空行は除く、末尾改行1つ、LF。"""
    parts: list[str] = []
    for block in blocks:
        lines: list[str] = []
        for line in block or []:
            for ln in str(line).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
                if ln.strip():
                    lines.append(ln.rstrip())
        if lines:
            parts.append("\n".join(lines))
    return "\n\n".join(parts) + "\n" if parts else ""


# ====================================================================================================
# 元 core/files.py
# アップロードファイルの保存・事前チェック・削除。
#
# 保存は分割して読みながら sha256 を計算する（ブック全体を解析してからハッシュを取らない）。
# Excel の事前チェックは openpyxl で開く前に、先頭バイトと zip の目次だけで行う。
# ====================================================================================================

logger = logging.getLogger(__name__)

CHUNK_SIZE = 1024 * 1024

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"
STRICT_NS = "http://purl.oclc.org/ooxml/spreadsheetml/main"

# zip 展開の上限（値は仮置き）
ZIP_MAX_TOTAL = 500 * 1024 * 1024
ZIP_MAX_PART = 200 * 1024 * 1024
ZIP_MAX_RATIO = 100
# 小さいパーツは圧縮率が高くても害がないので、この大きさ以上だけ圧縮率を見る
ZIP_RATIO_MIN_BYTES = 16 * 1024 * 1024
# 結合セルの面積（ブック全体の合計）の上限。openpyxl は開くときに結合範囲のセルを1つずつ作るので、
# シート全体の結合（A1:XFD1048576 など）が1つあるだけで、上のサイズ制限より前に固まる（数秒〜終わらない）。
# 行全体の結合（A1:XFD1 = 16,384 セル）や列全体の結合（A:A = 約105万セル）は通す。
# 一覧表は読み取り専用モードで開き、結合を展開しないので、この値で足りる
MAX_MERGED_CELLS = 2_000_000
# ハイパーリンク・コメントの範囲の面積（ブック全体の合計）の上限。openpyxl はリンク（<hyperlink ref>）とコメント
# （<comment ref>）の範囲のセルを1つずつ作る。
# シート全体を指す範囲1つで開くのが終わらなくなる（A1:XFD50 の約82万セルで一覧表は約28秒）。
# ふつうのリンク・コメントは1セルか数セルなので、1回あたり2秒以内に収まるこの値で断る。
MAX_LINKED_CELLS = 50_000
# セル数の上限（ブック全体の <c> の数）。openpyxl は開くときにセルを1つずつ作るので、
# 圧縮すると小さいが展開すると大量のセルがあるブック（64KB で 100万セルなど）は、上のサイズ制限を通っても固まる。
# 一覧表は tables.excel_source が別に上限（EXCEL_MAX_CELLS）を持ち「CSVで保存」と案内するので、ここは最後の砦の値。
MAX_CELLS = 1_000_000
# 図形（描画）の上限。openpyxl は開くときに、シートが参照する描画部品（xl/drawings/*.xml）をシートごとに読み直し、
# 図形（アンカー）を1つずつ作る。1つの大きな描画を多数のシートから参照させると、圧縮後は小さくても
# 開くたびに数十秒〜終わらない。そこで「描画部品の図形数 × 参照するシート数」と
# 「描画部品の展開後の大きさ × 参照するシート数」の合計を数えて上限を設ける。
# 普通の表は1シートに数十個程度の図形なので、十分に余裕のある値にしている。
MAX_DRAWING_ANCHORS = 10_000
MAX_DRAWING_BYTES = 20 * 1024 * 1024
# 図形（アンカー）として数える要素の名前（DrawingML の spreadsheetDrawing）
_ANCHOR_NAMES = ("absoluteAnchor", "oneCellAnchor", "twoCellAnchor")
_DRAWING_REL_SUFFIX = "/drawing"
# openpyxl がシートの数だけ読み直す部品の関係の種類。シートの部品（worksheet）は <sheet> 1つにつき1回、
# コメントの部品（comments）はそれを参照するシートを読むたびに読み直される。1つの部品を複数のシートから
# 参照させると（Excel は作らないが、手で作れば開ける）、結合セル・セル・行・コメントの範囲もその回数だけ
# 作られるので、部品1つ分だけ数えていた上限をいくらでもすり抜けられる。参照される回数を掛けて数える。
_REREAD_REL_SUFFIXES = ("/worksheet", "/comments")
_RELS_SUFFIX = ".rels"
_DRAWING_ERROR = "図形や画像の数が多すぎるため読み込めません。不要な図形・画像を削除して保存し直してください"
# セルとして数える要素の名前空間（SpreadsheetML 本体の <c> だけ。グラフの <c:chart> などは名前空間が違う）
_SHEET_NAMESPACES = ("http://schemas.openxmlformats.org/spreadsheetml/2006/main", STRICT_NS)
# workbook.xml の名前空間判定で読む先頭バイト数
_WORKBOOK_HEAD_BYTES = 64 * 1024
# 掴まれているファイルを消すときの再試行（ウイルス対策のスキャンなどは短時間で終わる）
REMOVE_RETRIES = 3
REMOVE_RETRY_WAIT = 0.05
# アップロードの保存先（UPLOAD_DIR 直下）と save_upload が付けるファイル名の形
UPLOAD_SUBDIRS = ("tables",)
_STORED_NAME = re.compile(r"[0-9a-f]{32}\.[A-Za-z0-9]+")


class UploadError(Exception):
    """利用者に見せる日本語メッセージを持つ例外。"""


@dataclass
class StoredFile:
    stored_path: str   # UPLOAD_DIR からの相対パス（/ 区切り）
    file_name: str     # 元のファイル名（表示・ダウンロード用）
    file_hash: str     # sha256
    size: int


def original_name(storage) -> str:
    # secure_filename は日本語を消してしまうので、表示用には元のファイル名（パス部分除去）を使う
    return Path((getattr(storage, "filename", None) or "").replace("\\", "/")).name


def _normalize_ext(ext: str) -> str:
    ext = ext.strip().lower()
    return ext if ext.startswith(".") else f".{ext}"


def _format_size(size: int) -> str:
    return f"{size / 1024 / 1024:.0f}MB" if size >= 1024 * 1024 else f"{size / 1024:.0f}KB"


def upload_path(stored_path: str) -> Path:
    """保存パス → 実ファイルのパス。UPLOAD_DIR の外を指すパスは拒否する。"""
    base = Path(current_app.config["UPLOAD_DIR"]).resolve()
    path = (base / stored_path).resolve()
    if path != base and base not in path.parents:
        raise UploadError("保存先のパスが不正です")
    return path


def _checked_name(storage, allowed: set[str]) -> tuple[str, str]:
    """元のファイル名と拡張子。名前が無い・扱えない拡張子なら UploadError。"""
    name = original_name(storage)
    if not name:
        raise UploadError("ファイルを選択してください")
    ext = Path(name).suffix.lower()
    allowed_exts = {_normalize_ext(e) for e in allowed}
    if ext not in allowed_exts:
        kinds = " / ".join(sorted(allowed_exts))
        hint = "（.xls はExcelで .xlsx に保存し直してください）" if ext == ".xls" else ""
        raise UploadError(f"{kinds} のファイルを選んでください{hint}")
    return name, ext


def save_upload(storage, subdir: str, allowed: set[str], max_bytes: int) -> StoredFile:
    """アップロードを UPLOAD_DIR/subdir に保存する。拡張子・サイズ・空ファイルを確認し、失敗時は消して UploadError。"""
    name, ext = _checked_name(storage, allowed)
    stored = f"{subdir.strip('/')}/{uuid4().hex}{ext}"
    dest = upload_path(stored)
    dest.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256()
    size = 0
    stream = getattr(storage, "stream", storage)
    try:
        with open(dest, "wb") as out:
            while True:
                chunk = stream.read(CHUNK_SIZE)
                if not chunk:
                    break
                size += len(chunk)
                if size > max_bytes:
                    raise UploadError(f"ファイルが大きすぎます（上限 {_format_size(max_bytes)}）")
                digest.update(chunk)
                out.write(chunk)
        if size == 0:
            raise UploadError("ファイルが空です")
    except BaseException:
        dest.unlink(missing_ok=True)
        raise
    return StoredFile(stored_path=stored, file_name=name, file_hash=digest.hexdigest(), size=size)


def precheck_excel(source, max_cells: int | None = None, max_merged: int | None = None,
                   max_rows: int | None = None) -> None:
    """Excel（.xlsx/.xlsm）として開いてよいかを、中身を展開せずに確かめる。問題があれば UploadError。

    source: ファイルのパス、またはブックの中身そのもの（bytes。zipfile はどちらも同じように読める）。
    max_cells: セル数の上限（省略時は MAX_CELLS）。
    max_merged: 結合セルの面積の上限（省略時は MAX_MERGED_CELLS）。
    max_rows: 行（<row>）の数の上限（省略時は max_merged と同じ値）。
    """
    if isinstance(source, (bytes, bytearray)):
        head, book = bytes(source[:8]), io.BytesIO(source)
    else:
        book = Path(source)
        with open(book, "rb") as f:
            head = f.read(8)
    if head.startswith(OLE_MAGIC[:4]):
        raise UploadError("パスワード付きのブック、または .xls 形式です。パスワードを外して .xlsx 形式で保存し直してください")
    if not head.startswith(ZIP_MAGIC):
        raise UploadError("Excelファイル（.xlsx / .xlsm）として読み込めません。Excelで開いて .xlsx 形式で保存し直してください")
    try:
        with zipfile.ZipFile(book) as zf:
            infos = zf.infolist()
            names = {i.filename for i in infos}
            if "xl/workbook.bin" in names:
                raise UploadError(".xlsb（バイナリブック）形式は対象外です。Excelで .xlsx 形式で保存し直してください")
            _check_zip_limits(infos)
            workbook = "xl/workbook.xml" if "xl/workbook.xml" in names else next(
                (n for n in sorted(names) if n.lower().endswith("workbook.xml")), None)
            if workbook is None:
                raise UploadError("Excelファイル（.xlsx / .xlsm）ではありません（ブックの情報が見つかりません）")
            with zf.open(workbook) as wb:
                text = wb.read(_WORKBOOK_HEAD_BYTES).decode("utf-8", errors="ignore")
            # シートの置き場所と名前（拡張子）は workbook.xml.rels で自由に決められる（sheet1.dat でも開ける）ので、
            # 名前で選ばずに部品をすべて見る（XML でない部品は、読み始めてすぐ読めなくなって終わる）
            max_merged = max_merged or MAX_MERGED_CELLS
            _check_sheet_parts(zf, sorted(names), max_cells or MAX_CELLS, max_merged, max_rows or max_merged)
    # zlib.error / EOFError: 圧縮データが壊れている（zipfile はこれらを包まずにそのまま投げる）
    except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError, RuntimeError, NotImplementedError,
            zlib.error, EOFError) as exc:
        # 例外の種類は画面に出さず（利用者には意味がない）、ログにだけ残す
        logger.warning("Excelファイルの事前チェックで読み込めませんでした: %s", exc.__class__.__name__)
        raise UploadError("Excelファイルとして読み込めません（ファイルが壊れている可能性があります）") from exc
    root = re.search(r"<(?:\w+:)?workbook\b[^>]*>", text)
    if STRICT_NS in (root.group(0) if root else text[:2048]):
        raise UploadError("Strict Open XML 形式のブックです。Excelの「名前を付けて保存」で通常の「Excel ブック (.xlsx)」を選んで保存し直してください")
    # シートが1つも無いブックは断る（読み取り先が無い）。先頭だけ読んでいるので、
    # シートの一覧の終わりが読んだ範囲に無いときは判断しない
    sheets_end = re.search(r"<(?:\w+:)?sheets\s*/>|</(?:\w+:)?sheets>", text)
    if sheets_end and not re.search(r"<(?:\w+:)?sheet\b", text[:sheets_end.end()]):
        raise UploadError("シートがないブックです。シートのあるブックを選んでください")


def _merged_area(ref: str | None) -> int:
    from openpyxl.utils.cell import range_boundaries

    try:
        min_col, min_row, max_col, max_row = range_boundaries(str(ref or "").strip().upper())
    except (ValueError, TypeError):
        return 0   # 読めない範囲は openpyxl 側で扱いが決まる（ここでは数えない）
    if None in (min_col, min_row, max_col, max_row):
        return 0
    return (abs(max_col - min_col) + 1) * (abs(max_row - min_row) + 1)


class _PartCounter:
    """XML の部品1つを分割して読み、結合範囲の面積・セル数・行数・リンクとコメントの範囲の面積を数える（xml.parsers.expat）。

    文字列を正規表現で探すのではなく XML として読むので、ref = "..."（= の前後の空白）、文字参照（&#88;）、
    どんな名前空間の接頭辞（<x.y:c> など）でも、openpyxl（xml.etree＝同じ expat）と同じ解釈で数える。
    """

    def __init__(self, merged: int, cells: int, max_cells: int, max_merged: int | None = None, *,
                 rows: int = 0, linked: int = 0, max_rows: int | None = None, factor: int = 1):
        from xml.parsers import expat

        self.merged, self.cells, self.max_cells = merged, cells, max_cells
        self.max_merged = max_merged or MAX_MERGED_CELLS
        self.rows, self.linked = rows, linked
        self.max_rows = max_rows or self.max_merged
        self.factor = max(int(factor), 1)   # この部品が openpyxl に読み直される回数（その回数だけ数える）
        self.anchors = 0              # この部品の図形（アンカー）の数
        self.drawing_targets = []     # この部品（.rels）が参照する描画部品の Target
        self.reread_targets = []      # この部品（.rels）が参照する、シートごとに読み直される部品の Target
        self.parser = expat.ParserCreate(namespace_separator=" ")
        self.parser.StartElementHandler = self._start
        # DTD で実体を定義して大量に展開させる細工は、正しいブックには無いので断る
        self.parser.EntityDeclHandler = self._entity

    def _start(self, name: str, attrs: dict) -> None:
        namespace, _, local = name.rpartition(" ")
        if local == "mergeCell":
            self.merged += _merged_area(attrs.get("ref")) * self.factor
            if self.merged > self.max_merged:
                raise UploadError("結合セルの範囲が大きすぎます（シート全体・列全体の結合など）。"
                                  "不要な結合を解除して保存し直してください")
        elif local == "c" and namespace in _SHEET_NAMESPACES:
            self.cells += self.factor
            if self.cells > self.max_cells:
                raise UploadError(f"セル数が上限（{self.max_cells:,} セル）を超えています。"
                                  "不要なシート・範囲を削除して保存し直してください")
        elif local == "row" and namespace in _SHEET_NAMESPACES:
            self.rows += self.factor
            if self.rows > self.max_rows:
                raise UploadError(f"行数が上限（{self.max_rows:,} 行）を超えています（高さなどの書式だけの行も数えます）。"
                                  "不要な行を削除して保存し直してください")
        elif local in ("hyperlink", "comment") and namespace in _SHEET_NAMESPACES:
            self.linked += _merged_area(attrs.get("ref")) * self.factor
            if self.linked > MAX_LINKED_CELLS:
                raise UploadError("ハイパーリンクまたはコメントの範囲が大きすぎます（シート全体・列全体を指すリンクなど）。"
                                  "不要なリンク・コメントを削除して保存し直してください")
        elif local in _ANCHOR_NAMES:
            self.anchors += 1
        elif local == "Relationship" and attrs.get("TargetMode") != "External":
            rel_type = str(attrs.get("Type", ""))
            if rel_type.endswith(_DRAWING_REL_SUFFIX):
                self.drawing_targets.append(str(attrs.get("Target", "")))
            else:
                for suffix in _REREAD_REL_SUFFIXES:
                    if rel_type.endswith(suffix):
                        self.reread_targets.append((suffix, str(attrs.get("Target", ""))))
                        break

    def _entity(self, *_args) -> None:
        raise UploadError("Excelファイルとして読み込めません（不正なファイルの可能性があります）")


def _check_sheet_parts(zf: zipfile.ZipFile, parts: list[str], max_cells: int, max_merged: int | None = None,
                       max_rows: int | None = None) -> None:
    """部品の XML を分割して読み、結合範囲の面積・セル数・行数・リンクとコメントの範囲の面積の合計を数え、上限を超えたら UploadError。

    openpyxl で開く前に確かめる（開いた時点で結合範囲のセルと、すべてのセルが作られてしまうため）。
    XML として読めなくなった部品は、そこまでの分だけ数える（画像などの XML でない部品はすぐ終わる）。
    openpyxl も同じ expat で読むので、読めない部品のその先の結合・セルは作られない（開くこと自体が失敗する）。
    """
    merged = 0
    cells = 0
    rows = 0
    linked = 0
    anchors: dict[str, int] = {}       # 部品 → 図形の数
    references: dict[str, int] = {}    # 描画部品 → 参照される回数（シートごとに読み直されるため）
    rereads = _reread_counts(zf, parts)
    for part in parts:
        counter = _PartCounter(merged, cells, max_cells, max_merged, rows=rows, linked=linked, max_rows=max_rows,
                               factor=rereads.get(part, 1))
        _parse_part(zf, part, counter)
        merged, cells, rows, linked = counter.merged, counter.cells, counter.rows, counter.linked
        anchors[part] = counter.anchors
        for target in counter.drawing_targets:
            drawing = _resolve_rel_target(part, target)
            # 参照元のシートの部品を複数の <sheet> が指していれば、その描画もその回数だけ読み直される
            references[drawing] = references.get(drawing, 0) + rereads.get(_rels_owner(part), 1)
    _check_drawings(zf, anchors, references)


def _parse_part(zf: zipfile.ZipFile, part: str, counter: "_PartCounter") -> None:
    """部品1つを分割して読み、counter に数えさせる（XML として読めなくなったら、そこまでの分で終わる）。"""
    from xml.parsers import expat

    with zf.open(part) as f:
        try:
            while True:
                chunk = f.read(CHUNK_SIZE)
                counter.parser.Parse(chunk, not chunk)
                if not chunk:
                    break
        except expat.ExpatError:
            pass


def _reread_counts(zf: zipfile.ZipFile, parts: list[str]) -> dict[str, int]:
    """部品 → openpyxl がそれを読む回数（シートの部品・コメントの部品だけ。ふつうのブックはどれも1回）。

    シートの部品は <sheet>（workbook.xml.rels の worksheet の関係）1つにつき1回、コメントの部品は
    それを参照するシートを読むたびに1回読まれる。関係を書いた .rels は小さいので、本体を数える前に先に読む。
    """
    targets: list[tuple[str, str, str]] = []    # (関係の種類, .rels の部品名, 参照先の部品名)
    for part in parts:
        if not part.endswith(_RELS_SUFFIX):
            continue
        counter = _PartCounter(0, 0, MAX_CELLS)
        _parse_part(zf, part, counter)
        targets += [(kind, part, _resolve_rel_target(part, t)) for kind, t in counter.reread_targets]
    counts: dict[str, int] = {}
    for kind, _rels_part, target in targets:
        if kind == "/worksheet":
            counts[target] = counts.get(target, 0) + 1
    # コメントの部品は、それを参照するシートを読むたびに読み直される（そのシートの部品を複数の <sheet> が
    # 指していれば、その回数だけ増える）ので、持ち主のシートの回数を足す
    for kind, rels_part, target in targets:
        if kind == "/comments":
            counts[target] = counts.get(target, 0) + max(counts.get(_rels_owner(rels_part), 0), 1)
    return counts


def _rels_owner(rels_part: str) -> str:
    """.rels の持ち主の部品（xl/worksheets/_rels/sheet1.xml.rels → xl/worksheets/sheet1.xml）。"""
    folder = posixpath.dirname(posixpath.dirname(rels_part))
    return posixpath.normpath(posixpath.join(folder, posixpath.basename(rels_part)[:-len(_RELS_SUFFIX)]))


def _resolve_rel_target(rels_part: str, target: str) -> str:
    """.rels の Target を zip 内のパスにする（xl/worksheets/_rels/sheet1.xml.rels の ../drawings/d.xml → xl/drawings/d.xml）。"""
    if target.startswith("/"):
        return target.lstrip("/")
    folder = posixpath.dirname(posixpath.dirname(rels_part))   # _rels の1つ上＝参照元の部品があるフォルダ
    return posixpath.normpath(posixpath.join(folder, target))


def _check_drawings(zf: zipfile.ZipFile, anchors: dict[str, int], references: dict[str, int]) -> None:
    """描画部品を読み直す回数（参照するシートの数）を掛けて、図形の数と大きさの合計が上限を超えたら UploadError。"""
    total_anchors = 0
    total_bytes = 0
    for drawing, count in references.items():
        if drawing not in anchors:
            continue   # 存在しない部品は openpyxl も読まない
        total_anchors += anchors[drawing] * count
        total_bytes += zf.getinfo(drawing).file_size * count
        if total_anchors > MAX_DRAWING_ANCHORS or total_bytes > MAX_DRAWING_BYTES:
            raise UploadError(_DRAWING_ERROR)


def _check_zip_limits(infos: list[zipfile.ZipInfo]) -> None:
    total = 0
    compressed = 0
    for info in infos:
        total += info.file_size
        compressed += info.compress_size
        if info.file_size > ZIP_MAX_PART:
            raise UploadError(f"ブックの中身が大きすぎます（1つの部品が展開後 {_format_size(info.file_size)}、上限 {_format_size(ZIP_MAX_PART)}）")
        if info.file_size >= ZIP_RATIO_MIN_BYTES and info.file_size > ZIP_MAX_RATIO * max(info.compress_size, 1):
            raise UploadError("ブックの圧縮率が異常に高いため読み込みを中止しました（壊れているか、不正なファイルの可能性があります）")
    if total > ZIP_MAX_TOTAL:
        raise UploadError(f"ブックの中身が大きすぎます（展開後の合計 {_format_size(total)}、上限 {_format_size(ZIP_MAX_TOTAL)}）")
    # 圧縮率はブック全体でも見る。部品ごとの確認は小さい部品を見ないので、上の大きさに足りない部品を
    # 並べるだけで、小さいファイルをいくらでも大きく展開させられる（700KB が 465MB ＝ 事前チェックだけで30秒以上）。
    # ふつうのブックは5〜7倍程度なので、部品ごとと同じ100倍で断る（小さいブックは見ない点も同じ）。
    if total >= ZIP_RATIO_MIN_BYTES and total > ZIP_MAX_RATIO * max(compressed, 1):
        raise UploadError("ブックの圧縮率が異常に高いため読み込みを中止しました（壊れているか、不正なファイルの可能性があります）")


def remove_upload(stored_path: str | None) -> bool:
    """アップロードしたファイルを消す。消せたら True。

    Windows では他のプロセス（ウイルス対策のスキャン、Excel で開いたまま、同期ソフト）や、解析に失敗した
    ブックを掴んだままの openpyxl のせいで消せないことがある。「ダウンロードしたら消えます」（design.md 3.3）と
    言い切っている以上、消せないときも中身だけは 0 バイトに切り詰めて、元のデータが残らないようにする。
    例外は投げない（ここで落とすと、日本語のエラー案内の代わりに 500 になる）。
    残った空ファイルは次の起動時に remove_orphan_uploads が片付ける。
    """
    if not stored_path:
        return True
    try:
        path = upload_path(stored_path)
    except UploadError:
        return False
    for attempt in range(REMOVE_RETRIES):
        try:
            path.unlink(missing_ok=True)
            return True
        except OSError:
            time.sleep(REMOVE_RETRY_WAIT * (attempt + 1))
    try:
        with open(path, "r+b") as f:
            f.truncate(0)
    except OSError:
        pass
    return False


def remove_orphan_import_dirs(tables_dir, known_ids) -> int:
    """DB に無い取り込みの imports/<id>/ フォルダを消す。消した件数を返す。

    取り込みを消し損ねたとき（削除の途中で落ちた・Windows でファイルを掴まれていた）に、読み込んだ行や
    作った Markdown が残り続けないように起動時に片付ける。作業中の取り込みは DB に行があるので消さない。
    """
    root = Path(tables_dir) / "imports"
    if not root.is_dir():
        return 0
    known = {int(i) for i in known_ids}
    removed = 0
    for path in sorted(root.iterdir()):
        if not path.is_dir() or not path.name.isdigit() or int(path.name) in known:
            continue
        shutil.rmtree(path, ignore_errors=True)
        if not path.exists():
            removed += 1
    return removed


# 起動時の片付けで、これより新しいファイルは消さない（保存してから DB の行を作るまでの間を守る）。
# STALE_AFTER（2分）と同じ長さ。
ORPHAN_GRACE_SECONDS = 120


def known_stored_paths(db) -> set[str]:
    """DB のどこかの表から参照されているアップロード済みファイルの一覧。

    「消してよいファイル」を決める唯一の場所。2か所に写していたころは、片方だけ直すと
    使用中のファイルを消す危険があった（stored_path の列を持つ表を全部見る）。
    """
    known: set[str] = set()
    for (table,) in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall():
        columns = {row[1] for row in db.execute(f'PRAGMA table_info("{table}")')}
        if "stored_path" in columns:
            known |= {row[0] for row in db.execute(f'SELECT stored_path FROM "{table}"') if row[0]}
    return known


def remove_orphan_uploads(base, known_paths) -> int:
    """DB のどこからも参照されていないアップロード済みファイルを消す。消した件数を返す。

    取り込みの途中で失敗し、消し損ねたファイルが残ることがある（画面からは消せない）ので起動時に片付ける。
    save_upload が作った名前（uuid + 拡張子）のファイルだけを対象にする。
    """
    base = Path(base)
    known = {str(p).replace("\\", "/").strip("/") for p in known_paths if p}
    removed = 0
    for subdir in UPLOAD_SUBDIRS:
        for path in sorted((base / subdir).glob("*")):
            if not path.is_file() or not _STORED_NAME.fullmatch(path.name):
                continue
            if f"{subdir}/{path.name}" in known:
                continue
            try:
                if time.time() - path.stat().st_mtime < ORPHAN_GRACE_SECONDS:
                    continue   # できたばかり: 別に起動しているアプリが取り込みの途中（DB の行を作る前）かもしれない
            except OSError:
                continue
            try:
                path.unlink()
            except OSError:
                continue
            removed += 1
    return removed


# ---- Windows のパスの長さ ------------------------------------------------------------
# 長いパスが有効でない Windows では、パス全体が 260 文字を超えるファイルを書けない。
# 置き場所（DATA_DIR 配下）に md を書く前に確かめるために使う（tables/pipeline.py）。

MAX_PATH_CHARS = 259


def _long_paths_enabled() -> bool:
    import os
    if os.name != "nt":
        return True
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem") as key:
            value, _kind = winreg.QueryValueEx(key, "LongPathsEnabled")
        return bool(value)
    except OSError:
        return False


def path_limit() -> int | None:
    """書けるファイルのパスの長さの上限（上限が無ければ None）。"""
    return None if _long_paths_enabled() else MAX_PATH_CHARS


# ====================================================================================================
# 元 core/jobs.py
# バックグラウンドジョブ（ワーカースレッド＋キュー。AI整形だけ別の列で動かす）。
#
# - 状態・進捗・一時停止/中止の要求は jobs テーブルに持つ（画面はポーリングで読む）。
# - fn(ctx) は start_job を呼んだアプリの app_context 内で実行する。fn 内で DB を使うときは connect()。
# - 実行中は一定間隔で heartbeat_at を更新する。起動時に recover_interrupted() で止まったジョブを「中断」にする。
#   動いている間も、このプロセスが持っていないジョブの heartbeat が古ければ、参照したときに「中断」にする
#   （アプリを閉じてすぐ起動し直したとき、前のプロセスのジョブが「処理中」のまま残らないように）。
# - AI整形（ai_format）は何時間もかかったり一時停止したりするので、読み込み・プレビュー・Markdown作成とは
#   別の列（lane）で動かす（止めている間に、ほかの取り込みの作業まで止まらないように）。
#   同じ取り込みのジョブは、列が違っても同時には動かさない（前のジョブが終わるまで待たせる）。
# - 社内LANのサーバーで数人が同時に使う（2026-09-20 の利用者の指示）ので、列ごとに数本のワーカーを持つ
#   （JOB_WORKERS。既定 3）。誰かが3万行の表を読んでいる間、ほかの人の読み込みが待たされないようにする。
#   取り込みごとの直列（同じ取り込みを2つのジョブが同時に触らない）は、ワーカーが増えても守る。
# ====================================================================================================

HEARTBEAT_INTERVAL = 30          # 秒。fn が呼ばなくても実行中はこの間隔で更新する
STALE_AFTER = timedelta(minutes=2)
PAUSE_POLL = 0.3                 # 一時停止中に再開・中止を確かめる間隔（秒）
# 進捗・メッセージ・heartbeat の書き込みが「database is locked」になったときの再試行（回数と待ち秒の単位）。
# 取れなければその1回の更新だけ見送る（進捗の表示のためにジョブ全体を失敗にしない）
SOFT_UPDATE_RETRIES = 2
SOFT_UPDATE_BACKOFF = 0.2

ACTIVE_STATUSES = ("queued", "running", "paused")
TERMINAL_STATUSES = ("done", "failed", "cancelled", "interrupted")

STATUS_LABELS = {
    "queued": "待機中",
    "running": "処理中",
    "paused": "一時停止中",
    "done": "完了",
    "failed": "エラー",
    "cancelled": "中止",
    "interrupted": "中断",
}

# ジョブの種類 → 動かす列（書いていない種類は "main"）
LANE_OF_KIND = {"ai_format": "ai"}

# 列ごとのワーカー本数（app.config["JOB_WORKERS"]）。数人が同時に使うので、1本だと
# 誰かの長い読み込みでほかの人が待たされる。増やしすぎても DB とディスクの取り合いになるだけ
DEFAULT_WORKERS = 3
MAX_WORKERS = 8
IDLE_POLL = 1.0                  # 待つものが何も無いときに、キューを見直す間隔（秒）

_queues: dict[str, queue.Queue] = {}
# キューから取り出したが、まだ始めていないジョブ（列ごと。同じ列のワーカーで分け合う）
_pending: dict[str, list] = {}
# (列, 何本目) → ワーカーのスレッド
_workers: dict[tuple[str, int], threading.Thread] = {}
_worker_lock = threading.Lock()
# 各ワーカーがいま実行している取り込み（ref_type, ref_id）。同じ取り込みのジョブを2つ同時に動かさないため
_executing: dict[tuple[str, int], tuple] = {}
# 各ワーカーがいま実行しているジョブ（DB のパス, job_id）。待機中のジョブに「何を待っているか」を出すため
_executing_job: dict[tuple[str, int], tuple[str, int]] = {}
# このプロセスが登録したジョブ（DB のパス, job_id）。これ以外で heartbeat が古いものは持ち主がいない
_owned: set[tuple[str, int]] = set()


class JobCancelled(Exception):
    """中止の要求を受けて fn の処理を打ち切るときに投げる。"""


class JobError(Exception):
    """利用者にそのまま見せてよい日本語メッセージを持つエラー。

    これを継承した例外（ai.AIJobError）は、そのメッセージを
    画面にそのまま出す。それ以外の例外は Python の例外名が画面に出ないようにし、内容はログにだけ残す。
    """


def error_message(exc: Exception) -> str:
    """ジョブの失敗を画面に出す日本語の文にする。詳しい内容（例外名）はログに残してあるので出さない。"""
    text = " ".join(str(exc).split())
    if isinstance(exc, JobError) and text:
        return text
    return "処理中にエラーが発生しました。もう一度実行してください（詳しい内容はアプリのログに記録しました）"


def _now() -> str:
    return now()


class JobContext:
    """fn に渡す操作口。スレッドをまたいで呼んでもよい（内部でロック）。"""

    def __init__(self, job_id: int, conn=None):
        self.job_id = job_id
        self._conn = conn or connect()
        self._db = _db_key()  # 待機中のジョブの heartbeat を更新するとき、このプロセスのものを選ぶため
        self._lock = threading.Lock()
        row = self._row()
        self._progress: dict = json.loads(row["progress_json"] or "{}") if row else {}
        self.params: dict = json.loads(row["params_json"] or "{}") if row else {}

    # ---- 内部 ----
    def _row(self):
        with self._lock:
            return self._conn.execute("SELECT * FROM jobs WHERE id = ?", (self.job_id,)).fetchone()

    def _update(self, sql_sets: str, args=()) -> None:
        with self._lock:
            self._conn.execute(f"UPDATE jobs SET {sql_sets} WHERE id = ?", (*args, self.job_id))
            self._conn.commit()

    def _soft_update(self, sql_sets: str, args=()) -> None:
        """進捗など、1回飛ばしても次の更新で追いつく書き込み。ロック中なら少し待って再試行し、だめなら見送る。"""
        for attempt in range(SOFT_UPDATE_RETRIES + 1):
            try:
                self._update(sql_sets, args)
                return
            except sqlite3.OperationalError as exc:
                text = str(exc).lower()
                if "locked" not in text and "busy" not in text:
                    raise
                with self._lock:
                    try:
                        self._conn.rollback()
                    except sqlite3.Error:
                        pass
                if attempt < SOFT_UPDATE_RETRIES:
                    time.sleep(SOFT_UPDATE_BACKOFF * (attempt + 1))
        logging.getLogger(__name__).warning("ジョブ %s の進捗の更新を見送りました（データベースが使用中）", self.job_id)

    def _flags(self) -> tuple[bool, bool]:
        row = self._row()
        if row is None:
            return True, False  # ジョブが消された → 中止扱い
        return bool(row["cancel_requested"]), bool(row["pause_requested"])

    def close(self) -> None:
        with self._lock:
            self._conn.close()

    # ---- 公開 ----
    def progress(self, **kw) -> None:
        """進捗を上書きマージして保存する（例: done=10, total=100, phase="読み込み"）。heartbeat も更新。"""
        self._progress.update(kw)
        ts = _now()
        self._soft_update("progress_json = ?, heartbeat_at = ?, updated_at = ?",
                          (json.dumps(self._progress, ensure_ascii=False, default=str), ts, ts))

    def message(self, text: str) -> None:
        self._soft_update("message = ?, updated_at = ?", (text, _now()))

    def heartbeat(self) -> None:
        self._soft_update("heartbeat_at = ?", (_now(),))

    def touch_queued(self) -> None:
        """このプロセスのキューで待っているジョブの heartbeat も更新する。

        待機中のジョブは自分では heartbeat を打たない。誤って2つ目を起動したとき、その起動時の
        recover_interrupted（別プロセス）が、1つ目で2分以上待っているジョブを「中断」にしないように。
        """
        with _worker_lock:
            owned = {job_id for db, job_id in _owned if db == self._db}
        if not owned:
            return
        with self._lock:
            ids = [r[0] for r in self._conn.execute("SELECT id FROM jobs WHERE status = 'queued'") if r[0] in owned]
            ts = _now()
            for i in range(0, len(ids), 500):  # SQLite の変数の上限を超えないよう分ける
                part = ids[i:i + 500]
                self._conn.execute(f"UPDATE jobs SET heartbeat_at = ? WHERE status = 'queued' AND id IN "
                                   f"({', '.join('?' for _ in part)})", (ts, *part))
            self._conn.commit()

    def is_cancelled(self) -> bool:
        return self._flags()[0]

    def should_stop(self) -> bool:
        """中止または一時停止が要求されていれば True（新しい処理を出さないための判定）。"""
        cancel, pause = self._flags()
        return cancel or pause

    def wait_if_paused(self) -> bool:
        """一時停止の要求があれば再開まで待つ。続けてよければ True、中止なら False。"""
        cancel, pause = self._flags()
        if cancel:
            return False
        if not pause:
            return True
        self._update("status = 'paused', heartbeat_at = ?, updated_at = ?", (_now(), _now()))
        last_beat = time.monotonic()
        while True:
            time.sleep(PAUSE_POLL)
            cancel, pause = self._flags()
            if cancel:
                return False
            if not pause:
                self._update("status = 'running', heartbeat_at = ?, updated_at = ?", (_now(), _now()))
                return True
            if time.monotonic() - last_beat >= HEARTBEAT_INTERVAL:
                self.heartbeat()
                last_beat = time.monotonic()

    def check_cancel(self) -> None:
        """一時停止中なら待ち、中止が要求されていれば JobCancelled を投げる。"""
        if not self.wait_if_paused():
            raise JobCancelled()


# ---- 実行 -----------------------------------------------------------------------

def start_job(kind: str, ref_type: str, ref_id: int | None, fn: Callable[[JobContext], dict | None],
              params: dict | None = None) -> int:
    """ジョブを登録してキューに入れ、job_id を返す。app_context 内で呼ぶ。"""
    app = current_app._get_current_object()
    ts = _now()
    conn = connect()
    try:
        cur = conn.execute(
            """INSERT INTO jobs (kind, ref_type, ref_id, status, params_json, progress_json, created_at, updated_at)
               VALUES (?, ?, ?, 'queued', ?, '{}', ?, ?)""",
            (kind, ref_type or "", ref_id, json.dumps(params or {}, ensure_ascii=False, default=str), ts, ts),
        )
        conn.commit()
        job_id = cur.lastrowid
    finally:
        conn.close()
    lane = LANE_OF_KIND.get(kind, "main")
    ref = (ref_type or "", ref_id) if ref_id is not None else None
    with _worker_lock:
        _owned.add((_db_key(), job_id))
        q = _queues.setdefault(lane, queue.Queue())
    q.put((app, job_id, fn, ref, _db_key()))
    _ensure_workers(lane)
    return job_id


def _db_key() -> str:
    return str(current_app.config["DATABASE"])


def worker_count() -> int:
    """1つの列で同時に動かすジョブの本数（app.config["JOB_WORKERS"]、既定 3）。"""
    try:
        value = int(current_app.config.get("JOB_WORKERS") or DEFAULT_WORKERS)
    except (KeyError, RuntimeError, TypeError, ValueError):
        value = DEFAULT_WORKERS
    return max(1, min(MAX_WORKERS, value))


def _ensure_workers(lane: str) -> None:
    """その列のワーカーを必要な本数まで立てる（落ちていたら立て直す）。"""
    count = worker_count()
    with _worker_lock:
        for index in range(count):
            key = (lane, index)
            worker = _workers.get(key)
            if worker is None or not worker.is_alive():
                worker = threading.Thread(target=_worker_loop, args=(lane, index), daemon=True,
                                          name=f"job-worker-{lane}-{index}")
                _workers[key] = worker
                worker.start()


def _next_item(lane: str, key: tuple[str, int]):
    """次に実行するものを取り出す。同じ取り込みを別のワーカーが実行中なら、その取り込みの分は後回しにする。

    先頭から見て最初に動けるものを選ぶので、同じ取り込みの後ろのジョブが前のジョブを追い越すことはない。
    待ち行列（_pending）は同じ列のワーカーで分け合うので、どのワーカーも必ず時間を区切って見直す
    （別のワーカーが取り出して置いた分に、いつまでも誰も手を付けない、が起きないように）。
    """
    q = _queues[lane]
    while True:
        with _worker_lock:
            pending = _pending.setdefault(lane, [])
            try:
                while True:
                    pending.append(q.get_nowait())
            except queue.Empty:
                pass
            busy = {ref for other, ref in _executing.items() if other != key and ref is not None}
            for i, item in enumerate(pending):
                if item[3] is None or item[3] not in busy:
                    _executing[key] = item[3]
                    _executing_job[key] = (item[4], item[1])
                    return pending.pop(i)
            waiting = bool(pending)
        try:
            item = q.get(timeout=PAUSE_POLL if waiting else IDLE_POLL)
        except queue.Empty:
            continue
        with _worker_lock:
            _pending.setdefault(lane, []).append(item)


def _worker_loop(lane: str = "main", index: int = 0) -> None:
    key = (lane, index)
    while True:
        app, job_id, fn, _ref, _db = _next_item(lane, key)
        try:
            _run(app, job_id, fn)
        except Exception:  # ワーカーは止めない
            try:
                app.logger.exception("ジョブ %s の実行管理でエラー", job_id)
            except Exception:
                pass
            _fail_left_open(app, job_id)
        finally:
            with _worker_lock:
                _executing.pop(key, None)
                _executing_job.pop(key, None)


def _fail_left_open(app, job_id: int) -> None:
    """実行管理の途中で落ちた（最後の状態の書き込みが DB のロックで失敗した など）ジョブを「エラー」で閉じる。

    閉じないと「処理中」のまま残り、中止も再実行もできなくなる。ここも失敗したら諦める（ログは残してある）。
    """
    try:
        with app.app_context():
            conn = connect()
            try:
                marks = ", ".join("?" for _ in ACTIVE_STATUSES)
                conn.execute(f"UPDATE jobs SET status = 'failed', pause_requested = 0, message = ?, updated_at = ? "
                             f"WHERE id = ? AND status IN ({marks})",
                             (error_message(RuntimeError()), _now(), job_id, *ACTIVE_STATUSES))
                conn.commit()
            finally:
                conn.close()
    except Exception:
        pass


class _Ticker:
    """fn が長い呼び出しで止まっていても heartbeat を更新し続ける（このプロセスの待機中のジョブの分も）。"""

    def __init__(self, ctx: JobContext):
        self._ctx = ctx
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._loop, name=f"job-{ctx.job_id}-heartbeat", daemon=True)

    def _loop(self) -> None:
        while not self._stop.wait(HEARTBEAT_INTERVAL):
            try:
                self._ctx.heartbeat()
                # 待機中のジョブは、このプロセスでいずれかのジョブが動いている（か一時停止している）間だけ
                # 待たされる。そのジョブの Ticker が代わりに更新する
                self._ctx.touch_queued()
            except Exception:
                return

    def __enter__(self):
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._thread.join(timeout=5)


def _run(app, job_id: int, fn: Callable[[JobContext], dict | None]) -> None:
    with app.app_context():
        ctx = JobContext(job_id)
        try:
            row = ctx._row()
            if row is None or row["status"] != "queued":
                return  # 待機中に中止・中断された
            if row["cancel_requested"]:
                ctx._update("status = 'cancelled', message = ?, updated_at = ?", ("中止しました", _now()))
                return
            ts = _now()
            with ctx._lock:  # 確認と開始の間に中止されていたら実行しない
                cur = ctx._conn.execute("UPDATE jobs SET status = 'running', heartbeat_at = ?, updated_at = ? "
                                        "WHERE id = ? AND status = 'queued'", (ts, ts, job_id))
                ctx._conn.commit()
            if cur.rowcount == 0:
                return
            try:
                with _Ticker(ctx):
                    result = fn(ctx)
            except JobCancelled:
                ctx._update("status = 'cancelled', pause_requested = 0, message = ?, updated_at = ?",
                            ("中止しました", _now()))
                return
            except Exception as exc:
                app.logger.exception("ジョブ %s（%s）でエラー", job_id, row["kind"])
                ctx._update("status = 'failed', pause_requested = 0, message = ?, updated_at = ?",
                            (error_message(exc), _now()))
                return
            if result is not None:
                ctx.progress(result=result)
            if ctx.is_cancelled():
                ctx._update("status = 'cancelled', pause_requested = 0, message = CASE WHEN message = '' "
                            "THEN '中止しました' ELSE message END, updated_at = ?", (_now(),))
            else:
                ctx._update("status = 'done', pause_requested = 0, updated_at = ?", (_now(),))
        finally:
            ctx.close()


# ---- 参照・操作（画面から） --------------------------------------------------------

def _decode(row) -> dict | None:
    if row is None:
        return None
    job = dict(row)
    job["params"] = json.loads(job.get("params_json") or "{}")
    job["progress"] = json.loads(job.get("progress_json") or "{}")
    job["result"] = job["progress"].get("result")
    job["status_label"] = STATUS_LABELS.get(job["status"], job["status"])
    job["finished"] = job["status"] in TERMINAL_STATUSES
    return job


def _with_conn(fn):
    conn = connect()
    try:
        return fn(conn)
    finally:
        conn.close()


INTERRUPTED_MESSAGE = "アプリの終了などで処理が中断されました。もう一度実行してください"


def _reap_orphan(conn, row):
    """持ち主のいないジョブ（このプロセスが登録しておらず、heartbeat が2分以上古い）を「中断」にして返す。

    起動時の recover_interrupted は2分以上古いものしか直さない（誤って2つ目を起動したときに、動いている
    1つ目のジョブを中断にしないため）。閉じてすぐ起動し直すと前のジョブが「処理中」のまま残るので、
    参照したときにもう一度確かめる。
    """
    if row is None or row["status"] not in ACTIVE_STATUSES:
        return row
    with _worker_lock:
        if (_db_key(), row["id"]) in _owned:
            return row
    cutoff = (datetime.now() - STALE_AFTER).isoformat(timespec="seconds")
    if (row["heartbeat_at"] or row["updated_at"] or row["created_at"] or "") >= cutoff:
        return row
    marks = ", ".join("?" for _ in ACTIVE_STATUSES)
    conn.execute(f"UPDATE jobs SET status = 'interrupted', pause_requested = 0, message = ?, updated_at = ? "
                 f"WHERE id = ? AND status IN ({marks})", (INTERRUPTED_MESSAGE, _now(), row["id"], *ACTIVE_STATUSES))
    conn.commit()
    return conn.execute("SELECT * FROM jobs WHERE id = ?", (row["id"],)).fetchone()


JOB_KIND_LABELS = {"ai_format": "AI整形"}


def _waiting_note(conn, job: dict | None) -> dict | None:
    """待機中のジョブに、何が終わるのを待っているかを付ける（job["waiting_for"] と、空なら job["message"]）。

    AI整形は一時停止している間も列と取り込みを使ったままなので、その後ろのジョブは「待機中」のまま動かない。
    理由が画面に出ないと、利用者は止まっているのか分からない（再開・中止すれば動き出すことも伝える）。
    """
    if job is None or job["status"] != "queued":
        return job
    lane = LANE_OF_KIND.get(job["kind"], "main")
    ref = (job.get("ref_type") or "", job.get("ref_id")) if job.get("ref_id") is not None else None
    db_key = _db_key()
    with _worker_lock:
        running = dict(_executing)
        running_jobs = dict(_executing_job)
        lane_size = sum(1 for k, w in _workers.items() if k[0] == lane and w.is_alive())
    mine = {k: v for k, v in running_jobs.items() if v[0] == db_key and v[1] != job["id"]}
    blocker, same_ref = None, False
    for key, owner in mine.items():        # 同じ取り込みを別のワーカーが実行中
        if ref is not None and running.get(key) == ref:
            blocker, same_ref = owner[1], True
            break
    if blocker is None:
        # 同じ列のワーカーが全部ふさがっている（空くまで動けない）
        busy_here = [owner[1] for key, owner in mine.items() if key[0] == lane]
        if busy_here and len(busy_here) >= max(1, lane_size):
            blocker = busy_here[0]
    if blocker is None:
        return job
    row = conn.execute("SELECT id, kind, ref_type, ref_id, status FROM jobs WHERE id = ?", (blocker,)).fetchone()
    if row is None or row["status"] not in ACTIVE_STATUSES:
        return job
    job["waiting_for"] = {"job_id": row["id"], "kind": row["kind"], "ref_type": row["ref_type"],
                          "ref_id": row["ref_id"], "status": row["status"], "same_ref": same_ref}
    if not job.get("message"):
        label = JOB_KIND_LABELS.get(row["kind"], "前の処理")
        whose = "この取り込みの" if same_ref else ("別の取り込みの" if row["kind"] in JOB_KIND_LABELS else "")
        if row["status"] == "paused":
            job["message"] = (f"{whose}{label}が一時停止中のため待っています。"
                              f"{label}の画面で再開するか中止すると始まります")
        else:
            job["message"] = f"{whose}{label}が終わるのを待っています"
    return job


def fits_row_id(value) -> bool:
    """SQLite の整数（8バイト）に収まる番号か。

    URL の番号をそのまま問い合わせに渡すので、桁が大きすぎると OverflowError で 500 になっていた。
    収まらない番号は「無い」として扱う（2026-09-26 の総ざらいで実測）。
    """
    return isinstance(value, int) and -2 ** 63 <= value < 2 ** 63


def get_job(job_id: int) -> dict | None:
    """ジョブ1件（params / progress / result / status_label / finished を付けて返す）。"""
    if not fits_row_id(job_id):
        return None
    return _with_conn(lambda c: _waiting_note(c, _decode(_reap_orphan(
        c, c.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()))))


def latest_job(ref_type: str, ref_id: int, kind: str | None = None) -> dict | None:
    sql, args = "SELECT * FROM jobs WHERE ref_type = ? AND ref_id = ?", [ref_type, ref_id]
    if kind:
        sql += " AND kind = ?"
        args.append(kind)
    return _with_conn(lambda c: _waiting_note(c, _decode(_reap_orphan(
        c, c.execute(sql + " ORDER BY id DESC LIMIT 1", args).fetchone()))))


def _request(job_id: int, sets: str, statuses=ACTIVE_STATUSES) -> bool:
    def run(conn):
        marks = ", ".join("?" for _ in statuses)
        cur = conn.execute(f"UPDATE jobs SET {sets}, updated_at = ? WHERE id = ? AND status IN ({marks})",
                           (_now(), job_id, *statuses))
        conn.commit()
        return cur.rowcount > 0
    return _with_conn(run)


def request_pause(job_id: int) -> bool:
    return _request(job_id, "pause_requested = 1", ("queued", "running"))


def request_resume(job_id: int) -> bool:
    return _request(job_id, "pause_requested = 0")


def request_cancel(job_id: int) -> bool:
    """中止を要求する。待機中のジョブはその場で中止にする。"""
    ok = _request(job_id, "cancel_requested = 1")
    _request(job_id, "status = 'cancelled', message = '中止しました'", ("queued",))
    return ok


def recover_interrupted() -> int:
    """起動時：実行中・待機中・一時停止中のまま heartbeat が2分以上古いジョブを「中断」にする。件数を返す。"""
    cutoff = (datetime.now() - STALE_AFTER).isoformat(timespec="seconds")

    def run(conn):
        marks = ", ".join("?" for _ in ACTIVE_STATUSES)
        cur = conn.execute(
            f"""UPDATE jobs SET status = 'interrupted', pause_requested = 0, message = ?, updated_at = ?
                WHERE status IN ({marks}) AND COALESCE(heartbeat_at, updated_at, created_at) < ?""",
            (INTERRUPTED_MESSAGE, _now(), *ACTIVE_STATUSES, cutoff),
        )
        conn.commit()
        return cur.rowcount
    return _with_conn(run)


def reap_orphan_jobs() -> int:
    """持ち主のいないジョブ（このプロセスが登録しておらず、生存印が2分以上古い）を「中断」にする。

    起動をすり抜けたジョブ（止めた直後に起動し直すと、生存印がまだ新しくて recover_interrupted に
    引っかからない）は、これまでブラウザがそのジョブを見に来たときしか閉じられなかった。
    見に来る人がいないと「実行中」のまま残り、その取り込みが片付けの対象から外れ続けて、
    元ファイル（最大200MB）と作った md が次の再起動まで消えなかった（2026-09-23 のレビュー）。
    このプロセスが動かしているジョブ（_owned）は、生存印が古くても閉じない。
    """
    cutoff = (datetime.now() - STALE_AFTER).isoformat(timespec="seconds")

    def run(conn):
        marks = ", ".join("?" for _ in ACTIVE_STATUSES)
        rows = conn.execute(
            f"""SELECT id FROM jobs WHERE status IN ({marks})
                AND COALESCE(heartbeat_at, updated_at, created_at) < ?""",
            (*ACTIVE_STATUSES, cutoff)).fetchall()
        with _worker_lock:
            ids = [r[0] for r in rows if (_db_key(), r[0]) not in _owned]
        if not ids:
            return 0
        conn.execute(
            f"""UPDATE jobs SET status = 'interrupted', pause_requested = 0, message = ?, updated_at = ?
                WHERE id IN ({", ".join("?" for _ in ids)})""",
            (INTERRUPTED_MESSAGE, _now(), *ids))
        conn.commit()
        return len(ids)
    return _with_conn(run)


def wait_job(job_id: int, timeout: float = 30.0, statuses=TERMINAL_STATUSES) -> dict | None:
    """指定の状態になるまで待つ（テスト・同期実行用）。時間切れなら最後の状態を返す。"""
    deadline = time.monotonic() + timeout
    while True:
        job = get_job(job_id)
        if job is None or job["status"] in statuses or time.monotonic() >= deadline:
            return job
        time.sleep(0.05)


# ====================================================================================================
# 元 core/purge.py
# 取り込んだデータを消す（design.md 3.3「データを残さない」）。
#
# このアプリはダウンロードが終わった時点で、その取り込みに属するものをすべて消す。
# 消すもの: アップロードした元のファイル、imports/<id>/（読み込んだ行・控え・作った md）、
#           DB の行（table_imports / jobs / ai_items / llm_calls とそれらの子）。
# 残すもの: ブラウザごとの AI接続（ai_connections）。これは「設定」でデータではない
#           （SETTINGS_TABLES。取り込みを指す列が後から足されても、ここからは消さない）。
#
# DB の行は表の名前を決め打ちせず、その取り込みを指す列（import_id / table_import_id）を持つ表を
# sqlite_master から探して消す（後から表が増えても消し残さないため）。
# ====================================================================================================

log = logging.getLogger(__name__)

IMPORT_REF_COLUMNS = ("import_id", "table_import_id")
# 取り込みを指す列があっても、まとめては消さない表。llm_calls は AI の生の応答で、同じ文面の行なら
# 別の取り込みの ai_items が同じ応答を使っていることがある（消すと再実行で再課金になる）。
# 持ち主が消えた分は _delete_orphan_llm_calls が「参照されていなければ消す」で扱う。
SHARED_TABLES = ("llm_calls",)
# 「設定」の表。取り込みを捨てる片付け（purge_* / sweep_stale / purge_session）では決して消さない。
# ai_connections はブラウザごとの AI接続（APIキー・接続先。利用者の指示 2026-09-21「ずっと保持」）で、
# 取り込みと同じ session_id を持ち主に持つが、その人の取り込みを捨てても残す。
SETTINGS_TABLES = ("ai_connections",)


def _tables_with_column(db, column: str) -> list[str]:
    names = []
    for (table,) in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall():
        if table.startswith("sqlite_") or table in SHARED_TABLES or table in SETTINGS_TABLES:
            continue
        if column in {row[1] for row in db.execute(f'PRAGMA table_info("{table}")')}:
            names.append(table)
    return names


def _delete_by_columns(db, columns, value) -> int:
    removed = 0
    for column in columns:
        for table in _tables_with_column(db, column):
            removed += db.execute(f'DELETE FROM "{table}" WHERE "{column}" = ?', (value,)).rowcount
    return removed


def _delete_jobs(db, ref_type: str, ref_id: int) -> int:
    return db.execute("DELETE FROM jobs WHERE ref_type = ? AND ref_id = ?", (ref_type, ref_id)).rowcount


_UNREFERENCED_LLM_CALLS = ("cache_key NOT IN (SELECT cache_key FROM ai_items WHERE cache_key IS NOT NULL)")
# 持ち主の取り込みがもう無い応答（持ち主の分からない古いDBの分＝NULL も含む）
NO_LIVE_OWNER = "(import_id IS NULL OR import_id NOT IN (SELECT id FROM table_imports))"


def _ai_work_alive(db) -> bool:
    """まだある取り込みに、AI整形の結果かAI整形のジョブ（一時停止・中断を含む）が残っているか。

    ai_items が覚えているのは最後に使った応答のキーだけで、再依頼で直した行の1回目の応答や、
    一時停止の直前に受け取った応答（ai_items の行がまだ無い）は、どこからも参照されない。
    それでも再実行・再開ではキャッシュとして引く（引けないと再課金になる）ので、
    AI整形が残っている取り込みがある間は、持ち主の分からない（import_id が NULL の古いDBの）
    参照されない応答を消さない。持ち主の分かる応答は、その取り込みを消すときに一緒に消える。
    """
    return db.execute(
        "SELECT 1 FROM ai_items WHERE import_id IN (SELECT id FROM table_imports) "
        "UNION ALL SELECT 1 FROM jobs WHERE kind = 'ai_format' AND ref_type = 'table_import' "
        "AND ref_id IN (SELECT id FROM table_imports) LIMIT 1").fetchone() is not None


def _delete_orphan_llm_calls(db, keys=()) -> int:
    """どの ai_items からも参照されていない AI の生の応答を消す（試し実行の分もここで消える）。

    keys: 消した取り込みが使っていた応答のキー。ほかの行が使っておらず、まだある取り込みが払った分でも
    なければ、いつでもすぐ消す。消す取り込みは、別の取り込みが払った応答をキャッシュとして引いている
    ことがある（同じ文面なら同じキーになる）ので、払った取り込みがまだあるうちは消さない
    （消すと、その取り込みの再開・再実行で同じ応答をもう一度買うことになる）。その応答は、払った
    取り込みを消すときに下の「持ち主の取り込みがもう無い分」で消えるので、残り続けることはない。
    もう無い取り込みが払った応答（llm_calls.import_id）も、ほかから使われていなければすぐ消す。
    ai_items は最後に使った応答のキーしか覚えていないので、再依頼で直した行の1回目の応答や、
    一時停止の直前に受け取った応答は keys に入らない。持ち主の列があれば、それも一緒に消せる。
    持ち主の分からない応答（古いDBの分）は、AI整形が残っている取り込みが1件も無いときだけ消す（_ai_work_alive）。
    """
    removed = 0
    keys = sorted({k for k in keys if k})
    for start in range(0, len(keys), 500):
        chunk = keys[start:start + 500]
        marks = ", ".join("?" for _ in chunk)
        removed += db.execute(f"DELETE FROM llm_calls WHERE cache_key IN ({marks}) "
                              f"AND {NO_LIVE_OWNER} AND {_UNREFERENCED_LLM_CALLS}", chunk).rowcount
    # 持ち主の取り込みがもう無い分（design.md 3.3「取り込んだデータはダウンロードが終わった時点で消す」）
    removed += db.execute(f"DELETE FROM llm_calls WHERE import_id IS NOT NULL "
                          f"AND import_id NOT IN (SELECT id FROM table_imports) "
                          f"AND {_UNREFERENCED_LLM_CALLS}").rowcount
    # まだ使われていて残した分は、持ち主が居なくなったので「持ち主なし」に戻す
    # （使っている取り込みを消すときに、その keys で消える）
    db.execute("UPDATE llm_calls SET import_id = NULL WHERE import_id IS NOT NULL "
               "AND import_id NOT IN (SELECT id FROM table_imports)")
    if not _ai_work_alive(db):
        removed += db.execute(f"DELETE FROM llm_calls WHERE {_UNREFERENCED_LLM_CALLS}").rowcount
    return removed


def delete_orphan_ai_items(db) -> int:
    """もう無い取り込みを指す AI整形の結果を消す（取り込みを消したあとに、動いていた AI の書き込みが残った分など）。

    import_id の無い行（古いDBの分。持ち主の取り込みが分からない）は、その取り込み設定の取り込みが
    1件も残っていなければ消す（残っていれば、その取り込みを消すときに purge_table_import が消す）。
    """
    removed = db.execute("DELETE FROM ai_items WHERE import_id IS NOT NULL "
                         "AND import_id NOT IN (SELECT id FROM table_imports)").rowcount
    removed += db.execute("DELETE FROM ai_items WHERE import_id IS NULL AND template_id NOT IN "
                          "(SELECT template_id FROM table_imports WHERE template_id IS NOT NULL)").rowcount
    return removed


def sweep_orphan_ai(db) -> int:
    """持ち主の無い AI整形の結果と、どこからも使われない AI の生の応答を消して縮める（戻り値: 消した件数）。

    取り込み設定を消したとき（ai_items は消えるが llm_calls が残る）や起動時の片付けで使う。
    """
    removed = delete_orphan_ai_items(db) + _delete_orphan_llm_calls(db)
    db.commit()
    if removed:
        _shrink(db)
    return removed


def _mark_incomplete() -> None:
    """この要求の中で、消し切れなかったもの（掴まれていて消せなかったファイル）があったと印を付ける。"""
    if has_app_context():
        g.purge_incomplete = True


def purge_incomplete() -> bool:
    """この要求の中で消したときに、消し切れなかったファイルがあったか（保存先フォルダの結果の画面で知らせるため）。

    消せなかったファイルも中身は 0 バイトに切り詰めてあり、次の起動時に片付く。例外にはしない
    （消すこと自体は DB の行まで終わっていて、画面を 500 にするより「消し切れなかった」と知らせる方がよい）。
    """
    return has_app_context() and bool(g.get("purge_incomplete"))


def _truncate_leftovers(folder: Path) -> None:
    """消し切れなかったフォルダに残ったファイルを 0 バイトに切り詰める（切り詰められないものは諦める）。"""
    for path in folder.rglob("*"):
        try:
            if path.is_file():
                with open(path, "r+b") as f:
                    f.truncate(0)
        except OSError:
            pass


def purge_after_send(response, fn, *args):
    """本文を最後まで送り終えてから消す（design.md 3.3）。

    送る前に消すと、ブラウザを閉じた・通信が切れた・保存先の空きが足りない、のどれでも
    「手元にファイルが無いのにサーバー側は消えている」ことになり、取り戻せない（再ダウンロードもできない）。
    本文を最後まで渡しきったときだけ消し、途中で切れたときは残す（もう一度ダウンロードできる）。

    「渡しきった」はサーバー（waitress）の送信の溜めに入れ終えたことで、相手が受け取ったことではない。
    溜めの上限は app.OUTBUF_HIGH_WATERMARK（16KB）で、OS の溜めと合わせて最後の数十KBは受け取られる前に
    消すことになる（数十KBより小さい md は、途中で切れても消える。design.md 3.3）。

    send_file の応答は direct_passthrough が立っていて、そのままだと WSGI が close のときの
    呼び出し（call_on_close）を拾わないので、ここで解除して本文を包み直す。
    """
    if response.status_code != 200:
        # 206（Range で一部だけ）・304 などは本文の全部を渡していない。ダウンロードマネージャーや
        # 途中からの再開、ウイルス対策のプロキシが送ってくることがある。消すと残りを取り戻せないので消さない。
        return response
    app = current_app._get_current_object()
    body = response.response
    sent: list[bool] = []

    def stream():
        try:
            for chunk in body:
                yield chunk
            sent.append(True)   # 最後まで渡した
        finally:
            close = getattr(body, "close", None)
            if close is not None:
                close()

    response.direct_passthrough = False
    response.response = stream()

    @response.call_on_close
    def _purge() -> None:
        if not sent:
            return   # 送信が途中で終わった: 消さずに残す
        try:
            with app.app_context():
                fn(*args)
        except Exception:  # 応答は送り終えているので、ここで落としても利用者に伝える先が無い
            log.exception("ダウンロードしたデータの削除に失敗しました")

    return response


# ---- 一覧表 ------------------------------------------------------------------------

def import_dir(import_id: int) -> Path:
    return Path(current_app.config["TABLES_DIR"]) / "imports" / str(int(import_id))


def purge_table_import(import_id: int) -> int:
    """一覧表の取り込み1件を消す（元のファイル・imports/<id>/・DB の行・AI整形の控え）。取り込み設定は残す。"""
    import_id = int(import_id)
    db = get_db()
    row = db.execute("SELECT stored_path, template_id FROM table_imports WHERE id = ?", (import_id,)).fetchone()
    if row is None:
        # DB の行が無くてもフォルダが残っていることがあるので、そこだけ片付ける
        shutil.rmtree(import_dir(import_id), ignore_errors=True)
        return 0
    # DB の行を先に消し、ファイルはそのあと消す（残ったフォルダ・ファイルは起動時に片付く）
    # この取り込みが使っていた AI の応答のキー（行を消す前に控える。消したあと、ほかで使われていなければ消す）
    keys = [r[0] for r in db.execute("SELECT cache_key FROM ai_items WHERE import_id = ? AND cache_key IS NOT NULL",
                                     (import_id,))]
    if row["template_id"] is not None:
        keys += [r[0] for r in db.execute("SELECT cache_key FROM ai_items WHERE template_id = ? AND import_id IS NULL "
                                          "AND cache_key IS NOT NULL", (row["template_id"],))]
    _delete_by_columns(db, IMPORT_REF_COLUMNS, import_id)
    _delete_jobs(db, "table_import", import_id)
    removed = db.execute("DELETE FROM table_imports WHERE id = ?", (import_id,)).rowcount
    # AI整形の結果の控え（行ごとの整形結果）は取り込んだデータそのものなので一緒に消す。この取り込みの分だけ
    # （ai_items.import_id）を消すので、同じ設定で作業中の別の取り込みの結果と応答キャッシュは残る。
    # import_id が無い行は古いDBの分（持ち主が分からない）なので、その設定の分をまとめて消す。
    if row["template_id"] is not None:
        db.execute("DELETE FROM ai_items WHERE template_id = ? AND import_id IS NULL", (row["template_id"],))
    delete_orphan_ai_items(db)
    _delete_orphan_llm_calls(db, keys)
    db.commit()
    if not remove_upload(row["stored_path"]):
        _mark_incomplete()
    folder = import_dir(import_id)
    shutil.rmtree(folder, ignore_errors=True)
    if folder.exists():
        # Windows で別のスレッド・プロセスがファイルを掴んでいると消し残る。黙って成功扱いにせず記録し、
        # remove_upload と同じく残ったファイルの中身を 0 バイトに切り詰める（名前は残っても読み込んだ行・md は残さない）。
        # DB の行はもう無いので、残ったフォルダは次の起動時に remove_orphan_import_dirs が片付ける。
        _truncate_leftovers(folder)
        _mark_incomplete()
        log.warning("取り込みのフォルダを消し切れませんでした（中身は空にし、次の起動時に片付けます）: imports/%s",
                    import_id)
    _shrink(db)
    return removed


def _shrink(db) -> None:
    """消したあとの後始末（design.md 3.3）。

    - wal_checkpoint(TRUNCATE): 消す前後の内容が残る app.db-wal を切り詰める（強制終了時に拾われないように）
    - VACUUM: 解放したページを手放してファイルの大きさも戻す（中身は PRAGMA secure_delete で消えている）
    どちらも他の接続が使っていると実行できない。そのときは次に消したときに縮む。

    VACUUM は DB 全体を書き直すあいだ書き込みを止める。大きな app.db だと busy_timeout より長くなり、
    動いている AI整形などのジョブの書き込みが「database is locked」で失敗する。なので、待機中・実行中・
    一時停止中のジョブがあるときは VACUUM をしない（消した中身は secure_delete で上書き済み。
    ファイルの大きさは、ジョブが無いときの次の削除か起動時の片付けで戻る）。

    さらに、空きページが少ないうちは VACUUM をしない。ジョブが1件も無ければ上の判定を素通りするので、
    .md を1件ダウンロードするだけでも毎回 DB 全体を書き直していた
    （118MB で1秒、その間ほかの人の自動保存が最長0.75秒待たされた。2026-09-23 のレビューで実測）。
    """
    forget_id_counters(db)
    # VACUUM → checkpoint の順。逆にすると VACUUM の結果が app.db-wal に残り、
    # instance/app.db 自体は縮まない（実測 121,260KB のまま。2026-09-23 のレビュー）
    vacuum = not _jobs_active(db) and _worth_vacuum(db)
    statements = (["VACUUM"] if vacuum else []) + ["PRAGMA wal_checkpoint(TRUNCATE)"]
    for statement in statements:
        try:
            db.execute(statement)
        except sqlite3.Error:
            pass


# 空きページがこの枚数とこの割合の両方を超えたときだけ VACUUM する（小さい削除で毎回書き直さない）
VACUUM_MIN_FREE_PAGES = 512          # 4KiB ページで約 2MB
VACUUM_MIN_FREE_RATIO = 0.25


def _worth_vacuum(db) -> bool:
    """縮める価値があるか（空きページが十分たまっているか）。"""
    try:
        free = db.execute("PRAGMA freelist_count").fetchone()[0]
        total = db.execute("PRAGMA page_count").fetchone()[0]
    except (sqlite3.Error, TypeError, IndexError):
        return True    # 数えられないときは今までどおり縮める
    return bool(free >= VACUUM_MIN_FREE_PAGES and total and free >= VACUUM_MIN_FREE_RATIO * total)


def _jobs_active(db) -> bool:
    try:
        row = db.execute("SELECT 1 FROM jobs WHERE status IN ('queued', 'running', 'paused') LIMIT 1").fetchone()
        return row is not None
    except sqlite3.Error:
        return True   # 確かめられないときは、動いているかもしれないジョブを止めない側に倒す


# ---- 番号の続きを忘れる ------------------------------------------------------------
# 取り込むたびに行を作り、ダウンロードで消す表。AUTOINCREMENT の表は、消したあとも sqlite_sequence に
# 「これまでに使った一番大きい番号」が残り、何件取り込んだかの記録になってしまう（design.md 3.3「履歴は持たない」）。
WORK_TABLES = ("table_imports", "jobs", "ai_items")
# 表が空になったら、続きの番号をこの範囲の乱数にする。最初の番号（1〜）とも前回の番号とも重ならないので、
# 開いたままの古い画面やブラウザに残った画面が、消した番号で別の取り込みを開いたり書き換えたりしない
# （重なる確率は 1回あたり 取り込み件数 / 約4.5×10^15。JavaScript で正確に扱える 2^53 未満に収める）。
_ID_BASE_MIN = 2 ** 31
_ID_BASE_MAX = 2 ** 52


def forget_id_counters(db) -> int:
    """空になった取り込みの表の、番号の続き（sqlite_sequence）を乱数に置き換える。戻り値: 置き換えた表の数。

    AUTOINCREMENT は外さない（外すと、消した番号がすぐ使い回され、古い画面が別の取り込みを指す）。
    行が残っている表は置き換えない（作業中の取り込みの番号が見えている間は、続きの番号を隠しても意味が無い）。
    """
    changed = 0
    try:
        for table in WORK_TABLES:
            base = _ID_BASE_MIN + secrets.randbelow(_ID_BASE_MAX - _ID_BASE_MIN)
            # 空かどうかの確認と置き換えを1文で行う（あいだに別のスレッドが行を足しても番号が重ならない）
            changed += db.execute(f'UPDATE sqlite_sequence SET seq = ? WHERE name = ? '
                                  f'AND NOT EXISTS (SELECT 1 FROM "{table}")', (base, table)).rowcount
        db.commit()
    except sqlite3.Error:
        # sqlite_sequence がまだ無い（AUTOINCREMENT の表に1行も入れていない新しいDB）か、ロック中。
        # 次に消したとき・次の起動時にもう一度行う
        db.rollback()
        return 0
    return changed


# ---- まとめて捨てる（作業中の表示を持たない） ----------------------------------------------
# この アプリは「作業中の一覧」も「ダウンロード待ちの一覧」も持たない（利用者の指示 2026-09-20）。
# 「その場でダウンロードしない限り、その場ですぐ捨てる」（利用者の指示 2026-09-20）ので、捨てる機会は4つ:
#   - 画面を離れたとき: その人が触っていた分を捨てる（purge_session / discard_table_imports）
#   - 新しいファイルを置いたとき: 同じ人の前の分を捨てる（同上）
#   - 動いている間: IDLE_HOURS さわられていないものを捨てる（sweep_stale）
#   - 起動時: ダウンロードしていない取り込みをすべて捨てる（purge_all_pending）
# あとの2つ（時間切れ・起動時）は、動いているジョブが付いているものには手を出さない（_busy_ids）。
# 残すのは「設定」（ブラウザごとの AI接続 = SETTINGS_TABLES）だけで、これは purge の対象ではない。
# AI接続だけは、クッキーの寿命（約1年。SESSION_LIFETIME）より長くさわられていない行を
# sweep_stale_ai_connections が消す（もう戻って来ないブラウザの APIキーを DB に残さない）。

# 放っておかれた取り込みを捨てるまでの時間。数人が同時に使う社内LANの置き方（2026-09-20）では、
# 画面を閉じた合図（sendBeacon）が届かないこと（ブラウザの強制終了・スリープ・LANの切断）があるので、
# 短めにして「残らないこと」を優先する。長くしたいときはこの数字だけを変える。
IDLE_HOURS = 2
STALE_HOURS = IDLE_HOURS   # 旧名（app.py が参照している）


def _import_ids(db, where: str = "", args=()) -> list[int]:
    return [row[0] for row in db.execute(f"SELECT id FROM table_imports {where}", args)]


def purge_all_pending() -> int:
    """ダウンロードしていない取り込みをすべて捨てる。戻り値: 捨てた件数。

    ダウンロードが終わったものはその時点で消えている（purge_after_send）ので、DB に残っている
    取り込みは「途中のもの」か「確定したがダウンロードしていないもの」しかない。
    続きを開く入口（作業中の一覧）を持たないので、起動時にまとめて捨てる。

    動いているジョブが付いているものには手を出さない。起動直後は _recover_jobs が動いていた
    ジョブを「中断」に直したあとなので、ふつうは1件も残らない。それでも、別のプロセスが同じ DB を
    見ているとき（JupyterLab のターミナルで二重に起動したとき）に、他方が処理中の取り込みを
    消してしまわないようにする。
    """
    db = get_db()
    busy_imports = _busy_ids(db, "table_import")
    tables = 0
    for import_id in _import_ids(db):
        if import_id not in busy_imports:
            tables += purge_table_import(import_id)
    return tables


def _stale_before(hours: float) -> str:
    from datetime import datetime, timedelta

    return (datetime.now() - timedelta(hours=hours)).isoformat(timespec="seconds")


# 動いているジョブ（待機中・実行中・一時停止中）が指している取り込みは捨てない。
# 大きな一覧表の読み込みや AI整形は24時間を超えることがあり、途中で消すとジョブが
# 「もう無い行」を書きに行って失敗する（利用者から見れば、放っておいたら処理が消えた、になる）。
# ジョブが終われば updated_at がその時刻になるので、次の回以降に改めて対象になる。
_BUSY = ("SELECT ref_id FROM jobs WHERE ref_type = ? AND ref_id IS NOT NULL "
         "AND status IN ('queued', 'running', 'paused')")


def _busy_ids(db, ref_type: str) -> set[int]:
    try:
        return {row[0] for row in db.execute(_BUSY, (ref_type,))}
    except sqlite3.Error:
        return set()   # 確かめられないときは何も捨てない側に倒せないので、呼び出し元で空集合＝全部対象


def _sweep_orphan_uploads() -> int:
    """行の無いアップロードファイル（置いたあと DB への登録が失敗したもの）を片付ける。

    以前は起動時にしか見ていなかったので、動かしている間にできた孤児（最大200MB）が
    次の再起動まで残っていた（2026-09-23 のレビュー）。
    """
    try:
        return remove_orphan_uploads(current_app.config["UPLOAD_DIR"], known_stored_paths(get_db()))
    except (sqlite3.Error, OSError, RuntimeError, KeyError):
        return 0     # 片付けは best effort。ここで見回り全体を止めない


def sweep_stale(hours: float = STALE_HOURS) -> int:
    """しばらくさわられていない取り込みを捨てる。戻り値: 捨てた件数。

    「最後にさわった日時」（途中保存・読み取り → 確定 → 取り込み の順に見る）で切る。
    動いているジョブが付いているものは、そのジョブが終わるまで残す。
    """
    try:
        reap_orphan_jobs()   # 書き込みなので database is locked で落ちうる。片付け全体は止めない
    except sqlite3.Error:
        pass
    _sweep_orphan_uploads()
    db = get_db()
    limit = _stale_before(hours)
    busy_imports = _busy_ids(db, "table_import")
    # 見るのは updated_at が先。confirmed_at は「最後に確定した時刻」で、そのあと直し続けても
    # 進まない（2026-09-23 のレビューで実測）
    tables = 0
    for import_id in _import_ids(
            db, "WHERE COALESCE(updated_at, confirmed_at, created_at) < ?", (limit,)):
        if import_id in busy_imports:
            continue
        tables += purge_table_import(import_id)
    sweep_stale_ai_connections()
    return tables


# ブラウザごとの AI接続を消すまでの日数。クッキーの寿命（SESSION_LIFETIME = 365日）を過ぎたブラウザは
# 同じ id で戻って来られないので、その行（APIキー）を持ち続ける理由が無い。保存・接続の確認・AI整形の
# 開始のたびに updated_at が進むので、使っている人の分は消えない。
AI_CONNECTION_KEEP_DAYS = 400


def sweep_stale_ai_connections(days: float = AI_CONNECTION_KEEP_DAYS) -> int:
    """クッキーの寿命より長くさわられていないブラウザの AI接続（APIキー）を消す。戻り値: 消した件数。"""
    db = get_db()
    limit = _stale_before(days * 24)
    try:
        removed = db.execute("DELETE FROM ai_connections WHERE updated_at < ?", (limit,)).rowcount
        db.commit()
    except sqlite3.Error:
        db.rollback()
        return 0
    return removed


# ---- 使っている人ごとに捨てる ---------------------------------------------------------
# 社内LANに置いて数人が同時に使う（利用者の指示 2026-09-20）ので、「作業中のもの」は
# 全員分がひとつの DB に混ざっている。画面を離れた人の分だけを捨てられるように、
# table_imports は持ち主（session_id。current_session_id() がブラウザごとに配る）を持つ。
#
# 持ち主の列がまだ無い古い DB でも動くようにしてある:
#   - 番号を指して捨てる（discard_table_imports）… 持ち主を確かめずに捨てる
#     （1人で使っていた頃と同じ動き。指された番号は、その画面が自分で取り込んだものしかない）
#   - まとめて捨てる（purge_session）… 誰のものか分からないので何も捨てない（他人の分を巻き込まない）

SESSION_COLUMN = "session_id"


def _has_session_column(db, table: str) -> bool:
    try:
        return SESSION_COLUMN in {row[1] for row in db.execute(f'PRAGMA table_info("{table}")')}
    except sqlite3.Error:
        return False


def _owned_ids(db, table: str, ids, session_id) -> list[int]:
    """ids のうち、いま DB にあって、その人のもの（持ち主の列が無ければ確かめない）。

    もう無い番号・他人の番号はここで落ちるので、呼び出しを何度繰り返しても2回目からは何もしない
    （画面を閉じる合図は同じものが2回届くことがある）。
    """
    wanted = sorted({int(i) for i in ids})
    if not wanted:
        return []
    marks = ", ".join("?" for _ in wanted)
    sql = f'SELECT id FROM "{table}" WHERE id IN ({marks})'
    args = list(wanted)
    if session_id and _has_session_column(db, table):
        sql += f' AND "{SESSION_COLUMN}" = ?'
        args.append(session_id)
    try:
        return [row[0] for row in db.execute(sql, args)]
    except sqlite3.Error:
        return []


def _session_ids(db, table: str, session_id) -> list[int]:
    if not session_id or not _has_session_column(db, table):
        return []
    try:
        return [row[0] for row in db.execute(f'SELECT id FROM "{table}" WHERE "{SESSION_COLUMN}" = ?', (session_id,))]
    except sqlite3.Error:
        return []


def discard_table_imports(import_ids, session_id=None) -> int:
    """指された一覧表の取り込みのうち、その人のもので、処理中でないものを捨てる。戻り値: 捨てた件数。"""
    db = get_db()
    ids = _owned_ids(db, "table_imports", import_ids, session_id)
    if not ids:
        return 0
    busy = _busy_ids(db, "table_import")
    removed = 0
    for import_id in ids:
        if import_id not in busy:
            removed += purge_table_import(import_id)
    return removed


def purge_session(session_id, *, include_busy: bool = False) -> int:
    """その人の、まだダウンロードしていない取り込みをすべて捨てる。戻り値: 捨てた件数。

    捨てるのは、元のファイル・imports/<id>/（読み込んだ行・控え・作った md）・DB の行・
    AI整形の控えと、ほかから使われなくなった AI の応答（purge_table_import と同じ）。
    残すのは設定（AI接続）だけ。

    動いているジョブ（読み込み・下書き・AI整形）が付いているものは、そのジョブが壊れるので捨てない。
    捨て損ねた分は、ジョブが終わったあと sweep_stale が IDLE_HOURS で片付ける。
    どうしても今すぐ捨てるときだけ include_busy=True（ジョブは次の書き込みで失敗して終わる）。
    """
    if not session_id:
        return 0
    db = get_db()
    import_ids = _session_ids(db, "table_imports", session_id)
    if not include_busy:
        busy = _busy_ids(db, "table_import")
        import_ids = [i for i in import_ids if i not in busy]
    tables = 0
    for import_id in import_ids:
        tables += purge_table_import(import_id)
    return tables


####################################################################################################
# 画面（旧 views.py）
####################################################################################################

# 先に土台（この上の部分）を読み終えてから extract を読む（extract.py は app を import するので、
# extract.py を単独で import すると順番が逆になって ImportError になる。起動は flask --app app）
from extract import (
    group_columns,
    MAX_GROUP_COLUMNS,
    count_levels,
    has_blocking,
    suggest_columns,
    ai_point_lines,
    people_index_for,
    record_block,
    open_source,
    spec_from_suggestions,
    validate_spec,
    build_download,
    create_import,
    get_import,
    import_files,
    import_source,
    issues_csv as build_issues_csv,
    load_issues,
    load_rows,
    load_rows_page,
    md_paths,
    md_text,
    preview_signature,
    ready_preview_files,
    save_spec,
    spec_for_import,
    start_preview_job,
    start_read_job,
    start_render_job,
    update_import,
)


# ====================================================================================================
# 元 views/__init__.py
# ====================================================================================================

# ---- 使っている人ごとの作業場所 ---------------------------------------------------------
# 社内LANのサーバーで動かし、数人が同時に別々のPCから使う（2026-09-20 の利用者の指示）。
# ログインは無いので「誰か」は分からないが、「どのブラウザか」はセッションのクッキーで分かる。
# 取り込んだ表はそのブラウザのものとして持ち主を記録し、ほかのブラウザからは
# 見えない・触れないようにする（views/tables.py の 404）。
# 「設定」のうち AI接続（APIキー・接続先）も、同じ id でブラウザごとに持つ
# （利用者の指示 2026-09-21「cookieでユーザー毎に登録内容をずっと保持」。ai.py・core.ai_connections）。
# そのためクッキーは約1年もたせる（SESSION_LIFETIME。create_app が PERMANENT_SESSION_LIFETIME に入れる）。
# 長くしても分かれ方は変わらない: id は uuid4 で、署名鍵（SECRET_KEY）付きの HttpOnly・SameSite=Lax のクッキー
# にしか無く、ほかのブラウザからは推測も持ち出しもできない。

SESSION_ID_KEY = "sid"
SESSION_LIFETIME = timedelta(days=365)


def current_session_id() -> str:
    """このブラウザの作業場所の id（クッキーに無ければ作る）。クッキーは約1年もたせる。"""
    sid = session.get(SESSION_ID_KEY)
    if not isinstance(sid, str) or len(sid) != 32:
        sid = uuid4().hex
        session[SESSION_ID_KEY] = sid
    session.permanent = True   # 有効期限（PERMANENT_SESSION_LIFETIME）付きのクッキーにする。要求のたびに延びる
    return sid


def owns(row) -> bool:
    """取り込みの行がこのブラウザのものか。

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
# このアプリへ POST する」ことは防げない（AI接続先の書き換え・取り込みの削除などができてしまう）。
# ブラウザは POST に必ず Origin を付け、最近のブラウザは Sec-Fetch-Site も付けるので、
# 他サイト発と分かる書き込みだけを断る（curl などヘッダの無い要求はアプリを直接たたく操作として許す）。

SAFE_METHODS = ("GET", "HEAD", "OPTIONS")
SAME_SITE_FETCH = ("same-origin", "none")

# GET でもデータを消すルート（zip をダウンロードすると、その取り込みを消す。design.md 3.3）。
# 他サイトのページの <img>・リンク・window.open からでも GET は出せるので、書き込みと同じく発火元を確かめる。
# アドレス欄に打った URL（Sec-Fetch-Site: none）とアプリ内のクリック（same-origin）は通す。
PURGING_ENDPOINTS = frozenset({"tables.download_zip"})


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


def render_part(template: str, part: str, **ctx) -> str:
    """画面のテンプレート（base.html）の中のマクロ part_* を1つだけ描いて、HTML の断片を返す。

    段の中身は fetch でそのつど取りに来るので、ページ全体ではなく断片だけを描く。マクロは render_template と同じ
    文脈（url_for・request・ctx の変数）を見る。マクロに引数があれば ctx の同じ名前の値を渡す。
    """
    current_app.update_template_context(ctx)
    module = current_app.jinja_env.get_template(template).make_module(ctx)
    fn = getattr(module, part)
    return str(fn(**{k: ctx[k] for k in fn.arguments if k in ctx}))


# ====================================================================================================
# 元 views/tables.py
# 表の取り込み（design.md 2.4）。1画面で全部できる（利用者の指示 2026-09-20）。
#
# 画面は /tables の1枚だけ。ファイルを置く → 読み取り方 → 表の範囲 → 列の対応づけ → AI整形（任意）
# → 内容の確認 → 確定してダウンロード（zip）を、同じ画面の「段」として順に開く。
# 段の中身はこの blueprint が HTML の断片（panel）として返し、保存・実行は JSON でやりとりする。
# 画面の移動（①→②→③）は無く、URL は変わらない（static/app.js）。
# 範囲の決まり（利用者の判断）: 期間の置き換え・投入済みとの差分・取り消しはしない。取り込みごとに
# その取り込みの記録だけから全 Markdown を作り、全ファイルを zip で渡す。クロス集計と名寄せ辞書は扱わない。
# ダウンロードした取り込みのデータは、zip を送り終えたあとに消す（design.md 3.3・purge_after_send）。
# 読み込み・Markdown 作成・AI整形は core.jobs のジョブ（tables.pipeline / ai.runner）で動かし、
# 同じ画面の中に進み具合を出す。表の範囲・列の段は取り込みの控え（tables.source_cache）を通して読む。
# ====================================================================================================

tables_bp = Blueprint("tables", __name__, url_prefix="/tables")

EXCEL_EXT = {".xlsx", ".xlsm"}
GRID_ROWS = 60
GRID_COLS = 30
GRID_CELL_CHARS = 40
DATA_PAGE = 100
ISSUES_SHOWN = 200

# 画面で選べる役割（列の対応づけ）。ここに無い役割（人名・数値・区分など）は候補のまま使い、画面では「その他」に見せる。
# 名前はどんな表にも当てはまる言い方にする（利用者の指示 2026-09-21。entity は「設備」、log は
# 「AI整形の対象（追記ログ）」と呼んでいたが、設備の記録以外の表には当てはまらなかった）。
# 中で使う名前（key/date/entity/log/attribute）は変えない（前に保存した取り込み設定がそのまま読める）。
SCREEN_ROLES = [("key", "識別番号"), ("date", "日付"), ("entity", "対象（顧客・案件・製品など）"),
                ("log", "経過の記録（1つのセルに日付ごとに書き足した列）"), ("attribute", "その他")]
SCREEN_ROLE_KEYS = {role for role, _label in SCREEN_ROLES}
ROW_KIND_LABELS = {"header": "見出し", "data": "データ", "subtotal": "小計・合計", "note": "注記", "continuation": "継続行",
               "excluded": "除外", "title": "表題", "blank": "空行"}
TABLE_KIND_LABELS = {"list": "一覧表", "crosstab": "クロス集計", "form_like": "項目名と値の縦並び", "unknown": "不明"}
SCOPE_LABELS = {"pending": "まだ整形していない行と、内容が変わった行", "errors": "エラーになった行だけ",
                "flagged": "要確認の行だけ", "all": "すべての行をやり直す"}
# CSV の文字コード・区切り文字は画面の選択肢だけを受け付ける（他の値は読み込みで落ちて画面が開けなくなるため）
# ダウンロードでデータが消えることの案内（design.md 3.3）
TABLES_DELETE_ON_DOWNLOAD_NOTE = ("zip をダウンロードすると、この取り込みのデータはサーバーから消えます"
                           "（もう一度ダウンロードすることはできません）。")
TABLES_DELETE_ON_DOWNLOAD_CONFIRM = ("ダウンロードすると、この取り込みのデータはサーバーから消えます。"
                              "もう一度ダウンロードすることはできません。")
ENCODING_CHOICES = [("utf-8-sig", "UTF-8（BOM付き）"), ("utf-8", "UTF-8"), ("cp932", "CP932（Shift_JIS）"),
                    ("euc_jp", "EUC-JP"), ("iso2022_jp", "ISO-2022-JP（JISコード）"),
                    ("shift_jis_2004", "Shift_JIS 2004"), ("utf-16", "UTF-16"),
                    ("utf-16-le", "UTF-16LE（BOMなし）"), ("utf-16-be", "UTF-16BE（BOMなし）")]
BUSY_MESSAGE = "処理中は変更できません。終わるか中止してから変更してください"
DELIMITER_CHOICES = [(",", "カンマ"), ("\t", "タブ"), (";", "セミコロン"), ("|", "縦棒")]


# ---- 共通 ---------------------------------------------------------------------------------

def _load_import(import_id: int) -> dict:
    """この画面（このブラウザ）の取り込みを返す。ほかの人の取り込みは「無い」として扱う（404）。

    社内LANで数人が同時に使うので、番号を打ち替えただけでほかの人の表を読めてはいけない。
    403 にすると「その番号の取り込みはある」ことが分かってしまうので、404 にそろえる。
    """
    imp = get_import(import_id)
    if imp is None or not owns(imp):
        abort(404)
    return imp


def _spec_for(imp: dict):
    return spec_for_import(imp)


def _is_excel(imp: dict) -> bool:
    return Path(imp["file_name"]).suffix.lower() in EXCEL_EXT


def _open(imp: dict):
    """控え付きの表ソース。元のファイルは控えにない範囲を読むときだけ開く。"""
    return import_source(imp)


def _sheet(imp: dict, source) -> str:
    if not _is_excel(imp):
        return source.file_name  # CSV はファイル名が1つのシート（全行を数えずに済ませる）
    src = imp.get("source") or {}
    names = [s.name for s in source.sheets()]
    if src.get("sheet") in names:
        return src["sheet"]
    visible = [s.name for s in source.sheets() if not s.hidden]
    if not (visible or names):
        raise UploadError("シートがないブックです。シートのあるブックを選んでください")
    return (visible or names)[0]


def _layout(imp: dict, source, sheet: str, header_rows=None, data_end=None, use_saved=True):
    src = imp.get("source") or {}
    if use_saved:
        header_rows = header_rows or [int(r) for r in src.get("header_rows") or [] if r] or None
        data_end = data_end or src.get("data_end_row") or None
    return source.layout(sheet, header_rows=header_rows or None, data_end=data_end or None)


def _after_read_panel(spec) -> str:
    """読み込みのあとに開く段。経過の記録の列があれば AI整形、なければ内容の確認。"""
    return "ai" if spec is not None and spec.log_stage is not None else "preview"


def _json_error(message: str, status: int = 400, **extra):
    return jsonify({"error": message, **extra}), status


def _tables_payload() -> dict:
    data = request.get_json(silent=True)
    return data if isinstance(data, dict) else {}


def _int(value, default=None):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


# 行番号の範囲「3-4」を広げる上限（見出し行の帯なので広くはならない）。static/app.js の parseRows と同じ規則
_MAX_ROW_SPAN = 10


def _row_no(value) -> int | None:
    """行番号の入力（全角数字も可）。数字だけのときだけ読む（static/app.js の parseEnd と同じ規則）。"""
    text = unicodedata.normalize("NFKC", str(value if value is not None else "")).strip()
    return int(text) if text.isascii() and text.isdigit() and int(text) > 0 else None


def _int_list(value) -> list[int]:
    """見出し行の入力（「1,2」「1、2」「1 2」「１，２」「1-2」）。static/app.js の parseRows と同じ規則で読む。"""
    if isinstance(value, (list, tuple)):
        items = [str(x) for x in value]
    else:
        items = re.split(r"[,、\s]+", unicodedata.normalize("NFKC", str(value or "")))
    out: set[int] = set()
    for item in items:
        m = re.fullmatch(r"(\d+)-(\d+)", item.strip())
        if m and item.isascii():
            a, b = int(m.group(1)), int(m.group(2))
            if 0 < a <= b and b - a < _MAX_ROW_SPAN:
                out.update(range(a, b + 1))
            continue
        n = _row_no(item)
        if n:
            out.add(n)
    return sorted(out)


def _ai_running(import_id: int) -> bool:
    job = latest_job("table_import", import_id, kind="ai_format")
    return bool(job and not job.get("finished"))


# AIの試し実行（同期の要求でジョブが無い）をしている取り込み → 実行中の数。実行中に取り込みを消すと、
# 終わった試し実行が AI の結果・応答を書き戻して残してしまう（design.md 3.3）
_TRIALS: Counter[int] = Counter()
_TRIALS_LOCK = threading.Lock()
TRIAL_BUSY_MESSAGE = "AIの試し実行中は削除・ダウンロード・保存できません。試し実行が終わってからもう一度押してください"


def _trial_running(import_id: int) -> bool:
    with _TRIALS_LOCK:
        return _TRIALS[import_id] > 0


def _processing(imp: dict) -> bool:
    """読み込み中・Markdown作成中・AI整形の実行中か。途中で設定やデータを変えると、古い設定の結果が残ったり、
    消したはずの AI の結果が書き戻されたりする（design.md 3.3）。"""
    return imp.get("status") in ("reading", "confirming") or _ai_running(imp["id"])


def _busy_error(imp: dict):
    """設定・データを変えられないとき（処理中・AIの試し実行中）の断り。変えてよければ None。

    試し実行を見ていなかったころは、断り書きで「保存できません」と言いながら保存が通り、
    行を入れ替えたあとに試し実行の結果が書き戻っていた（2026-09-26 の総ざらいで実測）。
    """
    if _processing(imp):
        return _json_error(BUSY_MESSAGE, 409)
    if _trial_running(imp["id"]):
        return _json_error(TRIAL_BUSY_MESSAGE, 409)
    return None


# ---- 1画面のやりとり（段の URL と、段の中身） ------------------------------------------------------

def _urls(import_id: int) -> dict:
    """画面（static/app.js）が使う URL。panel は末尾の NAME を段の名前に置き換えて使う。

    列の対応づけ・AI整形・プレビュー・ダウンロードの URL は、それぞれの段の HTML が
    data-* 属性や <a href> で持っているので、ここには入れない。
    """
    def u(endpoint: str, **kw) -> str:
        return url_for(endpoint, import_id=import_id, **kw)

    return {
        "panel": url_for("tables.panel", import_id=import_id, name="NAME"),
        "source": u("tables.save_source"), "detect": u("tables.layout_detect"), "layout": u("tables.save_layout"),
        "read": u("tables.reread"), "cancel": u("tables.cancel_job"), "delete": u("tables.delete_import"),
        "preview_start": u("tables.start_preview"), "confirm": u("tables.confirm"),
    }


def _job_info(job: dict | None, import_id: int | None = None) -> dict | None:
    if job is None:
        return None
    return {"id": job["id"], "kind": job.get("kind"), "status": job.get("status"),
            "finished": bool(job.get("finished")), "message": job.get("message") or "",
            "progress": job.get("progress") or {},
            "url": url_for("tables.api_job", job_id=job["id"]),
            "cancel_url": _cancel_url(job, import_id)}


def _cancel_url(job: dict, import_id: int | None) -> str | None:
    """そのジョブを止める口。止められないジョブでは None（画面は［中止］を出さない）。

    AI整形は別の口。Markdown の下書き（table_preview）は止める口が無いので、
    ボタンを出すと「中止できる処理がありません」と言うだけになる（2026-09-26 の総ざらいで実測）。
    """
    if not import_id:
        return None
    if job.get("kind") == "ai_format":
        return url_for("tables.ai_control", import_id=import_id, action="cancel")
    if job.get("kind") not in CANCELLABLE_JOBS:
        return None
    return url_for("tables.cancel_job", import_id=import_id)


def _running_job(imp: dict) -> dict | None:
    """読み込み中・確定処理中なら、その進み具合。終わっているのに状態が残っていれば直す。"""
    if imp.get("status") not in ("reading", "confirming"):
        return None
    job = get_job(imp["job_id"]) if imp.get("job_id") else None
    if job is None or job.get("finished"):
        fresh = get_import(imp["id"])
        if fresh and fresh.get("status") in ("reading", "confirming"):
            if fresh["status"] == "reading":
                stats = {**(fresh.get("stats") or {}), "error": "読み込みが途中で止まりました。もう一度読み込んでください"}
                update_import(imp["id"], status="failed", stats=stats)
            else:
                update_import(imp["id"], status="preview")
        return None
    return job


def _panel(html: str, **extra):
    return jsonify({"html": html, **extra})


def _locked(reason: str, **extra):
    """まだ使えない段（灰色で1行だけ理由を出す）。"""
    return jsonify({"html": "", "locked": reason, **extra})


@tables_bp.get("/")
@tables_bp.get("/new")
def new():
    """表の取り込みの画面（1枚）。段の中身はここでは出さず、ファイルを置いたあとに取りに来る。"""
    current_session_id()   # 画面を開いた時点で作業場所（クッキー）を決めておく（同時に置かれても取り違えない）
    return render_template("base.html", screen="tables")


@tables_bp.get("/imports/<int:import_id>/panel/<name>")
def panel(import_id: int, name: str):
    imp = _load_import(import_id)
    handler = _PANELS.get(name)
    if handler is None:
        abort(404)
    return handler(imp)


# ---- ファイルを置く ------------------------------------------------------------------------

@tables_bp.post("/upload", endpoint="upload")
def tables_upload():
    cfg = current_app.config
    try:
        stored = save_upload(request.files.get("file"), "tables", set(cfg["TABLE_ALLOWED_EXTENSIONS"]),
                             int(cfg["TABLE_MAX_UPLOAD_BYTES"]))
    except UploadError as exc:
        return _json_error(str(exc))
    path = upload_path(stored.stored_path)
    source_info: dict = {}
    try:
        if Path(stored.file_name).suffix.lower() in EXCEL_EXT:
            precheck_excel(path)
            source = open_source(path, stored.file_name, {"max_cells": cfg.get("EXCEL_MAX_CELLS", 500000)})
            sheets = [s for s in source.sheets() if not s.hidden] or source.sheets()
            if not sheets:
                raise UploadError("シートがないブックです。シートのあるブックを選んでください")
            # 開けても値を読めないブック（<v>NaN</v> など）は、取り込みを作る前にここで断る。
            # 先頭だけでなく全行を読む（下の方の値が読めないと、取り込みが読み込みの段から先へ進めない）
            for _row in source.rows(sheets[0].name):
                pass
            source_info = {"kind": "excel", "sheet": sheets[0].name}
        else:
            source = open_source(path, stored.file_name)
            sniff = source.sniff
            source_info = {"kind": "csv", "encoding": source.encoding, "delimiter": source.delimiter,
                           "errors": "strict"}
            if sniff is not None:
                source_info.update(bom=sniff.bom, preamble_rows=sniff.preamble_rows, sniff_warnings=sniff.warnings,
                                   confidence=sniff.confidence)
    except UploadError as exc:
        remove_upload(stored.stored_path)
        return _json_error(str(exc))
    except Exception:
        remove_upload(stored.stored_path)   # 思わぬエラーでもアップロードしたファイルを残さない（design.md 3.3）
        raise
    try:
        # DB への登録も同じ try に入れる。ここで落ちると（別の人の VACUUM 中の database is locked、
        # ディスク満杯など）、置いたファイル（最大200MB）だけが行の無い孤児として残り、
        # 再起動するまで消えなかった（2026-09-23 のレビュー）
        import_id = create_import(stored.file_name, stored.file_hash, stored.stored_path, source=source_info,
                                  session_id=current_session_id())
        if source_info.get("kind") == "excel":
            import_source(get_import(import_id), real=source).sheets()  # シート一覧を控えに入れる
    except Exception:
        remove_upload(stored.stored_path)
        raise
    return jsonify({"import_id": import_id, "file_name": stored.file_name, "urls": _urls(import_id)})


# ---- 画面を離れたので捨てる ------------------------------------------------------------------
# 「その場でダウンロードしない限り、その場ですぐ捨てる」（利用者の指示 2026-09-20）。
# 画面を閉じた・隠したときに static/app.js の ragDiscard がここへ「捨てて」と送ってくる。
# navigator.sendBeacon で届くので中身の型は text/plain（get_json(force=True) で読む）。
# 応答は読めず、やり直しもできないので、いつでも 204 を返す（もう無い番号・ほかの人の番号・
# 処理中のものは 片付け（purge_session など）の側で黙って外れる）。

@tables_bp.post("/discard", endpoint="discard")
def tables_discard():
    """この画面（このブラウザ）の、まだダウンロードしていない取り込みを捨てる。"""
    payload = request.get_json(force=True, silent=True)
    # sendBeacon は画面を閉じる途中で送るので、中身が辞書でないこともある（"x" や [1,2] など）
    ids = (payload.get("import_ids") or []) if isinstance(payload, dict) else []
    sid = current_session_id()
    try:
        if ids:
            discard_table_imports(ids, sid)
        else:
            purge_session(sid)
    except Exception as exc:   # 捨て損ねてもブラウザには伝えられない。時間切れの片付けに任せる
        current_app.logger.warning("取り込みの片付けに失敗しました: %s", exc.__class__.__name__)
    return "", 204


# ---- 読み取り方（文字コード・区切り・シート） ---------------------------------------------------------

def _source_note(imp: dict, src: dict) -> str:
    """段の見出しに出す1行（どう読むか）。"""
    if _is_excel(imp):
        return f"シート: {src.get('sheet') or ''}"
    delimiters = dict(DELIMITER_CHOICES)
    return f"{src.get('encoding') or ''}／{delimiters.get(src.get('delimiter'), src.get('delimiter') or '')}"


def _panel_source(imp: dict):
    import_id = imp["id"]
    src = imp.get("source") or {}
    error = None
    sheets, sheet, list_like = [], None, {}
    try:
        source_obj = _open(imp)
        sheets = source_obj.sheets() if _is_excel(imp) else []
        sheet = _sheet(imp, source_obj)
        if _is_excel(imp) and len([s for s in sheets if not s.hidden]) <= 8:
            # 「一覧表らしい」「クロス集計らしい（対応していません）」を出し分ける（表の範囲の判定と合わせる）
            list_like = {s.name: source_obj.list_kind(s.name) for s in sheets if not s.hidden}
    except UploadError as exc:
        error = str(exc)
    html = render_part("base.html", "part_source", imp=imp, src=src, error=error, sheets=sheets, sheet=sheet, list_like=list_like,
        encodings=ENCODING_CHOICES, delimiters=DELIMITER_CHOICES)
    note = _source_note(imp, {**src, "sheet": sheet or src.get("sheet")})
    return _panel(html, note=note, error=error, import_id=import_id)


@tables_bp.post("/imports/<int:import_id>/source")
def save_source(import_id: int):
    """読み取り方（シート・文字コード・区切り）。変えるたびに保存する。"""
    imp = _load_import(import_id)
    busy = _busy_error(imp)
    if busy is not None:
        return busy
    form = _tables_payload() or request.form
    src = dict(imp.get("source") or {})
    before = (src.get("sheet"), src.get("encoding"), src.get("delimiter"), src.get("errors"))
    if _is_excel(imp):
        if form.get("sheet"):
            src["sheet"] = form["sheet"]
    else:
        allowed = {"encoding": {v for v, _label in ENCODING_CHOICES},
                   "delimiter": {v for v, _label in DELIMITER_CHOICES}}
        for name in ("encoding", "delimiter"):
            value = form.get(name)
            if value and value not in allowed[name]:
                return _json_error("この画面にない文字コード・区切り文字は選べません")
            if value:
                src[name] = value
        src["errors"] = "replace" if form.get("replace_errors") in ("on", True, "true", "1") else "strict"
    columns: dict = {"source": src, "status": "uploaded"}   # 読み取り方を保存したら読み込みからやり直す
    if (src.get("sheet"), src.get("encoding"), src.get("delimiter"), src.get("errors")) != before:
        src.pop("header_rows", None)
        src.pop("data_end_row", None)
        # 置いたときの判定（文字コード・区切りの見立て）で出した注意書きと「前置き行」は、
        # 読み方を変えたら合わなくなる。捨てないと、文字コードを直したあとも
        # 「UTF-8 の行が混ざっています」が出続けて消せなかった（2026-09-23 のレビュー）
        # 捨てるのは置いたときの見立ての注意書きだけ。前置き行の数（preamble_rows）は
        # ③「表の範囲」が使う値なので残す（捨てると 0行になって実際と食い違う）
        for stale in ("sniff_warnings", "confidence"):
            src.pop(stale, None)
        # 見出しが変わるので、この取り込みの列の対応づけは捨てて決め直す
        columns.update(spec_json="", spec_hash="")
    update_import(import_id, **columns)
    return jsonify({"ok": True, "note": _source_note(imp, src), "next": "layout",
                    "reset": ["layout", "columns", "ai", "preview", "done"]})


# ---- 表の範囲（見出し行・データの終わり） ----------------------------------------------------------

def _grid(source, sheet: str, layout) -> tuple[list[dict], int]:
    classes = {rc.index: rc for rc in layout.row_classes}
    rows, width = [], 0
    for row in source.rows(sheet, 1, GRID_ROWS):
        cells = [(c.text or "")[:GRID_CELL_CHARS] for c in row.cells[:GRID_COLS]]
        while cells and not cells[-1]:
            cells.pop()
        width = max(width, len(cells))
        rc = classes.get(row.index)
        rows.append({"index": row.index, "cells": cells, "kind": rc.kind if rc else ("blank" if row.is_blank else "title"),
                     "reason": rc.reason if rc else "", "hidden": bool(row.hidden),
                     "strike": any(getattr(c, "strike", False) for c in row.cells if c.text)})
    return rows, max(width, 1)


def _layout_json(layout) -> dict:
    return {
        "sheet": layout.sheet, "table_kind": layout.table_kind,
        "table_kind_label": TABLE_KIND_LABELS.get(layout.table_kind, layout.table_kind),
        "header_rows": layout.header_rows, "data_start": layout.data_start, "data_end": layout.data_end,
        "headers": layout.headers, "warnings": layout.warnings, "errors": list(getattr(layout, "errors", [])),
        "counts": {ROW_KIND_LABELS.get(k, k): v for k, v in (layout.counts or {}).items()},
        "rows": {str(rc.index): {"kind": rc.kind, "reason": rc.reason} for rc in layout.row_classes},
    }


def _panel_layout(imp: dict):
    job = _running_job(imp)
    if job is not None and job.get("kind") == "table_read":
        return _panel("", job=_job_info(job, imp["id"]), reading=True)
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet)
        rows, width = _grid(source_obj, sheet, guess)
    except UploadError as exc:
        return _locked(str(exc))
    src = imp.get("source") or {}
    html = render_part("base.html", "part_layout", imp=imp, layout=guess, info=_layout_json(guess), rows=rows,
                           width=width, sheet=sheet, data_end_saved=src.get("data_end_row") or "")
    return _panel(html, note=f"{guess.data_start}〜{guess.data_end}行目" if guess.data_end >= guess.data_start else "")


@tables_bp.post("/imports/<int:import_id>/layout/detect")
def layout_detect(import_id: int):
    imp = _load_import(import_id)
    data = _tables_payload()
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet, header_rows=_int_list(data.get("header_rows")) or None,
                        data_end=_row_no(data.get("data_end")), use_saved=False)
    except UploadError as exc:
        return _json_error(str(exc))
    return jsonify(_layout_json(guess))


@tables_bp.post("/imports/<int:import_id>/layout")
def save_layout(import_id: int):
    """この範囲で読み込む。設定が無い・見出しが変わったときは列の対応づけへ、そうでなければ読み込みを始める。"""
    imp = _load_import(import_id)
    busy = _busy_error(imp)
    if busy is not None:
        return busy
    data = _tables_payload() or request.form
    header_rows = _int_list(data.get("header_rows"))
    data_end = _row_no(data.get("data_end_row"))
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet, header_rows=header_rows or None, data_end=data_end, use_saved=False)
    except UploadError as exc:
        return _json_error(str(exc))
    if guess.table_kind == "crosstab":
        return _json_error("月別集計のようなクロス集計の表には対応していません。1行＝1件の一覧表を選んでください")
    if guess.errors:
        return _json_error(guess.errors[0])
    if not guess.headers:
        return _json_error("見出し行とデータの範囲が見つかりません。見出し行の番号を指定してください")
    if guess.data_end < guess.data_start:
        # 見出しは見つかっている。直すのは見出し行ではなくデータの範囲なので、そう伝える
        last_header = guess.header_rows[-1] if guess.header_rows else 0
        if data_end:
            return _json_error(f"データの終わりに{data_end}行目を指定すると、データの行が1行もありません"
                               f"（見出しは{last_header}行目）。空欄にするか、見出し行より下の行を指定してください")
        return _json_error(f"見出し行（{last_header}行目）より下にデータの行が1行もありません。"
                           "見出し行の番号を確かめるか、データのある表を選んでください")
    src = dict(imp.get("source") or {})
    src.update(sheet=sheet, header_rows=guess.header_rows, header_row=None, data_end_row=data_end or None)
    update_import(import_id, source=src)
    # 取り込み設定は保存しないので、範囲を決めたら必ず「列の対応づけ」へ進む（利用者の指示 2026-09-20）
    return jsonify({"ok": True, "next": "columns", "reset": ["columns", "ai", "preview", "done"]})


# ---- 列の対応づけ ------------------------------------------------------------------------------
# 画面で決めるのは「使う・役割」だけ。キー・型・単位・md での扱い・空欄＝上と同じは、見出しと値から
# 候補づくり（tables.mapping.suggest_columns・tables.dictionary）が決める（利用者の指示 2026-09-20）。


def _screen_role(role: str) -> str:
    """候補の役割を画面の選択肢に寄せる（画面に出さない役割＝人名・数値・区分などは「その他」）。"""
    return role if role in SCREEN_ROLE_KEYS else "attribute"


def _suggest(imp: dict):
    """保存した範囲で見出しと先頭のデータを読み、列ごとの候補を作る。戻り値: (layout, suggestions)"""
    source_obj = _open(imp)
    sheet = _sheet(imp, source_obj)
    guess = _layout(imp, source_obj, sheet)
    samples = source_obj.sample_rows(sheet, guess, 200)
    return guess, suggest_columns(guess.headers, samples)


def _unused_note(sugg) -> str:
    """はじめから「使わない」にしてある列の、その理由（見出しの下に小さく出す。ふつうの列は空）。

    理由が読めないと、チェックの外れている列を入れ直してよいのか分からない。
    「知らせ」の列は 2026-09-21 にやめた（ほとんどの行で空で、表の上の行と同じことを書いていた）。
    型エラー・空欄の割合は表の上（_columns_todo）に出すので、ここには書かない。
    「値は「日付」らしい」も出さない（日付の役割は日付として読める列にしか出ないので、読んでも直せない）。
    """
    if sugg.md != "omit":
        return ""
    return ("空欄だけなので、はじめから使わない設定にしています" if sugg.omit_reason == "blank"
            else "記録に不要な管理用の列らしいので、はじめから使わない設定にしています")


def _column_row(sugg) -> dict:
    """画面の1行（使う・見出し・役割・値の例と、使わない設定にしている理由）。"""
    return {"index": sugg.index, "header": sugg.header, "use": sugg.md != "omit",
            "role": _screen_role(sugg.role), "type": sugg.type, "examples": list(sugg.examples or [])[:3],
            "note": _unused_note(sugg)}


DATE_TYPES = ("date", "datetime")


def _roles_for(row: dict) -> list[tuple[str, str]]:
    """その列で選べる役割。

    「日付」は値が日付として読める列にだけ出す（日付として読めない列を日付にすると保存が通らず、
    直しようのない行き止まりになる）。年月だけの列（2025-03・202503）は日付として読めるので出る
    ＝ その月のファイルに分かれる。
    """
    return [(role, label) for role, label in SCREEN_ROLES
            if role != "date" or row["type"] in DATE_TYPES or row["role"] == "date"]


def _has_date_column(pairs) -> bool:
    """日付として読める列があるか（1つも無ければ「日付の列を選べ」とは言わない）。"""
    return any(sugg.type in DATE_TYPES for _row, sugg in pairs)


# 「列の対応づけは決まっている」の決まり（利用者の問い 2026-09-20「列の対応付けを行う意味は？」）。
# この段で決められるのは「出す／出さない」と四つの役割だけで、ふつうの一覧表ならそのどちらも
# 候補づくり（tables.mapping）が見出しと値から決めている。決めることが無いのに22行の表を出すのは
# 意味がないので、次の条件をすべて満たすときは表をたたんで要約1行だけ出す。
#   1. 識別番号の列がちょうど1つで、見出しが辞書と完全一致している（matched_by == "dictionary"）
#   2. 日付の列がちょうど1つで、同じく完全一致している
#   3. 対象の列は0か1つ。1つなら完全一致している（0なら要約に「対象の列はありません」と書く）
#   4. 経過の記録の列は0か1つ。1つなら完全一致している（0なら要約にそう書く）
#   5. 出す列のどれにも、出すかどうかを決め直す理由が無い
#      ＝ 読めない値がある（type_error_rate > 0）／ほとんど空欄（blank_rate >= UNSURE_BLANK_RATE）
# 似た語で当たっただけ（matched_by == "similar"）や、値の並びから当てた（"none"）列が四つの役割に
# 付いていると 1〜4 で外れる。役割が本当に合っているかは人にしか決められないので、表を開く。
# 半分くらい空欄なのは決め直す理由にしない（出しても困らないため。ここに出さなければ画面のどこにも出ない）。
UNSURE_BLANK_RATE = 0.9

# 四つの役割（画面で決められるもの）と、要約・上の行に出す短い名前（プルダウンの但し書きまでは書かない）。
# 識別番号と日付は無いと決まらない
_DECIDED_ROLES = [("key", "識別番号", True), ("date", "日付", True),
                  ("entity", "対象", False), ("log", "経過の記録", False)]


def _role_columns(pairs, role: str) -> list[tuple[dict, object]]:
    return [pair for pair in pairs if pair[0]["use"] and pair[0]["role"] == role]


def _columns_todo(pairs) -> list[str]:
    """表を開いて決めてもらうことを並べる（空なら要約だけでよい）。pairs: [(画面の1行, 候補)]"""
    todo: list[str] = []
    for role, label, required in _DECIDED_ROLES:
        found = _role_columns(pairs, role)
        if len(found) > 1:
            todo.append(f"{label}の列が{len(found)}つあります。1つにしてください")
        elif not found:
            # 日付として読める列が1つも無い表では、選びようがないので求めない
            if required and (role != "date" or _has_date_column(pairs)):
                todo.append(f"{label}の列が決まっていません。1つ選んでください")
        elif found[0][1].matched_by != "dictionary":
            todo.append(f"「{found[0][0]['header']}」を{label}として読み取ります。これでよいか確かめてください")
    for row, sugg in pairs:
        if not row["use"]:
            continue
        if sugg.type_error_rate:
            todo.append(f"列「{row['header']}」に読み取れない値があります"
                        f"（{round(sugg.type_error_rate * 100, 1)}%）。出すかどうか決めてください")
        elif sugg.blank_rate >= UNSURE_BLANK_RATE:
            todo.append(f"列「{row['header']}」はほとんど空欄です"
                        f"（{int(round(sugg.blank_rate * 100))}%）。出すかどうか決めてください")
    return todo


def _columns_summary(pairs) -> str:
    """決まっているときに出す1行。見つけた役割と、出す列・出さない列の数を正直に書く。"""
    named, missing = [], []
    for role, label, _required in _DECIDED_ROLES:
        found = _role_columns(pairs, role)
        if found:
            named.append(f"{found[0][0]['header']}＝{label}")
        elif role == "date":
            missing.append("日付の列はありません。記録は「日付なし」の1ファイルにまとめます。")
        else:
            missing.append(f"{label}の列はありません。")
    used = sum(1 for row, _s in pairs if row["use"])
    left = len(pairs) - used
    return ("、".join(named) + "として読み取ります。" + "".join(missing)
            + f"{len(pairs)}列のうち{used}列を Markdown に出します"
            + (f"（残る{left}列は出しません）。" if left else "（出さない列はありません）。"))


def _default_table_name(imp: dict) -> str:
    """表の名前の初期値。ファイル名から作る（設定は残らないので、名前もこの取り込みだけのもの）。"""
    return Path(imp["file_name"]).stem


def _build_spec(payload: dict, guess, suggestions):
    """画面の選択（使う・役割）と列の候補から取り込み設定を作る。戻り値: (spec, errors)"""
    choices: dict[int, dict] = {}
    for row in payload.get("columns") or []:
        if isinstance(row, dict):
            choices[_int(row.get("index"), -1)] = row
    used: list[dict] = []
    date_errors: list[str] = []
    for sugg in suggestions:
        choice = choices.get(sugg.index)
        if not (bool(choice.get("use")) if choice else sugg.md != "omit"):
            continue
        role = str((choice or {}).get("role") or "")
        if role not in SCREEN_ROLE_KEYS:
            role = _screen_role(sugg.role)
        # 画面でそのままなら候補の役割（人名・数値・区分など）を活かす。選び直したときだけその役割にする
        role = sugg.role if role == _screen_role(sugg.role) else role
        if role == "date" and sugg.type not in DATE_TYPES:
            # 画面には出さない選び方だが、古い画面から送られたときに「型を日付にしてください」
            # （利用者には直せない指示）ではなく、何が起きているかを返す
            date_errors.append(f"列「{sugg.header}」の値は日付として読めないので、日付の列にはできません")
        type_ = "text" if role == "log" else sugg.type
        # 「使う」列は必ず md に出す（候補が「出さない」でも、チェックを入れたのだから出す）
        md = sugg.md if sugg.md != "omit" else ("body" if type_ == "text" else "attribute")
        used.append({"index": sugg.index, "header": sugg.header, "key": sugg.key, "display": sugg.display,
                     "type": type_, "role": role, "unit": sugg.unit, "md": md,
                     # 「対象」の空欄を上の値で埋めるのは値の形で決まる。画面で対象に選び直した列でも
                     # 同じにする（候補が対象だった列だけ埋めていた）
                     "fill_down_blank": bool(sugg.fill_down_blank
                                             or (role in ("entity", "entity_label")
                                                 and getattr(sugg, "looks_filled_down", False)))})
    name = str(payload.get("name") or "").strip()
    header_rows = list(getattr(guess, "header_rows", None) or [1])
    spec = spec_from_suggestions(name, {"table_kind": "list", "header_rows": header_rows}, used)
    group_errors: list[str] = []
    if payload.get("group_by") is not None:
        # ファイルの分け方。画面は列の位置（index）で送る。列のキーはここで決まるので、
        # 作った spec の列から引く（画面の段階ではキーがまだ無い列がある）
        order = [u["index"] for u in used]
        keys: list[str] = []
        for i in (_int(x, -1) for x in _as_list_payload(payload.get("group_by"))):
            if i in order and order.index(i) < len(spec.columns):
                keys.append(spec.columns[order.index(i)].key)
                continue
            # 「使う」を外した列を分け方に選んだまま保存すると、分け方が黙って無効になっていた
            head = next((s.header for s in suggestions if s.index == i), "")
            group_errors.append(f"ファイルの分け方に選んだ列「{head}」は「使う」にしてください"
                                if head else "ファイルの分け方に、この表に無い列が選ばれています")
        spec.markdown["group_by_columns"] = keys
    errors = group_errors + validate_spec(spec)
    if date_errors:
        # 「型を日付にしてください」は画面に直す場所が無いので、こちらの言い方に置き換える
        errors = date_errors + [e for e in errors if not e.startswith("日付の列「")]
    if sum(1 for u in used if u["role"] == "log") > 1:
        # 黙って最初の列だけを経過の記録にしない（2列目は日付ごとに分けられず1行につながれて出てしまう）
        errors.insert(0, "経過の記録の列は1つだけにしてください")
    return spec, errors


# 記録ファイルがこの数以上になったら、確認の画面で知らせる（止めはしない）
MANY_FILES_WARN = 200


def _as_list_payload(value) -> list:
    return list(value) if isinstance(value, (list, tuple)) else []


def _group_rows(rows: list[dict], suggestions, spec) -> list[dict]:
    """「ファイルの分け方」に出すチェックの一覧（出す列だけ。経過の記録は除く）。"""
    chosen = set((spec.markdown or {}).get("group_by_columns") or []) if spec is not None else set()
    fresh = spec is None
    out = []
    for row, sugg in zip(rows, suggestions):
        if not row["use"] or row["role"] == "log":
            continue
        is_date = sugg.type in DATE_TYPES or row["role"] == "date"
        out.append({"index": sugg.index, "label": sugg.header + ("（月ごと）" if is_date else ""),
                    "checked": (is_date if fresh else sugg.key in chosen)})
    return out


def _group_note(spec) -> str:
    """確認の画面に出す「いまの分け方」の一言。"""
    if spec is None:
        return ""
    cols = group_columns(spec)
    if not cols:
        return "分け方: 1つのファイルにまとめています"
    parts = [c.display + ("の月" if (c.type in DATE_TYPES or c.key == spec.date_key) else "") for c in cols]
    return "分け方: " + "・".join(parts) + "ごと"


def _panel_columns(imp: dict):
    import_id = imp["id"]
    job = _running_job(imp)
    if job is not None and job.get("kind") == "table_read":
        return _panel("", job=_job_info(job, import_id), reading=True)
    try:
        guess, suggestions = _suggest(imp)
    except UploadError as exc:
        return _locked(str(exc))
    if not guess.headers:
        return _locked("見出しが見つかりません。上の「表の範囲」で見出し行を指定してください")
    rows = [_column_row(s) for s in suggestions]
    spec = _spec_for(imp)
    if spec is not None:
        # この取り込みで一度保存していれば、そのときの「使う・役割」を残す（開き直しても選び直さずに済む）
        by_header = {}
        for col in spec.columns:
            for header in (col.headers or [col.display]):
                by_header[header] = col
        for row in rows:
            col = by_header.get(row["header"])
            row["use"] = col is not None
            if col is not None:
                row["role"] = _screen_role(col.role)
    for row in rows:
        row["roles"] = _roles_for(row)
    group_rows = _group_rows(rows, suggestions, spec)
    # 決めることが無ければ表をたたんで要約1行にする（表は隠すだけで残すので、保存で送る中身は同じ）
    pairs = list(zip(rows, suggestions))
    todo = _columns_todo(pairs)
    html = render_part("base.html", "part_columns", rows=rows, todo=todo, group_rows=group_rows,
        max_group=MAX_GROUP_COLUMNS, summary="" if todo else _columns_summary(pairs),
        name=spec.name if spec is not None else _default_table_name(imp),
        save_url=url_for("tables.save_columns", import_id=import_id))
    return _panel(html, note=f"{sum(1 for r in rows if r['use'])}／{len(rows)}列")


@tables_bp.post("/imports/<int:import_id>/columns")
def save_columns(import_id: int):
    imp = _load_import(import_id)
    busy = _busy_error(imp)
    if busy is not None:
        return busy
    try:
        guess, suggestions = _suggest(imp)
    except UploadError as exc:
        return _json_error(str(exc))
    spec, errors = _build_spec(_tables_payload(), guess, suggestions)
    if errors:
        return _json_error(errors[0], errors=errors)
    save_spec(import_id, spec)
    job_id = start_read_job(import_id)
    return jsonify({"ok": True, "next": _after_read_panel(spec), "reset": ["ai", "preview", "done"],
                    "job": _job_info(get_job(job_id), import_id), "reading": True})


@tables_bp.post("/imports/<int:import_id>/read")
def reread(import_id: int):
    """読み込み直し（失敗したとき・もう一度読むとき）。"""
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None:
        return _json_error("先に列の対応づけを保存してください")
    # AI整形の実行中・一時停止中に読み込み直すと、読み込みが AI整形の後ろで待ち続ける
    busy = _busy_error(imp)
    if busy is not None:
        return busy
    job_id = start_read_job(import_id)
    return jsonify({"ok": True, "next": _after_read_panel(spec), "reset": ["ai", "preview", "done"],
                    "job": _job_info(get_job(job_id), import_id), "reading": True})


# 中止できるジョブ（取り込みの job_id に入るもの）と、中止したときに戻す状態
CANCELLABLE_JOBS = {"table_read": "uploaded", "table_render": "preview"}


@tables_bp.post("/imports/<int:import_id>/cancel")
def cancel_job(import_id: int):
    """［中止］: 読み込み・Markdown作成のジョブを止める。"""
    imp = _load_import(import_id)
    job = get_job(imp["job_id"]) if imp.get("job_id") else None
    if job is None or job["kind"] not in CANCELLABLE_JOBS or job.get("finished"):
        return _json_error("中止できる処理がありません（すでに終わっている可能性があります）")
    request_cancel(job["id"])
    after = get_job(job["id"])
    # 取り込みの状態は読み直す（この間にジョブが成功して preview/confirmed を書いていたら、巻き戻してはいけない）
    fresh = get_import(import_id)
    if (after and after["status"] == "cancelled" and fresh is not None
            and fresh["status"] in ("reading", "confirming")):
        # 待機中のまま中止されたときは本体が動かないので、ここで取り込みの状態を戻す
        update_import(import_id, status=CANCELLABLE_JOBS[job["kind"]])
    # 止まるのは、動いている側が次に「中止の要求」を確かめたときなので、押した瞬間には止まっていない。
    # いつでも「中止しました」と答えていたころは、止まらなかったときも中止したと言っていた
    if after and after.get("finished"):
        return jsonify({"ok": True, "message": "処理を中止しました"})
    return jsonify({"ok": True, "message": "中止を要求しました（動いている処理が止まるまで少しかかります）"})


@tables_bp.post("/imports/<int:import_id>/delete")
def delete_import(import_id: int):
    """間違えて取り込んだファイルを消す（読み込んだ内容・作った Markdown も消える）。"""
    imp = _load_import(import_id)
    if imp["status"] in ("reading", "confirming"):
        return _json_error("処理中の取り込みは削除できません。終わるか中止してから削除してください")
    if _ai_running(import_id):
        return _json_error("AI整形の実行中は削除できません。AI整形を中止してから削除してください")
    if _trial_running(import_id):
        return _json_error(TRIAL_BUSY_MESSAGE)
    preview_job = latest_job("table_import", import_id, kind="table_preview")
    if preview_job is not None and not preview_job.get("finished"):
        # 下書きの作成中に消すと、開いているファイルが消え残ったり、消したあとに記録の md が書き戻されたりする
        return _json_error("Markdownの下書きを作っている間は削除できません。終わってから削除してください")
    purge_table_import(import_id)
    # ファイル名は出さない: 取引先名や「社外秘」を含みうる名前を画面に残さない（design.md 3.3）
    message = "取り込みを削除しました（元のファイル・読み込んだ内容・作った Markdown を消しました）"
    if purge_incomplete():
        # 使用中などで消し切れなかったファイルは中身を 0 バイトにしてある（次の起動時に片付く。design.md 3.3）
        message = "取り込みを削除しました（一部のファイルは使用中で消し切れず、中身を空にしました。次の起動時に片付きます）"
    return jsonify({"ok": True, "message": message, "reload": True})


def _not_ready_reason(imp: dict, spec) -> str:
    """読み込みが終わっていない・失敗したときに、段に出す1行。"""
    if spec is None:
        return "先に「列の対応づけ」を保存してください"
    if imp["status"] == "failed":
        return (imp.get("stats") or {}).get("error") or "表を読み込めませんでした"
    if imp["status"] == "uploaded":
        # ここに来るのは列の対応づけが保存済みのとき（spec is None は上で返している）。
        # 読み込みを中止したあとがこれなので、戻り口は④の［この対応づけで読み込む］
        return "表を読み込んでいません（中止したか、まだ読み込んでいません）。「列の対応づけ」の［この対応づけで読み込む］を押してください"
    return ""


# ---- AI整形（任意。経過の記録の列があるときだけ） --------------------------------------------------

def _ai_job(import_id: int) -> dict | None:
    return latest_job("table_import", import_id, kind="ai_format")


def _ai_connection_ctx() -> dict:
    """AI整形の段に要る AI接続の状態（設定そのものはヘッダーの「AI接続」。ai_header_ctx）。"""
    import ai

    ready = ai.is_configured()
    external, model = False, ""
    if ready:
        try:
            settings = ai.job_client_settings()
            external = not ai.is_local_endpoint(settings.get("chat_url") or "")
            model = settings.get("model") or ""
        except Exception:
            ready = False
    return {"ai_ready": ready, "external": external, "model": model, "ai_state": ai.connection_status()}


def _panel_ai(imp: dict):
    import_id = imp["id"]
    job = _running_job(imp)
    if job is not None:
        return _panel("", job=_job_info(job, import_id), reading=True)
    spec = _spec_for(imp)
    if spec is not None and spec.log_stage is None:
        return _locked("経過の記録の列がないので、この取り込みでは使いません")
    reason = _not_ready_reason(imp, spec)
    if reason:
        return _locked(reason)
    import ai

    log_key = spec.log_stage.column
    col = spec.column(log_key)
    row_choices = []
    for rec in load_rows(import_id):
        text = str((rec.get("values") or {}).get(log_key) or "").strip()
        if text:
            label = rec["key"] + "｜" + " ".join(text.split())[:40]
            row_choices.append((rec["key"], label))
            if len(row_choices) >= 200:
                break
    ai_job = _ai_job(import_id)
    html = render_part("base.html", "part_ai", imp=imp, log_display=col.display if col else log_key, row_choices=row_choices,
        trial_keys=[k for k, _ in row_choices[:10]], job=ai_job, job_active=bool(ai_job and not ai_job.get("finished")),
        job_url=url_for("tables.api_job", job_id=ai_job["id"]) if ai_job else None,
        counts=ai.counts(imp["template_id"], "log", import_id=import_id),
        scopes=SCOPE_LABELS, item_labels=ai.STATUS_LABELS, **_ai_connection_ctx())
    # AI整形のジョブは job を返さない。返すと画面側（loadPanel）が HTML を捨てて汎用の進捗箱に
    # 差し替えてしまい、この段にある一時停止・再開・内訳・［中止］が一度も出なかった
    # （そのうえ汎用の箱の［中止］は読み込み用の口なので AI整形は止まらなかった。2026-09-26 の実測）
    return _panel(html, note=f"経過の記録の列: {col.display if col else log_key}")


@tables_bp.post("/imports/<int:import_id>/ai/split-preview")
def ai_split_preview(import_id: int):
    import ai
    from ai import format_author, format_when, review_notes
    from ai import render_timeline
    import ai

    _load_import(import_id)
    row_key = str(_tables_payload().get("row_key") or "")
    try:
        data = ai.load_rows_for_ai(import_id)
        works = ai.prepare_works(data, ["log"], [row_key])
    except ai.AIJobError as exc:
        return _json_error(str(exc))
    if not works:
        return _json_error("行が見つかりません")
    w = works[0]
    parse = w.parse
    # id（s1, s2…）は AI との受け渡し用で画面には出さない。画面は no（順番1、順番2…）と、原文（text）の中の
    # 位置 start/end で「原文と並べて」見せる（利用者の指示 2026-09-21）
    segments = [{
        "id": s.id, "no": n, "body": s.body, "start": int(s.start), "end": int(s.end),
        "when": format_when(s.when), "when_estimated": bool(s.when and s.when.estimated),
        "author": format_author(s.author), "author_estimated": bool(s.author and s.author.estimated),
        "marks": list(s.marks or []), "identifiers": list(s.identifiers or []), "quantities": list(s.quantities or []),
        "plans": list(s.plans or []),
    } for n, s in enumerate(parse.segments, start=1)]
    return jsonify({
        "row_key": row_key, "kind": parse.kind, "order": parse.order, "text": parse.text, "segments": segments,
        "route": w.route, "reason": w.reason, "ai_ready": ai.is_configured(), "notes": review_notes(parse),
        "timeline": render_timeline(parse, w.entity_label, glossary=(w.stage.glossary or None) if w.stage else None),
        "sent_text": ai.messages_text(w.messages) if w.messages else "",
    })


@tables_bp.post("/imports/<int:import_id>/ai/trial")
def ai_trial(import_id: int):
    import ai

    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or spec.log_stage is None:
        return _json_error("経過の記録の列がありません")
    if imp["status"] == "confirmed":
        # 試し実行の結果は下書きに入るが、zip は確定したときの md を渡す。食い違わないよう断る
        return _json_error("確定後は試し実行できません（結果が確定した Markdown に入らないため）。"
                           "試すときは列の対応づけからもう一度読み込んでください")
    if _ai_running(import_id):
        # 全件実行と試し実行が同じ行に書き戻すので、別の設定の結果が混ざる（2026-09-23 のレビュー）
        return _json_error("AI整形が動いています（一時停止中を含む）。"
                           "終わるか中止してから試し実行してください")
    payload = _tables_payload()
    row_key = str(payload.get("row_key") or "")
    try:
        settings = ai.job_client_settings()
        if not ai.is_local_endpoint(settings.get("chat_url") or "") and not payload.get("confirm_external"):
            return _json_error("対応内容が外部のAIサービスに送信されます。確認のチェックを入れてください")
        with _TRIALS_LOCK:
            _TRIALS[import_id] += 1
        try:
            data = ai.load_rows_for_ai(import_id)
            result = ai.trial_row(import_id, row_key, stage_ids=["log"], data=data)
        finally:
            with _TRIALS_LOCK:
                _TRIALS[import_id] -= 1
                if _TRIALS[import_id] <= 0:
                    del _TRIALS[import_id]
    except ai.LLMNotConfigured:
        return _json_error("AI接続が設定されていません。画面右上の「AI接続」で設定してください")
    except ai.LLMCallError as exc:
        if exc.kind == "fatal":
            # 接続先に届かない・キーが拒否された。ヘッダーの表示もそれに合わせる（確認しに行く手間はかけない）
            ai.record_check(current_session_id(), False, [{"name": "試し実行", "ok": False, "detail": str(exc)}])
        return _json_error(str(exc))
    except ai.AIJobError as exc:
        return _json_error(str(exc))
    stage = next((s for s in result["stages"] if s.get("stage_id") == "log"), None) or {}
    accepted = stage.get("result") if stage.get("status") in ("ok", "flagged") else None
    record = next((r for r in data.rows if r["key"] == row_key), None)
    md = ""
    if record is not None:
        ai_results = {row_key: {"status": "ok", "result": accepted}} if accepted else {}
        md = "\n".join(record_block(record, spec, ai_results, people_index_for(spec, data.rows)))
    checks = stage.get("checks") or {}
    issues = [i.get("message", "") for i in (checks.get("issues") or []) if isinstance(i, dict)]
    return jsonify({
        "row_key": row_key, "status": stage.get("status"), "route": stage.get("route"), "reason": stage.get("reason"),
        "error": stage.get("error") or "", "points": ai_point_lines(accepted, spec.log_stage.glossary) if accepted else [],
        "timeline": stage.get("timeline") or [], "issues": issues, "markdown": md, "cached": stage.get("cached"),
        "stats": ai.trial_stats([result]),
    })


@tables_bp.post("/imports/<int:import_id>/ai/estimate")
def ai_estimate(import_id: int):
    import ai

    _load_import(import_id)
    data = _tables_payload()
    scope = data.get("scope") if data.get("scope") in SCOPE_LABELS else "pending"
    try:
        settings = ai.job_client_settings()
        result = ai.estimate(import_id, [s for s in data.get("trials") or [] if isinstance(s, dict)],
                                          scope=scope, concurrency=_int(data.get("concurrency")) or None,
                                          settings=settings, stage_ids=["log"])
    except ai.LLMNotConfigured:
        return _json_error("AI接続が設定されていません")
    except ai.AIJobError as exc:
        return _json_error(str(exc))
    return jsonify(result)


@tables_bp.post("/imports/<int:import_id>/ai/run")
def ai_run(import_id: int):
    import ai

    imp = _load_import(import_id)
    data = _tables_payload()
    if imp["status"] == "confirmed":
        # 「読み込みが終わってから」では理由が事実と違う（読み込みは終わっている）
        return _json_error("確定後はAI整形を実行できません（結果が確定した Markdown に入らないため）。"
                           "実行するときは「列の対応づけ」からもう一度読み込んでください")
    if imp["status"] != "preview":
        return _json_error("表の読み込みが終わってから実行してください")
    job = _ai_job(import_id)
    if job and not job.get("finished"):
        return _json_error("AI整形はすでに実行中です")
    scope = data.get("scope") if data.get("scope") in SCOPE_LABELS else "pending"
    concurrency = min(16, max(1, _int(data.get("concurrency"), 1) or 1))
    try:
        settings = ai.job_client_settings()
        if not ai.is_local_endpoint(settings.get("chat_url") or "") and not data.get("confirm_external"):
            return _json_error("対応内容が外部のAIサービスに送信されます。確認のチェックを入れてください")
        # 始める前に接続を確かめる（利用者の指示: AI整形を始めるときに確認）。ヘッダーの表示もここで更新される。
        # つながらなければジョブを作らずに理由を返す（何百行も失敗させてから気づかせない）
        ok, steps = ai.check_connection()
        ai.record_check(current_session_id(), ok, steps)
        if not ok:
            failed = next((s for s in steps if not s.get("ok")), steps[-1])
            return _json_error(f"AIにつながりません（{failed['name']}: {failed['detail']}）。"
                               "画面右上の「AI接続」を確認してください", ai_status=ai.connection_status())
        job_id = ai.start_ai_job(import_id, scope=scope, concurrency=concurrency, stage_ids=["log"])
    except ai.LLMNotConfigured:
        return _json_error("AI接続が設定されていません。画面右上の「AI接続」で設定してください")
    return jsonify({"ok": True, "job_id": job_id, "job_url": url_for("tables.api_job", job_id=job_id),
                    "ai_status": ai.connection_status()})


@tables_bp.post("/imports/<int:import_id>/ai/<action>")
def ai_control(import_id: int, action: str):
    _load_import(import_id)
    handlers = {"pause": request_pause, "resume": request_resume, "cancel": request_cancel}
    if action not in handlers:
        abort(404)
    job = _ai_job(import_id)
    ok = bool(job) and handlers[action](job["id"])
    if not ok:
        return _json_error("AI整形のジョブを操作できませんでした（すでに終わっている可能性があります）")
    return jsonify({"ok": True})


# ---- AI接続（ヘッダーの右上。どの画面でも同じパネルが開く。設定画面は無い） -----------------------------
# 利用者の指示（2026-09-21）:「AI接続の設定は、もともとの位置ヘッダーの画面右上『AI接続』に移動させる。
# 全部空欄にしておいてcookieでユーザー毎に登録内容をずっと保持させるようにしてほしい」
# 「AI接続はヘッダー上で、接続中 か 未接続 一目で分かるように」
# 設定はブラウザごと（current_session_id ごと。ai.py・core.ai_connections）。URL は昔のまま /tables/ の下に
# 置いてある（画面には出ない。base.html の data-* 属性で渡す）。

@tables_bp.app_context_processor
def ai_header_ctx() -> dict:
    """どの画面のヘッダーにも「AI接続」の状態を出す。確認しには行かず、覚えている結果を出すだけ。"""
    import ai

    try:
        status = ai.connection_status()
    except Exception as exc:   # DB が読めないときでも画面は出す
        status = {"state": "off", "state_label": "未接続", "ready": False, "checked_text": "", "check_detail": str(exc),
                  "sub_text": "クリックして設定", "panel_sub_text": "",
                  "chat_url": "", "models_url": "", "model": "", "models": [], "api_key_saved": False,
                  "api_key_source": "", "effective": {}, "fallback": {},
                  "placeholders": {"chat_url": "https://api.openai.com/v1/chat/completions",
                                   "models_url": "https://api.openai.com/v1/models", "model": ""}}
    return {"ai_header": status,
            "ai_urls": {"save": url_for("tables.save_ai_connection"), "test": url_for("tables.test_ai_connection"),
                        "models": url_for("tables.refresh_ai_models"), "clear": url_for("tables.clear_ai_key")}}


@tables_bp.post("/ai-connection")
def save_ai_connection():
    """ヘッダーの「AI接続」の［保存する］。保存したあと、その設定でつながるかを確かめる（結果はヘッダーに出る）。"""
    import ai

    data = _tables_payload()
    sid = current_session_id()
    try:
        ai.save_browser(sid, data)
    except ValueError as exc:
        return _json_error(str(exc))
    except sqlite3.Error as exc:
        return _json_error(f"保存できませんでした（{exc.__class__.__name__}）。少し待ってから、もう一度保存してください")
    ok, steps = ai.check_connection()
    ai.record_check(sid, ok, steps)
    status = ai.connection_status()
    if not status["ready"]:
        message = "保存しました（APIキーと接続先がそろうと接続を確かめます）"
    else:
        message = "保存しました。" + ("AIにつながりました" if ok else "AIにつながりません（下の結果を確認してください）")
    return jsonify({"ok": True, "message": message, "connected": ok, "steps": steps, "status": status})


@tables_bp.post("/ai-connection/clear")
def clear_ai_key():
    """［キーを消す］（共有PCで使い終わったとき）。このブラウザの APIキーだけ消す。接続先・モデルは残す。"""
    import ai

    status = ai.clear_browser_key(current_session_id())
    return jsonify({"ok": True, "message": "このブラウザのAPIキーを消しました", "status": status})


@tables_bp.post("/ai-connection/models")
def refresh_ai_models():
    """APIからモデル一覧を取得し、このブラウザの候補（入力欄の候補）として覚える。"""
    import ai

    try:
        catalog = ai.model_catalog(refresh=True)
    except Exception as exc:
        return _json_error(f"モデル一覧を取得できませんでした: {ai.friendly_error(exc)}")
    if catalog:
        save_ai_connection_row(current_session_id(), models=catalog)   # 入力欄の候補（datalist）として覚える
        ai.forget_browser()
    return jsonify({"ok": True, "models": catalog, "message": f"APIからモデル一覧を取得しました（{len(catalog)}件）",
                    "status": ai.connection_status()})


@tables_bp.post("/ai-connection/test")
def test_ai_connection():
    """接続の確認: モデル一覧の取得と、1回の短いチャット。パネルを開いたとき・［接続を確かめる］で呼ぶ。"""
    import ai

    ok, steps = ai.check_connection()
    ai.record_check(current_session_id(), ok, steps)
    return jsonify({"ok": ok, "steps": steps, "status": ai.connection_status()})


# ---- 内容の確認 ---------------------------------------------------------------------------

def _preview_job(import_id: int, imp: dict, spec, retry: bool = False) -> dict:
    """確認の段の md を作るジョブ。同じ入力のジョブがあればそれを見せ、なければ始める。

    retry: 同じ入力で失敗したジョブを作り直す（［もう一度作る］。ファイルの一時的なロックなどで失敗したとき）。
    """
    signature = preview_signature(import_id, imp, spec)
    job = latest_job("table_import", import_id, kind="table_preview")
    if job is not None and (job.get("params") or {}).get("signature") == signature:
        if not job.get("finished") or (job["status"] == "failed" and not retry):
            return job  # 動いている途中、または同じ入力で失敗したまま（開き直すたびに作り直さない）
    return get_job(start_preview_job(import_id, signature))


def _display_with_unit(col) -> str:
    """データ表の見出し。Markdown は「停止時間: 5分」と単位を付けて書くので、見出しにも単位を添える。"""
    unit = (col.unit or "").strip()
    if unit and unit not in col.display:
        return f"{col.display}（{unit}）"
    return col.display


def _panel_preview(imp: dict, page: int = 1):
    import_id = imp["id"]
    job = _running_job(imp)
    if job is not None:
        return _panel("", job=_job_info(job, import_id), reading=True)
    spec = _spec_for(imp)
    reason = _not_ready_reason(imp, spec)
    if reason:
        return _locked(reason, failed=imp["status"] == "failed",
                       retry_url=url_for("tables.reread", import_id=import_id) if spec is not None else None)
    if _ai_running(import_id):
        return _locked("AI整形が動いています（一時停止中を含む）。再開して終わらせるか、中止してから確認してください")
    files = ready_preview_files(import_id, imp, spec)
    if files is None:
        # 件数分の md を作るのに時間がかかるので、ジョブにして同じ画面に進み具合を出す
        draft = _preview_job(import_id, imp, spec)
        if draft.get("finished") and draft["status"] != "done":
            return _panel(render_part("base.html", "part_preview_failed", error=draft.get("message") or "Markdownの下書きを作れませんでした"),
                          failed=True)
        return _panel("", job=_job_info(draft, import_id), building=True)
    stats = imp.get("stats") or {}
    issues = load_issues(import_id)
    page = max(1, page)
    data_rows, total = load_rows_page(import_id, (page - 1) * DATA_PAGE, DATA_PAGE)
    total_pages = max(1, (total + DATA_PAGE - 1) // DATA_PAGE)
    if page > total_pages:
        page = total_pages
        data_rows, _total = load_rows_page(import_id, (page - 1) * DATA_PAGE, DATA_PAGE)
    blocking = has_blocking(issues)
    # 確定を止めているエラーを先に出す。行番号順のまま先頭200件を切っていたころは、
    # エラーが201件目以降だと「エラーが残っているため確定できません」と出るのに
    # 問題一覧にはエラーが1件も無く、何を直せばよいか分からなかった（2026-09-23 のレビューで実測）
    shown = sorted(issues, key=lambda i: 0 if i.get("level") == "error" else 1)
    html = render_part("base.html", "part_preview", imp=imp, spec=spec, stats=stats, issues=shown[:ISSUES_SHOWN],
        issue_total=len(issues), counts=count_levels(issues), blocking=blocking, files=files, data_rows=data_rows,
        page=page, total_pages=total_pages, columns=[(c.key, _display_with_unit(c)) for c in spec.columns],
        confirmed=imp["status"] == "confirmed", delete_note=TABLES_DELETE_ON_DOWNLOAD_NOTE,
        many_files=len(files) >= MANY_FILES_WARN, group_note=_group_note(spec))
    return _panel(html, note=f"{stats.get('records') or 0}件・{len(files)}ファイル", blocking=blocking,
                  confirmed=imp["status"] == "confirmed")


@tables_bp.post("/imports/<int:import_id>/preview")
def start_preview(import_id: int):
    """確認の段の Markdown の下書きを作り直す（失敗したときの［もう一度作る］）。"""
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        return _json_error("表の読み込みが終わってから確認してください")
    job = _preview_job(import_id, imp, spec, retry=True)
    return jsonify({"ok": True, "job": _job_info(job, import_id), "building": True})


@tables_bp.get("/imports/<int:import_id>/preview/file")
def preview_file(import_id: int):
    _load_import(import_id)
    text = md_text(import_files(import_id)["preview"], request.args.get("name", ""))
    if text is None:
        return _json_error("ファイルが見つかりません", 404)
    return jsonify({"name": request.args.get("name"), "text": text})


@tables_bp.get("/imports/<int:import_id>/issues.csv")
def issues_csv(import_id: int):
    imp = _load_import(import_id)
    data = build_issues_csv(load_issues(import_id))
    name = f"{Path(imp['file_name']).stem}_問題一覧.csv"
    return set_download_name(
        send_file(io.BytesIO(data), mimetype="text/csv", as_attachment=True, download_name=name), name, "issues")


@tables_bp.post("/imports/<int:import_id>/confirm", endpoint="confirm")
def tables_confirm(import_id: int):
    """確定して Markdown を作る（できたら同じ画面にダウンロードのボタンを出す）。"""
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        return _json_error("この取り込みはまだ確定できる状態ではありません")
    if _ai_running(import_id):
        return _json_error("AI整形の実行中は確定できません。終わるか中止してから確定してください")
    if has_blocking(load_issues(import_id)):
        return _json_error("エラーが残っているため確定できません。問題一覧を確認して、範囲や列の対応づけを直してください")
    job_id = start_render_job(import_id)
    return jsonify({"ok": True, "next": "done", "job": _job_info(get_job(job_id), import_id), "building": True})


# ---- 確定してダウンロード（zip） ------------------------------------------------------------------

def _download_base(imp: dict, spec) -> str:
    stem = safe_filename_part(spec.file_prefix if spec is not None else Path(imp["file_name"]).stem)
    return f"{stem}_{datetime.now().strftime('%Y-%m-%d')}"


def _panel_done(imp: dict):
    import_id = imp["id"]
    job = _running_job(imp)
    if job is not None:
        return _panel("", job=_job_info(job, import_id), building=True)
    if imp["status"] != "confirmed":
        return _locked("上の「内容の確認」で［確定してMarkdownを作成］を押してください")
    spec = _spec_for(imp)
    files = [{"name": p.name, "size": p.stat().st_size} for p in md_paths(import_id)]
    if not files:
        # 記録0件のまま確定された取り込み（渡せるものが無いので、ダウンロードのボタンは出さない）
        return _locked("この取り込みから作られた Markdown はありません（取り込める行がありませんでした）。"
                       "表の範囲か元のファイルを見直してください")
    html = render_part("base.html", "part_done", imp=imp, files=files, stats=imp.get("stats") or {},
                           delete_note=TABLES_DELETE_ON_DOWNLOAD_NOTE, delete_confirm=TABLES_DELETE_ON_DOWNLOAD_CONFIRM)
    return _panel(html, note=f"{len(files)}ファイル")


@tables_bp.get("/imports/<int:import_id>/download.zip")
def download_zip(import_id: int):
    """zip を渡し、渡し終えた取り込みのデータを消す（design.md 3.3）。

    zip は全体をメモリに作ってから消すので、消す処理で中身が欠けることはない。作れなかったときは何も消さない。
    """
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    # 断る理由はこの画面に直接書く。以前は知らせ（flash）を積んで転送していたが、別の段の描画が
    # 知らせを先に食べることがあり、理由の出ない空の画面に飛ばされていた（2026-09-23 のレビュー）
    reason = None
    if imp["status"] != "confirmed" or spec is None:
        reason = "先に [確定してMarkdownを作成] を押してください"
    elif not md_paths(import_id):
        # 確定はしたが記録が0件だった（押したばかりのボタンをもう一度押せ、とは言わない）
        reason = ("作成された Markdown がありません（取り込める行がありませんでした）。"
                  "表の範囲か元のファイルを見直してください")
    elif _ai_running(import_id):
        reason = "AI整形の実行中はダウンロードできません。終わるか中止してからダウンロードしてください"
    elif _trial_running(import_id):
        reason = TRIAL_BUSY_MESSAGE
    if reason:
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error=reason), 409
        return render_template("base.html", code=409, reason=reason), 409
    try:
        data = build_download(import_id)
    except FileNotFoundError:
        # 同時に押した別のダウンロードが渡し終えてデータを消した（ダブルクリックなど）。500 ではなく「ダウンロード済み」。
        return _handout_lost(import_id, "ダウンロード")
    name = f"{_download_base(imp, spec)}.zip"
    response = send_file(io.BytesIO(data), mimetype="application/zip", as_attachment=True, download_name=name,
                         conditional=False)   # Range でも全体を返す（一部だけ渡して消すことが無いように）
    set_download_name(response, name, "records")
    return purge_after_send(response, purge_table_import, import_id)


def _handout_lost(import_id: int, what: str):
    """渡す途中で md が消えたとき。取り込みが残っていて確定済みでなければ、別のタブで読み込み直しが始まった
    （「ダウンロード済み」の 404 は事実と違うので、画面に戻して知らせる）。消えていれば渡し終えた・削除した。"""
    now = get_import(import_id)
    if now is None or now["status"] == "confirmed":
        abort(404)
    flash(f"読み込み直しが始まったため、{what}を取りやめました。確定し直してからもう一度押してください", "error")
    return redirect(url_for("tables.new"))


# ---- 段の割り当て ---------------------------------------------------------------------------

def _panel_rows(imp: dict):
    """データ（100行ずつ）のページ送りだけ作り直す。"""
    return _panel_preview(imp, page=max(1, _int(request.args.get("page"), 1) or 1))


_PANELS = {"source": _panel_source, "layout": _panel_layout, "columns": _panel_columns, "ai": _panel_ai,
           "preview": _panel_rows, "done": _panel_done}


# ---- ジョブの進捗（JSON） ------------------------------------------------------------------------------

def _job_visible(job: dict) -> bool:
    """このブラウザの取り込みのジョブか（ほかの人のジョブは進み具合も見せない）。

    取り込みが消えている（ダウンロード済み・削除済み）ジョブは、持ち主が分からないので見せない。
    """
    if (job.get("ref_type") or "") != "table_import" or job.get("ref_id") is None:
        return False
    imp = get_import(job["ref_id"])
    return imp is not None and owns(imp)


def api_job(job_id: int):
    job = get_job(job_id)
    if job is None or not _job_visible(job):
        return _json_error("ジョブが見つかりません", 404)
    return jsonify({k: job.get(k) for k in ("id", "kind", "status", "status_label", "progress", "message", "result",
                                            "finished")})


@tables_bp.record_once
def _register_api(setup_state) -> None:
    # /api/jobs は /tables の外に置く（design.md 2.4）
    setup_state.app.add_url_rule("/api/jobs/<int:job_id>", "tables.api_job", api_job)


####################################################################################################
# create_app・エラー画面・起動コマンド
####################################################################################################

def create_app(overrides: dict | None = None) -> Flask:
    _refuse_debugger()
    app = Flask(__name__)
    app.config.from_object(Config)
    # クッキーは約1年もたせる（SESSION_LIFETIME。ヘッダーの「AI接続」の設定をブラウザごとに
    # 「ずっと保持」するため。利用者の指示 2026-09-21）。current_session_id が session.permanent を立てる
    app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                      PERMANENT_SESSION_LIFETIME=SESSION_LIFETIME)
    if overrides:
        app.config.update(overrides)
    if not app.config.get("SECRET_KEY"):
        app.config["SECRET_KEY"] = _secret_key()

    # DATA_DIR だけ差し替えたとき（テスト等）は一覧表の置き場所もそれに合わせる
    if not (overrides and "TABLES_DIR" in overrides) and "TABLES_DIR" not in os.environ:
        app.config["TABLES_DIR"] = Path(app.config["DATA_DIR"]) / "tables"
    for key in ("UPLOAD_DIR", "DATA_DIR", "TABLES_DIR"):
        app.config[key] = Path(app.config[key])
        app.config[key].mkdir(parents=True, exist_ok=True)
    init_app(app)
    if not app.config.get("ALLOWED_HOSTS"):
        app.config["ALLOWED_HOSTS"] = allowed_hosts()
    else:   # 設定から渡されたものも、Host と同じ形（ポート無し・小文字）に揃え、ループバックは必ず足す
        given = {h if h == "*" else (_hostname_only(h) or str(h).lower()) for h in app.config["ALLOWED_HOSTS"]}
        app.config["ALLOWED_HOSTS"] = {"*"} if "*" in given else given | {"127.0.0.1", "localhost", "::1"}

    app.register_blueprint(home_bp)
    app.register_blueprint(tables_bp)

    @app.before_request
    def _refuse_other_host():
        # 宛先の名前（Host）が、運用者が意図した名前かを確かめる（allowed_hosts のとおり）。
        # 待ち受けているアドレスを攻撃者のドメインが指していると、そのページと同一オリジンになり、
        # 画面の中身を読み取られてしまう（DNSリバインディング）。
        allowed = app.config["ALLOWED_HOSTS"]
        try:
            parts = urlsplit(f"//{request.host}")
            host, port = parts.hostname, parts.port   # ポート番号・[] を外したホスト名
        except ValueError:
            host, port = None, None
        host = (host or "").rstrip(".").lower()
        if "*" not in allowed and host not in allowed:
            # 汎用の 400 画面（ホームへ = 同じアドレスの / ）では、また断られるだけで開き方が分からない。
            # 開けるアドレスを示す（ポートは送られてきた Host のもの。読めなければ起動時の PORT）
            home_url = f"http://{_home_host()}:{port or PORT}/"
            text = f"このアドレスでは開けません。このアプリは {home_url} で開いてください"
            if request.accept_mimetypes.best == "application/json" or request.is_json:
                return jsonify(error=text), 400
            return render_template("base.html", code=400, host_refused=True, home_url=home_url), 400

    @app.before_request
    def _refuse_cross_site_write():
        # 他のサイトのページから、利用者のブラウザ経由で書き込ませない（is_cross_site_write のとおり）。
        if is_cross_site_write(request):
            abort(403)

    @app.after_request
    def _no_store(response):
        # 画面・JSON には取り込んだ値や Markdown が載る。ダウンロードでサーバー側から消しても（design.md 3.3）、
        # ブラウザのキャッシュ（戻る・再表示）に残らないよう、保存させない。静的ファイル（CSS・JS）は除く
        if request.endpoint != "static":
            response.headers["Cache-Control"] = "no-store"
        # ほかのサイトのページの中（iframe）に表示させない。枠の中でのクリックは同じサイトからの操作になり、
        # is_cross_site_write では防げない（削除ボタンを押させる、など）。静的ファイルも含めてすべての応答に付ける
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Content-Security-Policy"] = "frame-ancestors 'none'"
        return response

    @app.errorhandler(403)
    def _forbidden(_exc):
        # 画面からの JSON 送信（static/app.js の postJson）には JSON で返す。トーストに理由が出る。
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="ほかのサイトのページから送られてきた操作に見えたため受け付けませんでした。"
                                 "取り込みの画面からやり直してください"), 403
        return render_template("base.html", code=403), 403

    @app.errorhandler(404)
    def _not_found(_exc):
        # 403 と同じく、画面からの JSON 送信（static/app.js の postJson）には JSON で返す。
        # ダウンロード済みでデータが消えたあとに、開いたままの画面が途中保存・プレビューを送ると
        # ここに来る。HTML を返すとトーストに「通信に失敗しました」としか出ず、理由が伝わらない。
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="このデータはサーバーに残っていません（ダウンロード済みか、削除されています）。"
                                 "取り込みの画面からやり直してください"), 404
        return render_template("base.html", code=404), 404

    @app.errorhandler(413)
    def _too_large(_exc):
        # MAX_CONTENT_LENGTH を超えた送信。Werkzeug の英語の画面を出さない（design.md 2: エラー画面は日本語）
        limit = app.config.get("MAX_CONTENT_LENGTH") or 0
        text = f"{limit / 1024 / 1024:.0f}MB" if limit >= 1024 * 1024 else f"{limit / 1024:.0f}KB"
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error=f"送った内容が大きすぎます（1回に合計 {text} まで）。分けて送ってください"), 413
        return render_template("base.html", code=413, limit=text), 413

    @app.errorhandler(400)
    @app.errorhandler(405)
    def _bad_request(exc):
        # 送信専用の URL をアドレス欄から開いた（405）など。Werkzeug の英語の画面を出さない（design.md 2）
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="この操作は受け付けられませんでした。取り込みの画面から開き直してください"), exc.code
        return render_template("base.html", code=400), exc.code

    @app.errorhandler(500)
    def _server_error(_exc):
        # 画面の中からの送信に HTML を返すと、トーストに「通信に失敗しました」としか出ず
        # 理由が伝わらない（ほかのエラーと同じ扱いにする。2026-09-23 のレビュー）
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="サーバー側でエラーが発生しました。画面を開き直してもう一度お試しください"
                                 "（詳しい内容はアプリのログに記録しました）"), 500
        return render_template("base.html", code=500), 500

    @app.cli.command("serve")
    def _serve() -> None:
        """waitress で待ち受ける（flask run の代わり）。待ち受け先は環境変数 HOST・PORT。

        通常は `flask --app app run --host=0.0.0.0 --port=5000` で起動する。
        こちらは waitress を使いたいときだけ。どちらでも「途中で切れたダウンロードは消さない」は保てる
        （気づけない大きさは waitress で約48KB、開発サーバーで約128KB。2026-09-24 に実測）。
        """
        if DEBUG and not _is_loopback(HOST):
            # DEBUG=True の Flask はデバッガを開く。LAN に出す起動では絶対に開かせない（_NO_DEBUG_MSG と同じ理由）
            raise SystemExit("\n[app] DEBUG = True のまま LAN のアドレス（HOST=%s）では起動しません。\n"
                             "  理由: デバッガが開くと、例外が出たときにブラウザからこのサーバの Python を実行できます。\n"
                             "  詳しいエラーを見たいときは HOST を外して（127.0.0.1 で）起動してください。\n" % HOST)
        print(startup_notice())
        if DEBUG:
            app.run(host=HOST, port=PORT, debug=True, use_reloader=False)
            return
        try:
            from waitress import serve
        except ImportError:
            print("[app] waitress が無いため Flask の開発サーバで起動します（pip install waitress を推奨）")
            app.run(host=HOST, port=PORT, debug=False)
        else:
            serve(app, host=HOST, port=PORT, **WAITRESS_OPTIONS)

    _recover_jobs(app)
    _purge_pending(app)
    _cleanup_leftovers(app)
    _start_sweeper(app)
    if _start_command() == "run":
        # serve は自分で出すので run のときだけ。ログインが無いことの注意を必ず見せる
        print(startup_notice(), flush=True)
    # ヘッダーの行き先は「表の取り込み」と「解説」の2つだけ。使うAIモデルの選択は
    # ヘッダー右上の「AI接続」のパネルに移したので、共通の値は渡さない。
    return app


# 途中で放り出されたものを捨てる間隔（design.md 3.3）。起動時に全部捨て、動いている間は
# STALE_HOURS より古いものをこの間隔で捨てる（数人で使うサーバーに置きっぱなしになるため）。
# 捨てるまでの時間（STALE_HOURS）を短くしたときは、見回りもそれに合わせて短くする。
# 見回りの間隔だけ延びると「2時間で捨てます」と言いながら3時間残ることになるので、上限は10分にする。
SWEEP_INTERVAL_SECONDS = 10 * 60


def _sweep_interval(stale_hours: float) -> int:
    return max(60, min(SWEEP_INTERVAL_SECONDS, int(stale_hours * 3600 / 2) or SWEEP_INTERVAL_SECONDS))


def _purge_pending(app: Flask) -> None:
    """起動時：ダウンロードしていない取り込みをすべて捨てる（作業中の一覧を持たないため）。"""
    if app.config.get("TESTING"):
        return

    try:
        with app.app_context():
            tables_removed = purge_all_pending()
        if tables_removed:
            print(f"[app] 途中だった取り込みを捨てました（{tables_removed} 件）")
    except Exception as exc:  # 起動は止めない
        print(f"[app] 途中だった取り込みの片付けに失敗しました（{exc.__class__.__name__}: {exc}）")


def _start_sweeper(app: Flask) -> None:
    """動いている間：しばらくさわられていない取り込みを捨て続ける（daemon スレッド）。"""
    if app.config.get("TESTING"):
        return

    interval = _sweep_interval(STALE_HOURS)

    def loop() -> None:
        while True:
            _stop.wait(interval)
            try:
                with app.app_context():
                    removed = sweep_stale(STALE_HOURS)
                if removed:
                    print(f"[app] {STALE_HOURS}時間さわられていない取り込みを捨てました（{removed} 件）")
            except Exception as exc:   # 次の回でやり直す
                print(f"[app] 古い取り込みの片付けに失敗しました（{exc.__class__.__name__}: {exc}）")

    _stop = threading.Event()
    threading.Thread(target=loop, name="purge-sweeper", daemon=True).start()


def _cleanup_leftovers(app: Flask) -> None:
    """起動時：DB から参照されていない残骸（アップロードファイル・取り込みのフォルダ）を消す。

    データを残さない方針（design.md 3.3）でも、削除の途中で落ちたときやファイルを掴まれていたときに
    残ることがあるので、ここで片付ける。作業中のものは DB に行があるので消さない。
    """
    try:
        with app.app_context():
            db = get_db()
            removed = remove_orphan_uploads(app.config["UPLOAD_DIR"], known_stored_paths(db))
            import_ids = {row[0] for row in db.execute("SELECT id FROM table_imports")}
            removed_dirs = remove_orphan_import_dirs(app.config["TABLES_DIR"], import_ids)
            # もう無い取り込みの AI整形の結果と、どこからも使われない AI の応答も消す（データを残さない）。
            # ai_items が先に消えていて（取り込み設定の削除など）応答だけ残っていることもあるので、毎回見る。
            sweep_orphan_ai(db)
            # 前の版で空になった表に残っている番号の続き（何件取り込んだか）も忘れる（design.md 3.3「履歴は持たない」）
            if forget_id_counters(db):
                db.execute("PRAGMA wal_checkpoint(TRUNCATE)")   # 置き換える前の値を app.db-wal に残さない
        if removed:
            print(f"[app] 参照されていないアップロードファイルを {removed} 件片付けました")
        if removed_dirs:
            print(f"[app] 参照されていない取り込みのフォルダを {removed_dirs} 件片付けました")
    except Exception as exc:  # 起動は止めない
        print(f"[app] 残ったファイルの整理に失敗しました（{exc.__class__.__name__}: {exc}）")


def _recover_jobs(app: Flask) -> None:
    """前回の終了時に動いていたジョブを「中断」にする。"""
    try:
        with app.app_context():
            recover_interrupted()
    except Exception as exc:  # 起動は止めない
        print(f"[app] 中断したジョブの整理に失敗しました（{exc.__class__.__name__}: {exc}）")


# --- 起動の設定 -------------------------------------------------------------
# 待ち受け先（HOST・PORT）はファイルの上のほうで環境変数から読む。
THREADS = 8
# waitress が応答の本文を先読みして溜める上限（既定は 16MB。flask run の開発サーバーには無い設定）。
# 溜められる分はアプリ側では「送り終えた」ように見えるため、既定のままだと 16MB 未満の zip は
# 途中で通信が切れても消えてしまう（core/purge.purge_after_send・design.md 3.3「欠けないダウンロード」）。
# 小さくすると、送れた分だけ読み進めるので途中で切れたことに気づける。ループバックでは速度はほぼ変わらない
# （30MB の本文で 16MB: 約130ms、16KB: 約100ms）。それでも最後の数十KB（この上限＋OS の送受信の溜め。
# 2026-09-19 の測定で約48KB）は相手が受け取る前に送り終えたことになるので、それより小さい md・zip は
# 途中で切れても消える（design.md 3.3・8.0）。
OUTBUF_HIGH_WATERMARK = 16 * 1024
WAITRESS_OPTIONS = {"threads": THREADS, "outbuf_high_watermark": OUTBUF_HIGH_WATERMARK}
# エラー画面に詳細を出すか。通常は False のまま。
DEBUG = False
