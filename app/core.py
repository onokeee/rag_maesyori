"""帳票・一覧表の両フローで共通に使う土台。

SQLite のスキーマ・移行・データアクセス（database）・安全なファイル名（naming）・
Markdown テキスト処理（mdtext）・アップロードの保存と事前チェック（files）・
帳票登録が置いた Excel の一時的な覚え（workbook_cache。メモリだけ）・ジョブ実行（jobs）・
取り込んだデータの削除（purge）を、この順に1ファイルにまとめてある。
"""
from __future__ import annotations

import hashlib
import io
import json
import logging
import posixpath
import queue
import re
import secrets
import shutil
import sqlite3
import threading
import time
import unicodedata
import zipfile
import zlib
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import TYPE_CHECKING, Callable
from uuid import uuid4

from flask import current_app, g, has_app_context





# ====================================================================================================
# 元 models/database.py
# SQLite のスキーマ・マイグレーションとデータアクセス（帳票）。
#
# 一覧表・ジョブ・AI のテーブルもここで作る。一覧表のデータアクセスは tables/store.py、
# ジョブは core/jobs.py が持つ。
# ====================================================================================================

if TYPE_CHECKING:
    from app.extract import FieldDef, PatternDef, SheetDef

BUSY_TIMEOUT_MS = 5000

# 帳票の状態（導出値）: unread=読み取り前 / reviewing=確認中 / confirmed=確定済み / modified=修正中
# 画面に出す語は templates/base.html の STATE_LABELS が持つ
_STATE_SQL = """CASE
    WHEN d.data_json IS NULL THEN 'unread'
    WHEN d.confirmed_json IS NULL THEN 'reviewing'
    WHEN d.confirmed_json != d.data_json THEN 'modified'
    ELSE 'confirmed' END"""

# ---- スキーマ -----------------------------------------------------------------

# 第1版（作り直し前）のスキーマ。古いDBにも新規DBにも同じ形で適用できるよう IF NOT EXISTS
_SCHEMA_V1 = """
CREATE TABLE IF NOT EXISTS patterns (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    version TEXT NOT NULL DEFAULT 'v1',
    description TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'draft',          -- draft / active / inactive
    image_processing TEXT NOT NULL DEFAULT 'none', -- none / vision
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS pattern_sheets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_id INTEGER NOT NULL REFERENCES patterns(id) ON DELETE CASCADE,
    sheet_name TEXT NOT NULL
);

-- 古いDBにある required 列（項目の「必須」・シートの「必須」）はもう使わない。画面に必須の設定が
-- 無いので新しいDBでは作らず、読み書きもしない（古いDBには既定値のまま残る。あっても無くても同じ）
CREATE TABLE IF NOT EXISTS pattern_fields (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_id INTEGER NOT NULL REFERENCES patterns(id) ON DELETE CASCADE,
    sort_order INTEGER NOT NULL DEFAULT 0,
    field_name TEXT NOT NULL,
    display_name TEXT NOT NULL,
    candidates TEXT NOT NULL DEFAULT '[]',         -- JSON配列
    data_type TEXT NOT NULL DEFAULT 'string',
    extraction_rule TEXT NOT NULL DEFAULT '{}'     -- JSON {"direction", "columns", "section"}
);

CREATE TABLE IF NOT EXISTS pattern_samples (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_id INTEGER NOT NULL REFERENCES patterns(id) ON DELETE CASCADE,
    file_name TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    pattern_id INTEGER REFERENCES patterns(id) ON DELETE SET NULL,
    status TEXT NOT NULL DEFAULT 'uploaded',       -- 旧: uploaded / extracted / registered（使わない）
    data_json TEXT,
    markdown TEXT,                                 -- 使わない（互換のため残す）
    created_at TEXT NOT NULL,
    registered_at TEXT                             -- 使わない（confirmed_at へ移行）
);
"""

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


def _m1_base(conn: sqlite3.Connection) -> None:
    _run_script(conn, _SCHEMA_V1)


def _m2_forms(conn: sqlite3.Connection) -> None:
    _add_column(conn, "patterns", "title_fields", "TEXT NOT NULL DEFAULT '[]'")
    _add_column(conn, "patterns", "md_options", "TEXT NOT NULL DEFAULT '{}'")
    _add_column(conn, "patterns", "version_no", "INTEGER NOT NULL DEFAULT 1")
    _add_column(conn, "pattern_fields", "unit", "TEXT NOT NULL DEFAULT ''")
    _add_column(conn, "pattern_fields", "rag_output", "TEXT NOT NULL DEFAULT 'show'")
    _add_column(conn, "documents", "confirmed_json", "TEXT")
    _add_column(conn, "documents", "confirmed_at", "TEXT")
    _add_column(conn, "documents", "title", "TEXT NOT NULL DEFAULT ''")
    # 旧「登録済み」は確定済みとして引き継ぐ
    conn.execute("""UPDATE documents SET confirmed_json = data_json,
                        confirmed_at = COALESCE(registered_at, created_at)
                    WHERE status = 'registered' AND data_json IS NOT NULL AND confirmed_json IS NULL""")
    conn.execute("UPDATE documents SET confirmed_at = registered_at "
                 "WHERE confirmed_at IS NULL AND registered_at IS NOT NULL AND confirmed_json IS NOT NULL")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_hash ON documents(file_hash)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_pattern ON documents(pattern_id)")


def _m3_tables(conn: sqlite3.Connection) -> None:
    _run_script(conn, _SCHEMA_TABLES)


def _m4_form_batches(conn: sqlite3.Connection) -> None:
    """帳票のまとめ取り込み（1回の選択で複数ファイル）。同じ batch_id の帳票を順に確認して zip で渡す。"""
    _add_column(conn, "documents", "batch_id", "TEXT NOT NULL DEFAULT ''")
    _add_column(conn, "documents", "batch_order", "INTEGER NOT NULL DEFAULT 0")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_batch ON documents(batch_id)")


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

    ログインは無いので利用者は分からないが、セッションのクッキー（views.current_session_id）で
    ブラウザごとの作業場所は分けられる。ほかのブラウザの帳票・取り込みは見えない（404）。
    古いDBの行は NULL（持ち主が分からない）のままで、これまでどおり扱う（起動時の片付けで消える）。
    帳票の種類・取り込み設定は「設定」なので分けない（みんなで使う）。
    """
    _add_column(conn, "documents", "session_id", "TEXT")
    _add_column(conn, "table_imports", "session_id", "TEXT")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_session ON documents(session_id)")
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
    template_id / template_version_id は取り込み自身の id にそろえて残す（aiproc・core.purge が
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

    - 互換ビュー table_template_versions: _m9 が一時的に置いたもの。aiproc.runner.load_spec が
      table_imports.spec_json を直接読むようになったので要らない。
    - table_imports.period_json: 書く側がもう無く、常に '{}' のまま（期間は取り込み設定が持つ）。
    """
    conn.execute("DROP VIEW IF EXISTS table_template_versions")
    if "period_json" in _column_names(conn, "table_imports"):
        conn.execute("ALTER TABLE table_imports DROP COLUMN period_json")


def _m11_document_touch(conn: sqlite3.Connection) -> None:
    """帳票に「最後にさわった時刻」を持たせる（見回りが作業中の帳票を消さないように）。

    一覧表（table_imports）は updated_at を持っていて、design.md 3.3 の「2時間さわられて
    いないものを捨てる」どおりに動く。帳票だけ取り込んだ時刻で切られていた。
    """
    _add_column(conn, "documents", "updated_at", "TEXT")
    conn.execute("UPDATE documents SET updated_at = COALESCE(confirmed_at, created_at) WHERE updated_at IS NULL")


def _m12_drop_pattern_samples(conn: sqlite3.Connection) -> None:
    """見本の Excel の控え（pattern_samples）を落とす。

    利用者の指示（2026-09-21）:「帳票登録で、見本のExcelは置かずに、設定だけ保持するように
    してほしい」。置いた Excel はその場で読み取るだけで、保存も控えもしない（views/form_types.py）。
    控えが残っていると、もう無いファイルを指し続ける行になるので表ごと落とす。
    uploads/samples に残ったファイルは、起動時に core.files.remove_sample_dir がフォルダごと片付ける。
    """
    conn.execute("DROP TABLE IF EXISTS pattern_samples")


def _m13_ai_connections(conn: sqlite3.Connection) -> None:
    """AI接続（APIキー・接続先・モデル）を「ブラウザごと」に持つ（利用者の指示 2026-09-21）。

    利用者の指示:「AI接続の設定は、もともとの位置ヘッダーの画面右上『AI接続』に移動させる。全部空欄にしておいて
    cookieでユーザー毎に登録内容をずっと保持させるようにしてほしい」。
    これまでは data/model_settings.yaml 1つを全員で使っていたので、社内LANで数人が使うと互いのキーを
    上書きし合い、1つのキーを共有していた。持ち主は帳票・取り込みと同じ session_id（views.current_session_id）。
    この表は「設定」なので、取り込みを捨てる片付け（core.purge_session / sweep_stale / purge_all_pending）の
    対象にしない（core.SETTINGS_TABLES）。最後の接続確認の結果もここに持ち、ヘッダーはそれを表示するだけ
    （画面を開くたびに確認しに行かない）。
    """
    conn.execute("""CREATE TABLE IF NOT EXISTS ai_connections (
        session_id TEXT PRIMARY KEY,                   -- ブラウザの作業場所（views.current_session_id）
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
    """帳票の種類に、登録したときのシートの中身（覚えたシート）を持たせる（利用者の指示 2026-09-22）。

    利用者の指示:「再度シートを置かなくても、登録したときにシートのセル番地と文字情報を記憶しておけばだせるはず」。
    Excel のファイルは今までどおり保存しない（2026-09-21「見本のExcelは置かずに、設定だけ保持する」）。
    覚えるのはシートの中身だけ: シート名・セルの番地と文字（型のある値）・結合・太字・塗りの有る無し。
    セルの色は覚えない（同日の指示「セル色の情報は不要」。塗りの有る無しと「同じ塗りか」の番号だけ。
    forms.fill_group）。書き方は forms.WorkbookInfo.to_json。種類1つにつき1つ（最後に置いた Excel の分）で、
    種類を消せば一緒に消える（ON DELETE CASCADE）。この表は「設定」なので、取り込みを捨てる片付けの
    対象にしない（core.SETTINGS_TABLES）。
    """
    conn.execute("""CREATE TABLE IF NOT EXISTS pattern_books (
        pattern_id INTEGER PRIMARY KEY REFERENCES patterns(id) ON DELETE CASCADE,
        file_name TEXT NOT NULL,                       -- 置いた Excel のファイル名（表示用。ファイルそのものは無い）
        file_hash TEXT NOT NULL,                       -- その Excel の sha256（ブラウザが送る book_hash と同じもの）
        book_json TEXT NOT NULL,                       -- 覚えたシートの中身（forms.WorkbookInfo.to_json）
        saved_at TEXT NOT NULL                         -- 置いた日時
    )""")


# PRAGMA user_version = 適用済みの件数。追加は末尾にだけ行う
MIGRATIONS: list[Callable[[sqlite3.Connection], None]] = [_m1_base, _m2_forms, _m3_tables, _m4_form_batches,
                                                          _m5_purge_scope, _m6_ai_items_per_import,
                                                          _m7_llm_calls_owner, _m8_session_scope, _m9_import_spec,
                                                          _m10_drop_unused, _m11_document_touch,
                                                          _m12_drop_pattern_samples, _m13_ai_connections,
                                                          _m14_pattern_books]


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
    # ページに帳票の値・設備名・人名が平文で残り、DBファイル（と app.db-wal）から読めてしまう。
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


def _all(sql: str, args=()) -> list[dict]:
    return [dict(r) for r in get_db().execute(sql, args).fetchall()]


def _one(sql: str, args=()) -> dict | None:
    row = get_db().execute(sql, args).fetchone()
    return dict(row) if row else None


def _exec(sql: str, args=()) -> int:
    db = get_db()
    cur = db.execute(sql, args)
    db.commit()
    return cur.lastrowid


def _dumps(value) -> str:
    return value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)


# ---- patterns ---------------------------------------------------------------

def list_patterns() -> list[dict]:
    return _all("""
        SELECT p.*,
               (SELECT COUNT(*) FROM pattern_fields f WHERE f.pattern_id = p.id) AS field_count,
               (SELECT COUNT(*) FROM documents d WHERE d.pattern_id = p.id AND d.confirmed_json IS NOT NULL) AS document_count,
               (SELECT COUNT(*) FROM documents d WHERE d.pattern_id = p.id AND d.confirmed_json IS NULL) AS working_count
        FROM patterns p ORDER BY p.name, p.version
    """)


def count_active_patterns() -> int:
    return get_db().execute("SELECT COUNT(*) FROM patterns WHERE status = 'active'").fetchone()[0]


def create_pattern(name: str) -> int:
    ts = now()
    return _exec(
        "INSERT INTO patterns (name, version, description, created_at, updated_at) VALUES (?, 'v1', '', ?, ?)",
        (name, ts, ts),
    )


def _field_def(row: dict) -> FieldDef:
    from app.extract import FieldDef, PatternDef, SheetDef   # extract が core を読み込むので、ここで
    fd = FieldDef(
        field_name=row["field_name"],
        display_name=row["display_name"],
        candidates=json.loads(row["candidates"]),
        data_type=row["data_type"],
        direction=json.loads(row["extraction_rule"]).get("direction", "auto"),
    )
    # unit / rag_output は WP-forms が FieldDef に追加する。未追加でも属性として持たせる
    fd.unit = row.get("unit") or ""
    fd.rag_output = row.get("rag_output") or "show"
    rule = json.loads(row["extraction_rule"])
    fd.table_columns = list(rule.get("columns") or [])
    # 探す区画（発行側・回答側など、同じラベルが並ぶ帳票でどちらを読むか）
    fd.section = rule.get("section") or ""
    # クリックで作った項目の控え（見本でのシート名とセル番地）
    fd.sheet_name = rule.get("sheet_name") or ""
    fd.label_cell = rule.get("label_cell") or ""
    fd.cell = rule.get("cell") or ""
    fd.renamed = bool(rule.get("renamed"))
    return fd


def _extraction_rule(f: FieldDef) -> dict:
    """pattern_fields.extraction_rule の JSON。明細表は列見出しも持つ。"""
    rule = {"direction": f.direction}
    columns = getattr(f, "table_columns", None)
    if f.data_type == "table" and columns:
        rule["columns"] = list(columns)
    section = getattr(f, "section", "") or ""
    if section:
        rule["section"] = section
    for key in ("sheet_name", "label_cell", "cell"):
        value = getattr(f, key, "") or ""
        if value:
            rule[key] = value
    if getattr(f, "renamed", False):
        rule["renamed"] = 1   # 見出しを手で直した項目
    return rule


def load_pattern(pattern_id: int | None) -> PatternDef | None:
    from app.extract import FieldDef, PatternDef, SheetDef   # extract が core を読み込むので、ここで
    row = _one("SELECT * FROM patterns WHERE id = ?", (pattern_id,))
    if row is None:
        return None
    sheets = _all("SELECT * FROM pattern_sheets WHERE pattern_id = ? ORDER BY id", (pattern_id,))
    fields = _all("SELECT * FROM pattern_fields WHERE pattern_id = ? ORDER BY sort_order, id", (pattern_id,))
    pattern = PatternDef(
        id=row["id"],
        name=row["name"],
        version=row["version"],
        description=row["description"],
        status=row["status"],
        image_processing=row["image_processing"],
        sheets=[SheetDef(s["sheet_name"]) for s in sheets],
        fields=[_field_def(f) for f in fields],
    )
    pattern.title_fields = json.loads(row.get("title_fields") or "[]")
    pattern.md_options = json.loads(row.get("md_options") or "{}")
    pattern.version_no = row.get("version_no") or 1
    return pattern


def load_active_patterns() -> list[PatternDef]:
    ids = [r["id"] for r in _all("SELECT id FROM patterns WHERE status = 'active' ORDER BY name, version")]
    return [load_pattern(i) for i in ids]


def save_pattern(pattern: PatternDef, status: str) -> None:
    """種類のメタ情報を更新し、シート・項目定義を置き換える。保存ごとに version_no を+1。"""
    db = get_db()
    with db:
        db.execute(
            """UPDATE patterns SET name=?, version=?, description=?, image_processing=?, status=?,
                      title_fields=?, md_options=?, version_no=version_no + 1, updated_at=? WHERE id=?""",
            (pattern.name, pattern.version, pattern.description, pattern.image_processing, status,
             json.dumps(list(getattr(pattern, "title_fields", None) or []), ensure_ascii=False),
             json.dumps(dict(getattr(pattern, "md_options", None) or {}), ensure_ascii=False),
             now(), pattern.id),
        )
        db.execute("DELETE FROM pattern_sheets WHERE pattern_id = ?", (pattern.id,))
        db.execute("DELETE FROM pattern_fields WHERE pattern_id = ?", (pattern.id,))
        db.executemany(
            "INSERT INTO pattern_sheets (pattern_id, sheet_name) VALUES (?, ?)",
            [(pattern.id, s.sheet_name) for s in pattern.sheets],
        )
        db.executemany(
            """INSERT INTO pattern_fields
               (pattern_id, sort_order, field_name, display_name, candidates, data_type, extraction_rule,
                unit, rag_output)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            [
                (pattern.id, i, f.field_name, f.display_name, json.dumps(f.candidates, ensure_ascii=False),
                 f.data_type, json.dumps(_extraction_rule(f), ensure_ascii=False),
                 getattr(f, "unit", "") or "", getattr(f, "rag_output", "show") or "show")
                for i, f in enumerate(pattern.fields)
            ],
        )
    row = db.execute("SELECT version_no FROM patterns WHERE id = ?", (pattern.id,)).fetchone()
    if row is not None:
        try:
            pattern.version_no = row[0]
        except AttributeError:
            pass


def set_pattern_status(pattern_id: int, status: str) -> None:
    _exec("UPDATE patterns SET status = ?, updated_at = ? WHERE id = ?", (status, now(), pattern_id))


def delete_pattern(pattern_id: int) -> None:
    _exec("DELETE FROM patterns WHERE id = ?", (pattern_id,))


# 見本の Excel は保存しないので、その控え（add_sample / list_samples …）は持たない。
# 帳票の種類が持つのは設定（シート名・見出しのセル・値のセル・読み取る向き・項目名）と、
# 登録したときのシートの中身（pattern_books。下）だけ。


# ---- pattern_books（覚えたシート）-------------------------------------------------------
# 帳票の種類1つにつき、最後に置いた Excel のシートの中身（forms.WorkbookInfo.to_json）を1つ持つ。
# Excel のファイルではない（セルの番地と文字・結合・太字・塗りの有る無しだけ。色は無い）。

def save_pattern_book(pattern_id: int, file_name: str, file_hash: str, book_json: str) -> None:
    """覚えたシートを入れる（前のものがあれば入れ替える）。"""
    _exec(
        """INSERT INTO pattern_books (pattern_id, file_name, file_hash, book_json, saved_at)
           VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(pattern_id) DO UPDATE SET file_name = excluded.file_name, file_hash = excluded.file_hash,
                                                book_json = excluded.book_json, saved_at = excluded.saved_at""",
        (pattern_id, file_name, file_hash, book_json, now()),
    )


def load_pattern_book(pattern_id: int | None) -> dict | None:
    """覚えたシート（book_json を含む行）。無ければ None（覚える前に登録した種類）。"""
    return _one("SELECT * FROM pattern_books WHERE pattern_id = ?", (pattern_id,))


def pattern_book_meta(pattern_id: int | None) -> dict | None:
    """覚えたシートの見出し（ファイル名・sha256・置いた日時・JSON の大きさ）。中身は読まない。"""
    return _one("SELECT pattern_id, file_name, file_hash, saved_at, LENGTH(book_json) AS size "
                "FROM pattern_books WHERE pattern_id = ?", (pattern_id,))


def delete_pattern_book(pattern_id: int) -> None:
    """覚えたシートだけを捨てる（種類は残る。テスト・覚える前の状態に戻すとき用）。"""
    _exec("DELETE FROM pattern_books WHERE pattern_id = ?", (pattern_id,))


# ---- documents --------------------------------------------------------------
# 状態は導出する: data_json なし→unread、confirmed_json なし→reviewing、
# confirmed_json != data_json→modified、それ以外→confirmed

def create_document(file_name: str, file_hash: str, stored_path: str, pattern_id: int | None = None,
                    batch_id: str = "", batch_order: int = 0, session_id: str | None = None) -> int:
    """帳票を1件作る。session_id は取り込んだブラウザ（views.current_session_id）。"""
    return _exec(
        "INSERT INTO documents (file_name, file_hash, stored_path, pattern_id, batch_id, batch_order, session_id, "
        "created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (file_name, file_hash, stored_path, pattern_id, batch_id, batch_order, session_id, now(), now()),
    )


# 持ち主で絞る条件。持ち主の分からない行（この仕組みより前のDBの行）は、これまでどおり誰からでも扱える
_OWNED = "(d.session_id IS NULL OR d.session_id = ?)"


def _owner_where(session_id: str | None) -> tuple[str, list]:
    return (_OWNED, [session_id]) if session_id else ("", [])


def get_document(doc_id: int) -> dict | None:
    return _one(f"""
        SELECT d.*, {_STATE_SQL} AS state,
               p.name AS pattern_name, p.version AS pattern_version
        FROM documents d LEFT JOIN patterns p ON p.id = d.pattern_id
        WHERE d.id = ?
    """, (doc_id,))


_LIST_COLUMNS = f"""
        SELECT d.id, d.file_name, d.file_hash, d.stored_path, d.pattern_id, d.title, d.batch_id, d.batch_order,
               d.session_id, d.created_at, d.confirmed_at, {_STATE_SQL} AS state,
               p.name AS pattern_name, p.version AS pattern_version,
               -- 帳票の種類が削除されても、読み取ったときの種類名を表示に使う
               CASE WHEN json_valid(d.data_json) THEN json_extract(d.data_json, '$.pattern.name') END
                   AS extraction_pattern_name
        FROM documents d LEFT JOIN patterns p ON p.id = d.pattern_id"""


def list_batch_documents(batch_id: str, session_id: str | None = None) -> list[dict]:
    """まとめ取り込み（1回の選択で複数ファイル）の帳票を、選んだ順に返す。

    session_id を渡すと、そのブラウザの帳票だけを返す（ほかの人のまとまりは見えない）。
    """
    if not batch_id:
        return []
    owner, args = _owner_where(session_id)
    where = "WHERE d.batch_id = ?" + (f" AND {owner}" if owner else "")
    return _all(f"{_LIST_COLUMNS} {where} ORDER BY d.batch_order, d.id", (batch_id, *args))


def list_confirmed_documents(ids: list[int] | None = None, session_id: str | None = None) -> list[dict]:
    """確定済みの版を持つ帳票（修正中も確定済みの版を持つ）。ids・session_id 指定で絞り込み。"""
    sql = f"""SELECT d.*, {_STATE_SQL} AS state, p.name AS pattern_name, p.version AS pattern_version
              FROM documents d LEFT JOIN patterns p ON p.id = d.pattern_id
              WHERE d.confirmed_json IS NOT NULL"""
    owner, args = _owner_where(session_id)
    if owner:
        sql += f" AND {owner}"
    if ids is not None:
        if not ids:
            return []
        sql += f" AND d.id IN ({', '.join('?' for _ in ids)})"
        return _all(sql + " ORDER BY d.id", (*args, *ids))
    return _all(sql + " ORDER BY d.id", tuple(args))


def save_draft(doc_id: int, data_json, title: str | None = None) -> None:
    """作業中の値を保存する（dict も可）。title は検索用（タイトル項目の値）。"""
    if title is None:
        _exec("UPDATE documents SET data_json = ?, updated_at = ? WHERE id = ?",
              (_dumps(data_json), now(), doc_id))
    else:
        _exec("UPDATE documents SET data_json = ?, title = ?, updated_at = ? WHERE id = ?",
              (_dumps(data_json), title, now(), doc_id))


def confirm_document(doc_id: int, title: str | None = None) -> bool:
    """作業中の値を確定済みの版にする。読み取り前（data_json なし）なら False。"""
    sets, args = ["confirmed_json = data_json", "confirmed_at = ?", "updated_at = ?"], [now(), now()]
    if title is not None:
        sets.append("title = ?")
        args.append(title)
    db = get_db()
    cur = db.execute(f"UPDATE documents SET {', '.join(sets)} WHERE id = ? AND data_json IS NOT NULL",
                     (*args, doc_id))
    db.commit()
    return cur.rowcount > 0


def reset_document(doc_id: int, pattern_id: int | None, data_json) -> None:
    """選び直して再読み取り：作業中の値を置き換える（確定済みの版は残る→修正中になる）。"""
    _exec("UPDATE documents SET pattern_id = ?, data_json = ?, updated_at = ? WHERE id = ?",
          (pattern_id, None if data_json is None else _dumps(data_json), now(), doc_id))


def find_confirmed_by_hash(file_hash: str, exclude_id: int | None = None,
                           session_id: str | None = None) -> dict | None:
    """同じ中身の確定済みの帳票（取り込み直しの注意に出す）。ほかの人の帳票は知らせない。"""
    owner, args = _owner_where(session_id)
    sql = ("SELECT d.id, d.file_name, d.title FROM documents d WHERE d.file_hash = ? "
           "AND d.confirmed_json IS NOT NULL AND d.id != ?")
    if owner:
        sql += f" AND {owner}"
    return _one(sql + " ORDER BY d.id LIMIT 1",
                (file_hash, exclude_id if exclude_id is not None else -1, *args))


def update_document(doc_id: int, **columns) -> None:
    assignments = ", ".join(f"{name} = ?" for name in columns)
    _exec(f"UPDATE documents SET {assignments} WHERE id = ?", (*columns.values(), doc_id))


# 帳票を消すのは core/purge.py（元のファイルと DB の行をまとめて消す）


# ---- ai_connections（ブラウザごとの AI接続） ---------------------------------------------------
# 持ち主は session_id（views.current_session_id）。行は「保存」か「接続の確認」で初めてできる（読むだけでは作らない）。
# 取り込みの片付け（core.purge_*）では消えない「設定」。長く使われていない行だけ core.sweep_stale_ai_connections が消す。

AI_CONNECTION_FIELDS = ("api_key", "chat_url", "models_url", "model")


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


def save_ai_connection(session_id: str, *, api_key: str | None = None, chat_url: str | None = None,
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
# Markdown 出力のテキスト処理（帳票・一覧表で共通）。
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
# 丸数字は帳票の手順・項目番号でごく普通に使われるので、囲み文字だけは NFKC をかけずに残す。
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


def nfkc_value(text) -> str:
    """値の正規化。NFKC（囲み文字は残す）＋空白の畳み込み（改行は保持、行末の空白と前後の空行は除く）。"""
    if text is None:
        return ""
    s = nfkc_keep_enclosed(str(text)).replace("\r\n", "\n").replace("\r", "\n")
    lines = [_SPACES.sub(" ", line).strip() for line in s.split("\n")]
    return "\n".join(lines).strip("\n")


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
# 帳票登録の見本の Excel は保存しない（read_upload でメモリに読むだけ）ので、事前チェックは
# パスのほかに「中身そのもの（bytes）」も受け取れる。
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
MAX_MERGED_CELLS = 2_000_000
# 帳票（通常モードで開く）の結合セルの面積の上限。openpyxl は結合範囲ごとに中のセルを1つずつ片付け、
# 左上のセルに罫線があれば範囲の辺のセルを1つずつ作るので、面積に比例して遅くなる（行全体の結合 120 個
# ＝約200万セルで約16秒、罫線付きの列全体の結合1つで1分以上）。帳票は画面を開くたびにブックを開き直すので、
# 1回あたり2秒程度に収まるこの値で断る（行全体の結合なら12個まで通る。ふつうの帳票は数千セル）。
# 一覧表は読み取り専用モードで開き結合を展開しないので、上の MAX_MERGED_CELLS のまま。
FORM_MAX_MERGED_CELLS = 200_000
# ハイパーリンク・コメントの範囲の面積（ブック全体の合計）の上限。openpyxl はリンク（<hyperlink ref>）とコメント
# （<comment ref>）の範囲のセルを1つずつ作る（帳票の通常モードでも、一覧表の tables.excel_source でも）。
# シート全体を指す範囲1つで開くのが終わらなくなる（A1:XFD50 の約82万セルで一覧表は約28秒）。
# ふつうのリンク・コメントは1セルか数セルなので、1回あたり2秒以内に収まるこの値で断る。
MAX_LINKED_CELLS = 50_000
# セル数の上限（ブック全体の <c> の数）。openpyxl は開くときにセルを1つずつ作るので、
# 圧縮すると小さいが展開すると大量のセルがあるブック（64KB で 100万セルなど）は、上のサイズ制限を通っても固まる。
# 一覧表は tables.excel_source が別に上限（EXCEL_MAX_CELLS）を持ち「CSVで保存」と案内するので、ここは最後の砦の値。
MAX_CELLS = 1_000_000
# 図形（描画）の上限。openpyxl は開くときに、シートが参照する描画部品（xl/drawings/*.xml）をシートごとに読み直し、
# 図形（アンカー）を1つずつ作る。1つの大きな描画を多数のシートから参照させると、圧縮後は小さくても
# 開くたびに（書類の画面を開くたびにも）数十秒〜終わらない。そこで「描画部品の図形数 × 参照するシート数」と
# 「描画部品の展開後の大きさ × 参照するシート数」の合計を数えて上限を設ける。
# 普通の帳票は1シートに数十個程度の図形なので、十分に余裕のある値にしている。
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
# アップロードの保存先（UPLOAD_DIR 直下）と save_upload が付けるファイル名の形。
# 帳票登録の見本（samples）はもう保存しないので、ここには入れない（前の版の残骸は remove_sample_dir が消す）
UPLOAD_SUBDIRS = ("documents", "tables")
_STORED_NAME = re.compile(r"[0-9a-f]{32}\.[A-Za-z0-9]+")
# 前の版が見本の Excel を置いていたフォルダ（UPLOAD_DIR 直下）
SAMPLE_SUBDIR = "samples"


class UploadError(Exception):
    """利用者に見せる日本語メッセージを持つ例外。"""


@dataclass
class StoredFile:
    stored_path: str   # UPLOAD_DIR からの相対パス（/ 区切り）
    file_name: str     # 元のファイル名（表示・ダウンロード用）
    file_hash: str     # sha256
    size: int


@dataclass
class MemoryFile:
    """保存せずにメモリへ読んだアップロード（帳票登録の見本の Excel）。"""

    data: bytes        # ブックの中身そのもの
    file_name: str     # 元のファイル名（表示用）
    file_hash: str     # sha256

    @property
    def size(self) -> int:
        return len(self.data)


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


def read_upload(storage, allowed: set[str], max_bytes: int) -> MemoryFile:
    """アップロードをディスクに置かずにメモリへ読む。拡張子・サイズ・空ファイルを確認する。

    帳票登録の見本の Excel に使う。置いた Excel はサーバーに残さない（利用者の指示 2026-09-21
    「見本のExcelは置かずに、設定だけ保持するようにしてほしい」）ので、保存先を作らずに
    中身と sha256 だけを返す。呼び出し側は読み取った結果（WorkbookInfo）だけを短い間覚えておく。
    """
    name, _ext = _checked_name(storage, allowed)
    digest = hashlib.sha256()
    chunks: list[bytes] = []
    size = 0
    stream = getattr(storage, "stream", storage)
    while True:
        chunk = stream.read(CHUNK_SIZE)
        if not chunk:
            break
        size += len(chunk)
        if size > max_bytes:
            raise UploadError(f"ファイルが大きすぎます（上限 {_format_size(max_bytes)}）")
        digest.update(chunk)
        chunks.append(chunk)
    if size == 0:
        raise UploadError("ファイルが空です")
    return MemoryFile(data=b"".join(chunks), file_name=name, file_hash=digest.hexdigest())


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

    source: ファイルのパス、またはブックの中身そのもの（bytes）。帳票登録の見本の Excel は
      保存しないので bytes で渡す（zipfile はどちらも同じように読める）。
    max_cells: セル数の上限（省略時は MAX_CELLS）。
    max_merged: 結合セルの面積の上限（省略時は MAX_MERGED_CELLS。帳票は FORM_MAX_MERGED_CELLS を渡す）。
    max_rows: 行（<row>）の数の上限（省略時は max_merged と同じ値）。通常モードの openpyxl は、セルの無い
      <row ht="20" customHeight="1"/> のような行にも行の書式を1つずつ作るので、結合セルと同じくらい開くのが遅くなる
      （100万行で約9秒）。帳票は画面を開くたびに開き直すので、結合セルと同じ目安（1回あたり2秒程度）で断る。
      一覧表は読み取り専用モードで開き行の書式を作らないので、MAX_MERGED_CELLS のままで困らない。
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


def remove_sample_dir(base) -> int:
    """前の版が置いた見本の Excel（uploads/samples）をフォルダごと消す。消したファイルの件数を返す。

    帳票登録は見本の Excel を受け取っても置かなくなった（利用者の指示 2026-09-21）。
    前の版で置いたままのファイルが残ることがあるので、起動時に片付ける。
    """
    folder = Path(base) / SAMPLE_SUBDIR
    if not folder.is_dir():
        return 0
    count = sum(1 for p in folder.rglob("*") if p.is_file())
    shutil.rmtree(folder, ignore_errors=True)
    return count


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
# core.jobs.STALE_AFTER（2分）と同じ長さ。
ORPHAN_GRACE_SECONDS = 120


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
# 元 core/workbook_cache.py
# 置かれた Excel を「読み取った形」だけ、しばらくメモリに覚えておく置き場。
#
# 帳票登録は見本の Excel をサーバーに置かない（利用者の指示 2026-09-21「見本のExcelは置かずに、
# 設定だけ保持するようにしてほしい」）。ブラウザはファイルを選んだまま持っているので、セルを
# クリックするたびに同じ Excel を送り直してくる。そのたびにブックを開き直すと（結合セル・図形を
# 1つずつ作るので）数秒かかることがあるため、開いた結果（WorkbookInfo）だけをここに置く。
# 登録したときのシートの中身（セルの番地と文字）は帳票の種類と一緒に DB が覚える
# （pattern_books。2026-09-22）。ここはそれを開いた結果も同じように短い間だけ置く。
#
# 決まりごと:
#   - ディスクには何も書かない。Excel の中身（bytes）も持たない（開いた結果だけ）。
#   - アプリを終えれば消える（残るのは読み取りの設定と覚えたシートだけ、という約束を崩さない）。
#   - 鍵は「ブラウザの作業場所（session_id）＋ファイルの sha256」。ほかの人が置いたブックは見えない。
#   - 古いもの・多すぎるもの・大きすぎるものは、置いた順に落とす。
#   - waitress は1つのプロセスを複数のスレッドで回すので、出し入れは錠（Lock）の中で行う。
# ====================================================================================================

# 覚えておく時間。帳票登録は「置く → セルをクリックする → 使用開始」を続けて行う作業なので、
# 手が止まっても1回分の作業（30分）で足りる。切れてもブラウザが送り直すので読み直すだけ。
TTL_SECONDS = 30 * 60
# 覚えておく数と、その元になった Excel の大きさの合計。数人が同時に使っても RAM を食い尽くさない値。
MAX_ENTRIES = 8
MAX_BYTES = 64 * 1024 * 1024


@dataclass
class Book:
    """ブラウザが置いた Excel を読み取った結果（中身そのものは持たない）。

    帳票の種類が覚えているシート（pattern_books。views._stored_book が戻す）も同じ形で持つ。
    そのときの size は覚えた JSON の大きさ、saved_at は置いた日時。
    """

    file_name: str
    file_hash: str
    size: int                      # 元の Excel の大きさ（置ける量を数えるため）
    info: object                   # excel.workbook.WorkbookInfo
    used_at: float = field(default_factory=time.monotonic)
    saved_at: str = ""             # 帳票の種類と一緒に覚えたシートなら、その Excel を置いた日時


_books: dict[tuple[str, str], Book] = {}
_lock = threading.Lock()


def _expired(book: Book, now: float) -> bool:
    return now - book.used_at > TTL_SECONDS


def _evict(now: float) -> None:
    """古いもの → 入れた順、の順に落とす（錠の中から呼ぶ）。"""
    for key in [k for k, b in _books.items() if _expired(b, now)]:
        del _books[key]
    total = sum(b.size for b in _books.values())
    while _books and (len(_books) > MAX_ENTRIES or total > MAX_BYTES):
        key, dropped = next(iter(_books.items()))   # dict は入れた順に並ぶ＝いちばん古いもの
        del _books[key]
        total -= dropped.size


def put(session_id: str, book: Book) -> Book:
    """読み取った結果を覚える（同じブックを置き直したら入れ替える）。"""
    key = (session_id or "", book.file_hash)
    with _lock:
        book.used_at = time.monotonic()
        _books.pop(key, None)       # 入れ直して「いちばん新しい」にする
        _books[key] = book
        _evict(book.used_at)
    return book


def get(session_id: str, file_hash: str) -> Book | None:
    """覚えているブックを返す（無ければ None。呼び出し側はブラウザに置き直してもらう）。"""
    key = (session_id or "", file_hash or "")
    now = time.monotonic()
    with _lock:
        book = _books.get(key)
        if book is None:
            return None
        if _expired(book, now):
            del _books[key]
            return None
        book.used_at = now
        _books.pop(key)
        _books[key] = book          # 使ったものを「いちばん新しい」にする
        return book


def clear() -> None:
    """全部忘れる（テスト・片付け用）。"""
    with _lock:
        _books.clear()


def count() -> int:
    with _lock:
        return len(_books)


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

    これを継承した例外（aiproc.runner.AIJobError、tables.pipeline.PipelineError）は、そのメッセージを
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


KIND_LABELS = {"ai_format": "AI整形"}


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
        label = KIND_LABELS.get(row["kind"], "前の処理")
        whose = "この取り込みの" if same_ref else ("別の取り込みの" if row["kind"] in KIND_LABELS else "")
        if row["status"] == "paused":
            job["message"] = (f"{whose}{label}が一時停止中のため待っています。"
                              f"{label}の画面で再開するか中止すると始まります")
        else:
            job["message"] = f"{whose}{label}が終わるのを待っています"
    return job


def get_job(job_id: int) -> dict | None:
    """ジョブ1件（params / progress / result / status_label / finished を付けて返す）。"""
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
#           DB の行（documents / table_imports / jobs / ai_items / llm_calls とそれらの子）。
# 残すもの: 帳票の種類（patterns とその子）・ブラウザごとの AI接続（ai_connections）。これらは「設定」で
#           データではない（SETTINGS_TABLES。取り込みを指す列が後から足されても、ここからは消さない）。
#
# DB の行は表の名前を決め打ちせず、その取り込みを指す列（document_id / import_id）を持つ表を
# sqlite_master から探して消す（後から表が増えても消し残さないため）。
# ====================================================================================================

log = logging.getLogger(__name__)

# 取り込み1件を指す外部キーの列名（この列を持つ表は、その取り込みと一緒に消す）
DOCUMENT_REF_COLUMNS = ("document_id",)
IMPORT_REF_COLUMNS = ("import_id", "table_import_id")
# 取り込みを指す列があっても、まとめては消さない表。llm_calls は AI の生の応答で、同じ文面の行なら
# 別の取り込みの ai_items が同じ応答を使っていることがある（消すと再実行で再課金になる）。
# 持ち主が消えた分は _delete_orphan_llm_calls が「参照されていなければ消す」で扱う。
SHARED_TABLES = ("llm_calls",)
# 「設定」の表。取り込みを捨てる片付け（purge_* / sweep_stale / purge_session）では決して消さない。
# ai_connections はブラウザごとの AI接続（APIキー・接続先。利用者の指示 2026-09-21「ずっと保持」）で、
# 帳票・一覧表と同じ session_id を持ち主に持つが、その人の取り込みを捨てても残す。
# pattern_books は帳票の種類が覚えているシートの中身（セルの番地と文字。Excel のファイルではない。
# 利用者の指示 2026-09-22）で、種類と一緒にしか消えない。
SETTINGS_TABLES = ("patterns", "pattern_sheets", "pattern_fields", "pattern_books", "ai_connections")


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


# ---- 帳票 --------------------------------------------------------------------------

# 帳票を消したときに呼ぶ関数（画面側がメモリに持っている帳票ごとの目印を一緒に捨てるため）。
# 引数は消した帳票の id のリスト。
_DOCUMENT_PURGE_HOOKS: list = []


def on_documents_purged(fn):
    """帳票を消したあとに fn(doc_ids) を呼ぶよう登録する（同じ関数は1回だけ）。デコレーターとしても使える。"""
    if fn not in _DOCUMENT_PURGE_HOOKS:
        _DOCUMENT_PURGE_HOOKS.append(fn)
    return fn


def purge_documents(doc_ids) -> int:
    """帳票を消す（元のファイルと DB の行）。戻り値: 消した帳票の件数。"""
    ids = [int(i) for i in doc_ids]
    if not ids:
        return 0
    db = get_db()
    marks = ", ".join("?" for _ in ids)
    stored = [row[0] for row in db.execute(f"SELECT stored_path FROM documents WHERE id IN ({marks})", ids)]
    # DB の行を先に消す。ファイルを先に消すと、DB の削除が失敗したとき（ロック・強制終了）に
    # 「ファイルの無い帳票」が作業中として残ってしまう。逆なら残るのは行の無いファイルで、起動時に片付く。
    removed = 0
    for doc_id in ids:
        _delete_by_columns(db, DOCUMENT_REF_COLUMNS, doc_id)
        _delete_jobs(db, "document", doc_id)
        removed += db.execute("DELETE FROM documents WHERE id = ?", (doc_id,)).rowcount
    db.commit()
    for path in stored:
        if not remove_upload(path):
            _mark_incomplete()
    for hook in _DOCUMENT_PURGE_HOOKS:
        try:
            hook(ids)
        except Exception:  # 目印の片付けに失敗しても、消すこと自体は終わっている
            log.exception("帳票を消したあとの片付けに失敗しました")
    _shrink(db)
    return removed


def purge_batch(batch_id: str) -> int:
    """まとめ取り込み1回分（同じ batch_id の帳票）をすべて消す。"""
    if not batch_id:
        return 0
    ids = [row[0] for row in get_db().execute("SELECT id FROM documents WHERE batch_id = ?", (batch_id,))]
    return purge_documents(ids)


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
    # DB の行を先に消し、ファイルはそのあと（purge_documents と同じ理由。残ったフォルダ・ファイルは起動時に片付く）
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
    """
    forget_id_counters(db)
    statements = ["PRAGMA wal_checkpoint(TRUNCATE)"]
    if not _jobs_active(db):
        statements.append("VACUUM")
    for statement in statements:
        try:
            db.execute(statement)
        except sqlite3.Error:
            pass


def _jobs_active(db) -> bool:
    try:
        row = db.execute("SELECT 1 FROM jobs WHERE status IN ('queued', 'running', 'paused') LIMIT 1").fetchone()
        return row is not None
    except sqlite3.Error:
        return True   # 確かめられないときは、動いているかもしれないジョブを止めない側に倒す


# ---- 番号の続きを忘れる ------------------------------------------------------------
# 取り込むたびに行を作り、ダウンロードで消す表。AUTOINCREMENT の表は、消したあとも sqlite_sequence に
# 「これまでに使った一番大きい番号」が残り、何件取り込んだかの記録になってしまう（design.md 3.3「履歴は持たない」）。
WORK_TABLES = ("documents", "table_imports", "jobs", "ai_items")
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
#   - 画面を離れたとき: その人が触っていた分を捨てる（purge_session / discard_documents / discard_table_imports）
#   - 新しいファイルを置いたとき: 同じ人の前の分を捨てる（同上）
#   - 動いている間: IDLE_HOURS さわられていないものを捨てる（sweep_stale）
#   - 起動時: ダウンロードしていない帳票・一覧表をすべて捨てる（purge_all_pending）
# あとの2つ（時間切れ・起動時）は、動いているジョブが付いているものには手を出さない（_busy_ids）。
# 残すのは「設定」（帳票の種類・ブラウザごとの AI接続 = SETTINGS_TABLES）だけで、これは purge の対象ではない。
# AI接続だけは、クッキーの寿命（約1年。views.SESSION_LIFETIME）より長くさわられていない行を
# sweep_stale_ai_connections が消す（もう戻って来ないブラウザの APIキーを DB に残さない）。

# 放っておかれた取り込みを捨てるまでの時間。数人が同時に使う社内LANの置き方（2026-09-20）では、
# 画面を閉じた合図（sendBeacon）が届かないこと（ブラウザの強制終了・スリープ・LANの切断）があるので、
# 短めにして「残らないこと」を優先する。長くしたいときはこの数字だけを変える。
IDLE_HOURS = 2
STALE_HOURS = IDLE_HOURS   # 旧名（app.py が参照している）


def _import_ids(db, where: str = "", args=()) -> list[int]:
    return [row[0] for row in db.execute(f"SELECT id FROM table_imports {where}", args)]


def _document_ids(db, where: str = "", args=()) -> list[int]:
    return [row[0] for row in db.execute(f"SELECT id FROM documents {where}", args)]


def purge_all_pending() -> tuple[int, int]:
    """ダウンロードしていない帳票・一覧表をすべて捨てる。戻り値: (帳票の件数, 一覧表の件数)。

    ダウンロードが終わったものはその時点で消えている（purge_after_send）ので、DB に残っている
    帳票・取り込みは「途中のもの」か「確定したがダウンロードしていないもの」しかない。
    続きを開く入口（作業中の一覧）を持たないので、起動時にまとめて捨てる。

    動いているジョブが付いているものには手を出さない。起動直後は _recover_jobs が動いていた
    ジョブを「中断」に直したあとなので、ふつうは1件も残らない。それでも、別のプロセスが同じ DB を
    見ているとき（JupyterLab のターミナルで二重に起動したとき）に、他方が処理中の取り込みを
    消してしまわないようにする。
    """
    db = get_db()
    busy_docs = _busy_ids(db, "document")
    busy_imports = _busy_ids(db, "table_import")
    forms = purge_documents([i for i in _document_ids(db) if i not in busy_docs])
    tables = 0
    for import_id in _import_ids(db):
        if import_id not in busy_imports:
            tables += purge_table_import(import_id)
    return forms, tables


def purge_old_sample_files() -> int:
    """前の版が置いた見本の Excel（uploads/samples）を片付ける。戻り値: 消した件数。

    帳票登録は見本の Excel を置かなくなった（利用者の指示 2026-09-21「見本のExcelは置かずに、
    設定だけ保持するようにしてほしい」）。置いた Excel はその場で読み取るだけで保存しないので、
    ふつうは1件も無い。前の版で置いたままのファイルだけを、起動時にフォルダごと捨てる。
    """
    from flask import current_app


    return remove_sample_dir(current_app.config["UPLOAD_DIR"])


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


def sweep_stale(hours: float = STALE_HOURS) -> tuple[int, int]:
    """しばらくさわられていない帳票・一覧表を捨てる。戻り値: (帳票の件数, 一覧表の件数)。

    帳票も一覧表も「最後にさわった日時」（確定 → 途中保存・読み取り → 取り込み の順に見る）で切る。
    まとめ取り込みは、そのまとまりのどれか1件でも新しければ、まとまりごと残す（50件を上から順に
    見ていくと、まだ手が届いていない帳票だけが画面から消えてしまうため）。
    動いているジョブが付いているものは、そのジョブが終わるまで残す。
    """
    db = get_db()
    limit = _stale_before(hours)
    busy_docs = _busy_ids(db, "document")
    busy_imports = _busy_ids(db, "table_import")
    forms = purge_documents([i for i in _document_ids(
        db,
        "WHERE COALESCE(confirmed_at, updated_at, created_at) < ? "
        "AND (batch_id = '' OR NOT EXISTS (SELECT 1 FROM documents s WHERE s.batch_id = documents.batch_id "
        "AND COALESCE(s.confirmed_at, s.updated_at, s.created_at) >= ?))",
        (limit, limit)) if i not in busy_docs])
    tables = 0
    for import_id in _import_ids(
            db, "WHERE COALESCE(confirmed_at, updated_at, created_at) < ?", (limit,)):
        if import_id in busy_imports:
            continue
        tables += purge_table_import(import_id)
    sweep_stale_ai_connections()
    return forms, tables


# ブラウザごとの AI接続を消すまでの日数。クッキーの寿命（views.SESSION_LIFETIME = 365日）を過ぎたブラウザは
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
# documents / table_imports は持ち主（session_id。views.current_session_id() がブラウザごとに配る）を持つ。
#
# 持ち主の列がまだ無い古い DB でも動くようにしてある:
#   - 番号を指して捨てる（discard_documents / discard_table_imports）… 持ち主を確かめずに捨てる
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


def discard_documents(doc_ids, session_id=None) -> int:
    """指された帳票のうち、その人のもので、処理中でないものを捨てる。戻り値: 捨てた件数。

    画面を離れたとき・新しいファイルを置いたときに呼ぶ。もう無いもの・他人のもの・処理中のものは
    黙って飛ばす（何度呼んでも安全で、無駄な読み書きもしない）。
    """
    db = get_db()
    ids = _owned_ids(db, "documents", doc_ids, session_id)
    if not ids:
        return 0
    busy = _busy_ids(db, "document")
    return purge_documents([i for i in ids if i not in busy])


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


def purge_session(session_id, *, include_busy: bool = False,
                  documents: bool = True, tables: bool = True) -> tuple[int, int]:
    """その人の、まだダウンロードしていない帳票・一覧表をすべて捨てる。戻り値: (帳票の件数, 一覧表の件数)。

    捨てるのは、元のファイル・imports/<id>/（読み込んだ行・控え・作った md）・DB の行・
    AI整形の控えと、ほかから使われなくなった AI の応答（purge_documents / purge_table_import と同じ）。
    残すのは設定（帳票の種類・一覧表の取り込み設定・AI接続）だけ。

    documents / tables で片方だけにできる。帳票取り込みの画面を閉じた合図で
    表の取り込みの画面（同じブラウザの別のタブ）の作業まで巻き込まないために使う。

    動いているジョブ（読み込み・下書き・AI整形）が付いているものは、そのジョブが壊れるので捨てない。
    捨て損ねた分は、ジョブが終わったあと sweep_stale が IDLE_HOURS で片付ける。
    どうしても今すぐ捨てるときだけ include_busy=True（ジョブは次の書き込みで失敗して終わる）。
    """
    if not session_id:
        return 0, 0
    db = get_db()
    doc_ids = _session_ids(db, "documents", session_id) if documents else []
    import_ids = _session_ids(db, "table_imports", session_id) if tables else []
    if not include_busy:
        busy_docs = _busy_ids(db, "document")
        busy_imports = _busy_ids(db, "table_import")
        doc_ids = [i for i in doc_ids if i not in busy_docs]
        import_ids = [i for i in import_ids if i not in busy_imports]
    forms = purge_documents(doc_ids)
    tables = 0
    for import_id in import_ids:
        tables += purge_table_import(import_id)
    return forms, tables
