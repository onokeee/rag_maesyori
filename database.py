from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Callable

from flask import current_app, g




# ====================================================================================================
# 元 models/database.py
# SQLite のスキーマ・マイグレーションとデータアクセス（帳票）。
#
# 一覧表・ジョブ・AI のテーブルもここで作る。一覧表のデータアクセスは tables/store.py、
# ジョブは core/jobs.py が持つ。
# ====================================================================================================

if TYPE_CHECKING:
    from forms import FieldDef, PatternDef, SheetDef

BUSY_TIMEOUT_MS = 5000

# 帳票の状態（導出値）: unread=読み取り前 / reviewing=確認中 / confirmed=確定済み / modified=修正中
# 画面に出す語は templates/components/_ui.html の STATE_LABELS が持つ
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


# PRAGMA user_version = 適用済みの件数。追加は末尾にだけ行う
MIGRATIONS: list[Callable[[sqlite3.Connection], None]] = [_m1_base, _m2_forms, _m3_tables, _m4_form_batches,
                                                          _m5_purge_scope, _m6_ai_items_per_import,
                                                          _m7_llm_calls_owner, _m8_session_scope, _m9_import_spec,
                                                          _m10_drop_unused, _m11_document_touch,
                                                          _m12_drop_pattern_samples]


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
    from forms import FieldDef, PatternDef, SheetDef   # forms → core → database の順に読み込むので、ここで
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
    from forms import FieldDef, PatternDef, SheetDef   # forms → core → database の順に読み込むので、ここで
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
# 帳票の種類が持つのは設定だけ（シート名・見出しのセル・値のセル・読み取る向き・項目名）。


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
