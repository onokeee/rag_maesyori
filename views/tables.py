"""表の取り込み（design.md 2.4）。1画面で全部できる（利用者の指示 2026-09-20）。

画面は /tables の1枚だけ。ファイルを置く → 読み取り方 → 表の範囲 → 列の対応づけ → AI整形（任意）
→ 内容の確認 → 確定してダウンロード（zip）を、同じ画面の「段」として順に開く。
段の中身はこの blueprint が HTML の断片（panel）として返し、保存・実行は JSON でやりとりする。
画面の移動（①→②→③）は無く、URL は変わらない（static/tables.js）。
範囲の決まり（利用者の判断）: 期間の置き換え・投入済みとの差分・取り消しはしない。取り込みごとに
その取り込みの記録だけから全 Markdown を作り、全ファイルを zip で渡す。クロス集計と名寄せ辞書は扱わない。
ダウンロードした取り込みのデータは、zip を送り終えたあとに消す（design.md 3.3・core.purge.purge_after_send）。
読み込み・Markdown 作成・AI整形は core.jobs のジョブ（tables.pipeline / aiproc.runner）で動かし、
同じ画面の中に進み具合を出す。表の範囲・列の段は取り込みの控え（tables.source_cache）を通して読む。
"""
from __future__ import annotations

import io
import re
import threading
import time
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

from flask import Blueprint, abort, current_app, flash, jsonify, redirect, render_template, request, send_file, url_for

from core import jobs, purge
from core.files import UploadError, precheck_excel, remove_upload, save_upload, upload_path
from core.naming import safe_filename_part
from tables import outputs, pipeline, store
from tables.checks import count_levels, has_blocking
from tables.mapping import suggest_columns
from tables.markdown import ai_point_lines, people_index_for, record_block
from tables.source import open_source
from tables.spec import spec_from_suggestions, validate_spec
from views import current_session_id, owns, set_download_name

bp = Blueprint("tables", __name__, url_prefix="/tables")

EXCEL_EXT = {".xlsx", ".xlsm"}
GRID_ROWS = 60
GRID_COLS = 30
GRID_CELL_CHARS = 40
DATA_PAGE = 100
ISSUES_SHOWN = 200

# 画面で選べる役割（列の対応づけ）。ここに無い役割（人名・数値・区分など）は候補のまま使い、画面では「その他」に見せる
SCREEN_ROLES = [("key", "識別番号"), ("date", "日付"), ("entity", "設備"), ("log", "AI整形の対象（追記ログ）"),
                ("attribute", "その他")]
SCREEN_ROLE_KEYS = {role for role, _label in SCREEN_ROLES}
TYPE_LABELS = {"code": "コード", "string": "文字", "text": "長文", "date": "日付", "datetime": "日時", "time": "時刻",
               "number": "数値", "enum": "選択肢", "status": "状態"}
KIND_LABELS = {"header": "見出し", "data": "データ", "subtotal": "小計・合計", "note": "注記", "continuation": "継続行",
               "excluded": "除外", "title": "表題", "blank": "空行"}
TABLE_KIND_LABELS = {"list": "一覧表", "crosstab": "クロス集計", "form_like": "帳票らしい", "unknown": "不明"}
SCOPE_LABELS = {"pending": "まだ整形していない行と、内容が変わった行", "errors": "エラーになった行だけ",
                "flagged": "要確認の行だけ", "all": "すべての行をやり直す"}
# CSV の文字コード・区切り文字は画面の選択肢だけを受け付ける（他の値は読み込みで落ちて画面が開けなくなるため）
# ダウンロードでデータが消えることの案内（design.md 3.3）
DELETE_ON_DOWNLOAD_NOTE = ("zip をダウンロードすると、この取り込みのデータはサーバーから消えます"
                           "（もう一度ダウンロードすることはできません）。正規化CSVは zip の"
                           "「管理用_RAGには入れない」フォルダにも入っています。")
DELETE_ON_DOWNLOAD_CONFIRM = ("ダウンロードすると、この取り込みのデータはサーバーから消えます。"
                              "もう一度ダウンロードすることはできません。")
ENCODING_CHOICES = [("utf-8-sig", "UTF-8（BOM付き）"), ("utf-8", "UTF-8"), ("cp932", "CP932（Shift_JIS）"),
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
    imp = store.get_import(import_id)
    if imp is None or not owns(imp):
        abort(404)
    return imp


def _spec_for(imp: dict):
    return pipeline.spec_for_import(imp)


def _is_excel(imp: dict) -> bool:
    return Path(imp["file_name"]).suffix.lower() in EXCEL_EXT


def _open(imp: dict):
    """控え付きの表ソース。元のファイルは控えにない範囲を読むときだけ開く。"""
    return pipeline.import_source(imp)


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
    """読み込みのあとに開く段。追記ログ列があれば AI整形、なければ内容の確認。"""
    return "ai" if spec is not None and spec.log_stage is not None else "preview"


def _json_error(message: str, status: int = 400, **extra):
    return jsonify({"error": message, **extra}), status


def _payload() -> dict:
    data = request.get_json(silent=True)
    return data if isinstance(data, dict) else {}


def _int(value, default=None):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


# 行番号の範囲「3-4」を広げる上限（見出し行の帯なので広くはならない）。static/tables.js の parseRows と同じ規則
_MAX_ROW_SPAN = 10


def _row_no(value) -> int | None:
    """行番号の入力（全角数字も可）。数字だけのときだけ読む（static/tables.js の parseEnd と同じ規則）。"""
    text = unicodedata.normalize("NFKC", str(value if value is not None else "")).strip()
    return int(text) if text.isascii() and text.isdigit() and int(text) > 0 else None


def _int_list(value) -> list[int]:
    """見出し行の入力（「1,2」「1、2」「1 2」「１，２」「1-2」）。static/tables.js の parseRows と同じ規則で読む。"""
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
    job = jobs.latest_job("table_import", import_id, kind="ai_format")
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


# ---- 1画面のやりとり（段の URL と、段の中身） ------------------------------------------------------

def _urls(import_id: int) -> dict:
    """画面（static/tables.js）が使う URL。panel は末尾の NAME を段の名前に置き換えて使う。

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
            "cancel_url": url_for("tables.cancel_job", import_id=import_id) if import_id else None}


def _running_job(imp: dict) -> dict | None:
    """読み込み中・確定処理中なら、その進み具合。終わっているのに状態が残っていれば直す。"""
    if imp.get("status") not in ("reading", "confirming"):
        return None
    job = jobs.get_job(imp["job_id"]) if imp.get("job_id") else None
    if job is None or job.get("finished"):
        fresh = store.get_import(imp["id"])
        if fresh and fresh.get("status") in ("reading", "confirming"):
            if fresh["status"] == "reading":
                stats = {**(fresh.get("stats") or {}), "error": "読み込みが途中で止まりました。もう一度読み込んでください"}
                store.update_import(imp["id"], status="failed", stats=stats)
            else:
                store.update_import(imp["id"], status="preview")
        return None
    return job


def _panel(html: str, **extra):
    return jsonify({"html": html, **extra})


def _locked(reason: str, **extra):
    """まだ使えない段（灰色で1行だけ理由を出す）。"""
    return jsonify({"html": "", "locked": reason, **extra})


@bp.get("/")
@bp.get("/new")
def new():
    """表の取り込みの画面（1枚）。段の中身はここでは出さず、ファイルを置いたあとに取りに来る。"""
    current_session_id()   # 画面を開いた時点で作業場所（クッキー）を決めておく（同時に置かれても取り違えない）
    return render_template("tables/page.html")


@bp.get("/imports/<int:import_id>/panel/<name>")
def panel(import_id: int, name: str):
    imp = _load_import(import_id)
    handler = _PANELS.get(name)
    if handler is None:
        abort(404)
    return handler(imp)


# ---- ファイルを置く ------------------------------------------------------------------------

@bp.post("/upload")
def upload():
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
    import_id = store.create_import(stored.file_name, stored.file_hash, stored.stored_path, source=source_info,
                                    session_id=current_session_id())
    if source_info.get("kind") == "excel":
        pipeline.import_source(store.get_import(import_id), real=source).sheets()  # シート一覧を控えに入れる
    return jsonify({"import_id": import_id, "file_name": stored.file_name, "urls": _urls(import_id)})


# ---- 画面を離れたので捨てる ------------------------------------------------------------------
# 「その場でダウンロードしない限り、その場ですぐ捨てる」（利用者の指示 2026-09-20）。
# 画面を閉じた・隠したときに static/app.js の ragDiscard がここへ「捨てて」と送ってくる。
# navigator.sendBeacon で届くので中身の型は text/plain（get_json(force=True) で読む）。
# 応答は読めず、やり直しもできないので、いつでも 204 を返す（もう無い番号・ほかの人の番号・
# 処理中のものは core.purge 側で黙って外れる）。

@bp.post("/discard")
def discard():
    """この画面（このブラウザ）の、まだダウンロードしていない取り込みを捨てる。"""
    payload = request.get_json(force=True, silent=True) or {}
    ids = payload.get("import_ids") or []
    sid = current_session_id()
    try:
        if ids:
            purge.discard_table_imports(ids, sid)
        else:
            purge.purge_session(sid, documents=False)   # 帳票取り込み（別のタブ）は巻き込まない
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
    html = render_template(
        "tables/_p_source.html", imp=imp, src=src, error=error, sheets=sheets, sheet=sheet, list_like=list_like,
        encodings=ENCODING_CHOICES, delimiters=DELIMITER_CHOICES)
    note = _source_note(imp, {**src, "sheet": sheet or src.get("sheet")})
    return _panel(html, note=note, error=error, import_id=import_id)


@bp.post("/imports/<int:import_id>/source")
def save_source(import_id: int):
    """読み取り方（シート・文字コード・区切り）。変えるたびに保存する。"""
    imp = _load_import(import_id)
    if _processing(imp):
        return _json_error(BUSY_MESSAGE, 409)
    form = _payload() or request.form
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
        # 見出しが変わるので、この取り込みの列の対応づけは捨てて決め直す
        columns.update(spec_json="", spec_hash="")
    store.update_import(import_id, **columns)
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
        "headers": layout.headers, "warnings": layout.warnings,
        "counts": {KIND_LABELS.get(k, k): v for k, v in (layout.counts or {}).items()},
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
    html = render_template("tables/_p_layout.html", imp=imp, layout=guess, info=_layout_json(guess), rows=rows,
                           width=width, sheet=sheet, data_end_saved=src.get("data_end_row") or "")
    return _panel(html, note=f"{guess.data_start}〜{guess.data_end}行目" if guess.data_end >= guess.data_start else "")


@bp.post("/imports/<int:import_id>/layout/detect")
def layout_detect(import_id: int):
    imp = _load_import(import_id)
    data = _payload()
    try:
        source_obj = _open(imp)
        sheet = _sheet(imp, source_obj)
        guess = _layout(imp, source_obj, sheet, header_rows=_int_list(data.get("header_rows")) or None,
                        data_end=_row_no(data.get("data_end")), use_saved=False)
    except UploadError as exc:
        return _json_error(str(exc))
    return jsonify(_layout_json(guess))


@bp.post("/imports/<int:import_id>/layout")
def save_layout(import_id: int):
    """この範囲で読み込む。設定が無い・見出しが変わったときは列の対応づけへ、そうでなければ読み込みを始める。"""
    imp = _load_import(import_id)
    if _processing(imp):
        return _json_error(BUSY_MESSAGE, 409)
    data = _payload() or request.form
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
    store.update_import(import_id, source=src)
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


# 型の仲間。仲間が違うときだけ「値は…らしい」と知らせる（日付⇔日時、文字⇔長文は読み方が変わらない）
_TYPE_FAMILY = {"date": "日付", "datetime": "日付", "time": "日付", "number": "数値",
                "code": "文字", "string": "文字", "text": "文字", "enum": "文字", "status": "文字"}


def _warnings_of(sugg) -> list[str]:
    """その列で知らせることがあるときだけ出す1行（何も無ければ空）。"""
    out = []
    if sugg.type_error_rate:
        out.append(f"型エラー {round(sugg.type_error_rate * 100, 1)}%")
    if sugg.blank_rate >= 0.5:
        out.append(f"空欄 {int(round(sugg.blank_rate * 100))}%")
    if sugg.inferred_type and _TYPE_FAMILY.get(sugg.inferred_type) != _TYPE_FAMILY.get(sugg.type):
        out.append(f"値は「{TYPE_LABELS.get(sugg.inferred_type, sugg.inferred_type)}」らしい")
    if sugg.md == "omit":
        out.append("空欄だけなので、はじめから使わない設定にしています" if sugg.omit_reason == "blank"
                   else "記録に不要な管理用の列らしいので、はじめから使わない設定にしています")
    return out


def _column_row(sugg) -> dict:
    """画面の1行（使う・見出し・役割と、知らせること）。"""
    return {"index": sugg.index, "header": sugg.header, "use": sugg.md != "omit",
            "role": _screen_role(sugg.role), "examples": list(sugg.examples or [])[:3],
            "warnings": _warnings_of(sugg)}


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
    for sugg in suggestions:
        choice = choices.get(sugg.index)
        if not (bool(choice.get("use")) if choice else sugg.md != "omit"):
            continue
        role = str((choice or {}).get("role") or "")
        if role not in SCREEN_ROLE_KEYS:
            role = _screen_role(sugg.role)
        # 画面でそのままなら候補の役割（人名・数値・区分など）を活かす。選び直したときだけその役割にする
        role = sugg.role if role == _screen_role(sugg.role) else role
        type_ = "text" if role == "log" else sugg.type
        # 「使う」列は必ず md に出す（候補が「出さない」でも、チェックを入れたのだから出す）
        md = sugg.md if sugg.md != "omit" else ("body" if type_ == "text" else "attribute")
        used.append({"index": sugg.index, "header": sugg.header, "key": sugg.key, "display": sugg.display,
                     "type": type_, "role": role, "unit": sugg.unit, "md": md,
                     "fill_down_blank": bool(sugg.fill_down_blank)})
    name = str(payload.get("name") or "").strip()
    header_rows = list(getattr(guess, "header_rows", None) or [1])
    spec = spec_from_suggestions(name, {"table_kind": "list", "header_rows": header_rows}, used)
    errors = validate_spec(spec)
    if sum(1 for u in used if u["role"] == "log") > 1:
        # 黙って最初の列だけを AI整形の対象にしない（2列目は追記ログとして1行につながれて出てしまう）
        errors.insert(0, "AI整形の対象は1列だけにしてください")
    return spec, errors


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
    html = render_template(
        "tables/_p_columns.html", rows=rows, roles=SCREEN_ROLES,
        name=spec.name if spec is not None else _default_table_name(imp),
        save_url=url_for("tables.save_columns", import_id=import_id))
    return _panel(html, note=f"{sum(1 for r in rows if r['use'])}／{len(rows)}列")


@bp.post("/imports/<int:import_id>/columns")
def save_columns(import_id: int):
    imp = _load_import(import_id)
    if _processing(imp):
        return _json_error(BUSY_MESSAGE, 409)
    try:
        guess, suggestions = _suggest(imp)
    except UploadError as exc:
        return _json_error(str(exc))
    spec, errors = _build_spec(_payload(), guess, suggestions)
    if errors:
        return _json_error(errors[0], errors=errors)
    store.save_spec(import_id, spec)
    job_id = pipeline.start_read_job(import_id)
    return jsonify({"ok": True, "next": _after_read_panel(spec), "reset": ["ai", "preview", "done"],
                    "job": _job_info(jobs.get_job(job_id), import_id), "reading": True})


@bp.post("/imports/<int:import_id>/read")
def reread(import_id: int):
    """読み込み直し（失敗したとき・もう一度読むとき）。"""
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None:
        return _json_error("先に列の対応づけを保存してください")
    if _processing(imp):
        # AI整形の実行中・一時停止中に読み込み直すと、読み込みが AI整形の後ろで待ち続ける
        return _json_error(BUSY_MESSAGE, 409)
    job_id = pipeline.start_read_job(import_id)
    return jsonify({"ok": True, "next": _after_read_panel(spec), "reset": ["ai", "preview", "done"],
                    "job": _job_info(jobs.get_job(job_id), import_id), "reading": True})


# 中止できるジョブ（取り込みの job_id に入るもの）と、中止したときに戻す状態
CANCELLABLE_JOBS = {"table_read": "uploaded", "table_render": "preview"}


@bp.post("/imports/<int:import_id>/cancel")
def cancel_job(import_id: int):
    """［中止］: 読み込み・Markdown作成のジョブを止める。"""
    imp = _load_import(import_id)
    job = jobs.get_job(imp["job_id"]) if imp.get("job_id") else None
    if job is None or job["kind"] not in CANCELLABLE_JOBS or job.get("finished"):
        return _json_error("中止できる処理がありません（すでに終わっている可能性があります）")
    jobs.request_cancel(job["id"])
    after = jobs.get_job(job["id"])
    # 取り込みの状態は読み直す（この間にジョブが成功して preview/confirmed を書いていたら、巻き戻してはいけない）
    fresh = store.get_import(import_id)
    if (after and after["status"] == "cancelled" and fresh is not None
            and fresh["status"] in ("reading", "confirming")):
        # 待機中のまま中止されたときは本体が動かないので、ここで取り込みの状態を戻す
        store.update_import(import_id, status=CANCELLABLE_JOBS[job["kind"]])
    return jsonify({"ok": True, "message": "処理を中止しました"})


@bp.post("/imports/<int:import_id>/delete")
def delete_import(import_id: int):
    """間違えて取り込んだファイルを消す（読み込んだ内容・作った Markdown も消える）。"""
    imp = _load_import(import_id)
    if imp["status"] in ("reading", "confirming"):
        return _json_error("処理中の取り込みは削除できません。終わるか中止してから削除してください")
    if _ai_running(import_id):
        return _json_error("AI整形の実行中は削除できません。AI整形を中止してから削除してください")
    if _trial_running(import_id):
        return _json_error(TRIAL_BUSY_MESSAGE)
    preview_job = jobs.latest_job("table_import", import_id, kind="table_preview")
    if preview_job is not None and not preview_job.get("finished"):
        # 下書きの作成中に消すと、開いているファイルが消え残ったり、消したあとに記録の md が書き戻されたりする
        return _json_error("Markdownの下書きを作っている間は削除できません。終わってから削除してください")
    purge.purge_table_import(import_id)
    # ファイル名は出さない: 取引先名や「社外秘」を含みうる名前を画面に残さない（design.md 3.3）
    message = "取り込みを削除しました（元のファイル・読み込んだ内容・作った Markdown を消しました）"
    if purge.purge_incomplete():
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
        return "先に「表の範囲」で読み込んでください"
    return ""


# ---- AI整形（任意。追記ログの列があるときだけ） ----------------------------------------------------

def _ai_job(import_id: int) -> dict | None:
    return jobs.latest_job("table_import", import_id, kind="ai_format")


def _ai_connection_ctx() -> dict:
    """AI接続の状態（段の中の「AI接続」パネル）。"""
    from services import llm

    status = llm.admin_status()
    ready = llm.is_configured()
    external, model = False, ""
    if ready:
        try:
            settings = llm.job_client_settings()
            external = not llm.is_local_endpoint(settings.get("chat_url") or "")
            model = settings.get("model") or ""
        except Exception:
            ready = False
    return {"ai_status": status, "ai_ready": ready, "external": external, "model": model,
            "save_url": url_for("tables.save_ai_connection"), "test_url": url_for("tables.test_ai_connection"),
            "models_url": url_for("tables.refresh_ai_models")}


def _panel_ai(imp: dict):
    import_id = imp["id"]
    job = _running_job(imp)
    if job is not None:
        return _panel("", job=_job_info(job, import_id), reading=True)
    spec = _spec_for(imp)
    if spec is not None and spec.log_stage is None:
        return _locked("追記ログの列（AI整形の対象）がないので、この取り込みでは使いません")
    reason = _not_ready_reason(imp, spec)
    if reason:
        return _locked(reason)
    from aiproc import items as ai_items

    log_key = spec.log_stage.column
    col = spec.column(log_key)
    row_choices = []
    for rec in pipeline.load_rows(import_id):
        text = str((rec.get("values") or {}).get(log_key) or "").strip()
        if text:
            label = rec["key"] + "｜" + " ".join(text.split())[:40]
            row_choices.append((rec["key"], label))
            if len(row_choices) >= 200:
                break
    ai_job = _ai_job(import_id)
    html = render_template(
        "tables/_p_ai.html", imp=imp, log_display=col.display if col else log_key, row_choices=row_choices,
        trial_keys=[k for k, _ in row_choices[:10]], job=ai_job, job_active=bool(ai_job and not ai_job.get("finished")),
        job_url=url_for("tables.api_job", job_id=ai_job["id"]) if ai_job else None,
        counts=ai_items.counts(imp["template_id"], "log", import_id=import_id),
        scopes=SCOPE_LABELS, item_labels=ai_items.STATUS_LABELS, **_ai_connection_ctx())
    return _panel(html, note=f"対象の列: {col.display if col else log_key}",
                  job=_job_info(ai_job, import_id) if ai_job and not ai_job.get("finished") else None)


@bp.post("/imports/<int:import_id>/ai/split-preview")
def ai_split_preview(import_id: int):
    from aiproc import prompts, runner
    from logproc import format_author, format_when, review_notes
    from logproc import render_timeline
    from services import llm

    _load_import(import_id)
    row_key = str(_payload().get("row_key") or "")
    try:
        data = runner.load_rows_for_ai(import_id)
        works = runner.prepare_works(data, ["log"], [row_key])
    except runner.AIJobError as exc:
        return _json_error(str(exc))
    if not works:
        return _json_error("行が見つかりません")
    w = works[0]
    parse = w.parse
    segments = [{
        "id": s.id, "body": s.body, "when": format_when(s.when), "when_estimated": bool(s.when and s.when.estimated),
        "author": format_author(s.author), "author_estimated": bool(s.author and s.author.estimated),
        "marks": list(s.marks or []), "identifiers": list(s.identifiers or []), "quantities": list(s.quantities or []),
        "plans": list(s.plans or []),
    } for s in parse.segments]
    return jsonify({
        "row_key": row_key, "kind": parse.kind, "order": parse.order, "text": parse.text, "segments": segments,
        "route": w.route, "reason": w.reason, "ai_ready": llm.is_configured(), "notes": review_notes(parse),
        "timeline": render_timeline(parse, w.entity_label, glossary=(w.stage.glossary or None) if w.stage else None),
        "sent_text": prompts.messages_text(w.messages) if w.messages else "",
    })


@bp.post("/imports/<int:import_id>/ai/trial")
def ai_trial(import_id: int):
    from aiproc import runner
    from services import llm

    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or spec.log_stage is None:
        return _json_error("AI整形の対象の列がありません")
    if imp["status"] == "confirmed":
        # 試し実行の結果は下書きに入るが、zip は確定したときの md を渡す。食い違わないよう断る
        return _json_error("確定後は試し実行できません（結果が確定した Markdown に入らないため）。"
                           "試すときは列の対応づけからもう一度読み込んでください")
    payload = _payload()
    row_key = str(payload.get("row_key") or "")
    try:
        settings = llm.job_client_settings()
        if not llm.is_local_endpoint(settings.get("chat_url") or "") and not payload.get("confirm_external"):
            return _json_error("対応内容が外部のAIサービスに送信されます。確認のチェックを入れてください")
        with _TRIALS_LOCK:
            _TRIALS[import_id] += 1
        try:
            data = runner.load_rows_for_ai(import_id)
            result = runner.trial_row(import_id, row_key, stage_ids=["log"], data=data)
        finally:
            with _TRIALS_LOCK:
                _TRIALS[import_id] -= 1
                if _TRIALS[import_id] <= 0:
                    del _TRIALS[import_id]
    except llm.LLMNotConfigured:
        return _json_error("AI接続が設定されていません。この段の「AI接続」で設定してください")
    except llm.LLMCallError as exc:
        return _json_error(str(exc))
    except runner.AIJobError as exc:
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
        "stats": runner.trial_stats([result]),
    })


@bp.post("/imports/<int:import_id>/ai/estimate")
def ai_estimate(import_id: int):
    from aiproc import estimate as ai_estimate_mod
    from aiproc import runner
    from services import llm

    _load_import(import_id)
    data = _payload()
    scope = data.get("scope") if data.get("scope") in SCOPE_LABELS else "pending"
    try:
        settings = llm.job_client_settings()
        result = ai_estimate_mod.estimate(import_id, [s for s in data.get("trials") or [] if isinstance(s, dict)],
                                          scope=scope, concurrency=_int(data.get("concurrency")) or None,
                                          settings=settings, stage_ids=["log"])
    except llm.LLMNotConfigured:
        return _json_error("AI接続が設定されていません")
    except runner.AIJobError as exc:
        return _json_error(str(exc))
    return jsonify(result)


@bp.post("/imports/<int:import_id>/ai/run")
def ai_run(import_id: int):
    from aiproc import runner
    from services import llm

    imp = _load_import(import_id)
    data = _payload()
    if imp["status"] != "preview":
        return _json_error("表の読み込みが終わってから実行してください")
    job = _ai_job(import_id)
    if job and not job.get("finished"):
        return _json_error("AI整形はすでに実行中です")
    scope = data.get("scope") if data.get("scope") in SCOPE_LABELS else "pending"
    concurrency = min(16, max(1, _int(data.get("concurrency"), 1) or 1))
    try:
        settings = llm.job_client_settings()
        if not llm.is_local_endpoint(settings.get("chat_url") or "") and not data.get("confirm_external"):
            return _json_error("対応内容が外部のAIサービスに送信されます。確認のチェックを入れてください")
        job_id = runner.start_ai_job(import_id, scope=scope, concurrency=concurrency, stage_ids=["log"])
    except llm.LLMNotConfigured:
        return _json_error("AI接続が設定されていません")
    return jsonify({"ok": True, "job_id": job_id, "job_url": url_for("tables.api_job", job_id=job_id)})


@bp.post("/imports/<int:import_id>/ai/<action>")
def ai_control(import_id: int, action: str):
    _load_import(import_id)
    handlers = {"pause": jobs.request_pause, "resume": jobs.request_resume, "cancel": jobs.request_cancel}
    if action not in handlers:
        abort(404)
    job = _ai_job(import_id)
    ok = bool(job) and handlers[action](job["id"])
    if not ok:
        return _json_error("AI整形のジョブを操作できませんでした（すでに終わっている可能性があります）")
    return jsonify({"ok": True})


# ---- AI接続（AI整形の段の中。設定画面は無い） ---------------------------------------------------

@bp.post("/ai-connection")
def save_ai_connection():
    from services import llm

    data = _payload()
    models = [str(m).strip() for m in (data.get("models") or []) if str(m).strip()]
    models += [m.strip() for m in str(data.get("add_models") or "").splitlines() if m.strip()]
    try:
        llm.save_admin({
            "models": list(dict.fromkeys(models)),
            "default": str(data.get("default") or ""),
            "chat_url": str(data.get("chat_url") or ""),
            "models_url": str(data.get("models_url") or ""),
            "api_key": str(data.get("api_key") or ""),
            "api_key_clear": bool(data.get("api_key_clear")),
        })
    except ValueError as exc:
        return _json_error(str(exc))
    except OSError as exc:
        return _json_error(f"設定ファイルに書き込めませんでした（{exc.strerror or exc.__class__.__name__}）。"
                           "少し待ってから、もう一度保存してください")
    return jsonify({"ok": True, "message": "AI接続の設定を保存しました"})


@bp.post("/ai-connection/models")
def refresh_ai_models():
    """APIからモデル一覧を取得する。"""
    from services import llm

    try:
        catalog = llm.model_catalog(refresh=True)
    except Exception as exc:
        return _json_error(f"モデル一覧を取得できませんでした: {llm.friendly_error(exc)}")
    return jsonify({"ok": True, "models": catalog, "message": f"APIからモデル一覧を取得しました（{len(catalog)}件）"})


@bp.post("/ai-connection/test")
def test_ai_connection():
    """接続テスト: モデル一覧の取得と、1回の短いチャット。"""
    from services import llm

    if not llm.is_configured():
        return jsonify({"ok": False, "steps": [{"name": "設定", "ok": False,
                                                "detail": "APIキーまたは接続先が未設定です。"}]})
    steps = []
    started = time.monotonic()
    try:
        names = llm.fetch_api_models(refresh=True)
        steps.append({"name": "モデル一覧の取得", "ok": True,
                      "detail": f"{len(names)}件のモデルが見つかりました（{_ms(started)}ミリ秒）"})
    except Exception as exc:
        steps.append({"name": "モデル一覧の取得", "ok": False, "detail": llm.friendly_error(exc)})
    model = llm.current_model()
    started = time.monotonic()
    try:
        # AI整形のジョブと同じ呼び出し口（再試行なし・明示のタイムアウト）で確かめる
        result = llm.chat_raw(llm.job_client_settings(),
                              [{"role": "user", "content": "接続テストです。「OK」とだけ返してください。"}], max_tokens=20)
        reply = (result.text or "").strip()
        steps.append({"name": f"チャット（{model}）", "ok": True,
                      "detail": f"応答あり: {reply[:40] or '（本文なし）'}（{_ms(started)}ミリ秒）"})
    except Exception as exc:
        steps.append({"name": f"チャット（{model}）", "ok": False, "detail": llm.friendly_error(exc)})
    # チャットが通れば AI 整形は使える（モデル一覧APIが無い互換サーバもある）
    return jsonify({"ok": steps[-1]["ok"], "steps": steps})


def _ms(started: float) -> int:
    return int((time.monotonic() - started) * 1000)


# ---- 内容の確認 ---------------------------------------------------------------------------

def _preview_job(import_id: int, imp: dict, spec, retry: bool = False) -> dict:
    """確認の段の md を作るジョブ。同じ入力のジョブがあればそれを見せ、なければ始める。

    retry: 同じ入力で失敗したジョブを作り直す（［もう一度作る］。ファイルの一時的なロックなどで失敗したとき）。
    """
    signature = pipeline.preview_signature(import_id, imp, spec)
    job = jobs.latest_job("table_import", import_id, kind="table_preview")
    if job is not None and (job.get("params") or {}).get("signature") == signature:
        if not job.get("finished") or (job["status"] == "failed" and not retry):
            return job  # 動いている途中、または同じ入力で失敗したまま（開き直すたびに作り直さない）
    return jobs.get_job(pipeline.start_preview_job(import_id, signature))


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
    files = pipeline.ready_preview_files(import_id, imp, spec)
    if files is None:
        # 件数分の md を作るのに時間がかかるので、ジョブにして同じ画面に進み具合を出す
        draft = _preview_job(import_id, imp, spec)
        if draft.get("finished") and draft["status"] != "done":
            return _panel(render_template("tables/_p_preview_failed.html",
                                          error=draft.get("message") or "Markdownの下書きを作れませんでした"),
                          failed=True)
        return _panel("", job=_job_info(draft, import_id), building=True)
    stats = imp.get("stats") or {}
    issues = pipeline.load_issues(import_id)
    page = max(1, page)
    data_rows, total = pipeline.load_rows_page(import_id, (page - 1) * DATA_PAGE, DATA_PAGE)
    total_pages = max(1, (total + DATA_PAGE - 1) // DATA_PAGE)
    if page > total_pages:
        page = total_pages
        data_rows, _total = pipeline.load_rows_page(import_id, (page - 1) * DATA_PAGE, DATA_PAGE)
    blocking = has_blocking(issues)
    html = render_template(
        "tables/_p_preview.html", imp=imp, spec=spec, stats=stats, issues=issues[:ISSUES_SHOWN],
        issue_total=len(issues), counts=count_levels(issues), blocking=blocking, files=files, data_rows=data_rows,
        page=page, total_pages=total_pages, columns=[(c.key, _display_with_unit(c)) for c in spec.columns],
        confirmed=imp["status"] == "confirmed", delete_note=DELETE_ON_DOWNLOAD_NOTE)
    return _panel(html, note=f"{stats.get('records') or 0}件・{len(files)}ファイル", blocking=blocking,
                  confirmed=imp["status"] == "confirmed")


@bp.post("/imports/<int:import_id>/preview")
def start_preview(import_id: int):
    """確認の段の Markdown の下書きを作り直す（失敗したときの［もう一度作る］）。"""
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        return _json_error("表の読み込みが終わってから確認してください")
    job = _preview_job(import_id, imp, spec, retry=True)
    return jsonify({"ok": True, "job": _job_info(job, import_id), "building": True})


@bp.get("/imports/<int:import_id>/preview/file")
def preview_file(import_id: int):
    _load_import(import_id)
    text = pipeline.md_text(pipeline.import_files(import_id)["preview"], request.args.get("name", ""))
    if text is None:
        return _json_error("ファイルが見つかりません", 404)
    return jsonify({"name": request.args.get("name"), "text": text})


@bp.get("/imports/<int:import_id>/issues.csv")
def issues_csv(import_id: int):
    imp = _load_import(import_id)
    data = outputs.issues_csv(pipeline.load_issues(import_id))
    name = f"{Path(imp['file_name']).stem}_問題一覧.csv"
    return set_download_name(
        send_file(io.BytesIO(data), mimetype="text/csv", as_attachment=True, download_name=name), name, "issues")


@bp.post("/imports/<int:import_id>/confirm")
def confirm(import_id: int):
    """確定して Markdown を作る（できたら同じ画面にダウンロードのボタンを出す）。"""
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        return _json_error("この取り込みはまだ確定できる状態ではありません")
    if _ai_running(import_id):
        return _json_error("AI整形の実行中は確定できません。終わるか中止してから確定してください")
    if has_blocking(pipeline.load_issues(import_id)):
        return _json_error("エラーが残っているため確定できません。問題一覧を確認して、範囲や列の対応づけを直してください")
    job_id = pipeline.start_render_job(import_id)
    return jsonify({"ok": True, "next": "done", "job": _job_info(jobs.get_job(job_id), import_id), "building": True})


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
    files = [{"name": p.name, "size": p.stat().st_size} for p in pipeline.md_paths(import_id)]
    html = render_template("tables/_p_done.html", imp=imp, files=files, stats=imp.get("stats") or {},
                           delete_note=DELETE_ON_DOWNLOAD_NOTE, delete_confirm=DELETE_ON_DOWNLOAD_CONFIRM)
    return _panel(html, note=f"{len(files)}ファイル")


@bp.get("/imports/<int:import_id>/download.zip")
def download_zip(import_id: int):
    """zip を渡し、渡し終えた取り込みのデータを消す（design.md 3.3）。

    zip は全体をメモリに作ってから消すので、消す処理で中身が欠けることはない。作れなかったときは何も消さない。
    """
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if imp["status"] != "confirmed" or spec is None or not pipeline.md_paths(import_id):
        flash("先に [確定してMarkdownを作成] を押してください", "error")
        return redirect(url_for("tables.new"))
    if _ai_running(import_id):
        flash("AI整形の実行中はダウンロードできません。終わるか中止してからダウンロードしてください", "error")
        return redirect(url_for("tables.new"))
    if _trial_running(import_id):
        flash(TRIAL_BUSY_MESSAGE, "error")
        return redirect(url_for("tables.new"))
    try:
        data = pipeline.build_download(import_id, imp, spec)
    except FileNotFoundError:
        # 同時に押した別のダウンロードが渡し終えてデータを消した（ダブルクリックなど）。500 ではなく「ダウンロード済み」。
        return _handout_lost(import_id, "ダウンロード")
    name = f"{_download_base(imp, spec)}.zip"
    response = send_file(io.BytesIO(data), mimetype="application/zip", as_attachment=True, download_name=name,
                         conditional=False)   # Range でも全体を返す（一部だけ渡して消すことが無いように）
    set_download_name(response, name, "records")
    return purge.purge_after_send(response, purge.purge_table_import, import_id)


def _handout_lost(import_id: int, what: str):
    """渡す途中で md が消えたとき。取り込みが残っていて確定済みでなければ、別のタブで読み込み直しが始まった
    （「ダウンロード済み」の 404 は事実と違うので、画面に戻して知らせる）。消えていれば渡し終えた・削除した。"""
    now = store.get_import(import_id)
    if now is None or now["status"] == "confirmed":
        abort(404)
    flash(f"読み込み直しが始まったため、{what}を取りやめました。確定し直してからもう一度押してください", "error")
    return redirect(url_for("tables.new"))


@bp.get("/imports/<int:import_id>/normalized.csv")
def download_csv(import_id: int):
    imp = _load_import(import_id)
    spec = _spec_for(imp)
    if spec is None or imp["status"] not in ("preview", "confirmed"):
        abort(404)
    data = outputs.normalized_csv(spec, pipeline.load_rows(import_id))
    name = f"{_download_base(imp, spec)}_正規化データ.csv"
    return set_download_name(
        send_file(io.BytesIO(data), mimetype="text/csv", as_attachment=True, download_name=name), name, "records")


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
    imp = store.get_import(job["ref_id"])
    return imp is not None and owns(imp)


def api_job(job_id: int):
    job = jobs.get_job(job_id)
    if job is None or not _job_visible(job):
        return _json_error("ジョブが見つかりません", 404)
    return jsonify({k: job.get(k) for k in ("id", "kind", "status", "status_label", "progress", "message", "result",
                                            "finished")})


@bp.record_once
def _register_api(setup_state) -> None:
    # /api/jobs は /tables の外に置く（design.md 2.4）
    setup_state.app.add_url_rule("/api/jobs/<int:job_id>", "tables.api_job", api_job)
