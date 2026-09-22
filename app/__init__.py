"""Flask アプリ（create_app）と設定（env ファイルを読み込む。旧 config.py）。起動は run.py。

アプリ本体はこの app/ フォルダに全部ある: core.py（土台。DBもここ）・forms.py（帳票）・tables.py（一覧表）・
ai.py（AI整形・AI接続・経過の記録）・views.py（画面）と templates/・static/。"""
import os
import secrets
import socket
import sys
import threading
from functools import lru_cache
from pathlib import Path
from urllib.parse import urlsplit

from dotenv import load_dotenv
from flask import Blueprint, Flask, abort, jsonify, redirect, render_template, request, url_for

from app import core
from app.views import SESSION_LIFETIME, form_types_bp, forms_bp, is_cross_site_write, tables_bp



# ====================================================================================================
# 設定（元 config.py）
# ====================================================================================================

BASE_DIR = Path(__file__).resolve().parent.parent   # プロジェクトの根（env・instance/・data/・uploads/ の置き場所）

# 接続設定は aiagent_minimal_rag_tougou と同じく、ドット無しの "env" ファイル（export KEY="..." 形式も可）から読む
load_dotenv(BASE_DIR / "env")


def _split(value: str) -> list[str]:
    return [v.strip() for v in value.replace(",", ";").split(";") if v.strip()]


def _optional_float(name: str) -> float | None:
    raw = os.environ.get(name, "").strip()
    return float(raw) if raw else None


def _optional_int(name: str) -> int | None:
    raw = os.environ.get(name, "").strip()
    return int(raw) if raw else None


class Config:
    # セッション署名鍵は create_app が FLASK_SECRET_KEY または .flask_secret ファイルから設定する
    SECRET_KEY = None
    DATABASE = Path(os.environ.get("DATABASE", BASE_DIR / "instance" / "app.db"))
    UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", BASE_DIR / "uploads"))
    # 画面から保存する設定（model_settings.yaml）の置き場所
    DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data"))
    # 一覧表の行データ・状態ファイルの置き場所（DATA_DIR/tables。create_app で DATA_DIR に合わせて決め直す）
    TABLES_DIR = Path(os.environ.get("TABLES_DIR", DATA_DIR / "tables"))
    # 1回のリクエストの上限（一覧表の大きいCSV/Excelを想定）
    MAX_CONTENT_LENGTH = 200 * 1024 * 1024
    # 帳票（1ファイル＝1件）
    ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
    # 一覧表（Excel/CSV・1行＝1件）
    TABLE_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".tsv", ".txt"}
    TABLE_MAX_UPLOAD_BYTES = int(os.environ.get("TABLE_MAX_UPLOAD_BYTES", str(200 * 1024 * 1024)))
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
    OPENAI_MAX_TOKENS = _optional_int("OPENAI_MAX_TOKENS")
    LLM_RATE_LIMIT_RETRIES = int(os.environ.get("LLM_RATE_LIMIT_RETRIES", "3"))
    LLM_RATE_LIMIT_MAX_WAIT = float(os.environ.get("LLM_RATE_LIMIT_MAX_WAIT", "20"))


# ====================================================================================================
# アプリ（元 app.py）
# ====================================================================================================

_SECRET_FILE = BASE_DIR / ".flask_secret"

# --- 「/」は帳票取り込みへ ------------------------------------------------------------
# ホーム画面は無い。画面は「帳票取り込み」「表の取り込み」「帳票登録」の3つだけで、
# 作業中の一覧もダウンロード待ちの一覧も持たない（利用者の指示 2026-09-20）。
# ここに残るのは「/」を開いたときの転送だけ（古いブックマーク・開いたままの画面の戻り先も拾う）。
home_bp = Blueprint("home", __name__)


@home_bp.get("/", endpoint="index")
def to_forms():
    return redirect(url_for("forms.new"))


@home_bp.get("/guide", endpoint="guide")
def guide():
    """解説（帳票の Markdown がどう作られるか）。上部タブの4つ目。読むだけの画面で、データには触れない
    （利用者の求め 2026-09-22）。中身は templates/guide.html。"""
    from flask import render_template

    return render_template("guide.html")


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
    "  通常の起動: python run.py\n"
    "  詳しいエラーを見たいとき: app/__init__.py の DEBUG を True にして python run.py\n"
)


def _refuse_debugger() -> None:
    if str(os.environ.get("FLASK_DEBUG") or "").strip().lower() in _TRUE:
        raise SystemExit(_NO_DEBUG_MSG)
    if any(a in _DEBUG_ARGS for a in sys.argv[1:]):
        raise SystemExit(_NO_DEBUG_MSG)


# --- 待ち受け先と、受け付ける宛先の名前 -------------------------------------------------
# サーバ（JupyterLab のターミナルなど）で起動し、社内LANの他のPCから数人で開いて使う（design.md 0）。
# 既定はこのサーバの中からだけ開ける 127.0.0.1。LAN に出すときは起動時に環境変数で渡す:
#   HOST=0.0.0.0 PORT=5000 python run.py
_ALL_ADDRESSES = ("", "0.0.0.0", "::")


def _env_port(default: int = 5000) -> int:
    raw = (os.environ.get("PORT") or "").strip()
    if not raw:
        return default
    try:
        port = int(raw)
    except ValueError:
        raise SystemExit(f"\n[app] PORT が数字ではありません: {raw!r}\n  例: PORT=5000 python run.py\n") from None
    if not 1 <= port <= 65535:
        raise SystemExit(f"\n[app] PORT が範囲外です: {port}（1〜65535）\n")
    return port


HOST = (os.environ.get("HOST") or "127.0.0.1").strip() or "127.0.0.1"
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
        f"    いま取り込んでいる帳票・一覧表の中身も、開かれれば見えます。\n"
        f"    信頼できる社内LANの中だけで使ってください。\n"
        f"  ・ダウンロードするとサーバーからデータが消えます。ダウンロードしていないものも\n"
        f"    しばらくすると捨てます。要るものはその場でダウンロードしてください。\n"
        f"  ・止めるとき: Ctrl+C（nohup で動かしているときは kill <PID>）\n"
        f"{line}\n"
    )


def create_app(overrides: dict | None = None) -> Flask:
    _refuse_debugger()
    app = Flask(__name__)
    app.config.from_object(Config)
    # クッキーは約1年もたせる（views.SESSION_LIFETIME。ヘッダーの「AI接続」の設定をブラウザごとに
    # 「ずっと保持」するため。利用者の指示 2026-09-21）。views.current_session_id が session.permanent を立てる
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
    core.init_app(app)
    if not app.config.get("ALLOWED_HOSTS"):
        app.config["ALLOWED_HOSTS"] = allowed_hosts()
    else:   # 設定から渡されたものも、Host と同じ形（ポート無し・小文字）に揃え、ループバックは必ず足す
        given = {h if h == "*" else (_hostname_only(h) or str(h).lower()) for h in app.config["ALLOWED_HOSTS"]}
        app.config["ALLOWED_HOSTS"] = {"*"} if "*" in given else given | {"127.0.0.1", "localhost", "::1"}

    app.register_blueprint(home_bp)
    for bp in (forms_bp, form_types_bp, tables_bp):
        app.register_blueprint(bp)

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
            return render_template("error.html", code=400, host_refused=True, home_url=home_url), 400

    @app.before_request
    def _refuse_cross_site_write():
        # 他のサイトのページから、利用者のブラウザ経由で書き込ませない（views.is_cross_site_write のとおり）。
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
        return render_template("error.html", code=403), 403

    @app.errorhandler(404)
    def _not_found(_exc):
        # 403 と同じく、画面からの JSON 送信（static/app.js の postJson）には JSON で返す。
        # ダウンロード済みでデータが消えたあとに、開いたままの画面が途中保存・プレビューを送ると
        # ここに来る。HTML を返すとトーストに「通信に失敗しました」としか出ず、理由が伝わらない。
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="このデータはサーバーに残っていません（ダウンロード済みか、削除されています）。"
                                 "取り込みの画面からやり直してください"), 404
        return render_template("error.html", code=404), 404

    @app.errorhandler(413)
    def _too_large(_exc):
        # MAX_CONTENT_LENGTH を超えた送信。Werkzeug の英語の画面を出さない（design.md 2: エラー画面は日本語）
        limit = app.config.get("MAX_CONTENT_LENGTH") or 0
        text = f"{limit / 1024 / 1024:.0f}MB" if limit >= 1024 * 1024 else f"{limit / 1024:.0f}KB"
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error=f"送った内容が大きすぎます（1回に合計 {text} まで）。分けて送ってください"), 413
        return render_template("error.html", code=413, limit=text), 413

    @app.errorhandler(400)
    @app.errorhandler(405)
    def _bad_request(exc):
        # 送信専用の URL をアドレス欄から開いた（405）など。Werkzeug の英語の画面を出さない（design.md 2）
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="この操作は受け付けられませんでした。取り込みの画面から開き直してください"), exc.code
        return render_template("error.html", code=400), exc.code

    @app.errorhandler(500)
    def _server_error(_exc):
        return render_template("error.html", code=500), 500

    _recover_jobs(app)
    _purge_pending(app)
    _cleanup_leftovers(app)
    _start_sweeper(app)
    # ヘッダーは「帳票取り込み／表の取り込み／帳票登録」の3つだけ。使うAIモデルの選択は
    # 「表の取り込み」画面の AI整形の段の中（AI接続）に移したので、共通の値は渡さない。
    return app


# 途中で放り出されたものを捨てる間隔（design.md 3.3）。起動時に全部捨て、動いている間は
# core.purge.STALE_HOURS より古いものをこの間隔で捨てる（数人で使うサーバーに置きっぱなしになるため）。
# 捨てるまでの時間（STALE_HOURS）を短くしたときは、見回りもそれに合わせて短くする。
# 見回りの間隔だけ延びると「2時間で捨てます」と言いながら3時間残ることになるので、上限は10分にする。
SWEEP_INTERVAL_SECONDS = 10 * 60


def _sweep_interval(stale_hours: float) -> int:
    return max(60, min(SWEEP_INTERVAL_SECONDS, int(stale_hours * 3600 / 2) or SWEEP_INTERVAL_SECONDS))


def _purge_pending(app: Flask) -> None:
    """起動時：ダウンロードしていない帳票・一覧表をすべて捨てる（作業中の一覧を持たないため）。"""
    if app.config.get("TESTING"):
        return
    from app import core

    try:
        with app.app_context():
            forms_removed, tables_removed = core.purge_all_pending()
            samples_removed = core.purge_old_sample_files()
        if forms_removed or tables_removed:
            print(f"[app] 途中だった取り込みを捨てました（帳票 {forms_removed} 件・一覧表 {tables_removed} 件）")
        if samples_removed:
            print(f"[app] 前の版が置いた見本のExcelを捨てました（{samples_removed} 件・見本はもう置きません）")
    except Exception as exc:  # 起動は止めない
        print(f"[app] 途中だった取り込みの片付けに失敗しました（{exc.__class__.__name__}: {exc}）")


def _start_sweeper(app: Flask) -> None:
    """動いている間：しばらくさわられていない帳票・一覧表を捨て続ける（daemon スレッド）。"""
    if app.config.get("TESTING"):
        return
    from app import core

    interval = _sweep_interval(getattr(core, "STALE_HOURS", 24))

    def loop() -> None:
        while True:
            _stop.wait(interval)
            try:
                with app.app_context():
                    forms_removed, tables_removed = core.sweep_stale(core.STALE_HOURS)
                if forms_removed or tables_removed:
                    print(f"[app] {core.STALE_HOURS}時間さわられていない取り込みを捨てました"
                          f"（帳票 {forms_removed} 件・一覧表 {tables_removed} 件）")
            except Exception as exc:   # 次の回でやり直す
                print(f"[app] 古い取り込みの片付けに失敗しました（{exc.__class__.__name__}: {exc}）")

    _stop = threading.Event()
    threading.Thread(target=loop, name="purge-sweeper", daemon=True).start()


def _cleanup_leftovers(app: Flask) -> None:
    """起動時：DB から参照されていない残骸（アップロードファイル・取り込みのフォルダ）を消す。

    データを残さない方針（design.md 3.3）でも、削除の途中で落ちたときやファイルを掴まれていたときに
    残ることがあるので、ここで片付ける。作業中のものは DB に行があるので消さない。
    """
    from app import core
    from app.core import remove_orphan_import_dirs, remove_orphan_uploads
    from app.core import get_db

    try:
        with app.app_context():
            db = get_db()
            known: set[str] = set()
            for (table,) in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall():
                columns = {row[1] for row in db.execute(f'PRAGMA table_info("{table}")')}
                if "stored_path" in columns:
                    known |= {row[0] for row in db.execute(f'SELECT stored_path FROM "{table}"') if row[0]}
            removed = remove_orphan_uploads(app.config["UPLOAD_DIR"], known)
            import_ids = {row[0] for row in db.execute("SELECT id FROM table_imports")}
            removed_dirs = remove_orphan_import_dirs(app.config["TABLES_DIR"], import_ids)
            # もう無い取り込みの AI整形の結果と、どこからも使われない AI の応答も消す（データを残さない）。
            # ai_items が先に消えていて（取り込み設定の削除など）応答だけ残っていることもあるので、毎回見る。
            core.sweep_orphan_ai(db)
            # 前の版で空になった表に残っている番号の続き（何件取り込んだか）も忘れる（design.md 3.3「履歴は持たない」）
            if core.forget_id_counters(db):
                db.execute("PRAGMA wal_checkpoint(TRUNCATE)")   # 置き換える前の値を app.db-wal に残さない
        if removed:
            print(f"[app] 参照されていないアップロードファイルを {removed} 件片付けました")
        if removed_dirs:
            print(f"[app] 参照されていない取り込みのフォルダを {removed_dirs} 件片付けました")
    except Exception as exc:  # 起動は止めない
        print(f"[app] 残ったファイルの整理に失敗しました（{exc.__class__.__name__}: {exc}）")


def _recover_jobs(app: Flask) -> None:
    """前回の終了時に動いていたジョブを「中断」にする。"""
    from app.core import recover_interrupted

    try:
        with app.app_context():
            recover_interrupted()
    except Exception as exc:  # 起動は止めない
        print(f"[app] 中断したジョブの整理に失敗しました（{exc.__class__.__name__}: {exc}）")


# --- 起動の設定 -------------------------------------------------------------
# 待ち受け先（HOST・PORT）はファイルの上のほうで環境変数から読む。
THREADS = 8
# waitress が応答の本文を先読みして溜める上限（既定は 16MB）。
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
