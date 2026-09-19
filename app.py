import os
import secrets
import sys
from pathlib import Path
from urllib.parse import urlsplit

from flask import Flask, abort, jsonify, render_template, request

from config import BASE_DIR, Config
from models import database
from services import llm
from views import form_types, forms, home, is_cross_site_write, settings, tables

_SECRET_FILE = BASE_DIR / ".flask_secret"


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
    "  通常の起動: python app.py\n"
    "  詳しいエラーを見たいとき: app.py の DEBUG を True にして python app.py\n"
)
_NO_REMOTE_MSG = (
    "\n[app] このアプリにはログイン機能が無いため、このPC以外から届くアドレスでは起動しません。\n"
    "  いまの設定: HOST = {host}\n"
    "  他のPCから開くと、APIキーの設定を含むすべての画面を誰でも操作できてしまいます。\n"
)


def _is_loopback(host) -> bool:
    # 空文字は 0.0.0.0 と同じで全てのネットワークから届くので loopback に含めない
    return str(host or "").strip().lower() in ("127.0.0.1", "localhost", "::1")


def _refuse_debugger() -> None:
    if str(os.environ.get("FLASK_DEBUG") or "").strip().lower() in _TRUE:
        raise SystemExit(_NO_DEBUG_MSG)
    if any(a in _DEBUG_ARGS for a in sys.argv[1:]):
        raise SystemExit(_NO_DEBUG_MSG)


def create_app(overrides: dict | None = None) -> Flask:
    _refuse_debugger()
    app = Flask(__name__)
    app.config.from_object(Config)
    app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax")
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
    database.init_app(app)

    for module in (home, forms, form_types, tables, settings):
        app.register_blueprint(module.bp)

    @app.before_request
    def _refuse_other_host():
        # このPCからしか開けないこと（＝ログインが要らない前提）を Host でも確かめる。
        # 127.0.0.1 で待ち受けても、攻撃者のドメインが 127.0.0.1 を指していれば
        # そのページと同一オリジンになり、画面の中身を読み取られてしまう（DNSリバインディング）。
        try:
            parts = urlsplit(f"//{request.host}")
            host, port = parts.hostname, parts.port   # ポート番号・[] を外したホスト名
        except ValueError:
            host, port = None, None
        if not _is_loopback(host):
            # 汎用の 400 画面（ホームへ = 同じアドレスの / ）では、また断られるだけで開き方が分からない。
            # 開けるアドレスを示す（ポートは送られてきた Host のもの。読めなければ起動時の PORT）
            home_url = f"http://127.0.0.1:{port or PORT}/"
            text = f"このアドレスでは開けません。このアプリは {home_url} で開いてください"
            if request.accept_mimetypes.best == "application/json" or request.is_json:
                return jsonify(error=text), 400
            return render_template("errors/400.html", host_refused=True, home_url=home_url), 400

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
        return response

    @app.errorhandler(403)
    def _forbidden(_exc):
        # 画面からの JSON 送信（static/app.js の postJson）には JSON で返す。トーストに理由が出る。
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="ほかのサイトのページから送られてきた操作に見えたため受け付けませんでした。"
                                 "ホームからやり直してください"), 403
        return render_template("errors/403.html"), 403

    @app.errorhandler(404)
    def _not_found(_exc):
        # 403 と同じく、画面からの JSON 送信（static/app.js の postJson）には JSON で返す。
        # ダウンロード済みでデータが消えたあとに、開いたままの画面が途中保存・プレビューを送ると
        # ここに来る。HTML を返すとトーストに「通信に失敗しました」としか出ず、理由が伝わらない。
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="このデータはこのPCに残っていません（ダウンロード済みか、削除されています）。"
                                 "ホームからやり直してください"), 404
        return render_template("errors/404.html"), 404

    @app.errorhandler(413)
    def _too_large(_exc):
        # MAX_CONTENT_LENGTH を超えた送信。Werkzeug の英語の画面を出さない（design.md 2: エラー画面は日本語）
        limit = app.config.get("MAX_CONTENT_LENGTH") or 0
        text = f"{limit / 1024 / 1024:.0f}MB" if limit >= 1024 * 1024 else f"{limit / 1024:.0f}KB"
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error=f"送った内容が大きすぎます（1回に合計 {text} まで）。分けて送ってください"), 413
        return render_template("errors/413.html", limit=text), 413

    @app.errorhandler(400)
    @app.errorhandler(405)
    def _bad_request(exc):
        # 送信専用の URL をアドレス欄から開いた（405）など。Werkzeug の英語の画面を出さない（design.md 2）
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="この操作は受け付けられませんでした。ホームから開き直してください"), exc.code
        return render_template("errors/400.html"), exc.code

    @app.errorhandler(500)
    def _server_error(_exc):
        return render_template("errors/500.html"), 500

    _recover_jobs(app)
    _cleanup_leftovers(app)

    @app.context_processor
    def inject_model_picker():
        # ヘッダーのモデル選択（全画面共通）
        return {"model_picker": {"current": llm.current_model(), "models": llm.available(),
                                 "llm_ready": llm.is_configured()}}

    return app


def _cleanup_leftovers(app: Flask) -> None:
    """起動時：DB から参照されていない残骸（アップロードファイル・取り込みのフォルダ）を消す。

    データを残さない方針（design.md 3.3）でも、削除の途中で落ちたときやファイルを掴まれていたときに
    残ることがあるので、ここで片付ける。作業中のものは DB に行があるので消さない。
    """
    from core import purge
    from core.files import remove_orphan_import_dirs, remove_orphan_uploads
    from models.database import get_db

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
            purge.sweep_orphan_ai(db)
        if removed:
            print(f"[app] 参照されていないアップロードファイルを {removed} 件片付けました")
        if removed_dirs:
            print(f"[app] 参照されていない取り込みのフォルダを {removed_dirs} 件片付けました")
    except Exception as exc:  # 起動は止めない
        print(f"[app] 残ったファイルの整理に失敗しました（{exc.__class__.__name__}: {exc}）")


def _recover_jobs(app: Flask) -> None:
    """前回の終了時に動いていたジョブを「中断」にする。"""
    from core.jobs import recover_interrupted

    try:
        with app.app_context():
            recover_interrupted()
    except Exception as exc:  # 起動は止めない
        print(f"[app] 中断したジョブの整理に失敗しました（{exc.__class__.__name__}: {exc}）")


# --- 待ち受け先 -------------------------------------------------------------
# ログイン機能が無い（1人で使う）ため、このPCからだけ開ける 127.0.0.1 で起動する。
HOST = "127.0.0.1"
PORT = 5000
THREADS = 8
# エラー画面に詳細を出すか。通常は False のまま。
DEBUG = False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(f"不明な引数: {sys.argv[1]}（起動は引数なしの python app.py）")
    if not _is_loopback(HOST):
        raise SystemExit(_NO_REMOTE_MSG.format(host=HOST))

    application = create_app()
    if DEBUG:
        application.run(host=HOST, port=PORT, debug=True, use_reloader=False)
    else:
        try:
            from waitress import serve
        except ImportError:
            print("[app] waitress が無いため Flask の開発サーバで起動します（pip install waitress を推奨）")
            application.run(host=HOST, port=PORT, debug=False)
        else:
            print(f"[app] http://localhost:{PORT} で起動しました")
            serve(application, host=HOST, port=PORT, threads=THREADS)
