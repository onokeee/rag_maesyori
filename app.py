import os
import secrets
import sys
from pathlib import Path

from flask import Flask

from config import BASE_DIR, Config
from models import database
from services import llm
from views import form_types, forms, history, home, settings, tables

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

    for module in (home, forms, form_types, tables, settings, history):
        app.register_blueprint(module.bp)

    _recover_jobs(app)

    @app.context_processor
    def inject_model_picker():
        # ヘッダーのモデル選択（全画面共通）
        return {"model_picker": {"current": llm.current_model(), "models": llm.available(),
                                 "llm_ready": llm.is_configured()}}

    return app


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
