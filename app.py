import os
import secrets
import sys
import threading
from pathlib import Path
from urllib.parse import urlsplit

from flask import Blueprint, Flask, abort, jsonify, redirect, render_template, request, url_for

from config import BASE_DIR, Config
from models import database
from views import form_types, forms, is_cross_site_write, tables

_SECRET_FILE = BASE_DIR / ".flask_secret"

# --- 「/」は帳票取り込みへ ------------------------------------------------------------
# ホーム画面は無い。画面は「帳票取り込み」「表の取り込み」「帳票登録」の3つだけで、
# 作業中の一覧もダウンロード待ちの一覧も持たない（利用者の指示 2026-09-20）。
# ここに残るのは「/」を開いたときの転送だけ（古いブックマーク・開いたままの画面の戻り先も拾う）。
home_bp = Blueprint("home", __name__)


@home_bp.get("/", endpoint="index")
def to_forms():
    return redirect(url_for("forms.new"))


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

    app.register_blueprint(home_bp)
    for module in (forms, form_types, tables):
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
        return render_template("errors/403.html"), 403

    @app.errorhandler(404)
    def _not_found(_exc):
        # 403 と同じく、画面からの JSON 送信（static/app.js の postJson）には JSON で返す。
        # ダウンロード済みでデータが消えたあとに、開いたままの画面が途中保存・プレビューを送ると
        # ここに来る。HTML を返すとトーストに「通信に失敗しました」としか出ず、理由が伝わらない。
        if request.accept_mimetypes.best == "application/json" or request.is_json:
            return jsonify(error="このデータはこのPCに残っていません（ダウンロード済みか、削除されています）。"
                                 "取り込みの画面からやり直してください"), 404
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
            return jsonify(error="この操作は受け付けられませんでした。取り込みの画面から開き直してください"), exc.code
        return render_template("errors/400.html"), exc.code

    @app.errorhandler(500)
    def _server_error(_exc):
        return render_template("errors/500.html"), 500

    _recover_jobs(app)
    _purge_pending(app)
    _cleanup_leftovers(app)
    _start_sweeper(app)
    # ヘッダーは「帳票取り込み／表の取り込み／帳票登録」の3つだけ。使うAIモデルの選択は
    # 「表の取り込み」画面の AI整形の段の中（AI接続）に移したので、共通の値は渡さない。
    return app


# 途中で放り出されたものを捨てる間隔（design.md 3.3）。起動時に全部捨て、動いている間は
# SWEEP_HOURS より古いものを SWEEP_INTERVAL ごとに捨てる（ブラウザを閉じたまま開きっぱなしのサーバ向け）。
SWEEP_INTERVAL_SECONDS = 60 * 60


def _purge_pending(app: Flask) -> None:
    """起動時：ダウンロードしていない帳票・一覧表をすべて捨てる（作業中の一覧を持たないため）。"""
    if app.config.get("TESTING"):
        return
    from core import purge

    try:
        with app.app_context():
            forms_removed, tables_removed = purge.purge_all_pending()
        if forms_removed or tables_removed:
            print(f"[app] 途中だった取り込みを捨てました（帳票 {forms_removed} 件・一覧表 {tables_removed} 件）")
    except Exception as exc:  # 起動は止めない
        print(f"[app] 途中だった取り込みの片付けに失敗しました（{exc.__class__.__name__}: {exc}）")


def _start_sweeper(app: Flask) -> None:
    """動いている間：しばらくさわられていない帳票・一覧表を捨て続ける（daemon スレッド）。"""
    if app.config.get("TESTING"):
        return
    from core import purge

    def loop() -> None:
        while True:
            _stop.wait(SWEEP_INTERVAL_SECONDS)
            try:
                with app.app_context():
                    forms_removed, tables_removed = purge.sweep_stale(purge.STALE_HOURS)
                if forms_removed or tables_removed:
                    print(f"[app] {purge.STALE_HOURS}時間さわられていない取り込みを捨てました"
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
            # 前の版で空になった表に残っている番号の続き（何件取り込んだか）も忘れる（design.md 3.3「履歴は持たない」）
            if purge.forget_id_counters(db):
                db.execute("PRAGMA wal_checkpoint(TRUNCATE)")   # 置き換える前の値を app.db-wal に残さない
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
            serve(application, host=HOST, port=PORT, **WAITRESS_OPTIONS)
