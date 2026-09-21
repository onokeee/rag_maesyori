"""起動: python run.py（待ち受け先は環境変数 HOST・PORT で渡す。詳しくは README）。

アプリ本体は app/ にある。ここは本番用サーバー（waitress）で create_app() を動かすだけ。
"""
import sys

from app import DEBUG, HOST, PORT, WAITRESS_OPTIONS, _is_loopback, create_app, startup_notice

if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(f"不明な引数: {sys.argv[1]}（起動は引数なしの python run.py。"
                 f"待ち受け先は環境変数 HOST・PORT で渡します）")
    if DEBUG and not _is_loopback(HOST):
        # DEBUG=True の Flask はデバッガを開く。LAN に出す起動では絶対に開かせない（_NO_DEBUG_MSG と同じ理由）
        raise SystemExit("\n[app] DEBUG = True のまま LAN のアドレス（HOST=%s）では起動しません。\n"
                         "  理由: デバッガが開くと、例外が出たときにブラウザからこのサーバの Python を実行できます。\n"
                         "  詳しいエラーを見たいときは HOST を外して（127.0.0.1 で）起動してください。\n" % HOST)

    application = create_app()
    print(startup_notice())
    if DEBUG:
        application.run(host=HOST, port=PORT, debug=True, use_reloader=False)
    else:
        try:
            from waitress import serve
        except ImportError:
            print("[app] waitress が無いため Flask の開発サーバで起動します（pip install waitress を推奨）")
            application.run(host=HOST, port=PORT, debug=False)
        else:
            serve(application, host=HOST, port=PORT, **WAITRESS_OPTIONS)
