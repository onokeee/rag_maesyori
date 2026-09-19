"""保存先フォルダへの保存の画面側（帳票・一覧表で共通）。

保存は POST で受け、結果の画面をその応答でそのまま出す（リダイレクトしない）。ファイル名は結果の画面にだけ出し、
flash（セッションクッキー）には入れない（design.md 3.3「画面のメッセージにファイル名を出さない」）。
保存がすべて成功したときだけ、ダウンロードを渡し終えたときと同じ削除（core.purge）を行う。
"""
from __future__ import annotations

import logging

from flask import make_response, render_template

from core import purge
from services import output_folder

log = logging.getLogger(__name__)

def folder_ctx() -> dict:
    """ボタンを出すかどうかと、ボタンの説明に出す保存先。"""
    return {"save_folder": output_folder.configured_folder()}


def _page(mode: str, status: int, **ctx):
    response = make_response(render_template("save_result.html", mode=mode, **ctx), status)
    response.headers["Cache-Control"] = "no-store"
    return response


def save_and_purge(md_files, *, purge_fn, purge_args=(), admin_files=None, admin_stem="", back_url: str,
                   what: str, next_url: str = "", remaining: int = 0, remaining_url: str = "",
                   detail_url: str = "", admin_skipped: bool = False):
    """ファイルを保存し、すべて書けたら purge_fn(*purge_args) で消して結果の画面を返す。

    md_files / admin_files: [(名前, 中身)]（ダウンロードの md・zip と同じ中身）
    what: 画面に出す対象の呼び名（「帳票」「まとめ取り込み」「一覧表の取り込み」）
    remaining / remaining_url: 確定済みの分だけ保存したとき、まとまりに残った未確定の帳票の件数と、次に確認する帳票
    detail_url: 消し損ねたときに、残ったデータを削除できる画面（帳票の詳細画面など。無ければホームを案内する）
    admin_skipped: 一覧表で、設定により管理用ファイルを保存しなかったとき True（結果の画面でそのことを知らせる）
    """
    try:
        result = output_folder.write_files(md_files, admin_files, admin_stem)
    except output_folder.SaveConflict as exc:
        return _page("conflict", 409, folder=exc.folder, existing=exc.existing, parsed=exc.parsed,
                     back_url=back_url, what=what)
    except output_folder.SaveError as exc:
        return _page("error", exc.status, message=str(exc), leftovers=exc.leftovers, back_url=back_url, what=what)
    purge_failed = False
    try:
        purge_fn(*purge_args)
    except Exception:   # 保存は終わっている。消し損ねは画面で知らせる（残ったものはホームに出る）
        log.exception("保存先フォルダに保存したデータの削除に失敗しました")
        purge_failed = True
    # 例外にならなくても、掴まれていて消せなかったファイルがあれば「消しました」とは言わない
    purge_failed = purge_failed or purge.purge_incomplete()
    return _page("saved", 200, result=result, purge_failed=purge_failed, what=what, next_url=next_url,
                 remaining=remaining, remaining_url=remaining_url, detail_url=detail_url,
                 admin_skipped=admin_skipped)
