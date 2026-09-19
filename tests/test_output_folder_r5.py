r"""保存先フォルダ・設定ファイルまわりの不具合の修正（5巡目）。

- 保存先に、アプリのデータのフォルダを別の書き方（\\?\・\\localhost\C$）で指定できない
- 保存先に、LightRAG の __parsed__ や管理用のフォルダを指定できない（スキャンで読まれない）
- 断られたあとの「いまの設定」は保存済みの値を出す
- 確定済みの分だけ保存したとき、残った未確定の帳票と［次の帳票へ］を出す
- 消し損ねたとき、「消しました」と言わない
- 保存したあとの 404 は、保存先フォルダへの保存も理由に挙げる
- 設定ファイルが UTF-8 以外でも画面が出る。読み書きが重なっても 500 にならない
"""
import os
import threading
from pathlib import Path

import pytest
import yaml

from models import database as db
from services import output_folder, settings_store
from tests.test_output_folder import _files, _set_folder, out_dir  # noqa: F401  (fixture)
from tests.test_retention import _add_confirmed_document


# ---- SEC5-2: 別の書き方のパス ---------------------------------------------------------

@pytest.mark.skipif(os.name != "nt", reason="Windows のパスの書き方")
@pytest.mark.parametrize("spelling", ["device", "admin_share"])
def test_app_data_folder_is_refused_even_when_spelled_differently(app, spelling):
    with app.test_request_context("/"):
        data_dir = Path(app.config["DATA_DIR"]).resolve()
        (data_dir / "sub").mkdir(parents=True, exist_ok=True)
        plain = str(data_dir / "sub")
        if spelling == "device":
            text = "\\\\?\\" + plain
        else:
            text = "\\\\localhost\\" + plain[0] + "$" + plain[2:]
            if not Path(text).exists():
                pytest.skip("管理共有（C$）が使えない環境")
        _path, errors = output_folder.check_folder(text)
        assert errors and "このアプリがデータを置くフォルダの中" in errors[0]


# ---- BH5-2: __parsed__ と管理用のフォルダ ---------------------------------------------------

@pytest.mark.parametrize("name", [output_folder.LIGHTRAG_PARSED_DIR, output_folder.ADMIN_SUBDIR])
def test_parsed_and_admin_folders_cannot_be_the_save_folder(app, tmp_path, name):
    inside = tmp_path / "inputs" / name / "deeper"
    inside.mkdir(parents=True)
    with app.test_request_context("/"):
        for folder in (inside.parent, inside):
            _path, errors = output_folder.check_folder(str(folder))
            assert errors and "INPUT_DIR" in errors[0], folder
        # その1つ上（INPUT_DIR）は使える
        assert output_folder.check_folder(str(tmp_path / "inputs"))[1] == []


# ---- R5UX-1: 断られたあとの「いまの設定」 -----------------------------------------------------

def test_refused_folder_is_not_shown_as_the_current_setting(app, client, out_dir, tmp_path):  # noqa: F811
    assert _set_folder(client, out_dir).status_code == 302
    missing = tmp_path / "missing_folder"
    res = _set_folder(client, missing)
    page = res.get_data(as_text=True)
    assert res.status_code == 400
    assert f'<code id="savedFolder">{out_dir.resolve()}</code>' in page
    assert f'<code id="savedFolder">{missing}' not in page
    assert f'value="{missing}"' in page   # 入力欄には打った値が残る


# ---- R5UX-2: 確定済みの分だけ保存したとき ------------------------------------------------------

def test_saving_only_confirmed_forms_points_to_the_remaining_ones(app, client, out_dir):  # noqa: F811
    _set_folder(client, out_dir)
    _add_confirmed_document(app, "1.xlsx", batch_id="B", order=0)
    with app.app_context():
        pending = db.create_document("2.xlsx", "0" * 64, "documents/2.xlsx", batch_id="B", batch_order=1)
    res = client.post("/forms/batches/B/save-to-folder", data={"confirmed_only": "1"})
    page = res.get_data(as_text=True)
    assert res.status_code == 200
    assert "未確定の1件は残っています" in page and "確定済みの1件のデータは" in page
    assert "このアプリからはデータを消しました" not in page
    assert f'class="btn primary" href="/forms/{pending}/' in page and "次の帳票へ" in page


def test_saving_a_whole_batch_keeps_the_usual_result(app, client, out_dir):  # noqa: F811
    _set_folder(client, out_dir)
    _add_confirmed_document(app, "1.xlsx", batch_id="B", order=0)
    page = client.post("/forms/batches/B/save-to-folder").get_data(as_text=True)
    assert "このアプリからはデータを消しました" in page and "次の帳票へ" not in page


# ---- R5UX-4: 消し損ねたとき ---------------------------------------------------------------

def test_a_failed_purge_does_not_claim_the_data_was_removed(app, client, out_dir, monkeypatch):  # noqa: F811
    from core import purge
    _set_folder(client, out_dir)
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")

    def broken(*_args, **_kwargs):
        raise OSError("ロックされています")

    monkeypatch.setattr(purge, "purge_documents", broken)
    page = client.post(f"/forms/{doc_id}/save-to-folder").get_data(as_text=True)
    assert "消し切れませんでした" in page
    assert "このアプリからはデータを消しました" not in page
    assert f'href="/forms/{doc_id}"' in page and "この帳票を削除" in page


# ---- R5UX-5: 保存したあとの 404 -----------------------------------------------------------

def test_404_after_saving_mentions_the_save_folder(app, client, out_dir):  # noqa: F811
    _set_folder(client, out_dir)
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    assert client.post(f"/forms/{doc_id}/save-to-folder").status_code == 200
    res = client.get(f"/forms/{doc_id}/done")
    assert res.status_code == 404 and "保存先フォルダに保存済み" in res.get_data(as_text=True)


# ---- BH5-1: UTF-8 以外の設定ファイル -----------------------------------------------------------

@pytest.mark.parametrize("encoding", ["utf-16", "cp932", "utf-8-sig"])
def test_settings_file_in_another_encoding_does_not_break_pages(app, client, out_dir, encoding):  # noqa: F811
    path = Path(app.config["DATA_DIR"]) / output_folder.SETTINGS_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump({"folder": str(out_dir.resolve()), "on_conflict": "stop"}, allow_unicode=True),
                    encoding=encoding)
    for url in ("/", "/settings/output"):
        res = client.get(url)
        assert res.status_code == 200, (url, encoding)
    assert str(out_dir.resolve()) in client.get("/settings/output").get_data(as_text=True)


def test_unreadable_bytes_in_settings_file_are_treated_as_empty(app, client):
    path = Path(app.config["DATA_DIR"]) / output_folder.SETTINGS_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"folder: \x81\x7f\xff\xfe\xfd")
    assert client.get("/").status_code == 200 and client.get("/settings/output").status_code == 200


# ---- R5C-2: 読み書きの重なり ---------------------------------------------------------------

def test_reads_and_writes_of_settings_do_not_collide(app):
    """画面を出すスレッドが設定を読んでいる間に保存しても、置き換えが WinError 5 で失敗しない。"""
    errors, stop = [], threading.Event()

    def reader():
        with app.app_context():
            while not stop.is_set():
                if settings_store.read_yaml("prefs.yaml") == {} and \
                        (Path(app.config["DATA_DIR"]) / "prefs.yaml").exists():
                    errors.append("empty")

    with app.app_context():
        settings_store.write_yaml("prefs.yaml", {"n": 0})
        threads = [threading.Thread(target=reader) for _ in range(2)]
        for t in threads:
            t.start()
        try:
            for n in range(150):
                settings_store.write_yaml("prefs.yaml", {"n": n})
        except OSError as exc:
            errors.append(repr(exc))
        finally:
            stop.set()
            for t in threads:
                t.join()
    assert errors == []


def test_a_settings_file_that_stays_locked_gives_a_message_not_a_500(app, client, out_dir, monkeypatch):  # noqa: F811
    monkeypatch.setattr(settings_store, "REMOVE_RETRY_WAIT", 0)

    def locked(_src, _dst):
        raise PermissionError(13, "アクセスが拒否されました")

    monkeypatch.setattr(settings_store.os, "replace", locked)
    res = _set_folder(client, out_dir)
    page = res.get_data(as_text=True)
    assert res.status_code == 500 and "設定ファイルに書き込めませんでした" in page
    assert f'value="{out_dir}"' in page
    assert not list(Path(app.config["DATA_DIR"]).glob("*.tmp"))
    res = client.post("/settings/ai", data={"models": "m1", "default": "m1"})
    assert res.status_code == 302
