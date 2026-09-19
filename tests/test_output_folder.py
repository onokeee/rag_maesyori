"""保存先フォルダへの保存（design.md 2.6・3.3）。

- 設定: 絶対パス・ある・フォルダ・書き込める、を保存のときに確かめる
- 保存: ダウンロードと同じ中身を1ファイルずつ書く。一覧表の管理用ファイルは設定でオンのときだけサブフォルダへ
- 同じ名前: 「止める」なら何も書かない。「上書きする」なら上書きして知らせる
- 失敗: この回に書いたものを消し（上書きしたものは戻し）、データは消さない
- 成功: ダウンロードを渡し終えたときと同じくデータを消す
"""
import io
import os
import zipfile
from pathlib import Path

import pytest
import yaml

from models import database as db
from services import output_folder
from tables import pipeline, store
from tests.test_retention import _add_confirmed_document, _confirmed_import, _rows_for, _uploaded_files


# ---- 共通 -------------------------------------------------------------------------

@pytest.fixture
def out_dir(tmp_path) -> Path:
    folder = tmp_path / "lightrag_inputs"
    folder.mkdir()
    return folder


def _set_folder(client, folder, *, save_admin=False, on_conflict="stop"):
    form = {"folder": str(folder), "on_conflict": on_conflict}
    if save_admin:
        form["save_admin"] = "on"
    return client.post("/settings/output", data=form)


def _files(folder: Path) -> list[str]:
    """フォルダ内のファイル（サブフォルダの中も。フォルダからの相対パス、/ 区切り）。"""
    return sorted(p.relative_to(folder).as_posix() for p in folder.rglob("*") if p.is_file())


def _single_markdown(app, doc_id):
    from views.forms import _single_markdown as build
    with app.test_request_context("/"):
        return build(doc_id)


def _doc_gone(app, doc_id) -> bool:
    return _rows_for(app, "documents", ("document_id",), doc_id) == {}


# ---- 設定 -------------------------------------------------------------------------

def test_settings_tab_is_listed_and_empty_by_default(client):
    page = client.get("/settings/output").get_data(as_text=True)
    assert "保存先フォルダ" in page and "_管理用_RAGには入れない" in page
    assert "保存を止めて知らせる" in page and "上書きする" in page
    # 設定のタブ・ナビから開ける
    assert 'href="/settings/output"' in client.get("/settings/ai").get_data(as_text=True)
    assert 'href="/settings/output"' in client.get("/").get_data(as_text=True)


def test_saving_a_valid_folder_stores_it_under_data(app, client, out_dir):
    res = _set_folder(client, f'"{out_dir}"', save_admin=True, on_conflict="overwrite")   # エクスプローラーの " 付き
    assert res.status_code == 302
    saved = yaml.safe_load((Path(app.config["DATA_DIR"]) / "output_settings.yaml").read_text(encoding="utf-8"))
    assert saved == {"folder": str(out_dir.resolve()), "save_admin": True, "on_conflict": "overwrite"}
    page = client.get("/settings/output").get_data(as_text=True)
    assert str(out_dir.resolve()) in page
    # 書き込めるかの確かめで作った一時ファイルは残らない
    assert _files(out_dir) == []


def test_an_empty_folder_turns_the_feature_off(app, client, out_dir):
    _set_folder(client, out_dir)
    assert _set_folder(client, "").status_code == 302
    with app.test_request_context("/"):
        assert output_folder.configured_folder() == ""


@pytest.mark.parametrize("value, message", [
    ("inputs\\lightrag", "相対パスは使えません"),
    ("missing", "このフォルダがありません"),
    ("a_file", "これはファイルです"),
])
def test_unusable_folders_are_refused(app, client, tmp_path, value, message):
    (tmp_path / "a_file").write_text("x", encoding="utf-8")
    folder = value if value.startswith("inputs") else str(tmp_path / value)
    res = _set_folder(client, folder)
    page = res.get_data(as_text=True)
    assert res.status_code == 400 and message in page
    assert folder in page   # 打った値は残る
    assert not (Path(app.config["DATA_DIR"]) / "output_settings.yaml").exists()


def test_a_folder_that_cannot_be_written_is_refused(app, client, out_dir, monkeypatch):
    """アクセス権で断られたら1回で諦める（tempfile.mkstemp は Windows で名前を変えながら試し続け、戻ってこない）。"""
    real_open = os.open
    probes = []

    def refuse(path, *args, **kwargs):
        if ".rag_probe_" in str(path):
            probes.append(path)
            raise PermissionError(13, "アクセスが拒否されました")
        return real_open(path, *args, **kwargs)

    monkeypatch.setattr(output_folder.os, "open", refuse)
    res = _set_folder(client, out_dir)
    assert res.status_code == 400 and "書き込めません" in res.get_data(as_text=True)
    assert len(probes) == 1
    assert not (Path(app.config["DATA_DIR"]) / "output_settings.yaml").exists()
    # [フォルダを確認する] も同じ（止まらずに答える）
    result = client.post("/settings/output/check", json={"folder": str(out_dir)}).get_json()
    assert not result["ok"] and "書き込めません" in result["errors"][0] and len(probes) == 2


def test_a_folder_path_too_long_for_windows_is_reported_as_such(app, client, out_dir, monkeypatch):
    """長いパスが無効な Windows では、確かめ用のファイルも書けない長さなら「長すぎる」と言う（権限の問題とは言わない）。"""
    monkeypatch.setattr(output_folder, "path_limit", lambda: len(str(out_dir.resolve())) + 20)
    res = _set_folder(client, out_dir)
    page = res.get_data(as_text=True)
    assert res.status_code == 400 and "フォルダのパスが長すぎます" in page and "権限" not in page
    # 書けるが、名前に使える文字数が少ないときは注意を出す
    monkeypatch.setattr(output_folder, "path_limit", lambda: len(str(out_dir.resolve())) + 60)
    assert _set_folder(client, out_dir).status_code == 302
    assert "ファイル名に使えるのは59文字までです" in client.get("/settings/output").get_data(as_text=True)
    result = client.post("/settings/output/check", json={}).get_json()
    assert result["ok"] and "ファイル名に使えるのは59文字まで" in result["warning"]
    monkeypatch.setattr(output_folder, "path_limit", lambda: None)
    assert "ファイル名に使える" not in client.get("/settings/output").get_data(as_text=True)


def test_the_apps_own_data_folder_is_refused(app, client):
    res = _set_folder(client, Path(app.config["UPLOAD_DIR"]))
    assert res.status_code == 400 and "このアプリがデータを置くフォルダ" in res.get_data(as_text=True)


def test_check_button_revalidates_the_folder(client, out_dir):
    (out_dir / "__parsed__").mkdir()
    (out_dir / "既存.md").write_text("x", encoding="utf-8")
    result = client.post("/settings/output/check", json={"folder": str(out_dir)}).get_json()
    assert result["ok"] and result["folder"] == str(out_dir.resolve())
    assert result["md_count"] == 1 and result["parsed_dir"] is True
    # 入力が空なら保存済みのフォルダを確かめる。消えていれば使えないと出る
    # （__parsed__ は保存先にできないので、別のサブフォルダで試す）
    (out_dir / "sub").mkdir()
    assert _set_folder(client, out_dir / "sub").status_code == 302
    (out_dir / "sub").rmdir()
    result = client.post("/settings/output/check", json={}).get_json()
    assert not result["ok"] and "このフォルダがありません" in result["errors"][0]


# ---- ボタン -------------------------------------------------------------------------

def test_save_buttons_are_hidden_until_a_folder_is_set(app, client, out_dir):
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    first, _p = _add_confirmed_document(app, "1.xlsx", value="EQ-101", batch_id="B", order=0)
    for url in (f"/forms/{doc_id}/done", f"/forms/{doc_id}", f"/forms/{first}/done", "/"):
        page = client.get(url).get_data(as_text=True)
        assert "save-to-folder" not in page and "保存先フォルダに保存" not in page, url

    _set_folder(client, out_dir)
    single = client.get(f"/forms/{doc_id}/done").get_data(as_text=True)
    assert f'action="/forms/{doc_id}/save-to-folder"' in single
    assert "保存先フォルダに保存すると、この帳票のデータはこのPCのアプリから消えます" in single
    # 主ボタンは保存、ダウンロードは次の選択肢として残る
    assert 'class="btn primary" title="保存先:' in single
    assert f'href="/forms/{doc_id}/download.md"' in single
    assert f'action="/forms/{doc_id}/save-to-folder"' in client.get(f"/forms/{doc_id}").get_data(as_text=True)
    batch = client.get(f"/forms/{first}/done").get_data(as_text=True)
    assert 'action="/forms/batches/B/save-to-folder"' in batch
    home = client.get("/").get_data(as_text=True)
    assert f'action="/forms/{doc_id}/save-to-folder"' in home and 'action="/forms/batches/B/save-to-folder"' in home
    # まとまりの中の1件ずつ（「.mdだけ」の横）にも出る
    assert f'action="/forms/{first}/save-to-folder"' in home


def test_table_save_button_follows_the_setting(app, client, out_dir):
    import_id = _confirmed_import(app, client)
    assert "save-to-folder" not in client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)
    _set_folder(client, out_dir)
    page = client.get(f"/tables/imports/{import_id}/done").get_data(as_text=True)
    assert f'action="/tables/imports/{import_id}/save-to-folder"' in page
    assert f'action="/tables/imports/{import_id}/save-to-folder"' in client.get("/").get_data(as_text=True)


def test_batch_save_confirm_follows_the_review_autosave(app, client, out_dir):
    """確認画面の途中保存で修正中になったら、保存のフォームの確認文も書き替える（review.js の data-batch-save）。"""
    from tests.test_forms_fixes4 import _confirmed_doc, _state
    _set_folder(client, out_dir)
    first = _confirmed_doc(app, "1.xlsx", batch_id="B", order=0)
    _confirmed_doc(app, "2.xlsx", batch_id="B", order=1)
    page = client.get(f"/forms/{first}/review").get_data(as_text=True)
    assert "data-batch-save" in page and 'action="/forms/batches/B/save-to-folder"' in page
    client.post(f"/forms/{first}/draft", json={"values": {"equipment_name": "直した名前"}})
    assert _state(app, first) == "modified"
    summary = client.post(f"/forms/{first}/preview", json={"values": {}}).get_json()
    assert "修正中の帳票が1件あります" in summary["batch"]["save_confirm"]
    assert "保存先フォルダに保存すると" in summary["batch"]["save_confirm"]


# ---- 帳票 ---------------------------------------------------------------------------

def test_saving_one_form_writes_the_download_bytes_and_purges(app, client, out_dir):
    _set_folder(client, out_dir)
    doc_id, path = _add_confirmed_document(app, "報告書.xlsx")
    name, body = _single_markdown(app, doc_id)

    res = client.post(f"/forms/{doc_id}/save-to-folder")
    page = res.get_data(as_text=True)
    assert res.status_code == 200 and "保存先フォルダに保存しました" in page
    assert name in page and str(out_dir.resolve()) in page and "/documents/scan" in page
    assert res.headers["Cache-Control"] == "no-store"
    with client.session_transaction() as session:   # ファイル名を flash（セッションクッキー）に入れない
        assert name not in str(dict(session)) and Path(name).stem not in str(dict(session))
    assert _files(out_dir) == [name]
    assert (out_dir / name).read_bytes() == body
    # ダウンロードを渡し終えたときと同じく、何も残らない
    assert not path.exists() and _uploaded_files(app) == [] and _doc_gone(app, doc_id)
    assert client.post(f"/forms/{doc_id}/save-to-folder").status_code == 404


def test_saving_a_batch_writes_the_same_files_as_the_zip(app, client, out_dir):
    for batch in ("A", "B"):
        for order, value in enumerate(("EQ-001", "EQ-002", "EQ-003")):
            # 元ファイル名も同じにする（md の出典に出るので、A の zip と B の保存で中身がそろう）
            _add_confirmed_document(app, f"{order}.xlsx", value=value, batch_id=batch, order=order)
    with zipfile.ZipFile(io.BytesIO(client.get("/forms/batches/A/download.zip").data)) as zf:
        expected = {n: zf.read(n) for n in zf.namelist()}

    _set_folder(client, out_dir)
    res = client.post("/forms/batches/B/save-to-folder")
    assert res.status_code == 200 and "Markdown 3件を保存しました" in res.get_data(as_text=True)
    assert _files(out_dir) == sorted(expected)   # zip ではなく1ファイルずつ、フォルダの直下に
    for name, data in expected.items():
        assert (out_dir / name).read_bytes() == data
    with app.app_context():
        assert db.list_batch_documents("B") == []
    assert _uploaded_files(app) == []


def test_saving_only_the_confirmed_forms_of_a_batch_keeps_the_rest(app, client, out_dir):
    _set_folder(client, out_dir)
    done_id, _p = _add_confirmed_document(app, "1.xlsx", batch_id="B", order=0)
    with app.app_context():
        pending = db.create_document("2.xlsx", "0" * 64, "documents/2.xlsx", batch_id="B", batch_order=1)
    # 未確定が残っているのに「全部」を押したら断る（何も書かない）
    res = client.post("/forms/batches/B/save-to-folder")
    assert res.status_code == 302 and _files(out_dir) == []
    res = client.post("/forms/batches/B/save-to-folder", data={"confirmed_only": "1"})
    assert res.status_code == 200 and len(_files(out_dir)) == 1
    assert _doc_gone(app, done_id)
    with app.app_context():
        assert db.get_document(pending) is not None


# ---- 同じ名前 --------------------------------------------------------------------------

def test_a_name_clash_stops_the_save_and_keeps_the_data(app, client, out_dir):
    _set_folder(client, out_dir)
    doc_id, path = _add_confirmed_document(app, "報告書.xlsx")
    name, _body = _single_markdown(app, doc_id)
    (out_dir / name).write_bytes(b"old")

    res = client.post(f"/forms/{doc_id}/save-to-folder")
    page = res.get_data(as_text=True)
    assert res.status_code == 409 and "同じ名前のファイルがあります" in page and name in page
    assert "削除してから" in page and "ファイル名" in page
    assert (out_dir / name).read_bytes() == b"old" and _files(out_dir) == [name]
    assert path.exists() and not _doc_gone(app, doc_id)


def test_a_file_already_taken_in_by_lightrag_also_stops_the_save(app, client, out_dir):
    """LightRAG は取り込み終えたファイルを __parsed__ に移す。同じ名前で入れてもスキャンで取り込まれず、__parsed__ に移されるだけ。"""
    _set_folder(client, out_dir)
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    name, _body = _single_markdown(app, doc_id)
    (out_dir / "__parsed__").mkdir()
    (out_dir / "__parsed__" / name).write_bytes(b"old")
    res = client.post(f"/forms/{doc_id}/save-to-folder")
    assert res.status_code == 409 and "__parsed__" in res.get_data(as_text=True)
    assert _files(out_dir) == [f"__parsed__/{name}"] and not _doc_gone(app, doc_id)


def test_lightrag_key_matches_lightrags_notion_of_the_same_document():
    """1.5.7 は末尾の .[ヒント] を外した名前で文書を見分け、__parsed__ では _001 などの番号付きで残す。"""
    key = output_folder.lightrag_key
    assert key("T_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)].md") == key("T_2026-08.md")
    assert key("t_2026-08.MD") == key("T_2026-08.md")   # Windows では同じファイル
    assert key("T_2026-08_001.md", archived=True) == key("T_2026-08.md")
    assert key("T_2026-08_001.md") != key("T_2026-08.md")   # 直下では番号付きは別の名前
    assert key("T_2026-08.md") != key("T_2026-09.md")


def test_a_hint_variant_or_an_archived_copy_also_counts_as_a_clash(app, client, out_dir):
    _set_folder(client, out_dir)
    with app.test_request_context("/"):
        (out_dir / "T_2026-08.md").write_bytes(b"old")
        with pytest.raises(output_folder.SaveConflict) as info:
            output_folder.write_files([("T_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)].md", b"new")])
        assert info.value.existing == ["T_2026-08.md"]
        (out_dir / "T_2026-08.md").unlink()
        (out_dir / "__parsed__").mkdir()
        (out_dir / "__parsed__" / "T_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)]_001.md").write_bytes(b"old")
        with pytest.raises(output_folder.SaveConflict) as info:
            output_folder.write_files([("T_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)].md", b"new")])
        assert info.value.parsed == ["T_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)]_001.md"]
        # 2つが LightRAG で同じ文書になる組み合わせは書かない
        with pytest.raises(output_folder.SaveError):
            output_folder.write_files([("A.md", b"1"), ("a.[legacy-R].md", b"2")])
    assert _files(out_dir) == ["__parsed__/T_2026-08.[legacy-R(chunk_ts=1500,chunk_ol=0)]_001.md"]


def test_an_archived_name_that_itself_ends_in_digits_counts_as_a_clash(app, client, out_dir):
    """名前がもともと _123 で終わる文書（No. 123 など）は、LightRAG が番号を足さずに __parsed__ に移す。"""
    _set_folder(client, out_dir)
    (out_dir / "__parsed__").mkdir()
    (out_dir / "__parsed__" / "設備修理報告書_No._123.md").write_bytes(b"old")
    with app.test_request_context("/"):
        with pytest.raises(output_folder.SaveConflict) as info:
            output_folder.write_files([("設備修理報告書_No._123.md", b"new")])
        assert info.value.parsed == ["設備修理報告書_No._123.md"]
        # 番号付きで移されたもの（2回目の取り込み）も今までどおり重なり
        (out_dir / "__parsed__" / "設備修理報告書_No._123.md").rename(out_dir / "__parsed__" / "設備修理報告書_No._123_001.md")
        with pytest.raises(output_folder.SaveConflict):
            output_folder.write_files([("設備修理報告書_No._123.md", b"new")])
    assert _files(out_dir) == ["__parsed__/設備修理報告書_No._123_001.md"]


def test_an_archived_digit_name_stops_a_form_save_end_to_end(app, client, out_dir, monkeypatch):
    _set_folder(client, out_dir)
    doc_id, path = _add_confirmed_document(app, "報告書.xlsx")
    monkeypatch.setattr("views.forms.markdown_filename", lambda _doc, _ext: "点検記録_001.md")
    (out_dir / "__parsed__").mkdir()
    (out_dir / "__parsed__" / "点検記録_001.md").write_bytes(b"old")
    res = client.post(f"/forms/{doc_id}/save-to-folder")
    page = res.get_data(as_text=True)
    assert res.status_code == 409 and "点検記録_001.md" in page
    # __parsed__ だけの重なり: 削除のときにファイルも消す案内をし、「上書きする」は勧めない
    assert "delete_file=true" in page and "__parsed__" in page and "「上書きする」にします" not in page
    assert _files(out_dir) == ["__parsed__/点検記録_001.md"] and path.exists() and not _doc_gone(app, doc_id)


def test_overwrite_policy_replaces_and_reports(app, client, out_dir):
    _set_folder(client, out_dir, on_conflict="overwrite")
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    name, body = _single_markdown(app, doc_id)
    (out_dir / name).write_bytes(b"old")
    res = client.post(f"/forms/{doc_id}/save-to-folder")
    page = res.get_data(as_text=True)
    assert res.status_code == 200 and "上書きしました" in page and name in page
    assert (out_dir / name).read_bytes() == body and _files(out_dir) == [name]
    assert _doc_gone(app, doc_id)
    # 保存したあとで LightRAG の削除にファイルも消させると、いま保存したファイルも消える（delete_file_variants_by_file_path）
    assert "チェックを入れないで" in page and "delete_file=false" in page


def test_a_plain_save_does_not_show_the_delete_file_warning(app, client, out_dir):
    _set_folder(client, out_dir)
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    page = client.post(f"/forms/{doc_id}/save-to-folder").get_data(as_text=True)
    assert "delete_file=false" not in page
    # 次に行うことの手順には、削除するときのチェックの向きを書く
    assert "チェックを入れません" in page


def test_overwriting_a_read_only_file_leaves_no_backup(app, client, out_dir):
    import stat
    _set_folder(client, out_dir, on_conflict="overwrite")
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    name, body = _single_markdown(app, doc_id)
    (out_dir / name).write_bytes(b"old")
    os.chmod(out_dir / name, stat.S_IREAD)
    try:
        res = client.post(f"/forms/{doc_id}/save-to-folder")
        assert res.status_code == 200
        assert _files(out_dir) == [name] and (out_dir / name).read_bytes() == body   # .rag_….bak が残らない
    finally:
        for p in out_dir.iterdir():
            os.chmod(p, stat.S_IWRITE)


# ---- 失敗したとき ------------------------------------------------------------------------

def test_a_failed_third_write_removes_what_was_written_and_keeps_the_data(app, client, out_dir, monkeypatch):
    _set_folder(client, out_dir, on_conflict="overwrite")
    ids = [_add_confirmed_document(app, f"{i}.xlsx", value=f"EQ-00{i}", batch_id="B", order=i)[0] for i in range(3)]
    names = [_single_markdown(app, i)[0] for i in ids]
    (out_dir / names[0]).write_bytes(b"old")   # 1件目は上書きになる。失敗したら元に戻る

    real_replace = os.replace

    def failing_replace(src, dst):
        if Path(dst).name == names[2]:
            raise OSError(28, "ディスクの空き領域が不足しています")
        return real_replace(src, dst)

    monkeypatch.setattr(output_folder.os, "replace", failing_replace)
    res = client.post("/forms/batches/B/save-to-folder")
    page = res.get_data(as_text=True)
    assert res.status_code == 500 and "保存できませんでした" in page and "ディスクの空き領域" in page
    assert f"「{names[2]}」を書き込めませんでした" in page   # どのファイルか分かる
    monkeypatch.setattr(output_folder.os, "replace", real_replace)
    # 書いた2件目は消え、上書きした1件目は元の中身に戻り、一時ファイルも残らない
    assert _files(out_dir) == [names[0]] and (out_dir / names[0]).read_bytes() == b"old"
    for doc_id in ids:
        assert not _doc_gone(app, doc_id)
    # 直せばもう一度保存できる
    assert client.post("/forms/batches/B/save-to-folder").status_code == 200
    assert sorted(_files(out_dir)) == sorted(names)


def test_files_that_cannot_be_removed_after_a_failure_are_listed(app, client, out_dir, monkeypatch):
    """書き込みの失敗のあと、掴まれていて消せなかったファイルは「消した」と言わずに名前を出す。"""
    monkeypatch.setattr(output_folder, "REMOVE_RETRY_WAIT", 0)
    _set_folder(client, out_dir)
    ids = [_add_confirmed_document(app, f"{i}.xlsx", value=f"EQ-10{i}", batch_id="B", order=i)[0] for i in range(3)]
    names = [_single_markdown(app, i)[0] for i in ids]
    real_replace, real_unlink = os.replace, Path.unlink

    def failing_replace(src, dst):
        if Path(dst).name == names[2]:
            raise OSError(28, "ディスクの空き領域が不足しています")
        return real_replace(src, dst)

    def locked_unlink(self, *args, **kwargs):
        if self.suffix == ".md":
            raise PermissionError(32, "別のプロセスが使用中です")
        return real_unlink(self, *args, **kwargs)

    monkeypatch.setattr(output_folder.os, "replace", failing_replace)
    monkeypatch.setattr(Path, "unlink", locked_unlink)
    res = client.post("/forms/batches/B/save-to-folder")
    monkeypatch.setattr(Path, "unlink", real_unlink)
    monkeypatch.setattr(output_folder.os, "replace", real_replace)
    page = res.get_data(as_text=True)
    assert res.status_code == 500 and "この回に書いた分は消し" not in page
    assert "2件を消せませんでした" in page and "スキャンする前に" in page
    for name in names[:2]:
        assert str(out_dir.resolve() / name) in page
    assert sorted(_files(out_dir)) == sorted(names[:2])
    for doc_id in ids:
        assert not _doc_gone(app, doc_id)


@pytest.mark.skipif(os.name != "nt", reason="ジャンクションは Windows だけ")
def test_a_junction_for_the_admin_subfolder_is_refused(app, client, out_dir, tmp_path):
    import _winapi
    outside = tmp_path / "outside"
    outside.mkdir()
    _winapi.CreateJunction(str(outside), str(out_dir / output_folder.ADMIN_SUBDIR))
    try:
        import_id = _confirmed_import(app, client)
        _set_folder(client, out_dir, save_admin=True)
        res = client.post(f"/tables/imports/{import_id}/save-to-folder")
        assert res.status_code == 400 and "ジャンクション" in res.get_data(as_text=True)
        assert list(outside.iterdir()) == []
        assert [p.name for p in out_dir.iterdir()] == [output_folder.ADMIN_SUBDIR]   # md も書いていない
        assert _rows_for(app, "table_imports", ("import_id", "table_import_id"), import_id) != {}
    finally:
        os.rmdir(out_dir / output_folder.ADMIN_SUBDIR)   # ジャンクションだけを外す（先の中身は消さない）


def test_names_too_long_for_windows_are_refused_before_writing(app, client, out_dir, monkeypatch):
    _set_folder(client, out_dir)
    ids = [_add_confirmed_document(app, f"{i}.xlsx", value=f"EQ-20{i}", batch_id="B", order=i)[0] for i in range(2)]
    monkeypatch.setattr("views.forms.markdown_filename", lambda doc, _ext: "設備修理報告書_" + "長い名前" * 5 + f"_{doc['id']:05d}.md")
    # 確かめ用の一時ファイル（31文字）は入るが、md の名前は入らない
    monkeypatch.setattr(output_folder, "path_limit", lambda: len(str(out_dir.resolve())) + 35)
    res = client.post("/forms/batches/B/save-to-folder")
    page = res.get_data(as_text=True)
    assert res.status_code == 400 and "パスが長すぎて" in page
    assert _files(out_dir) == []
    for doc_id in ids:
        assert not _doc_gone(app, doc_id)


def test_a_name_that_points_outside_the_folder_is_refused(app, client, out_dir, monkeypatch):
    _set_folder(client, out_dir)
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    monkeypatch.setattr("views.forms.markdown_filename", lambda _doc, _ext: "..\\escape.md")
    res = client.post(f"/forms/{doc_id}/save-to-folder")
    assert res.status_code == 400 and "保存しませんでした" in res.get_data(as_text=True)
    assert not (out_dir.parent / "escape.md").exists() and _files(out_dir) == []
    assert not _doc_gone(app, doc_id)


@pytest.mark.parametrize("name", ["../x.md", "sub/x.md", "..", "C:x.md", ""])
def test_write_files_refuses_names_outside_the_folder(app, client, out_dir, name):
    _set_folder(client, out_dir)
    with app.test_request_context("/"):
        with pytest.raises(output_folder.SaveError):
            output_folder.write_files([("ok.md", b"a"), (name, b"b")])
    assert _files(out_dir) == [] and not (out_dir.parent / "x.md").exists()


def test_another_site_cannot_trigger_a_save(app, client, out_dir):
    _set_folder(client, out_dir)
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    for headers in ({"Origin": "https://evil.example"}, {"Sec-Fetch-Site": "cross-site"}):
        res = client.post(f"/forms/{doc_id}/save-to-folder", headers=headers)
        assert res.status_code == 403
    assert client.post("/forms/batches/B/save-to-folder", headers={"Origin": "https://evil.example"}).status_code == 403
    assert _files(out_dir) == [] and not _doc_gone(app, doc_id)


def test_saving_without_a_folder_setting_is_refused(app, client):
    doc_id, _path = _add_confirmed_document(app, "報告書.xlsx")
    res = client.post(f"/forms/{doc_id}/save-to-folder")
    assert res.status_code == 400 and "保存先フォルダが設定されていません" in res.get_data(as_text=True)
    assert not _doc_gone(app, doc_id)


# ---- 一覧表 --------------------------------------------------------------------------

def _table_state(app, import_id):
    with app.app_context():
        imp = store.get_import(import_id)
        return imp, Path(app.config["UPLOAD_DIR"]) / imp["stored_path"], pipeline.import_dir(import_id)


def _expected_table_files(app, import_id):
    with app.app_context():
        imp = store.get_import(import_id)
        return pipeline.build_download_files(import_id, imp, pipeline.spec_for_import(imp))


def test_saving_a_table_import_writes_only_the_md_files(app, client, out_dir):
    import_id = _confirmed_import(app, client)
    _set_folder(client, out_dir)
    md_files, _extras = _expected_table_files(app, import_id)
    _imp, stored, folder = _table_state(app, import_id)

    res = client.post(f"/tables/imports/{import_id}/save-to-folder")
    assert res.status_code == 200 and f"Markdown {len(md_files)}件を保存しました" in res.get_data(as_text=True)
    assert _files(out_dir) == sorted(n for n, _d in md_files)   # 管理用のファイルは保存しない（設定がオフ）
    for name, data in md_files:
        assert (out_dir / name).read_bytes() == data
    assert not stored.exists() and not folder.exists()
    assert _rows_for(app, "table_imports", ("import_id", "table_import_id"), import_id) == {}


def test_saving_a_table_import_with_admin_files_uses_the_subfolder(app, client, out_dir):
    import_id = _confirmed_import(app, client)
    _set_folder(client, out_dir, save_admin=True)
    md_files, extras = _expected_table_files(app, import_id)
    # zip の中身と同じ（RAG投入用/ と 管理用_RAGには入れない/）
    with app.app_context():
        imp = store.get_import(import_id)
        with zipfile.ZipFile(io.BytesIO(pipeline.build_download(import_id, imp, pipeline.spec_for_import(imp)))) as zf:
            zipped = {n: zf.read(n) for n in zf.namelist()}

    res = client.post(f"/tables/imports/{import_id}/save-to-folder")
    assert res.status_code == 200
    top = sorted(p.name for p in out_dir.iterdir() if p.is_file())
    assert top == sorted(n for n, _d in md_files)
    admin_dirs = list((out_dir / output_folder.ADMIN_SUBDIR).iterdir())
    assert len(admin_dirs) == 1 and admin_dirs[0].name.startswith("トラブル対応一覧_")
    assert sorted(p.name for p in admin_dirs[0].iterdir()) == sorted(extras)
    for name in extras:
        assert (admin_dirs[0] / name).read_bytes() == zipped[f"管理用_RAGには入れない/{name}"]
    for name, _d in md_files:
        assert (out_dir / name).read_bytes() == zipped[f"RAG投入用/{name}"]
    assert _rows_for(app, "table_imports", ("import_id", "table_import_id"), import_id) == {}


def test_a_failed_table_save_purges_nothing(app, client, out_dir, monkeypatch):
    import_id = _confirmed_import(app, client)
    _set_folder(client, out_dir, save_admin=True)
    _imp, stored, folder = _table_state(app, import_id)
    real_replace = os.replace
    calls = []

    def failing_replace(src, dst):
        calls.append(dst)
        if Path(dst).name == "問題一覧.csv":
            raise OSError(5, "アクセスが拒否されました")
        return real_replace(src, dst)

    monkeypatch.setattr(output_folder.os, "replace", failing_replace)
    res = client.post(f"/tables/imports/{import_id}/save-to-folder")
    assert res.status_code == 500
    assert _files(out_dir) == []   # md も、作ったサブフォルダも消えている
    assert stored.exists() and folder.exists()
    assert _rows_for(app, "table_imports", ("import_id", "table_import_id"), import_id) != {}
