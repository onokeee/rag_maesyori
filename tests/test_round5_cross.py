"""5巡目の修正で、担当の範囲をまたいで直したところの確認。

- ホームの一覧表の［保存先フォルダに保存］の確認文も、管理用ファイルを保存しない設定ならそう書く（R5UX-3）
- 一覧表を保存した結果の画面で、管理用ファイルを保存しなかったことを知らせる（R5UX-3）
- 帳票・帳票の見本のアップロードは、結合セルの面積の上限を帳票用（FORM_MAX_MERGED_CELLS）で確かめる（R5-FUZZ-2）
- 取り込み設定の JSON の確認で思わぬ例外が出ても 500 にしない（R5T-4）
- ログの区切りの目印・日付ではない書き方の正規表現には、行の先頭の一部だけを渡す（SEC5-3）
- ヘッダーのモデル選択で設定ファイルが書けなくても 500 の HTML にしない（R5C-2）
"""
from __future__ import annotations

import io
import json
import time
from datetime import date

import pytest

from core import files
from logproc.models import SplitOptions
from logproc.people import PeopleIndex
from logproc.segment import parse_log
from services import llm, settings_store
from tests.test_core_files import _rewrite_merges, _xlsx_with_merge
from tests.test_output_folder import _set_folder
from tests.test_retention import _confirmed_import
from views.tables import SAVE_ADMIN_OFF_NOTE


# ---- R5UX-3 ---------------------------------------------------------------------------

def test_home_table_save_confirm_mentions_admin_files_when_not_saved(app, client, tmp_path):
    folder = tmp_path / "out"
    folder.mkdir()
    _confirmed_import(app, client)
    _set_folder(client, folder)
    assert SAVE_ADMIN_OFF_NOTE in client.get("/").get_data(as_text=True)
    _set_folder(client, folder, save_admin=True)
    assert SAVE_ADMIN_OFF_NOTE not in client.get("/").get_data(as_text=True)


@pytest.mark.parametrize("save_admin", [False, True])
def test_table_save_result_says_whether_admin_files_were_saved(app, client, tmp_path, save_admin):
    folder = tmp_path / "out"
    folder.mkdir()
    import_id = _confirmed_import(app, client)
    _set_folder(client, folder, save_admin=save_admin)
    page = client.post(f"/tables/imports/{import_id}/save-to-folder").get_data(as_text=True)
    assert "保存先フォルダに保存しました" in page
    assert ("管理用ファイル</dt><dd>保存していません" in page) is (not save_admin)


# ---- R5-FUZZ-2 ------------------------------------------------------------------------

def _row_merged_book(tmp_path, rows: int):
    path = _xlsx_with_merge(tmp_path / "rows.xlsx", "A1:XFD1")
    path.write_bytes(_rewrite_merges(path, [f"A{r}:XFD{r}" for r in range(1, rows + 1)]))
    return path


def test_form_upload_refuses_many_whole_row_merges_quickly(app, client, tmp_path):
    path = _row_merged_book(tmp_path, 120)
    started = time.monotonic()
    res = client.post("/forms/upload", data={"file": (io.BytesIO(path.read_bytes()), "rows.xlsx")},
                      content_type="multipart/form-data", follow_redirects=True)
    assert time.monotonic() - started < 10
    assert "結合セルの範囲が大きすぎます" in res.get_data(as_text=True)


def test_form_type_sample_upload_refuses_many_whole_row_merges(app, client, sample_dir, tmp_path):
    from tests.test_endpoints_settings import _make_form_type

    pattern_id = _make_form_type(app, client, sample_dir)
    path = _row_merged_book(tmp_path, 120)
    res = client.post(f"/settings/form-types/{pattern_id}/samples",
                      data={"samples": (io.BytesIO(path.read_bytes()), "rows.xlsx")},
                      content_type="multipart/form-data", follow_redirects=True)
    page = res.get_data(as_text=True)
    assert "結合セルの範囲が大きすぎます" in page and "追加しました" not in page
    assert files.FORM_MAX_MERGED_CELLS < files.MAX_MERGED_CELLS   # 一覧表の上限（読み取り専用で開く）は別


# ---- R5T-4 -----------------------------------------------------------------------------

def test_template_json_is_refused_when_its_check_raises(client, monkeypatch):
    import views.settings as settings_view

    def boom(spec):
        raise TypeError("想定外")

    monkeypatch.setattr(settings_view, "validate_spec", boom)
    col = {"key": "a", "header": "a", "display": "a", "type": "string", "role": "attribute"}
    data = {"spec": {"name": "設定", "columns": [col]}}
    res = client.post("/settings/table-templates/import",
                      data={"file": (io.BytesIO(json.dumps(data).encode()), "t.json")},
                      content_type="multipart/form-data", follow_redirects=True)
    assert res.status_code == 200 and settings_view.NOT_A_TEMPLATE_JSON in res.get_data(as_text=True)


# ---- SEC5-3 ----------------------------------------------------------------------------

@pytest.mark.parametrize("field", ["extra_anchors", "not_date_patterns"])
def test_slow_user_pattern_on_a_long_line_finishes_quickly(field):
    text = "2024/04/01 開始\n" + "1" * 5000 + "\n2024/04/02 終了"
    started = time.monotonic()
    parsed = parse_log(text, date(2024, 4, 1), PeopleIndex(), SplitOptions(**{field: [r"\d*\d*\d*x"]}))
    assert time.monotonic() - started < 5
    assert len(parsed.segments) == 2


def test_user_anchor_still_matches_at_the_line_start():
    parsed = parse_log("◎ 1件目\n◎ 2件目", date(2024, 4, 1), PeopleIndex(), SplitOptions(extra_anchors=[r"◎"]))
    assert len(parsed.segments) == 2


# ---- R5C-2（ヘッダーのモデル選択） ------------------------------------------------------------

def test_model_choice_that_cannot_be_saved_returns_a_japanese_json_error(client, monkeypatch):
    import views.settings as settings_view

    monkeypatch.setattr(llm, "available", lambda: ["m1"])

    def locked(*a, **k):
        raise OSError("locked")

    monkeypatch.setattr(settings_store, "write_yaml", locked)
    res = client.post("/api/models", json={"model": "m1"})
    assert res.status_code == 500 and res.get_json() == {"error": settings_view.MODEL_PREF_WRITE_ERROR}
