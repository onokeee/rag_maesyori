"""5巡目の修正で、担当の範囲をまたいで直したところの確認。

- 帳票・帳票の見本のアップロードは、結合セルの面積の上限を帳票用（FORM_MAX_MERGED_CELLS）で確かめる（R5-FUZZ-2）
- ログの区切りの目印・日付ではない書き方の正規表現には、行の先頭の一部だけを渡す（SEC5-3）
- AI接続の保存で設定ファイルが書けなくても 500 の HTML にしない（R5C-2）
"""
from __future__ import annotations

import io
import time
from datetime import date

import pytest

from core import files
from logproc.models import SplitOptions
from logproc.people import PeopleIndex
from logproc.segment import parse_log
from services import settings_store
from tests.test_core_files import _rewrite_merges, _xlsx_with_merge


# ---- R5-FUZZ-2 ------------------------------------------------------------------------

def _row_merged_book(tmp_path, rows: int):
    path = _xlsx_with_merge(tmp_path / "rows.xlsx", "A1:XFD1")
    path.write_bytes(_rewrite_merges(path, [f"A{r}:XFD{r}" for r in range(1, rows + 1)]))
    return path


def test_form_upload_refuses_many_whole_row_merges_quickly(app, client, tmp_path):
    path = _row_merged_book(tmp_path, 120)
    started = time.monotonic()
    res = client.post("/forms/upload", data={"file": (io.BytesIO(path.read_bytes()), "rows.xlsx")},
                      content_type="multipart/form-data")
    assert time.monotonic() - started < 10
    assert res.status_code == 400 and "結合セルの範囲が大きすぎます" in res.get_json()["error"]


def test_form_type_sample_upload_refuses_many_whole_row_merges(app, client, sample_dir, tmp_path):
    """帳票登録の見本ファイルも、帳票と同じ上限で断る（読み書きできるように開くため）。"""
    res = client.post("/form-types/new",
                      data={"name": "設備修理報告書",
                            "samples": ((sample_dir / "standard.xlsx").open("rb"), "standard.xlsx")},
                      content_type="multipart/form-data")
    assert res.status_code == 200, res.get_json()
    pattern_id = res.get_json()["pattern_id"]

    path = _row_merged_book(tmp_path, 120)
    res = client.post(f"/form-types/{pattern_id}/samples",
                      data={"samples": (io.BytesIO(path.read_bytes()), "rows.xlsx")},
                      content_type="multipart/form-data")
    assert res.status_code == 400
    assert "結合セルの範囲が大きすぎます" in res.get_json()["error"]
    assert files.FORM_MAX_MERGED_CELLS < files.MAX_MERGED_CELLS   # 一覧表の上限（読み取り専用で開く）は別


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


# ---- R5C-2（AI接続の保存。設定画面は無くなり、表の取り込み画面の AI整形の段から送る） ------------------

def test_ai_settings_that_cannot_be_saved_return_a_japanese_json_error(client, monkeypatch):
    def locked(*a, **k):
        raise OSError("locked")

    monkeypatch.setattr(settings_store, "write_yaml", locked)
    res = client.post("/tables/ai-connection", json={"models": ["m1"], "default": "m1", "api_key": ""})
    assert res.status_code == 400
    error = res.get_json()["error"]
    assert "設定ファイルに書き込めませんでした" in error and "もう一度保存してください" in error
    assert "Traceback" not in res.get_data(as_text=True)
