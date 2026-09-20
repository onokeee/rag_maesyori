"""一覧表で渡すファイルを「RAG に入れる Markdown だけ」にした変更（7巡目）の確認。

利用者の判断（2026-09-20）: 件数・順位・推移のような集計は RAG の仕組みに向いていないので作らない。
データセット説明も作らない。管理用の CSV も渡さない。
"""
from __future__ import annotations

import io
import zipfile

from tables.markdown import render_all
from tables.spec import spec_from_dict, validate_spec
from tests.tables_helpers import confirmed, panel_html
from tests.test_tables_pipe import _rec, list_spec_dict


# ---- R7-1: 作る md は記録ファイルだけ ---------------------------------------------------------

def test_only_record_files_are_made():
    """集計（月次・設備別年度）とデータセット説明は作らない。"""
    spec = spec_from_dict(list_spec_dict())
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="CMP-101", downtime=30),
               _rec("A-2", record_no="A-2", occurred_at="2026-09-04", equipment_id="CVD-201", downtime=60)]
    files = render_all(spec, records, {})
    assert [f.name for f in files] == ["トラブル対応一覧_2026-08.md", "トラブル対応一覧_2026-09.md"]
    assert {f.kind for f in files} == {"records"}
    text = "\n".join(f.text for f in files)
    for word in ("集計", "データセット説明", "合計は", "件です。", "上位"):
        assert word not in text, word


def test_saved_settings_for_the_summaries_are_read_and_dropped():
    """前の版の取り込み設定（dataset_card・summaries）が残っていても、読み飛ばして記録ファイルだけを作る。"""
    d = list_spec_dict(dataset_card=True, summaries=[{"id": "month", "metrics": ["count"]}])
    spec = spec_from_dict(d)
    assert "dataset_card" not in spec.markdown and "summaries" not in spec.markdown
    assert validate_spec(spec) == []
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="CMP-101")]
    assert [f.name for f in render_all(spec, records, {})] == ["トラブル対応一覧_2026-08.md"]


# ---- R7-2: zip は md だけ（フォルダ分けも管理用CSVも無い） ---------------------------------------

def test_the_zip_is_flat_markdown_only(app, client):
    import_id = confirmed(app, client, "トラブル一覧.csv", "トラブル対応一覧")
    res = client.get(f"/tables/imports/{import_id}/download.zip")
    assert res.status_code == 200
    names = zipfile.ZipFile(io.BytesIO(res.data)).namelist()
    assert names and all(n.endswith(".md") and "/" not in n for n in names)


def test_the_screens_do_not_offer_a_csv_download(app, client):
    """確認・確定の段に正規化CSVのボタンを出さない（渡すのは zip だけ）。"""
    import_id = confirmed(app, client, "トラブル一覧.csv", "トラブル対応一覧")
    for name in ("preview", "done"):
        html = panel_html(client, import_id, name)
        assert "正規化CSV" not in html and "normalized.csv" not in html
    assert "RAG に入れる Markdown（.md）だけ" in panel_html(client, import_id, "done")
    assert client.get(f"/tables/imports/{import_id}/normalized.csv").status_code == 404
