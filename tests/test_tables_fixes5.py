"""一覧表の不具合修正（5巡目）の確認。"""
from __future__ import annotations

import time
from pathlib import Path

import pytest
from openpyxl import Workbook

from tables import pipeline
from tables.markdown import _Names, record_title, render_all
from tables.spec import spec_from_dict, validate_spec
from tests.test_tables_pipe import FakeCtx, _new_import, _rec, list_spec_dict


# ---- R5T-1: 大文字・小文字だけ違う設備のファイル名が Windows で上書きし合わない ------------------------------

def test_names_differing_only_in_case_get_a_suffix():
    names = _Names()
    a = names.make(["a", "ETC-302号機"])
    b = names.make(["a", "Etc-302号機"])
    assert a.casefold() != b.casefold()


def test_entities_differing_only_in_case_keep_every_record(tmp_path):
    d = list_spec_dict(group_by="entity_month")
    for c in d["columns"]:
        if c["key"] == "equipment_id":
            c["normalize"] = ["nfkc"]   # 大文字にそろえない設定
    d["columns"] = [c for c in d["columns"] if c["key"] != "equipment_name"]
    spec = spec_from_dict(d)
    records = [_rec("A-1", record_no="A-1", occurred_at="2026-08-03", equipment_id="ETC-302号機"),
               _rec("A-2", record_no="A-2", occurred_at="2026-08-04", equipment_id="Etc-302号機")]
    files = render_all(spec, records, {}, {})
    assert len({f.name.casefold() for f in files}) == len(files)
    target = tmp_path / "imp"
    target.mkdir()
    pipeline._write_md_dir(target / "md", files)
    written = list((target / "md").glob("*.md"))
    assert len(written) == len(files)
    text = "".join(p.read_text(encoding="utf-8") for p in written)
    assert "【A-1】" in text and "【A-2】" in text


# ---- R5T-2: CSV の「1:30」（[h]:mm を書き出した値）を分として読む ------------------------------------------

# ---- R5-MD-4: 1行目が札と日付だけの現象は、次の行を見出しに使う ----------------------------------------------

@pytest.mark.parametrize("symptom", [
    "発生:2024-04-28 14:50(休日)\n設備:CMP-103\n現象:MESとの通信が断続的に切断",
    "【発生】R06.04.28 14:50\n【現象】MESとの通信が断続的に切断",
])
def test_title_uses_the_next_line_when_the_first_is_only_a_date(symptom):
    spec = spec_from_dict(list_spec_dict())
    title = record_title({"record_no": "CA-1", "occurred_at": "2024-05-01", "equipment_id": "CMP-103",
                          "symptom": symptom}, spec)
    assert "MESとの通信が断続的に切断" in title and "設備:" not in title


def test_title_of_a_plain_symptom_is_unchanged():
    spec = spec_from_dict(list_spec_dict())
    title = record_title({"record_no": "CA-1", "occurred_at": "2024-05-01", "equipment_id": "CMP-103",
                          "symptom": "MES通信断\n詳細は別紙"}, spec)
    assert "MES通信断" in title and "詳細" not in title


# ---- SEC5-3: 選択（|）を含むグループに量指定子が付く区切りの正規表現を断る -------------------------------------

def test_overlapping_alternation_pattern_is_refused():
    from tests.test_tables_fixes4 import _log_spec

    assert any("遅く" in e for e in validate_spec(_log_spec(splitter={"extra_anchors": [r"(?:\d|\d)*年"]})))
    assert any("遅く" in e for e in validate_spec(_log_spec(splitter={"not_date_patterns": [r"(a|ab)+c"]})))
    assert validate_spec(_log_spec(splitter={"not_date_patterns": [r"\d+\.\d+\s*(?:mm|MPa)"]})) == []




# ---- R5-FUZZ-1: 遠くのセル1つ（XFD1 / XFD1048576）で読み込みが止まらない ---------------------------------------

def _far_book(path: Path, far: str, rows: int) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["管理No", "発生日", "設備", "現象", "処置"])
    for i in range(rows):
        ws.append([f"A-{i}", f"2026/08/{i % 28 + 1:02d}", "CMP-1", f"現象{i}", "処置"])
    ws[far] = "メモ"
    wb.save(path)
    return path


def test_far_stray_header_cell_does_not_widen_the_table(tmp_path):
    from tables.detect import guess_layout
    from tables.source import open_source

    path = _far_book(tmp_path / "far.xlsx", "XFD1", 3000)
    started = time.monotonic()
    source = open_source(path, path.name)
    assert sum(1 for _ in source.rows("S")) == 3001
    layout = guess_layout(source, "S")
    assert time.monotonic() - started < 10
    assert len(layout.headers) == 5 and any("XFD" in w for w in layout.warnings)


def test_far_last_cell_in_a_near_column_does_not_pad_every_row(tmp_path):
    from tables.source import open_source

    # 256列目（PAD_MAX_COLUMNS 以下）の遠いセル。行×列で埋めると100秒近くかかっていた
    path = _far_book(tmp_path / "far3.xlsx", "IV1048576", 20)
    started = time.monotonic()
    source = open_source(path, path.name)
    widths = {len(row.cells) for row in source.rows("S")}
    assert time.monotonic() - started < 15
    assert widths == {0, 5, 256}


def test_far_last_cell_upload_finishes_and_reads_the_table(app, client, tmp_path):
    from tests.tables_helpers import upload

    path = _far_book(tmp_path / "far2.xlsx", "XFD1048576", 50)
    started = time.monotonic()
    res = upload(client, path.read_bytes(), "far2.xlsx")
    assert time.monotonic() - started < 30
    assert res.status_code == 200 and res.get_json()["import_id"]


# ---- R5C-1: 確定の処理の途中で渡し終えて消えた取り込みのフォルダを作り直さない ---------------------------------

def test_render_does_not_recreate_a_purged_import(app, monkeypatch):
    from core import purge

    with app.app_context():
        _tid, import_id = _new_import(app)
        pipeline.run_read(FakeCtx(), import_id)
        original = pipeline.load_rows

        def load_then_purge(*a, **k):
            rows = original(*a, **k)
            purge.purge_table_import(import_id)   # 別のタブの保存・ダウンロードが渡し終えた
            return rows

        monkeypatch.setattr(pipeline, "load_rows", load_then_purge)
        with pytest.raises(pipeline.PipelineError):
            pipeline.run_render(FakeCtx(), import_id)
        assert not pipeline.import_dir(import_id).exists()


def test_write_md_dir_does_not_create_a_missing_import_folder(tmp_path):
    with pytest.raises(OSError):
        pipeline._write_md_dir(tmp_path / "gone" / "md", [])
    assert not (tmp_path / "gone").exists()
