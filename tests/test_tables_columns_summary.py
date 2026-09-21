"""「列の対応づけ」の段は、決めることが無ければ要約1行だけにする。

利用者の問い（2026-09-20）「列の対応付けを行う意味は？」への答え。この段で決められるのは
「出す／出さない」と四つの役割（識別番号・日付・対象・経過の記録）だけで、ふつうの一覧表なら
どちらも見出しと値から決まっている。決まっているときは22行の表を出さず、要約1行と［変更する］にする。
決まっていないとき（識別番号が無い・読み取れない値がある など）は、今までどおり表を開く。
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

from tables import store
from tests.tables_helpers import (CSV_TEXT, columns_payload, csv_source, editor_body, panel_html, save_columns,
                                  save_layout, spec_import, upload_csv, wait_import_job)

# 識別番号になる見出しが無い表（「作業メモ」は辞書に無い）
NO_KEY_CSV = "作業メモ,発生日,設備番号,対応内容\r\n" + "".join(
    f'メモ{i},2026-08-{i:02d},EQ-0{i % 3 + 1},"8/{i} 10:00 田中: 確認。\n8/{i} 11:00 佐藤: 復旧。"\r\n'
    for i in range(1, 13))


def _ready(app, client, text: str, name: str = "一覧.csv") -> int:
    """読み取り方と範囲まで決めて、「列の対応づけ」の段を開ける取り込みを作る。"""
    import_id = upload_csv(client, name, text)
    assert csv_source(client, import_id)["next"] == "layout"
    assert save_layout(client, import_id).status_code == 200
    return import_id


def _summary(html: str) -> str:
    found = re.search(r'<p class="col-summary">(.*?)</p>', html, re.S)
    return found.group(1).strip() if found else ""


def _todo(html: str) -> list[str]:
    if "data-columns-todo" not in html:
        return []
    block = html.split("data-columns-todo", 1)[1].split("</ul>", 1)[0]
    return [t.strip() for t in re.findall(r"<li>(.*?)</li>", block, re.S)]


# ---- 決まっているとき: 要約1行 -----------------------------------------------------------------

def test_the_summary_replaces_the_table_when_nothing_needs_deciding(app, client):
    import_id = _ready(app, client, CSV_TEXT)
    html = panel_html(client, import_id, "columns")

    # 四つの役割を名指しし、出す列と出さない列の数も書く
    assert _summary(html) == ("管理No＝識別番号、発生日＝日付、設備番号＝対象、対応内容＝経過の記録として読み取ります。"
                              "8列のうち8列を Markdown に出します（出さない列はありません）。")
    assert _todo(html) == []
    # 表は隠すだけで DOM に残す（［変更する］で開ける・保存で送る中身は変わらない）
    assert "data-columns-table hidden" in html and html.count("data-col ") == 8
    assert "変更する" in html and "この対応づけで読み込む" in html


def test_the_summary_says_so_when_there_is_no_equipment_column(app, client):
    """対象の列が無い表では、あるふりをせず「ありません」と書く（SPEC_CSV の設備名は設備番号ではない）。"""
    import_id = spec_import(app, client)
    summary = _summary(panel_html(client, import_id, "columns"))
    assert summary.startswith("管理No＝識別番号、発生日＝日付、対応内容＝経過の記録として読み取ります。")
    assert "対象の列はありません。" in summary
    assert "8列のうち8列を Markdown に出します（出さない列はありません）。" in summary


# ---- 決まっていないとき: 今までどおり表 ----------------------------------------------------------

def test_the_table_opens_when_the_record_number_is_missing(app, client):
    import_id = _ready(app, client, NO_KEY_CSV, "メモ.csv")
    html = panel_html(client, import_id, "columns")

    assert _summary(html) == ""
    assert _todo(html) == ["識別番号の列が決まっていません。1つ選んでください"]
    assert "data-columns-table hidden" not in html and html.count("data-col ") == 4


def test_the_table_opens_when_a_column_has_values_that_cannot_be_read(app, client):
    text = CSV_TEXT.replace(",60,田中", ",不明,田中")   # 停止時間（数値）に読み取れない値を混ぜる
    import_id = _ready(app, client, text)
    html = panel_html(client, import_id, "columns")

    assert _summary(html) == ""
    assert _todo(html) == ["列「停止時間(分)」に読み取れない値があります（8.3%）。出すかどうか決めてください"]


# ---- ［変更する］は表を出すだけ（送る中身は変わらない） ---------------------------------------------

@pytest.mark.skipif(shutil.which("node") is None, reason="node がない")
def test_the_change_button_only_unhides_the_table():
    js = (Path(__file__).resolve().parents[1] / "static" / "tables.js").read_text(encoding="utf-8")
    start = js.index("function openColumnsTable(")
    end = js.index("function onEditorChange(")
    script = js[start:end] + """
const summary = { hidden: false }, table = { hidden: true };
const editor = { querySelector: (s) => (s === "[data-columns-summary]" ? summary : table) };
openColumnsTable(editor);
openColumnsTable(null);
process.stdout.write(JSON.stringify({ summary: summary.hidden, table: table.hidden }));
"""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "check.mjs"
        path.write_text(script, encoding="utf-8")
        out = subprocess.run(["node", str(path)], capture_output=True, check=True, timeout=30, encoding="utf-8")
    assert json.loads(out.stdout) == {"summary": True, "table": False}


def test_the_spec_saved_through_the_summary_is_the_spec_saved_through_the_table(app, client):
    """要約のまま保存しても、表を開いて何も変えずに保存しても、取り込み設定は1文字も変わらない。"""
    # 要約の段から集めた中身（表は隠れているが DOM にあるので、画面と同じものが集まる）
    summary_id = _ready(app, client, CSV_TEXT, "要約.csv")
    assert "data-columns-table hidden" in panel_html(client, summary_id, "columns")
    body = editor_body(client, summary_id)
    assert save_columns(client, summary_id, body).status_code == 200
    wait_import_job(app, summary_id)

    # 表から集めたのと同じ中身（tests.tables_helpers.columns_payload が作る、画面の表そのままの送り方）
    table_id = _ready(app, client, CSV_TEXT, "表.csv")
    assert save_columns(client, table_id, columns_payload("要約", ai_role="log")).status_code == 200
    wait_import_job(app, table_id)

    with app.app_context():
        assert body["name"] == "要約"
        assert store.get_import(summary_id)["spec_json"] == store.get_import(table_id)["spec_json"]
