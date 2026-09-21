"""「列の対応づけ」の言葉づかい（利用者の指示 2026-09-21）。

1. 表の5列目「知らせ」をやめる。ほとんどの行で空のうえ、書いていたことは表の上の行と重なっていた。
   はじめから「使わない」にしている理由だけは、見出しの下に小さく残す（無いと、チェックの外れた列を
   入れ直してよいのか分からない）。「値は「日付」らしい」は出さない（日付の役割は日付として読める列に
   しか出ないので、読んでも直しようがない）。
2. 役割の名前を、設備の記録以外の表にも当てはまる言い方にする（設備 → 対象、AI整形の対象（追記ログ）
   → 経過の記録）。中で使う名前（key/date/entity/log/attribute）は変えないので、保存した取り込み設定は
   そのまま読める。
3. AI整形の段（⑤）からも「追記ログ」という言葉をなくす。
"""
from __future__ import annotations

import re
from types import SimpleNamespace

from tables import store
from tests.tables_helpers import (CSV_TEXT, csv_source, editor_body, imported, panel, panel_html, save_columns,
                                  save_layout, upload_csv, wait_import_job)

# 空欄だけの列（予備）と、型が合っていない列（状態コード＝数字だがコードの列）がある表
MIXED_CSV = "管理No,発生日,設備番号,予備,状態コード,対応内容,停止時間\r\n" + "".join(
    f'TR-{i:03d},2026-08-{i % 28 + 1:02d},EQ-0{i % 3 + 1},,{i % 3},'
    f'"8/{i % 28 + 1} 10:00 田中: 確認した。",{i * 5}\r\n' for i in range(1, 29))


def _columns_html(client, text: str = MIXED_CSV, name: str = "言葉.csv") -> str:
    """読み取り方と範囲まで決めて、「列の対応づけ」の段の HTML を返す。"""
    import_id = upload_csv(client, name, text)
    assert csv_source(client, import_id)["next"] == "layout"
    assert save_layout(client, import_id).status_code == 200
    return panel_html(client, import_id, "columns")


def _row(html: str, header: str) -> str:
    """その見出しの行（<tr>…</tr>）だけを取り出す。"""
    found = re.search(r'<tr data-col [^>]*data-header="' + re.escape(header) + r'".*?</tr>', html, re.S)
    assert found, f"列「{header}」の行が無い"
    return found.group(0)


def _head_cell(row_html: str) -> str:
    """その行の「見出し」のセル（2つ目の <td>）。"""
    cells = re.findall(r"<td[^>]*>.*?</td>", row_html, re.S)
    assert len(cells) == 4, cells   # 使う・見出し・役割・値の例の4つだけ
    return cells[1]


# ---- 1. 「知らせ」の列をやめる -----------------------------------------------------------------

def test_the_column_table_has_four_columns_and_no_notice_column(client):
    """表の見出しは「使う／見出し／役割／値の例」の4つだけ。"""
    html = _columns_html(client)
    head = re.search(r"<thead>.*?</thead>", html, re.S).group(0)
    assert re.findall(r"<th>(.*?)</th>", head, re.S) == ["使う", "見出し（そのまま項目名になります）", "役割", "値の例"]
    assert "知らせ" not in html


def test_the_reason_a_column_starts_unchecked_is_under_its_name(client):
    """チェックの外れている列は、その理由を見出しの下に小さく出す（無いと入れ直してよいか分からない）。"""
    html = _columns_html(client)
    row = _row(html, "予備")
    assert "is-unused" in row and "checked" not in row
    cell = _head_cell(row)
    assert "<strong>予備</strong>" in cell
    assert "空欄だけなので、はじめから使わない設定にしています" in cell
    assert 'class="muted small"' in cell   # 見出しより小さく、薄い字で出す

    # ふつうの列（出す列）には何も足さない
    assert _head_cell(_row(html, "対応内容")) == "<td><strong>対応内容</strong></td>"


def test_the_note_says_which_of_the_two_reasons_it_is():
    """理由は2つ（空欄だけ／記録に要らない管理用の列）。出す列には何も書かない。"""
    from views.tables import _unused_note

    assert _unused_note(SimpleNamespace(md="omit", omit_reason="blank")) == \
        "空欄だけなので、はじめから使わない設定にしています"
    assert _unused_note(SimpleNamespace(md="omit", omit_reason="dictionary")) == \
        "記録に不要な管理用の列らしいので、はじめから使わない設定にしています"
    assert _unused_note(SimpleNamespace(md="attribute", omit_reason="")) == ""


def test_the_table_no_longer_guesses_the_type_of_a_column(client):
    """「値は「日付」らしい」は出さない（tests/test_tables_flow.py から移したテスト）。

    日付の役割は日付として読める列にしか出ないので、読んでも人には直せない行き止まりだった。
    型エラー・空欄の割合は、表の上の行（data-columns-todo）に出す。
    """
    html = _columns_html(client)
    assert not re.findall(r"値は「([^」]+)」らしい", html)
    assert "型エラー" not in html and "空欄 " not in html


def test_a_column_that_cannot_be_read_is_still_reported_above_the_table(client):
    """知らせの列をやめても、読み取れない値があることは表の上の行で分かる。"""
    html = _columns_html(client, CSV_TEXT.replace(",60,田中", ",不明,田中"), "読めない.csv")
    assert "data-columns-todo" in html
    todo = html.split("data-columns-todo", 1)[1].split("</ul>", 1)[0]
    assert "列「停止時間(分)」に読み取れない値があります（8.3%）。出すかどうか決めてください" in todo


# ---- 2. 役割の名前 ---------------------------------------------------------------------------

def test_the_role_choices_use_words_that_fit_any_table(client):
    """プルダウンの役割は、設備の記録だけでなくどんな表にも当てはまる言い方にする。"""
    html = _columns_html(client)
    # 日付として読める列（発生日）だけが五つとも選べる（ほかの列に「日付」は出さない）
    select = re.search(r'<select data-field="role".*?</select>', _row(html, "発生日"), re.S).group(0)
    labels = re.findall(r"<option [^>]*>(.*?)</option>", select, re.S)
    assert labels == ["識別番号", "日付", "対象（設備・製品・顧客など）",
                      "経過の記録（1つのセルに日付ごとに書き足した列）", "その他"]
    # 画面のどこにも古い言い方を出さない（「設備」は値の例・見出しには出てよい）
    assert ">設備</option>" not in html and "追記ログ" not in html and "AI整形の対象" not in html


def test_the_screen_explains_what_the_two_new_roles_are_for(client):
    """名前を変えただけでは分からないので、何に使う役割なのかを表と一緒に出す。"""
    html = _columns_html(client)
    block = html.split("data-role-help", 1)[1].split("</ul>", 1)[0]
    assert "その記録が何についてのものかを表す列" in block and "長い記録を分けたときも各かたまりに書きます" in block
    assert "例: 対応内容、対応履歴、経過、対応メモ" in block
    assert "日付ごとに切り分けて時系列で出します（AIを使わなくても出ます）" in block
    # 説明は表と一緒に出し入れする（要約1行のときは表ごと隠れる）
    assert html.index("data-columns-table") < html.index("data-role-help")


def test_the_summary_and_the_todo_lines_use_the_new_words(app, client):
    """要約1行と、表の上の「決めてください」の行も新しい言い方にそろえる。"""
    html = _columns_html(client, CSV_TEXT, "要約.csv")
    summary = re.search(r'<p class="col-summary">(.*?)</p>', html, re.S).group(1)
    assert "設備番号＝対象" in summary and "対応内容＝経過の記録" in summary
    assert "設備" not in summary.replace("設備番号", "") and "追記ログ" not in summary

    # 対象・経過の記録の列が無い表では、その名前で「ありません」と書く
    plain = _columns_html(client, "管理No,発生日,数量\r\n" + "".join(
        f"TR-{i:03d},2026-08-{i % 28 + 1:02d},{i}\r\n" for i in range(1, 29)), "数量.csv")
    line = re.search(r'<p class="col-summary">(.*?)</p>', plain, re.S).group(1)
    assert "対象の列はありません。経過の記録の列はありません。" in line


def test_the_role_keys_in_the_saved_spec_do_not_change(app, client):
    """画面の名前を変えても、取り込み設定に入る役割の名前（entity / log）は同じ。

    保存した取り込み設定（spec_json）が読めなくなると、開き直したときに選び直しになる。
    """
    import_id = upload_csv(client, "役割.csv", CSV_TEXT)
    csv_source(client, import_id)
    save_layout(client, import_id)
    body = editor_body(client, import_id)
    roles = {col["header"]: col["role"] for col in body["columns"]}
    assert roles["設備番号"] == "entity" and roles["対応内容"] == "log"   # 画面が送るのは中の名前
    assert save_columns(client, import_id, body).status_code == 200
    wait_import_job(app, import_id)
    with app.app_context():
        spec = store.get_import(import_id)["spec_json"]
    assert '"role": "entity"' in spec and '"role": "log"' in spec


def test_two_log_columns_are_refused_with_the_new_words(app, client):
    """経過の記録は1列だけ。断るときも新しい言い方にする。"""
    import_id = upload_csv(client, "2列.csv", CSV_TEXT)
    csv_source(client, import_id)
    save_layout(client, import_id)
    body = editor_body(client, import_id)
    for col in body["columns"]:
        if col["header"] in ("現象", "対応内容"):
            col["role"] = "log"
    res = save_columns(client, import_id, body)
    assert res.status_code == 400
    assert res.get_json()["error"] == "経過の記録の列は1つだけにしてください"


# ---- 3. AI整形の段（⑤） ----------------------------------------------------------------------

def test_the_ai_step_says_what_it_does_without_the_old_word(app, client):
    """⑤の1行目は「経過の記録」の列をどうするかを書く（「追記ログ」とは言わない）。"""
    import_id = imported(app, client, "ai.csv", "AI", ai_role="log")
    data = panel(client, import_id, "ai")
    html = data["html"]
    assert "追記ログ" not in html and "AI整形の対象" not in html
    assert "④で「経過の記録」にした列「対応内容」" in html
    assert "ルールで日付・記入者ごとに分けて時系列にします" in html
    assert "AI なしで Markdown に出ます" in html
    assert "「対応の要点」" in html and "記録の種別" in html
    assert data["note"] == "経過の記録の列: 対応内容"


def test_the_ai_step_is_locked_with_the_new_words(app, client):
    """経過の記録の列が無い取り込みでは、その名前で「使いません」と出す。"""
    import_id = imported(app, client, "なし.csv", "なし")
    assert panel(client, import_id, "ai")["locked"] == "経過の記録の列がないので、この取り込みでは使いません"


def test_no_step_of_the_table_import_says_log_append(app, client):
    """表の取り込みのどの段にも「追記ログ」「AI整形の対象」を出さない。"""
    import_id = imported(app, client, "通し.csv", "通し", ai_role="log")
    pages = [client.get("/tables").get_data(as_text=True)]
    for step in ("source", "layout", "columns", "ai", "done"):
        data = panel(client, import_id, step)
        pages.append(data["html"] + (data.get("locked") or ""))
    for text in pages:
        assert "追記ログ" not in text and "AI整形の対象" not in text
