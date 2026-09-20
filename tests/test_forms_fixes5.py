"""帳票まわりの修正（5巡目の確認で見つかった不具合）の回帰テスト。"""
from tests.test_forms_fixes import _confirmed_doc, _review_html


# ---- R5F-2: 読み取り結果の入力欄で Enter を押しても、何も送信されない ----------------------------------

def test_enter_in_the_review_form_submits_nothing(app):
    doc_id = _confirmed_doc(app, "1.xlsx")
    html = _review_html(app, doc_id)
    form = html[html.index('id="reviewForm"'):html.index("</form>", html.index('id="reviewForm"'))]
    # 暗黙の送信（Enter）を止める。送信ボタンも action も持たせない（値は途中保存で送る）
    assert '<form id="reviewForm" onsubmit="return false">' in html
    assert 'type="submit"' not in form and "formaction" not in form and "action=" not in form


# ---- R5-MD-5: タイトル項目が設備だけのとき、出典にもタイトルに足した番号を書く ------------------------

def test_source_line_names_the_work_number_when_the_title_is_equipment_only():
    from export.formats import build_markdown
    from tests.test_forms_fixes import _inspection

    doc = {"id": 1, "file_name": "点検.xlsx", "file_hash": "0" * 64}
    md = build_markdown(doc, _inspection("W-0101", "2026-04-01"))
    assert md.split("\n", 1)[0].endswith("｜W-0101")
    assert "- 出典: 点検.xlsx（作業No. W-0101）" in md.split("\n")
    # 番号が無ければ日付、どちらも無ければ今までどおり元ファイル名だけ
    assert "- 出典: 点検.xlsx（作業日 2026-04-08）" in build_markdown(doc, _inspection("", "2026-04-08")).split("\n")
    assert "- 出典: 点検.xlsx" in build_markdown(doc, _inspection("", None)).split("\n")
