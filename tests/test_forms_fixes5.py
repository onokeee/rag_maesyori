"""帳票まわりの修正（5巡目の確認で見つかった不具合）の回帰テスト。"""
import re

from tests.test_forms_fixes import _confirmed_doc


# ---- R5F-2: 確認画面の入力欄で Enter を押しても「AIで空欄を探す」や確定を送らない ----------------------

def test_enter_in_the_review_form_does_not_press_the_ai_button(app, client, monkeypatch):
    from services import llm

    monkeypatch.setattr(llm, "is_configured", lambda: True)  # AIボタンが出る状態
    doc_id = _confirmed_doc(app, "1.xlsx")
    page = client.get(f"/forms/{doc_id}/review").get_data(as_text=True)
    form = page[page.index('id="reviewForm"'):page.index("</form>", page.index('id="reviewForm"'))]
    assert "AIで空欄を探す" in form
    # 暗黙の送信（Enter）はフォームの最初の送信ボタンを押す。それが無効なら送信されない
    first = re.search(r"<button\b[^>]*>", form).group(0)
    assert 'type="submit"' in first and "disabled" in first and "formaction" not in first


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
