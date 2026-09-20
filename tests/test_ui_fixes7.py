"""画面まわりの直し（7巡目）: ③の移動ボタン・押しても何も起きないボタン・隠れるシートタブ。

読み取り結果（③）は帳票を縦に並べるだけで、1件ずつ確定するボタンが無い。確定した件数で
動かしていた［次の未確定の帳票へ］は、どれも未確定のままなので必ず1件目に戻っていた。
静的ファイルはサーバを通さないので、node で本物の static/review.js の一部を動かして確かめる。
"""
from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / "static"

pytestmark = pytest.mark.skipif(shutil.which("node") is None, reason="node が無い")


def _run_node(script: str):
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "check.mjs"
        path.write_text(script, encoding="utf-8")
        out = subprocess.run(["node", str(path)], capture_output=True, check=True, timeout=30, encoding="utf-8")
    return json.loads(out.stdout)


def _batch_bar_source() -> str:
    """static/review.js の「進み具合の1行」の部分だけ取り出す（本物のコードを動かす）。"""
    js = (STATIC / "review.js").read_text(encoding="utf-8")
    start = js.index("  // ---- まとめて置いたときの進み具合")
    return js[start:js.index("  function gotoBlock(")]


# 画面を持たない node で動かすための、最小限の DOM の代わり。
# 帳票は縦に GAP px ずつ並び、［次の帳票へ］で飛ぶとその帳票が画面のいちばん上に来る（＝実際の動き）
_STUB = """
const GAP = 900;
function textNode() { return { textContent: "" }; }
function button() { return { disabled: false, textContent: "" }; }
function reviewBody(text, next) {
  const handlers = [];
  return {
    addEventListener: (type, fn) => handlers.push(fn),
    querySelector: (s) => (s === "[data-batch-text]" ? text : s === "[data-next-doc]" ? next : null),
    click: () => handlers.forEach((fn) => fn({ target: { closest: () => ({}) } })),
  };
}
"""


def _script(count: int, at: int, presses: int, body: str) -> str:
    """count 件が並ぶ③の at 件目を見ているとき、［次の帳票へ］を presses 回押した動き。"""
    docs = [{"id": 10 + i, "state": "reviewing"} for i in range(count)]
    return _STUB + """
const docs = %s;
let scroll = %d * GAP;
const blocks = new Map(docs.map((d, i) => [d.id, {
  id: d.id, index: i, root: { getBoundingClientRect: () => ({ top: i * GAP - scroll }) },
}]));
const text = textNode();
const next = button();
const el = { reviewBody: reviewBody(text, next) };
const issueCount = () => 0;
const jumped = [];
function gotoBlock(id) { jumped.push(id); scroll = blocks.get(id).index * GAP; }
const api = new Function("docs", "blocks", "el", "issueCount", "gotoBlock",
  %s + "\\nreturn { updateBar };")(docs, blocks, el, issueCount, gotoBlock);
api.updateBar();
const first = { text: text.textContent, label: next.textContent, disabled: next.disabled };
for (let i = 0; i < %d; i += 1) el.reviewBody.click();
console.log(JSON.stringify({ first, jumped, after: text.textContent, label: next.textContent }));
""" % (json.dumps(docs), at, json.dumps(body), presses)


# ---- ③ の［次の帳票へ］は、押すたびに次へ進む ---------------------------------------------------

def test_the_next_button_walks_through_the_forms_instead_of_returning_to_the_first():
    """3件のうち1件目を見ているとき、押すたびに 2件目 → 3件目 → 1件目 と進む。

    まとめ取り込みの③・④には1件ずつ確定するボタンが無いので、確定済みかどうかで次を探すと
    どの帳票も「未確定」のままになり、いつも1件目へ戻っていた。
    """
    result = _run_node(_script(3, at=0, presses=3, body=_batch_bar_source()))
    assert result["jumped"] == [11, 12, 10]          # 最後まで行ったら先頭に戻る
    assert result["first"]["text"] == "3件中1件目を表示中"
    assert result["first"]["disabled"] is False and result["first"]["label"] == "次の帳票へ"
    assert result["after"] == "3件中1件目を表示中"   # 帯の数字も付いていく


def test_the_next_button_starts_from_the_form_at_the_top_of_the_screen():
    """2件目まで手でスクロールしていたら、次は3件目（1件目には戻らない）。"""
    result = _run_node(_script(3, at=1, presses=1, body=_batch_bar_source()))
    assert result["jumped"] == [12] and result["after"] == "3件中3件目を表示中"


def test_the_bar_says_where_you_are_and_the_button_wraps_at_the_end():
    """最後の帳票にいるときは［先頭の帳票へ］になる（「すべて確定しました」で止めない）。"""
    result = _run_node(_script(3, at=2, presses=1, body=_batch_bar_source()))
    assert result["first"]["text"] == "3件中3件目を表示中"
    assert result["first"]["label"] == "先頭の帳票へ" and result["first"]["disabled"] is False
    assert result["jumped"] == [10]


def test_the_next_button_is_off_when_only_one_form_is_shown():
    """③に1件しか並んでいなければ押せない（行き先が無い）。"""
    result = _run_node(_script(1, at=0, presses=2, body=_batch_bar_source()))
    assert result["first"]["disabled"] is True and result["jumped"] == []


# ---- ③に並んでいない帳票のボタンは、理由を出す ---------------------------------------------------

def test_going_to_a_form_that_was_not_read_says_why():
    """読み取れなかった帳票の［読み取り結果を見る］は、黙って何も起きないのではなく理由を出す。"""
    js = (STATIC / "review.js").read_text(encoding="utf-8")
    start = js.index("  function gotoBlock(")
    body = js[start:js.index("\n  }", start) + 4]
    script = """
const blocks = new Map();
const docs = [{ id: 7, file_name: "点検記録表.xlsx" }];
const docOf = (id) => docs.find((d) => d.id === id) || null;
const toasts = [];
const toast = (m, kind) => toasts.push([m, kind]);
const sections = { open() {} };
const api = new Function("blocks", "docOf", "toast", "sections",
  %s + "\\nreturn { gotoBlock };")(blocks, docOf, toast, sections);
api.gotoBlock(7);
console.log(JSON.stringify({ toasts }));
""" % json.dumps(body)
    result = _run_node(script)
    assert len(result["toasts"]) == 1
    message, kind = result["toasts"][0]
    assert "点検記録表.xlsx" in message and "読み取れていません" in message and kind == "err"


# ---- ダウンロードのあと、渡しに行った分も片付けの対象にする ---------------------------------------

def test_the_download_keeps_the_handed_over_ids_for_the_close_time_cleanup():
    """通信が切れても「消えました」と言い切らず、画面を閉じるときに捨てられるようにする。

    purge_after_send は本文を渡しきれなかったときは消さない（tests/test_retention.py）ので、
    渡した番号を docs から抜いたままにすると、どの片付けにも載らなくなる。
    """
    js = (STATIC / "review.js").read_text(encoding="utf-8")
    assert "let handedOver = [];" in js
    # 画面を閉じるときの送り先に、渡しに行った分も載せる
    assert "doc_ids: docs.map((d) => d.id).concat(handedOver)" in js
    # 渡しに行った分を控える（ダウンロードのボタンを押したところ）
    assert "handedOver = handedOver.concat(docs.filter(isDone).map((d) => d.id));" in js
    # 言い切らない文言にする
    assert "渡した分のデータはサーバーから消えました" not in js
    assert "ダウンロードしました。渡し終えた分はサーバーから消えます" in js


# ---- 帳票1件のとき、ファイル名の帯がシートタブを隠さない ------------------------------------------

def test_the_sticky_sheet_pane_clears_the_file_name_bar_for_a_single_form():
    """1件だけ置いたときも「元のシート」の枠をファイル名の帯の分だけ下げる。

    まとめ置き（.review-stack.is-batch）のときだけずらしてあり、いちばん普通の使い方である
    1件のときはシートのタブの上3分の1が帯に隠れて押せなかった。
    """
    css = (STATIC / "style.css").read_text(encoding="utf-8")
    single = ".review-stack .split-pane.sticky { top: calc(var(--topbar-h) + 3rem);"
    assert single in css
    # まとめ置きのずらし（より細かい指定）はそのまま残す
    batch = ".review-stack.is-batch .split-pane.sticky { top: calc(var(--topbar-h) + 5.6rem);"
    assert batch in css
    assert css.index(single) < css.index(batch)


# ---- 登録画面: 項目を削除するとき、見ている見本を送る ---------------------------------------------

def test_deleting_a_field_posts_the_sample_that_is_open():
    """見本が2つ以上あるとき、削除で1つ目の見本の表示に戻らないようにする。"""
    js = (STATIC / "form_types.js").read_text(encoding="utf-8")
    handler = js[js.index('const field = event.target.closest("[data-delete-field]");'):]
    handler = handler[:handler.index("const sampleDel")]
    assert 'data.append("sample", (root && root.dataset.sample) || "");' in handler
    assert "new FormData()" in handler and "postForm(field.dataset.deleteField, data)" in handler


# ---- 登録画面: 失敗したあとに置き直したファイルの名前を使う ---------------------------------------

def test_dropping_another_file_after_a_failure_replaces_the_auto_filled_name():
    """1回目に断られたあと別のファイルを置いたら、前のファイル名のまま登録しない。"""
    js = (STATIC / "form_types.js").read_text(encoding="utf-8")
    block = js[js.index('file.addEventListener("change"'):]
    block = block[:block.index("newForm.addEventListener")]
    # 自動で入れた名前かどうかを覚えておき、そのときだけ置き直したファイルの名前に入れ替える
    assert 'name.dataset.autofill' in block
    assert "if (!name.value.trim() || name.value === name.dataset.autofill) {" in block
    assert 'name.dataset.autofill = ""' in block      # ファイルを外したら覚えも消す
