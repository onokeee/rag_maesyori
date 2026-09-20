"""6巡目の「別の担当の領域だったので残した」修正の確認。

- 試し実行の片付けが、まだある取り込みが払った応答まで消さないこと（core.purge と同じ持ち主の確認）
- 列の対応づけの保存エラーを1件だけにしないこと（static/app.js の postJson → static/tables.js の表示）
- AI整形の［見積もる］を連打できないこと（static/tables.js）
"""
from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

from aiproc.runner import _ensure_trial_import
from core.jobs import JobError
from models import database as db
from tables import store

STATIC = Path(__file__).resolve().parents[1] / "static"


# ---- 試し実行の片付け: 持ち主のいる応答は消さない ---------------------------------------------------

def _llm_call(conn, cache_key: str, import_id: int | None = None) -> None:
    conn.execute("INSERT INTO llm_calls (cache_key, raw_text, created_at, import_id) "
                 "VALUES (?, 'あ', '2026-09-19', ?)", (cache_key, import_id))


def _llm_keys(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute("SELECT cache_key FROM llm_calls"))


def test_a_trial_on_a_deleted_import_keeps_a_response_another_import_paid_for(app):
    """試し実行の間に取り込みが消えても、まだある別の取り込みが払った応答は消さない。

    試し実行は同じ文面ならほかの取り込みの応答をキャッシュとして引く。巻き添えで消すと、払った
    取り込みの再開・再実行で同じ応答をもう一度買うことになる（design.md 3.3）。
    持ち主のいない応答（この試し実行が払った分）は、これまでどおりすぐ消す。
    """
    with app.app_context():
        alive = store.create_import("作業中.csv", "1" * 64, "tables/other.csv", {"kind": "csv"})
        conn = db.get_db()
        _llm_call(conn, "K-paid-by-alive", alive)   # まだある取り込みが払った（誰も参照していない）
        _llm_call(conn, "K-no-owner", None)         # 持ち主のいない分
        conn.commit()

        gone = alive + 1000   # もう無い取り込み（試し実行の最中に消された）
        with pytest.raises(JobError):
            _ensure_trial_import(gone, keys=["K-paid-by-alive", "K-no-owner"])

        assert _llm_keys(db.get_db()) == ["K-paid-by-alive"]




# ---- 画面側（node で確かめる） ------------------------------------------------------------------
# 画面は3つの1枚ページになり、段の中身は fetch で入れ替わる（2026-09-20 の作り直し）。
# static/tables.js は window.ragFetch（static/app.js）を通してサーバとやりとりするので、
# ragFetch は本物を切り出して使い、DOM だけを最小限の代わりで置き換える。

def _run_node(script: str):
    """node -e はコマンドラインの長さに上限があるので、ファイルに書いてから動かす。"""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "check.mjs"
        path.write_text(script, encoding="utf-8")
        out = subprocess.run(["node", str(path)], capture_output=True, check=True, timeout=30, encoding="utf-8")
    return json.loads(out.stdout)


def _rag_fetch() -> str:
    """static/app.js の ragFetch の部分だけ取り出す（断りの errors をそのまま渡すところ）。"""
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    return js[js.index("async function ragFetch("):js.index("// ---- ragSections")]


def _tables_js() -> str:
    return (STATIC / "tables.js").read_text(encoding="utf-8")


# 画面を持たない node で static/tables.js を動かすための、最小限の DOM の代わり
_DOM_STUB = """
class N {
  constructor(tag) {
    this.tagName = tag; this.children = []; this.dataset = {}; this.handlers = {};
    this._text = ""; this.disabled = false; this.className = ""; this.value = ""; this.hidden = false;
    this.sel = new Set(); this.parent = null; this.one = {}; this.many = {};
    this.classList = { add() {}, remove() {}, toggle() {}, contains: () => false };
  }
  get textContent() { return this.children.length ? this.children.map((c) => c.textContent).join("") : this._text; }
  set textContent(v) { this._text = v; this.children = []; }
  setAttribute(k, v) { this.dataset[k] = v; }
  _adopt(kids) { kids.forEach((k) => { if (k instanceof N) k.parent = this; }); }
  append(...kids) { this._adopt(kids); this.children.push(...kids); }
  replaceChildren(...kids) { this._adopt(kids); this.children = kids; this._text = ""; }
  addEventListener(type, fn) { (this.handlers[type] = this.handlers[type] || []).push(fn); }
  fire(type, event) { return (this.handlers[type] || []).map((fn) => fn(event)); }
  querySelector(s) { return this.one[s] || null; }
  querySelectorAll(s) { return this.many[s] || []; }
  closest(s) { let n = this; while (n) { if (n.sel.has(s)) return n; n = n.parent; } return null; }
}
globalThis.Node = N;
globalThis.toasts = [];
globalThis.window = {
  location: {},
  App: { toast: (m) => globalThis.toasts.push(m), bindProgressBox() {} },
  ragSections: { el: () => null, body: () => new N("div"), open() {}, close() {}, done() {}, reset() {},
                 working() {} },
};
globalThis.document = {
  createElement: (tag) => new N(tag),
  createTextNode: (t) => { const n = new N("#text"); n.textContent = t; return n; },
  querySelector: (s) => globalThis.roots[s] || null,
  querySelectorAll: (s) => [],
  addEventListener() {},
};
globalThis.roots = {};
function node(tag, selectors = [], dataset = {}) {
  const n = new N(tag);
  selectors.forEach((s) => n.sel.add(s));
  Object.assign(n.dataset, dataset);
  return n;
}
"""


def _columns_editor_script(fetch_body: str) -> str:
    """列の対応づけの段だけを置いた画面で、［この対応づけで読み込む］を1回押す。"""
    return _DOM_STUB + f"""
{fetch_body}
{_rag_fetch()}
globalThis.window.ragFetch = ragFetch;
const errors = node("ul");
const editor = node("div", ["[data-columns-editor]"],
                     {{ saveUrl: "/tables/imports/1/columns", headerRowsCount: "1" }});
editor.one = {{ "[data-editor-errors]": errors }};
editor.many = {{ "[data-setting]": [], "tr[data-col]": [] }};
const saveButton = node("button", ["[data-editor-save]"]);
saveButton.parent = editor;
const page = node("div", ["[data-tables-page]"], {{ newUrl: "/tables/new" }});
page.one = {{ "[data-columns-editor]": editor }};
globalThis.roots["[data-tables-page]"] = page;
{_tables_js()}
(async () => {{
  await Promise.all(page.fire("click", {{ target: saveButton }}));
  process.stdout.write(JSON.stringify({{
    shown: errors.children.map((c) => c.textContent),
    toast: globalThis.toasts[globalThis.toasts.length - 1],
    reenabled: saveButton.disabled === false,
  }}));
}})();
"""


def _json_response(status: int, body: str) -> str:
    return f"""
globalThis.fetch = async () => ({{
  ok: {str(status == 200).lower()}, status: {status},
  headers: {{ get: () => "application/json" }},
  json: async () => ({body}),
}});
"""


@pytest.mark.skipif(shutil.which("node") is None, reason="node がない")
def test_every_save_error_of_the_column_editor_is_shown():
    """列の対応づけの保存が複数の理由で断られたら、全部並べる（1件ずつ直させない）。"""
    body = ('{ error: "キーの列を1つ選んでください", '
            'errors: ["キーの列を1つ選んでください", "日付の列がありません", "見出しが重なっています"] }')
    got = _run_node(_columns_editor_script(_json_response(400, body)))
    assert got["shown"] == ["キーの列を1つ選んでください", "日付の列がありません", "見出しが重なっています"]
    assert got["toast"] == "3件の問題があります"
    assert got["reenabled"]   # 直して保存し直せる


@pytest.mark.skipif(shutil.which("node") is None, reason="node がない")
def test_the_column_editor_shows_a_single_error_as_before():
    """理由が1つだけ（errors を返さない断り方）なら、今までどおりその1件を出す。"""
    got = _run_node(_columns_editor_script(_json_response(409, '{ error: "ほかの処理の最中です" }')))
    assert got["shown"] == ["ほかの処理の最中です"]
    assert got["toast"] == "ほかの処理の最中です"


@pytest.mark.skipif(shutil.which("node") is None, reason="node がない")
@pytest.mark.parametrize("ok", [True, False])
def test_the_estimate_button_cannot_be_pressed_twice(ok):
    """［見積もる］は答えが返るまで押せない（連打すると同じ計算が何度も走る）。終わったら押せる。"""
    script = _DOM_STUB + f"""
let release;
const pending = new Promise((r) => {{ release = r; }});
globalThis.fetch = async () => {{
  await pending;
  return {{
    ok: {"true" if ok else "false"}, status: {"200" if ok else "400"},
    headers: {{ get: () => "application/json" }},
    json: async () => ({{ ai_rows: 3, calls: 3, cached: 0, rule_only_rows: 1, tokens_in: 10, tokens_out: 5,
                        minutes: 1, error: "見積もれません" }}),
  }};
}};
{_rag_fetch()}
globalThis.window.ragFetch = ragFetch;
const button = node("button", ["[data-estimate-run]"]);
const out = node("p");
const aiPage = node("div", ["[data-ai-page]"], {{ estimateUrl: "/tables/imports/1/ai/estimate" }});
aiPage.one = {{ "[data-estimate-run]": button, "[data-estimate]": out }};
button.parent = aiPage;
const page = node("div", ["[data-tables-page]"], {{ newUrl: "/tables/new" }});
page.one = {{ "[data-ai-page]": aiPage }};
globalThis.roots["[data-tables-page]"] = page;
{_tables_js()}
(async () => {{
  const running = Promise.all(page.fire("click", {{ target: button }}));
  await new Promise((r) => setTimeout(r, 0));
  const whileRunning = button.disabled;      // 答えを待っている間
  release();
  await running;
  process.stdout.write(JSON.stringify({{ whileRunning, after: button.disabled }}));
}})();
"""
    got = _run_node(script)
    assert got["whileRunning"] is True    # 待っている間は押せない
    assert got["after"] is False          # 成功でも失敗でも、終われば押せる
