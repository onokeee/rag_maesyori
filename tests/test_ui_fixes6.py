"""画面まわりの不具合修正（6巡目）の確認。"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
from pathlib import Path

import pytest

from tables import store
from views import TABLE_IMPORT_ACTIVE


def _delete_form_shown(html: str, import_id: int) -> bool:
    return f'action="/tables/imports/{import_id}/delete"' in html


# ---- UX6-4: 処理中（読み込み中・確定処理中）の取り込みには、どの画面でも［削除］を出さない --------------------

@pytest.mark.parametrize("url", ["/", "/tables/new"])
def test_busy_imports_have_no_delete_button_on_any_screen(app, client, url):
    with app.app_context():
        import_id = store.create_import("一覧.xlsx", "hash", "tables/一覧.xlsx", {"sheet": "一覧"})

    for status in ("reading", "confirming"):
        with app.app_context():
            store.update_import(import_id, status=status)
        html = client.get(url).get_data(as_text=True)
        assert not _delete_form_shown(html, import_id), f"{url} / {status}"

    # 処理中でなければ、どちらの画面にも［削除］は出る
    for status in ("uploaded", "preview", "failed"):
        with app.app_context():
            store.update_import(import_id, status=status)
        html = client.get(url).get_data(as_text=True)
        assert _delete_form_shown(html, import_id), f"{url} / {status}"


# ---- UX6-5: ホームの「作業中の一覧表」の副題が、そこに出る状態をすべて言っている -------------------------------

def test_home_subtitle_names_every_working_table_state(client):
    html = client.get("/").get_data(as_text=True)
    subtitle = re.search(r"作業中の一覧表</h2>\s*<p class=\"card-subtitle\">([^<]*)</p>", html).group(1)
    for word in ("読み込み前", "読み込み中", "確認中", "確定処理中", "失敗"):
        assert word in subtitle
    # 状態が増えたら副題も直す（TABLE_IMPORT_ACTIVE と数を合わせる）
    assert len(subtitle.replace("のもの", "").split("・")) == len(TABLE_IMPORT_ACTIVE)


# ---- R6-C3: 待ち画面のポーリングは、ジョブが消えた（404）ら止まって画面を読み直す -------------------------------

@pytest.mark.skipif(shutil.which("node") is None, reason="node がない")
def test_polling_stops_and_reloads_when_the_job_is_gone():
    js = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")
    start = js.index("async function getJson(")
    end = js.index("const JOB_LABELS")
    script = """
let calls = 0, reloads = 0, status = 404;
globalThis.fetch = async () => {
  calls++;
  return { ok: status === 200, status, json: async () => ({ error: "このジョブはありません" }) };
};
globalThis.window = { location: { reload() { reloads++; } } };
""" + js[start:end] + """
(async () => {
  const run = async (code) => {
    calls = 0; reloads = 0; status = code;
    const poll = pollJob("/api/jobs/4", () => {}, { interval: 5, maxInterval: 20 });
    await new Promise((r) => setTimeout(r, 150));
    poll.stop();
    return { calls, reloads };
  };
  const gone = await run(404);
  const broken = await run(500);
  process.stdout.write(JSON.stringify({ gone, broken }));
})();
"""
    out = subprocess.run(["node", "-e", script], capture_output=True, check=True, timeout=30, encoding="utf-8")
    got = json.loads(out.stdout)
    # 404 は一度で止めて読み直す
    assert got["gone"] == {"calls": 1, "reloads": 1}
    # ほかの通信エラーは今までどおり間隔を広げて再試行する（読み直さない）
    assert got["broken"]["calls"] > 1 and got["broken"]["reloads"] == 0
