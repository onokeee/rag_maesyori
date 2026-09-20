"""画面まわりの不具合修正（6巡目）の確認。

画面が3つになり、取り込みの一覧は無くなった（2026-09-20 の作り直し）。
［削除］はその取り込みの段の中にしか無いので、「処理中は削除させない」は送り先で確かめる。
"""
from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest

from tables import store


# ---- UX6-4: 処理中（読み込み中・確定処理中）の取り込みは削除できない --------------------------------

@pytest.mark.parametrize("status", ["reading", "confirming"])
def test_a_busy_import_cannot_be_deleted(app, client, status):
    with app.app_context():
        import_id = store.create_import("一覧.xlsx", "hash", "tables/一覧.xlsx", {"sheet": "一覧"})
        store.update_import(import_id, status=status)
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 400
    assert "処理中の取り込みは削除できません" in res.get_json()["error"]
    with app.app_context():
        assert store.get_import(import_id) is not None   # 消えていない


@pytest.mark.parametrize("status", ["uploaded", "preview", "failed"])
def test_an_idle_import_can_be_deleted(app, client, status):
    with app.app_context():
        import_id = store.create_import("一覧.xlsx", "hash", "tables/一覧.xlsx", {"sheet": "一覧"})
        store.update_import(import_id, status=status)
    res = client.post(f"/tables/imports/{import_id}/delete")
    assert res.status_code == 200 and res.get_json()["ok"] is True
    with app.app_context():
        assert store.get_import(import_id) is None


# ---- R6-C3: 段のポーリングは、ジョブが消えた（404）ら止まって画面を読み直す ----------------------------

def _run_node(script: str):
    """node -e はコマンドラインの長さに上限があるので、ファイルに書いてから動かす。"""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "check.mjs"
        path.write_text(script, encoding="utf-8")
        out = subprocess.run(["node", str(path)], capture_output=True, check=True, timeout=30, encoding="utf-8")
    return json.loads(out.stdout)


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
    got = _run_node(script)
    # 404 は一度で止めて読み直す
    assert got["gone"] == {"calls": 1, "reloads": 1}
    # ほかの通信エラーは今までどおり間隔を広げて再試行する（読み直さない）
    assert got["broken"]["calls"] > 1 and got["broken"]["reloads"] == 0
