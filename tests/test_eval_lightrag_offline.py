"""scripts/eval/lightrag_offline_eval.py の計測補助（LightRAG を使わない部分）。"""
from core.naming import LIGHTRAG_HINT_RECORDS, md_filename
from scripts.eval import lightrag_offline_eval as ev

RECORDS_MD = (
    "# 一覧 2024年5月の記録\n\n- データ種別: 一覧（1行＝1件）の記録\n\n"
    "## 【TR-001】CMP 1号機（CMP-101）異音｜2024-05-01\n- 管理No: TR-001\n- 設備: CMP 1号機（CMP-101）\n"
    "- 発生日: 2024-05-01\n\n"
    "## 【TR-002】CVD 2号機（CVD-202）停止｜2024-05-02\n- 管理No: TR-002\n- 設備: CVD 2号機（CVD-202）\n"
    "- 発生日: 2024-05-02\n"
)
MANIFEST = [{"name": "a.md", "kind": "records", "records": [
    {"id": "TR-001", "entities": ["CMP-101"], "date": "2024-05-01"},
    {"id": "TR-002", "entities": ["CVD-202"], "date": "2024-05-02"}]}]


class _Tok:
    def encode(self, text):
        return list(text)


def _chunks(parts):
    return [{"content": p, "tokens": len(p), "chunk_order_index": i} for i, p in enumerate(parts)]


def test_record_blocks_and_entities():
    blocks = ev._record_blocks(RECORDS_MD)
    assert len(blocks) == 2 and blocks[1].startswith("## 【TR-002】")
    est = ev.estimate_entities(blocks[0])
    assert est["records"] == 1 and est["est"] >= 3  # 記録＋TR-001＋CMP-101＋設備名


def test_measure_route_counts_cut_records_and_missing_context():
    whole = ev.measure_route(MANIFEST, {"a.md": RECORDS_MD}, lambda text, name: _chunks(text.split("\n\n")), _Tok())
    assert whole["records_total"] == 2 and whole["records_cut"] == 0
    assert whole["chunks"] == 4 and whole["chunks_without_record_id"] == 2  # 見出しとファイル説明のチャンク

    def cut_in_middle(text, name):
        i = text.index("- 設備: CVD")
        return _chunks([text[:i], text[i:]])
    cut = ev.measure_route(MANIFEST, {"a.md": RECORDS_MD}, cut_in_middle, _Tok())
    assert cut["records_cut"] == 1
    assert cut["chunks_without_record_id"] == 1  # 2つ目のチャンクに管理Noがない
    assert cut["llm_calls_est"] == 4


def test_per_record_file_counts_as_one_record():
    manifest = [{"name": "b.md", "kind": "records", "records": [{"id": "TR-001", "entities": [], "date": ""}]}]
    text = "# 【TR-001】異音\n\n- 管理No: TR-001\n"
    res = ev.measure_route(manifest, {"b.md": text}, lambda t, n: _chunks(t.split("\n\n")), _Tok())
    assert res["records_total"] == 1 and res["records_cut"] == 1 and res["headings_only_chunks"] == 1


def test_hint_regex_matches_app_hint_only():
    hinted = md_filename(["トラブル対応一覧", "2024-05"], LIGHTRAG_HINT_RECORDS)
    m = ev.HINT_RE.search(hinted)
    assert m and m.group(1) == LIGHTRAG_HINT_RECORDS
    assert not ev.HINT_RE.search(md_filename(["報告書.[legacy-F]", "A"]))
    assert not ev.FORBIDDEN.search(hinted)
