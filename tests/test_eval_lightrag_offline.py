"""scripts/eval/lightrag_offline_eval.py の計測補助（LightRAG を使わない部分）。"""
from core.naming import md_filename
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


def test_app_filenames_never_look_like_a_parser_hint():
    """ファイル名にヒント（.[...]）は付けない。元の値にヒントらしい書き方があっても名前には残さない。"""
    name = md_filename(["トラブル対応一覧", "2024-05"])
    assert not ev.HINT_RE.search(name) and not ev.FORBIDDEN.search(name)
    assert not ev.HINT_RE.search(md_filename(["報告書.[legacy-F]", "A"]))


def test_split_records_keep_the_metadata_of_the_record_they_came_from():
    """「（続きn/m）」に分かれた記録の各部分も、元の記録の管理No・日付で照合する。"""
    from tables.markdown import record_title
    from tables.spec import spec_from_dict

    spec = spec_from_dict({"name": "x", "columns": [
        {"key": "record_no", "display": "管理No", "type": "code", "role": "key"},
        {"key": "occurred_at", "display": "発生日", "type": "date", "role": "date"}],
        "record": {"key": ["record_no"]}, "period": {"date_column": "occurred_at"}})
    recs = [{"values": {"record_no": "TR-001", "occurred_at": "2024-05-01"}},
            {"values": {"record_no": "TR-002", "occurred_at": "2024-05-02"}}]
    titles = [record_title(r["values"], spec) for r in recs]
    text = (f"# 一覧\n\n## {titles[0]}（1/2）\n- 管理No: TR-001\n\n"
            f"## {titles[0]}（続き2/2）\n- 管理No: TR-001\n\n## {titles[1]}\n- 管理No: TR-002\n")
    files = [{"name": "a.md", "kind": "records", "text": text}]
    ev._attach_table_meta(files, recs, spec)
    assert [m["id"] for m in files[0]["records"]] == ["TR-001", "TR-001", "TR-002"]
