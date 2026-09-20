"""アプリが実際に出す Markdown を、LightRAG 1.5.7 のチャンク分割にオフラインでかけて投入適性を測る（LLM・API は呼ばない）。

    # 生成（プロジェクトの venv）→ 計測（LightRAG の venv を子プロセスで起動）をまとめて実行
    .venv\\Scripts\\python.exe scripts/eval/lightrag_offline_eval.py --lightrag <lightrag_src> --out <出力先>
    # 生成だけ / 計測だけ / 表にまとめるだけ（metrics*.json → metrics_all.json と Markdown 表）
    .venv\\Scripts\\python.exe scripts/eval/lightrag_offline_eval.py generate --out <出力先>
    <lightrag_src>\\venv\\Scripts\\python.exe scripts/eval/lightrag_offline_eval.py measure --out <出力先> --lightrag <lightrag_src>

    python scripts/eval/lightrag_offline_eval.py summary --out <出力先>
    経路は --routes F_<chunk>_<overlap>,R_<chunk>_<overlap>,P_<size> で指定（既定 F_1200_100,F_600_50,P_2000）。
    生成する variant は --variants month,entity_month,per_record,single で絞れる。
    大きい variant は --only と --metrics を分けて並列に計測できる。

<lightrag_src> は LightRAG 1.5.7 の venv（venv/）と tiktoken のキャッシュ（tiktoken_cache/）を持つフォルダ。
生成物: <出力先>/md/<variant>/*.md と manifest.json（ファイルごとの記録ID・設備・日付）、計測結果: <出力先>/metrics.json
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

SAMPLES = PROJECT_ROOT / "samples"
# 一覧表: (ファイル名, 取り込み設定名, ログ役割にする列見出し)
TABLES = {
    "T1": ("T1_トラブル対応一覧_2023-2026.xlsx", "トラブル対応一覧", "対応内容"),
    "T2": ("T2_設備故障履歴_システム出力.csv", "設備故障履歴", None),
    "T5": ("T5_是正処置管理台帳.csv", "是正処置管理台帳", None),
}
DATE_RE = re.compile(r"(?:19|20)\d{2}[-/年.]\s?\d{1,2}[-/月.]\s?\d{1,2}|(?:19|20)\d{2}年\d{1,2}月|R0?\d[./]\d{1,2}[./]\d{1,2}")


# =============================================================================
# 生成（プロジェクトの venv で実行）
# =============================================================================
def _write_variant(out: Path, variant: str, files: list[dict]) -> None:
    d = out / "md" / variant
    shutil.rmtree(d, ignore_errors=True)
    d.mkdir(parents=True)
    manifest = []
    for f in files:
        (d / f["name"]).write_bytes(f["text"].encode("utf-8"))
        manifest.append({k: v for k, v in f.items() if k != "text"})
    (d / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"  {variant}: {len(files)} files, {sum(len(f['text']) for f in files):,} chars", flush=True)


def generate_forms(out: Path) -> None:
    """帳票: 見本3件で種類を作り（evaluate_forms と同じ）、30ファイルを抽出して build_markdown する。"""
    from excel.extractor import extract_document
    from excel.workbook import load_workbook_info
    from export.formats import build_markdown, markdown_filename
    from pattern.builder import suggest_rows
    from pattern.forms import rows_to_pattern
    from pattern.matcher import match_pattern
    from scripts.samples.evaluate_forms import pick_samples

    for folder in sorted(p for p in (SAMPLES / "forms").iterdir() if (p / "_expected.jsonl").exists()):
        entries = [json.loads(x) for x in (folder / "_expected.jsonl").open(encoding="utf-8") if x.strip()]
        infos = {e["file"]: load_workbook_info(folder / e["file"]) for e in entries}
        sheet_rows, field_rows = suggest_rows([infos[e["file"]] for e in pick_samples(entries, 3)])
        type_name = folder.name.split("_", 1)[1]  # 利用者が付ける種類名の想定（F1_ などの番号は外す）
        pattern = rows_to_pattern(1, {"name": type_name, "version": "1"}, sheet_rows, field_rows)
        files, used = [], Counter()
        for i, e in enumerate(entries, start=1):
            info = infos[e["file"]]
            extraction = extract_document(info, pattern, match_pattern(info, pattern).sheet_names)
            doc = {"id": i, "file_name": e["file"], "file_hash": hashlib.sha256((folder / e["file"]).read_bytes()).hexdigest()}
            text = build_markdown(doc, extraction)
            name = markdown_filename(doc, extraction)
            used[name] += 1
            if used[name] > 1:  # 同名（上書き）の数を数えるため別名で保存する
                name = name[:-3] + f"_dup{used[name]}.md"
            vals = e.get("values") or {}
            files.append({"name": name, "text": text, "kind": "form", "records": [{
                "id": str(vals.get("report_id") or ""),
                "entities": [str(vals[k]) for k in ("equipment_id", "equipment_name") if vals.get(k)],
                "date": str(vals.get("occurred_at") or "")[:10]}]})
        _write_variant(out, f"forms_{folder.name.split('_')[0]}", files)


def _table_spec(code: str):
    from tables.detect import guess_layout, sample_data_rows
    from tables.mapping import suggest_columns
    from tables.source import open_source
    from tables.spec import LogStageSpec, spec_from_suggestions

    file_name, name, log_header = TABLES[code]
    path = SAMPLES / "tables" / file_name
    source = open_source(path, path.name)
    sheets = [s for s in source.sheets() if not s.hidden] or source.sheets()
    sheet = sheets[0].name
    layout = guess_layout(source, sheet)
    suggestions = suggest_columns(layout.headers, sample_data_rows(source, sheet, layout, 200))
    for s in suggestions:
        if log_header and s.header == log_header:
            s.role, s.type, s.md = "log", "text", "body"
        elif s.role == "log" and not (log_header and s.header == log_header):
            s.role = "text"  # 指定した列以外はログ扱いにしない（画面の既定操作に合わせる）
    spec = spec_from_suggestions(name, layout, suggestions, {"group_by": "month"})
    if log_header and spec.log_stage is None:
        spec.log_stage = LogStageSpec(column=next(c for c in spec.columns if c.role == "log").key)
    return source, sheet, layout, spec


def _record_meta(rec: dict | None, spec) -> dict:
    from tables.records import entity_columns, entity_display

    if rec is None:
        return {"id": "", "entities": [], "date": ""}
    values = rec.get("values") or {}
    key_col = spec.first_role("key")
    entity, label = entity_columns(spec)
    names = [str(values.get(c.key)) for c in (entity, label) if c is not None and values.get(c.key)]
    if entity is not None:
        # md の本文は「設備名（設備番号）」に分けて書くので、分けたあとの番号・名前も照合の対象にする
        # （元のセルが「ETC-302(OXIDEエッチャ 2号機)」のような台帳で、本文に設備が無いと誤判定しないため）
        names += [x for x in entity_display(values, spec)[:2] if x]
    return {"id": str(values.get(key_col.key) or "") if key_col else "",
            "entities": sorted(set(names)),
            "date": str(values.get(spec.date_key) or "")[:10]}


PART_SUFFIX = re.compile(r"（(?:続き)?\d+/\d+）$")   # 大きい記録を分けた見出しの「（1/3）」「（続き2/3）」


def _attach_table_meta(files: list[dict], recs: list[dict], spec) -> None:
    """記録ファイルの ## 見出しごとに、同じ見出しを持つ記録のメタを順に割り当てる。

    大きい記録は「（続きn/m）」に分かれて同じ見出しが続くので、その分は同じ記録のメタにする。
    """
    from tables.markdown import record_title

    by_title: dict[str, list[dict]] = defaultdict(list)
    for r in recs:
        by_title[record_title(r.get("values") or {}, spec)].append(r)
    cursor: Counter = Counter()
    for f in files:
        meta = []
        if f["kind"] == "records":
            for line in f["text"].split("\n"):
                if not line.startswith("## "):
                    continue
                heading = line[3:]
                title = PART_SUFFIX.sub("", heading)
                if heading.startswith(title + "（続き"):
                    index = max(0, cursor[title] - 1)      # 続きの部分は1つ前（同じ記録）のまま
                else:
                    index = cursor[title]
                    cursor[title] += 1
                cand = by_title.get(title) or []
                meta.append(_record_meta(cand[index % len(cand)] if cand else None, spec))
        f["records"] = meta


def generate_tables(out: Path, codes: list[str], variants: list[str] | None = None) -> None:
    from core.naming import md_filename
    from tables.markdown import _Block, join_file, people_index_for, record_block, render_all
    from tables.normalize import read_records

    want = set(variants or ("month", "entity_month", "per_record", "single"))

    for code in codes:
        t0 = time.perf_counter()
        source, sheet, layout, spec = _table_spec(code)
        records, issues, _stats = read_records(source, {"sheet": sheet, "file_name": TABLES[code][0]}, layout, spec)
        recs = [r.to_dict() for r in records]
        print(f"  {code}: {len(recs):,} records, key={getattr(spec.first_role('key'), 'display', None)}, "
              f"entity={getattr(spec.first_role('entity'), 'display', None)}, date={spec.date_key}, "
              f"log={spec.log_stage.column if spec.log_stage else None} ({time.perf_counter() - t0:.0f}s)", flush=True)
        (out / "md").mkdir(parents=True, exist_ok=True)
        (out / "md" / f"{code}_spec.json").write_text(json.dumps(
            {"records": len(recs), "issues": len(issues),
             "columns": [[c.key, c.display, c.type, c.role, c.md] for c in spec.columns]},
            ensure_ascii=False, indent=1), encoding="utf-8")

        for group_by in ("month", "entity_month"):
            if group_by not in want:
                continue
            spec.markdown["group_by"] = group_by
            files = [{"name": f.name, "text": f.text, "kind": f.kind} for f in render_all(spec, recs, {})]
            _attach_table_meta(files, recs, spec)
            _write_variant(out, f"{code}_{group_by}", files)

        if not ({"per_record", "single"} & want):
            continue
        # 1件1ファイル（シミュレーション: 記録ブロックの ## を # にして単独ファイルにする）
        people = people_index_for(spec, recs) if spec.log_stage else None
        prefix = spec.file_prefix
        files, names, blocks = [], Counter(), []
        for rec in recs:
            block = record_block(rec, spec, {}, people)
            blocks.append(block)
            meta = _record_meta(rec, spec)
            parts = [prefix, meta["id"] or rec["key"]]
            names[md_filename(parts)] += 1
            if names[md_filename(parts)] > 1:
                parts.append(str(names[md_filename(parts)]))
            text = join_file([_Block("#" + block[0][2:], [f"- データ種別: {spec.name}（1行＝1件）の記録"] + block[1:])])
            files.append({"name": md_filename(parts), "text": text, "kind": "records", "records": [meta]})
        if "per_record" in want:
            _write_variant(out, f"{code}_per_record", files)
        if "single" not in want:
            continue
        # 全件1ファイル
        header = _Block(f"# {spec.name} 全記録", [f"- データ種別: {spec.name}（1行＝1件）の記録",
                                                 f"- このファイルの記録: {len(recs):,}件"])
        _write_variant(out, f"{code}_single", [{
            "name": md_filename([prefix, "全件"]), "text": join_file([header] + blocks), "kind": "records",
            "records": [_record_meta(r, spec) for r in recs]}])


# =============================================================================
# 計測（LightRAG 1.5.7 の venv で実行）
# =============================================================================
ROUTES = {
    "F_1200_100": "legacy-F（LIGHTRAG_PARSER 未設定時の既定。chunk 1200 / overlap 100）",
    "F_600_50": "legacy-F（狭い窓の場合。chunk 600 / overlap 50）",
    "P_2000": "native-P（CHUNK_P_SIZE 2000）",
}
ENTITY_LIMIT = 40
# 抽出1チャンクあたりの LLM 入力の上乗せ（prompt.py を tiktoken で数えた値。テキスト形式・既定の例）
PROMPT_SYSTEM_TOKENS = 1309 + 58
PROMPT_USER_TOKENS = 268
PROMPT_GLEANING_TOKENS = 405
ASSUMED_OUTPUT_TOKENS = 400  # 1回の抽出出力の仮定（実測ではない）
CODE_RE = re.compile(r"(?<![A-Za-z0-9])[A-Z][A-Z0-9]{0,6}(?:-[A-Z0-9]{1,8}){1,3}(?![A-Za-z0-9])")
NAME_BULLET_RE = re.compile(r"^- ([^:：\n]{1,20})[:：]\s*(.{1,60})$", re.M)
PART_RE = re.compile(r"[（(]([A-Z0-9][A-Z0-9-]{3,})[)）]\s*[×x]\s*\d+")
HINT_RE = re.compile(r"\.\[([^\]]*)\](\.[^.]+)$")  # routing.py:62 と同じ
FORBIDDEN = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
# 固有名になりやすい箇条書きの項目名（分類値・設備・部署など）
NAME_LABELS = ("設備", "工程", "ライン", "故障区分", "部位", "原因区分", "対象設備", "起票部署", "部署", "発見区分",
               "重要度", "状態", "判定", "記録区分", "分類", "区分")


def estimate_entities(chunk: str) -> dict:
    """LLM を使わない粗い推定（下限寄り）: 記録（## ）＋ 識別コード ＋ 固有名になりやすい項目値 の異なり数。"""
    records = len(re.findall(r"^#{1,2} ", chunk, flags=re.M))
    codes = set(CODE_RE.findall(chunk)) | set(PART_RE.findall(chunk))
    names = {m.group(2).strip() for m in NAME_BULLET_RE.finditer(chunk) if any(k in m.group(1) for k in NAME_LABELS)}
    return {"records": len(re.findall(r"^## ", chunk, flags=re.M)), "est": records + len(codes) + len(names)}


def _pct(values: list, q: float):
    if not values:
        return 0
    s = sorted(values)
    return s[min(len(s) - 1, int(round(q * (len(s) - 1))))]


def _record_blocks(text: str) -> list[str]:
    return [p for p in re.split(r"(?m)^(?=## )", text) if p.startswith("## ")]


def make_chunkers(tok, routes: list[str], work: Path) -> dict:
    from lightrag.chunker import chunking_by_recursive_character, chunking_by_token_size
    from lightrag.constants import DEFAULT_R_SEPARATORS

    out = {}
    for route in routes:
        kind, _, rest = route.partition("_")
        size, _, overlap = rest.partition("_")
        size, overlap = int(size or 0), int(overlap or 0)
        if kind == "F":  # F_<chunk>_<overlap>
            out[route] = (lambda s, o: lambda text, name: chunking_by_token_size(
                tok, text, split_by_character=None, split_by_character_only=False,
                chunk_overlap_token_size=o, chunk_token_size=s))(size, overlap)
            ROUTES.setdefault(route, f"legacy-F（chunk {size} / overlap {overlap}）")
        elif kind == "R":  # R_<chunk>_<overlap>
            out[route] = (lambda s, o: lambda text, name: chunking_by_recursive_character(
                tok, text, s, chunk_overlap_token_size=o, separators=list(DEFAULT_R_SEPARATORS)))(size, overlap)
            ROUTES.setdefault(route, f"legacy-R（chunk_ts={size},chunk_ol={overlap}）")
        elif kind == "P":  # P_<CHUNK_P_SIZE>
            p = _p_chunker(tok, work, size or 2000)
            if p is not None:
                out[route] = p
                ROUTES.setdefault(route, f"native-P（CHUNK_P_SIZE {size or 2000}）")
    return out


def _p_chunker(tok, work: Path, size: int = 2000):
    """native Markdown パーサ（IR → *.parsed/ の blocks.jsonl）＋ paragraph_semantic(2000)。動かなければ None。"""
    try:
        import tempfile

        from lightrag.chunker.paragraph_semantic import chunking_by_paragraph_semantic
        from lightrag.parser.markdown.parser import NativeMarkdownParser
        from lightrag.sidecar import write_sidecar
    except Exception as exc:  # noqa: BLE001
        print(f"  P route unavailable: {exc}", flush=True)
        return None
    work.mkdir(parents=True, exist_ok=True)

    def run(text, name):
        parser = NativeMarkdownParser()
        blocks, _warnings, meta = parser._extract_text(text, bundle_root=None)
        stem = hashlib.md5(name.encode("utf-8")).hexdigest()[:12]
        ir = parser.build_ir(blocks, document_name=name, asset_dir_name=stem + ".blocks.assets", metadata=meta)
        with tempfile.TemporaryDirectory(dir=work) as tmp:
            pdata = write_sidecar(ir, parsed_dir=Path(tmp) / (stem + ".parsed"), doc_id="doc-" + "0" * 32,
                                  engine="native", clean_parsed_dir=True)
            return chunking_by_paragraph_semantic(tok, pdata["content"], size, blocks_path=pdata["blocks_path"],
                                                  chunk_overlap_token_size=100)
    return run


_TOKEN_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9_\-./]*")
_NORM_RE = re.compile(r"[\s\\]")  # 空白とエスケープの \ を除いて比べる（native パーサの再出力に合わせる）


def _contains_any(s: str, needles: set[str]) -> bool:
    if not needles:
        return False
    if len(needles) <= 8:
        return any(n in s for n in needles)
    return bool(needles & set(_TOKEN_RE.findall(s)))


def measure_route(manifest: list[dict], texts: dict, fn, tok) -> dict:
    chunk_tokens, recs_per_chunk, ents = [], [], []
    n_chunks = fffd = no_id = no_ctx = headings_only = over_limit = cut = total = 0
    id_in_heading_only = 0
    for m in manifest:
        text = texts[m["name"]]
        chunks = fn(text, m["name"])
        n_chunks += len(chunks)
        contents = [c["content"] for c in chunks]
        ids = {r["id"] for r in m["records"] if r.get("id")}
        ent_names = {e for r in m["records"] for e in r.get("entities", []) if e}
        if len(ent_names) > 8:  # 大きいファイルは設備コード（英数字）だけで判定する
            ent_names = {e for e in ent_names if _TOKEN_RE.fullmatch(e)}
        for s, c in zip(contents, chunks):
            chunk_tokens.append(c.get("tokens") or len(tok.encode(s)))
            fffd += "\ufffd" in s
            if ids and not _contains_any(s, ids):
                no_id += 1
                heading = json.dumps(c.get("heading") or {}, ensure_ascii=False)
                id_in_heading_only += _contains_any(heading, ids)
            has_ent = _contains_any(s, ent_names) if ent_names else True
            no_ctx += not (has_ent and DATE_RE.search(s))
            body = [ln for ln in s.split("\n") if ln.strip()]
            headings_only += bool(body) and all(ln.lstrip().startswith("#") for ln in body)
            e = estimate_entities(s)
            recs_per_chunk.append(e["records"])
            ents.append(e["est"])
            over_limit += e["est"] > ENTITY_LIMIT
        if m["kind"] == "form" or (m["kind"] == "records" and not _record_blocks(text)):
            total += 1  # 帳票・1件1ファイルはファイル全体が1記録
            cut += len(chunks) > 1
        elif m["kind"] == "records":
            # 記録（## ブロック）が1つのチャンクに丸ごと入っているか。見出し行で候補チャンクを絞る
            norm = [_NORM_RE.sub("", x) for x in contents]
            by_heading: dict[str, list[int]] = defaultdict(list)
            for i, x in enumerate(contents):
                for ln in x.split("\n"):
                    if ln.startswith("## "):
                        by_heading[_NORM_RE.sub("", ln)].append(i)
            for block in _record_blocks(text):
                total += 1
                b = _NORM_RE.sub("", block)
                head = _NORM_RE.sub("", block.split("\n", 1)[0])
                cut += not any(b in norm[i] for i in by_heading.get(head, []))
    # 入力トークンの概算: 1回目（system+user+chunk）＋ gleaning（system+user+chunk+1回目の出力+継続指示）
    input_est = sum(2 * (PROMPT_SYSTEM_TOKENS + PROMPT_USER_TOKENS + t) + ASSUMED_OUTPUT_TOKENS + PROMPT_GLEANING_TOKENS
                    for t in chunk_tokens)
    return {
        "chunks": n_chunks, "llm_calls_est": n_chunks * 2, "llm_input_tokens_est": input_est,
        "llm_output_tokens_est": n_chunks * 2 * ASSUMED_OUTPUT_TOKENS,
        "tokens": {"p10": _pct(chunk_tokens, .1), "p50": _pct(chunk_tokens, .5), "p90": _pct(chunk_tokens, .9),
                   "max": max(chunk_tokens, default=0), "sum": sum(chunk_tokens)},
        "records_total": total, "records_cut": cut, "records_cut_share": cut / total if total else 0,
        "chunks_fffd": fffd, "chunks_without_record_id": no_id, "chunks_id_only_in_heading_meta": id_in_heading_only,
        "chunks_without_entity_or_date": no_ctx, "headings_only_chunks": headings_only,
        "records_per_chunk": {"p50": _pct(recs_per_chunk, .5), "p90": _pct(recs_per_chunk, .9),
                              "max": max(recs_per_chunk, default=0)},
        "entities_est": {"p50": _pct(ents, .5), "p90": _pct(ents, .9), "max": max(ents, default=0),
                         "chunks_over_40": over_limit},
    }


def measure_variant(vdir: Path, chunkers: dict, tok) -> dict:
    manifest = json.loads((vdir / "manifest.json").read_text(encoding="utf-8"))
    texts = {m["name"]: (vdir / m["name"]).read_text(encoding="utf-8") for m in manifest}
    res: dict = {"files": len(manifest), "chars": sum(len(t) for t in texts.values()),
                 "files_by_kind": dict(Counter(m["kind"] for m in manifest))}
    res["tokens_total"] = sum(len(tok.encode(t)) for t in texts.values())
    names = [m["name"] for m in manifest]
    res["filenames"] = {
        "forbidden_chars": sum(1 for n in names if FORBIDDEN.search(n)),
        "hint_like_in_plain": sum(1 for n in names if HINT_RE.search(n)),  # ヒント記法に見える名前（0 であること）
        "max_len": max((len(n) for n in names), default=0),
        "collisions": sum(1 for n in names if "_dup" in n),
    }
    # LightRAG 1.5.7 は内容のハッシュ（doc-md5）で同一文書を重複として捨てる
    try:
        from lightrag.utils_pipeline import compute_text_content_hash
    except ImportError:  # pragma: no cover
        def compute_text_content_hash(t):
            return hashlib.md5(t.encode("utf-8")).hexdigest()
    hashes = Counter(compute_text_content_hash(t) for t in texts.values())
    res["duplicate_content_files"] = sum(c - 1 for c in hashes.values() if c > 1)
    # 定型行（半数以上のファイルに同じ行）の割合。3ファイル未満は対象外
    line_df: Counter = Counter()
    for t in texts.values():
        line_df.update({ln for ln in t.split("\n") if ln.strip()})
    thr = max(3, len(texts) // 2)
    lines = boiler = tok_all = tok_boiler = 0
    if len(texts) >= 3:
        for t in texts.values():
            for ln in t.split("\n"):
                if not ln.strip():
                    continue
                n = len(tok.encode(ln))
                lines += 1
                tok_all += n
                if line_df[ln] >= thr:
                    boiler += 1
                    tok_boiler += n
    res["boilerplate"] = {"threshold_files": thr, "line_share": boiler / lines if lines else 0,
                          "token_share": tok_boiler / tok_all if tok_all else 0,
                          "top": [ln for ln, c in line_df.most_common(10) if c >= thr and len(texts) >= 3]}
    rec_tokens = []
    for m in manifest:
        t = texts[m["name"]]
        if m["kind"] == "form":
            rec_tokens.append(len(tok.encode(t)))
        elif m["kind"] == "records":
            rec_tokens += [len(tok.encode(b)) for b in (_record_blocks(t) or [t])]
    # 記録の中の定型行（全記録の10%以上に同じ行）。例: 「- 状態: 完了」「- 再発: 無」
    rec_df: Counter = Counter()
    rec_lines_tok = rec_boiler_tok = 0
    blocks = [b for m in manifest if m["kind"] in ("records", "form")
              for b in (_record_blocks(texts[m["name"]]) or [texts[m["name"]]])]
    for b in blocks:
        rec_df.update({ln for ln in b.split("\n")[1:] if ln.strip()})
    n_blocks = len(blocks)
    for b in blocks[:20000]:
        for ln in b.split("\n")[1:]:
            if ln.strip():
                n = len(tok.encode(ln))
                rec_lines_tok += n
                if rec_df[ln] >= max(3, n_blocks // 10):
                    rec_boiler_tok += n
    res["record_boilerplate"] = {"blocks": n_blocks, "token_share": rec_boiler_tok / rec_lines_tok if rec_lines_tok else 0,
                                 "top": [(ln, c) for ln, c in rec_df.most_common(12) if c >= max(3, n_blocks // 10)]}
    res["record_tokens"] = {"n": len(rec_tokens), "p50": _pct(rec_tokens, .5), "p90": _pct(rec_tokens, .9),
                            "p99": _pct(rec_tokens, .99), "max": max(rec_tokens, default=0),
                            "over_800": sum(x > 800 for x in rec_tokens), "over_1200": sum(x > 1200 for x in rec_tokens)}
    res["routes"] = {}
    for route, fn in chunkers.items():
        t0 = time.perf_counter()
        res["routes"][route] = measure_route(manifest, texts, fn, tok)
        res["routes"][route]["seconds"] = round(time.perf_counter() - t0, 1)
    return res


def measure(out: Path, lightrag_dir: Path | None, routes: list[str], only: list[str] | None,
            metrics_name: str = "metrics.json") -> dict:
    if lightrag_dir is not None and (lightrag_dir / "tiktoken_cache").exists():
        os.environ.setdefault("TIKTOKEN_CACHE_DIR", str(lightrag_dir / "tiktoken_cache"))
    import logging

    from lightrag.utils import TiktokenTokenizer

    logging.getLogger("lightrag").setLevel(logging.WARNING)
    tok = TiktokenTokenizer("gpt-4o-mini")
    chunkers = make_chunkers(tok, routes, out / "work")
    path = out / metrics_name
    results = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {"variants": {}}
    results.setdefault("routes", {}).update({k: ROUTES[k] for k in chunkers})
    for vdir in sorted(d for d in (out / "md").iterdir() if d.is_dir()):
        if only and not any(vdir.name.startswith(o) for o in only):
            continue
        t0 = time.perf_counter()
        results["variants"][vdir.name] = measure_variant(vdir, chunkers, tok)
        print(f"  measured {vdir.name} ({time.perf_counter() - t0:.0f}s)", flush=True)
        path.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    return results


def summarize(out: Path) -> str:
    """metrics*.json をまとめ、報告書用の Markdown 表（variant × 経路）を返す。"""
    variants: dict = {}
    for path in sorted(out.glob("metrics*.json")):
        if path.name == "metrics_all.json":
            continue
        for name, v in json.loads(path.read_text(encoding="utf-8")).get("variants", {}).items():
            if name in variants:  # 同じ variant を別の経路で計測したファイルは経路だけ足す
                variants[name]["routes"].update(v["routes"])
            else:
                variants[name] = v
    (out / "metrics_all.json").write_text(json.dumps({"variants": variants}, ensure_ascii=False, indent=1), encoding="utf-8")
    rows = ["| variant | 経路 | ファイル | チャンク | 記録 | 途中切断 | U+FFFD | ID無し | 設備/日付無し | 見出しだけ | "
            "tok p50/p90/max | 記録/チャンク p50/max | 推定エンティティ p90/max（>40） | LLM呼出 | 入力tok概算 |",
            "|" + "---|" * 15]
    for name, v in variants.items():
        for route, r in v["routes"].items():
            t, rp, e = r["tokens"], r["records_per_chunk"], r["entities_est"]
            rows.append(
                f"| {name} | {route} | {v['files']:,} | {r['chunks']:,} | {r['records_total']:,} | "
                f"{r['records_cut']:,}（{r['records_cut_share']:.0%}） | {r['chunks_fffd']:,} | {r['chunks_without_record_id']:,} | "
                f"{r['chunks_without_entity_or_date']:,} | {r['headings_only_chunks']:,} | {t['p50']}/{t['p90']}/{t['max']} | "
                f"{rp['p50']}/{rp['max']} | {e['p90']}/{e['max']}（{e['chunks_over_40']}） | {r['llm_calls_est']:,} | "
                f"{r.get('llm_input_tokens_est', 0) / 1e6:.1f}M |")
    rows += ["", "| variant | ファイル | 総tok | 記録tok p50/p90/max | >800 | >1200 | 定型行tok（ファイル間） | 定型行tok（記録内） | "
             "内容重複 | 名前衝突 | 禁止文字 | ヒント記法 | 名前最大長 |", "|" + "---|" * 13]
    for name, v in variants.items():
        rt, fn = v["record_tokens"], v["filenames"]
        rows.append(
            f"| {name} | {v['files']:,} | {v['tokens_total']:,} | {rt['p50']}/{rt['p90']}/{rt['max']} | {rt['over_800']:,} | "
            f"{rt['over_1200']:,} | {v['boilerplate']['token_share']:.1%} | {v.get('record_boilerplate', {}).get('token_share', 0):.1%} | "
            f"{v['duplicate_content_files']} | {fn['collisions']} | {fn['forbidden_chars']} | "
            f"{fn['hint_like_in_plain']} | {fn['max_len']} |")
    return "\n".join(rows)


def main(argv=None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass
    ap = argparse.ArgumentParser(description="LightRAG 1.5.7 オフライン評価（生成・計測）")
    ap.add_argument("command", nargs="?", default="all", choices=("all", "generate", "measure", "summary"))
    ap.add_argument("--lightrag", help="LightRAG 1.5.7 のフォルダ（venv/ と tiktoken_cache/ を持つ）")
    ap.add_argument("--out", required=True, help="出力先フォルダ")
    ap.add_argument("--only", help="対象の絞り込み（generate: forms,T1,T2,T5 / measure: variant 名の先頭）")
    ap.add_argument("--routes", default="F_1200_100,F_600_50,P_2000")
    ap.add_argument("--variants", help="生成する variant（month,entity_month,per_record,single）")
    ap.add_argument("--metrics", default="metrics.json", help="計測結果のファイル名（並列実行時に分ける）")
    args = ap.parse_args(argv)
    out = Path(args.out)
    only = [x.strip() for x in args.only.split(",")] if args.only else None
    if args.command in ("all", "generate"):
        if only is None or "forms" in only:
            generate_forms(out)
        codes = [c for c in TABLES if only is None or c in only]
        if codes:
            generate_tables(out, codes, [x.strip() for x in args.variants.split(",")] if args.variants else None)
    if args.command == "measure":
        measure(out, Path(args.lightrag) if args.lightrag else None, args.routes.split(","), only, args.metrics)
    if args.command == "summary":
        print(summarize(out))
    if args.command == "all":
        if not args.lightrag:
            ap.error("all には --lightrag が必要です")
        py = Path(args.lightrag) / "venv" / "Scripts" / "python.exe"
        if not py.exists():
            py = Path(args.lightrag) / "venv" / "bin" / "python"
        subprocess.run([str(py), __file__, "measure", "--out", str(out), "--lightrag", args.lightrag,
                        "--routes", args.routes], check=True)
        print(summarize(out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
