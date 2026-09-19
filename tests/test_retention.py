"""データを残さないこと（design.md 3.3）の確認。

ダウンロードが終わったら、その取り込みに属するものは何も残らない:
アップロードしたファイル・imports/<id>/ のファイル・DB の行（どの表でも）。
残るのは帳票の種類と一覧表の取り込み設定（＝設定であってデータではない）。
"""
import io
import json
import zipfile
from pathlib import Path

from core import jobs, purge
from models import database as db
from tables import pipeline, store
from tests.test_tables_flow import COLUMNS, CSV_TEXT, _preview_page, _wait_import_job

EXTRACTION = {
    "pattern": {"id": 1, "name": "設備修理報告書", "version": "v1"},
    "values": {"equipment_id": "EQ-001"},
    "fields": [{"field_name": "equipment_id", "display_name": "設備番号", "data_type": "string", "value": "EQ-001",
                "sheet": "修理報告書", "label_cell": "A4", "value_cell": "B4", "edited": False, "ai_filled": False}],
    "missing_required": [], "attachments": [], "sheets": ["修理報告書"],
}


# ---- 共通 -------------------------------------------------------------------------

def _extraction(value="EQ-001") -> str:
    data = json.loads(json.dumps(EXTRACTION))
    data["values"]["equipment_id"] = value
    data["fields"][0]["value"] = value
    return json.dumps(data, ensure_ascii=False)


def _add_confirmed_document(app, name, *, value="EQ-001", batch_id="", order=0) -> tuple[int, Path]:
    """確定済みの帳票を1件作る（アップロードしたファイルの実体も置く）。"""
    with app.app_context():
        stored = f"documents/{name}"
        path = Path(app.config["UPLOAD_DIR"]) / stored
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"dummy-excel")
        doc_id = db.create_document(name, "0" * 64, stored, batch_id=batch_id, batch_order=order)
        db.update_document(doc_id, data_json=_extraction(value), confirmed_json=_extraction(value),
                           title=f"{value} 確定")
    return doc_id, path


def _rows_for(app, table_of_id: str, ref_columns: tuple[str, ...], value) -> dict[str, int]:
    """その取り込みを指す行が残っている表を {表.列: 件数} で返す。

    表の名前を決め打ちせず sqlite_master を見るので、あとから増えた表も見つかる。
    """
    hits: dict[str, int] = {}
    with app.app_context():
        conn = db.get_db()
        for (table,) in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall():
            if table.startswith("sqlite_"):
                continue
            columns = {row[1] for row in conn.execute(f'PRAGMA table_info("{table}")')}
            checks = [(c, value) for c in ref_columns if c in columns]
            if table == table_of_id and "id" in columns:
                checks.append(("id", value))
            for column, val in checks:
                count = conn.execute(f'SELECT COUNT(*) FROM "{table}" WHERE "{column}" = ?', (val,)).fetchone()[0]
                if count:
                    hits[f"{table}.{column}"] = count
    return hits


def _uploaded_files(app) -> list[Path]:
    base = Path(app.config["UPLOAD_DIR"])
    return [p for p in base.rglob("*") if p.is_file()]


# ---- 帳票（1件ずつ） ---------------------------------------------------------------------

def test_downloading_a_form_removes_its_file_and_every_row(app, client):
    doc_id, path = _add_confirmed_document(app, "報告書.xlsx")
    assert _rows_for(app, "documents", ("document_id",), doc_id)

    res = client.get(f"/forms/{doc_id}/download.md")
    assert res.status_code == 200 and "EQ-001" in res.get_data(as_text=True)

    assert not path.exists()
    assert _uploaded_files(app) == []
    assert _rows_for(app, "documents", ("document_id",), doc_id) == {}
    assert client.get(f"/forms/{doc_id}").status_code == 404
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 404


def test_a_form_that_cannot_be_built_is_not_deleted(app, client):
    """確定していない帳票はダウンロードできず、そのときデータも消さない。"""
    with app.app_context():
        doc_id = db.create_document("未確定.xlsx", "0" * 64, "documents/未確定.xlsx")
        db.save_draft(doc_id, _extraction())
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 404
    with app.app_context():
        assert db.get_document(doc_id) is not None


def test_deleting_a_form_by_hand_removes_the_file_too(app, client):
    doc_id, path = _add_confirmed_document(app, "消す.xlsx")
    res = client.post(f"/forms/{doc_id}/delete", data={"next": "/"})
    assert res.status_code == 302 and res.headers["Location"].endswith("/")
    assert not path.exists() and _rows_for(app, "documents", ("document_id",), doc_id) == {}


def test_the_form_type_is_kept_when_the_form_is_deleted(app, client):
    """帳票の種類は設定なので、データを消しても残る。"""
    with app.app_context():
        pattern_id = db.create_pattern("設備修理報告書")
    doc_id, _path = _add_confirmed_document(app, "種類つき.xlsx")
    with app.app_context():
        db.update_document(doc_id, pattern_id=pattern_id)
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 200
    with app.app_context():
        assert db.load_pattern(pattern_id) is not None


# ---- 帳票（まとめ取り込み） ----------------------------------------------------------------

def _upload_many(client, sample_dir, names):
    files = [(io.BytesIO((sample_dir / n).read_bytes()), n) for n in names]
    return client.post("/forms/upload", data={"file": files}, content_type="multipart/form-data")


def test_uploading_several_forms_makes_one_batch(app, client, sample_dir):
    res = _upload_many(client, sample_dir, ["standard.xlsx", "shifted.xlsx", "table.xlsx"])
    assert res.status_code == 302 and "/type" in res.headers["Location"]
    with app.app_context():
        rows = db.list_documents()
        assert len(rows) == 3
        batch_ids = {r["batch_id"] for r in rows}
        assert len(batch_ids) == 1 and "" not in batch_ids
    # 先頭の帳票の画面に進み具合（1/3）と案内が出る
    page = client.get(res.headers["Location"]).get_data(as_text=True)
    assert "まとめ取り込み 1/3件目" in page and "zip" in page


def test_a_single_file_upload_still_has_no_batch(app, client, sample_dir):
    res = client.post("/forms/upload", data={"file": (io.BytesIO((sample_dir / "standard.xlsx").read_bytes()),
                                                      "standard.xlsx")}, content_type="multipart/form-data")
    assert res.status_code == 302 and "/type" in res.headers["Location"]
    with app.app_context():
        assert db.list_documents()[0]["batch_id"] == ""
    assert "まとめ取り込み" not in client.get(res.headers["Location"]).get_data(as_text=True)


def test_batch_zip_is_refused_until_every_form_is_confirmed(app, client):
    doc_id, path = _add_confirmed_document(app, "1.xlsx", batch_id="B", order=0)
    with app.app_context():
        pending = db.create_document("2.xlsx", "0" * 64, "documents/2.xlsx", batch_id="B", batch_order=1)
    # 完了画面では「次の帳票へ」を出し、まとめてのダウンロードはまだ出さない
    page = client.get(f"/forms/{doc_id}/done").get_data(as_text=True)
    assert "まとめ取り込み 1/2件目" in page and "確定済み 1/2" in page
    assert "次の帳票へ（残り1件）" in page and f"/forms/{pending}/type" in page
    assert "まとめてMarkdownをダウンロード（zip）" not in page

    res = client.get("/forms/batches/B/download.zip")
    assert res.status_code == 302 and f"/forms/{pending}/type" in res.headers["Location"]
    assert path.exists()
    with app.app_context():
        assert db.get_document(doc_id) is not None


def test_downloading_the_batch_zip_removes_the_whole_batch(app, client):
    first, path1 = _add_confirmed_document(app, "1.xlsx", value="EQ-001", batch_id="B", order=0)
    second, path2 = _add_confirmed_document(app, "2.xlsx", value="EQ-002", batch_id="B", order=1)

    # すべて確定したら、完了画面から zip でまとめてダウンロードできる（押す前に消えることを知らせる）
    page = client.get(f"/forms/{second}/done").get_data(as_text=True)
    assert "まとめ取り込み 2/2件目" in page and "まとめてMarkdownをダウンロード（zip）" in page
    assert "/forms/batches/B/download.zip" in page
    assert "もう一度ダウンロードすることはできません" in page

    res = client.get("/forms/batches/B/download.zip")
    assert res.status_code == 200 and res.mimetype == "application/zip"
    with zipfile.ZipFile(io.BytesIO(res.data)) as zf:
        names = zf.namelist()
        texts = [zf.read(n).decode("utf-8") for n in names]
    assert len(names) == 2 and all(n.endswith(".md") for n in names)
    assert any("EQ-001" in t for t in texts) and any("EQ-002" in t for t in texts)

    assert not path1.exists() and not path2.exists()
    assert _uploaded_files(app) == []
    for doc_id in (first, second):
        assert _rows_for(app, "documents", ("document_id",), doc_id) == {}
    assert client.get("/forms/batches/B/download.zip").status_code == 404


# ---- 一覧表 ------------------------------------------------------------------------------

def _confirmed_import(app, client) -> int:
    """CSV を取り込んで確定まで進める（AI整形は使わない）。"""
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "トラブル一覧.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    client.post(f"/tables/imports/{import_id}/source",
                data={"encoding": "cp932", "delimiter": ",", "template": "new", "new_template_name": "トラブル対応一覧"})
    client.post(f"/tables/imports/{import_id}/layout", data={"header_rows": "1", "data_end_row": ""})
    payload = {"name": "トラブル対応一覧", "group_by": "month", "max_records_per_file": 300, "omit_person": True,
               "columns": [{"index": i, "header": h, "use": True, "key": k, "display": h, "type": t, "role": r,
                            "unit": "", "md": "attribute", "fill_down_blank": False, "ai": False, "description": ""}
                           for i, (k, h, t, r) in enumerate(COLUMNS)]}
    assert client.post(f"/tables/imports/{import_id}/columns", json=payload).status_code == 200
    _wait_import_job(app, import_id)
    _preview_page(app, client, import_id)
    client.post(f"/tables/imports/{import_id}/confirm")
    assert _wait_import_job(app, import_id)["status"] == "confirmed"
    return import_id


def test_downloading_the_table_zip_removes_everything(app, client):
    import_id = _confirmed_import(app, client)
    with app.app_context():
        imp = store.get_import(import_id)
        template_id = imp["template_id"]
        stored = Path(app.config["UPLOAD_DIR"]) / imp["stored_path"]
        folder = pipeline.import_dir(import_id)
        # AI整形の控え（行ごとの結果と生の応答）も取り込みと一緒に消える
        conn = db.get_db()
        conn.execute("INSERT INTO llm_calls (cache_key, raw_text, created_at) VALUES ('K', 'あ', '2026-09-19')")
        conn.execute("""INSERT INTO ai_items (template_id, stage_id, row_key, cache_key, status, updated_at)
                        VALUES (?, 'log', 'TR-001', 'K', 'ok', '2026-09-19')""", (template_id,))
        conn.commit()
    assert stored.exists() and folder.exists()

    res = client.get(f"/tables/imports/{import_id}/download.zip")
    assert res.status_code == 200 and res.mimetype == "application/zip"
    assert "RAG投入用/トラブル対応一覧_2026-08.md" in zipfile.ZipFile(io.BytesIO(res.data)).namelist()

    assert not stored.exists() and not folder.exists()
    assert _uploaded_files(app) == []
    assert _rows_for(app, "table_imports", ("import_id", "table_import_id"), import_id) == {}
    with app.app_context():
        conn = db.get_db()
        assert conn.execute("SELECT COUNT(*) FROM ai_items").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM jobs WHERE ref_id = ?", (import_id,)).fetchone()[0] == 0
        # 取り込み設定は残る（設定であってデータではない）
        assert store.get_template(template_id) is not None
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 404


def test_a_table_zip_that_cannot_be_built_deletes_nothing(app, client):
    """確定していない取り込みは zip を作れない。そのときデータも消さない。"""
    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "途中.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    res = client.get(f"/tables/imports/{import_id}/download.zip")
    assert res.status_code == 302
    with app.app_context():
        imp = store.get_import(import_id)
        assert imp is not None and (Path(app.config["UPLOAD_DIR"]) / imp["stored_path"]).exists()


# ---- 起動時の片付け -------------------------------------------------------------------------

def test_startup_removes_orphan_import_folders_but_keeps_work_in_progress(app, client, tmp_path):
    from app import create_app
    from tests.conftest import make_config

    res = client.post("/tables/upload", data={"file": (io.BytesIO(CSV_TEXT.encode("cp932")), "作業中.csv")},
                      content_type="multipart/form-data")
    import_id = int(res.headers["Location"].split("/")[3])
    with app.app_context():
        working = pipeline.import_dir(import_id)
        working.mkdir(parents=True, exist_ok=True)
        (working / "rows.jsonl.gz").write_bytes(b"x")
        orphan = working.parent / "99999"
        orphan.mkdir(parents=True, exist_ok=True)
        (orphan / "rows.jsonl.gz").write_bytes(b"x")
        config = {"DATABASE": app.config["DATABASE"], "UPLOAD_DIR": app.config["UPLOAD_DIR"],
                  "DATA_DIR": app.config["DATA_DIR"], "TABLES_DIR": app.config["TABLES_DIR"]}

    create_app(make_config(tmp_path, **config))  # もう一度起動する
    assert not orphan.exists()          # DB に無い取り込みのフォルダは消える
    assert working.exists()             # 作業中のものは残る


def test_purge_is_safe_when_there_is_nothing_left(app):
    """二重にダウンロードされても（すでに消えていても）落ちない。"""
    with app.app_context():
        assert purge.purge_documents([]) == 0
        assert purge.purge_documents([999]) == 0
        assert purge.purge_batch("") == 0
        assert purge.purge_table_import(999) == 0
        assert jobs.recover_interrupted() == 0


# ---- 消したデータの残りかす（DBファイル・アップロードファイル） -----------------------------------

SECRET = "ヒミツ商事_山田一郎"


def _db_bytes(app) -> bytes:
    """instance/app.db 本体と、まだ書き戻されていない -wal の生バイト。"""
    base = Path(app.config["DATABASE"])
    data = b""
    for path in (base, base.with_name(base.name + "-wal")):
        if path.exists():
            data += path.read_bytes()
    return data


def test_downloaded_values_do_not_stay_in_the_database_file(app, client):
    """行を消すだけでは、解放されたページに帳票の値が平文で残る（design.md 3.3）。"""
    doc_id, _path = _add_confirmed_document(app, "残りかす.xlsx", value=SECRET)
    assert SECRET.encode("utf-8") in _db_bytes(app)      # 確定した時点では当然ある

    assert client.get(f"/forms/{doc_id}/download.md").status_code == 200
    assert SECRET.encode("utf-8") not in _db_bytes(app)  # ダウンロードしたら DB ファイルからも読めない


def test_downloaded_table_values_do_not_stay_in_the_database_file(app, client):
    import_id = _confirmed_import(app, client)
    with app.app_context():
        conn = db.get_db()
        conn.execute("UPDATE table_imports SET file_name = ? WHERE id = ?", (f"{SECRET}.csv", import_id))
        conn.commit()
    assert SECRET.encode("utf-8") in _db_bytes(app)

    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    assert SECRET.encode("utf-8") not in _db_bytes(app)


# ---- 送信できなかったダウンロード ----------------------------------------------------------

def test_a_download_that_is_cut_off_keeps_the_data(app, client):
    """本文を受け取れなかったとき（ブラウザを閉じた・通信が切れた）は消さない。消したら取り戻せないため。"""
    doc_id, path = _add_confirmed_document(app, "切れた.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md", buffered=False)
    res.close()  # 本文を1バイトも受け取らずに切れた
    assert path.exists()
    with app.app_context():
        assert db.get_document(doc_id) is not None

    # もう一度ダウンロードすれば受け取れる（そのときは消える）
    res = client.get(f"/forms/{doc_id}/download.md")
    assert res.status_code == 200 and "EQ-001" in res.get_data(as_text=True)
    assert not path.exists()


def test_a_table_download_that_is_cut_off_keeps_the_data(app, client):
    import_id = _confirmed_import(app, client)
    client.get(f"/tables/imports/{import_id}/download.zip", buffered=False).close()
    with app.app_context():
        assert store.get_import(import_id) is not None
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    with app.app_context():
        assert store.get_import(import_id) is None


# ---- まとめ取り込み: 確定済みだけをダウンロードする ------------------------------------------------

def test_only_the_confirmed_forms_of_a_batch_can_be_downloaded(app, client):
    first, path1 = _add_confirmed_document(app, "1.xlsx", value="EQ-001", batch_id="B", order=0)
    with app.app_context():
        pending = db.create_document("2.xlsx", "0" * 64, "documents/2.xlsx", batch_id="B", batch_order=1)

    # 完了画面から「確定済みだけ」を選べる（未確定が残っていても作業が止まらない）
    page = client.get(f"/forms/{first}/done").get_data(as_text=True)
    assert "確定済み1件だけをダウンロード（zip）" in page

    res = client.get("/forms/batches/B/download.zip?confirmed_only=1")
    assert res.status_code == 200 and res.mimetype == "application/zip"
    assert len(zipfile.ZipFile(io.BytesIO(res.data)).namelist()) == 1

    assert not path1.exists()                       # 渡した分だけ消える
    with app.app_context():
        assert db.get_document(first) is None
        assert db.get_document(pending) is not None  # 未確定の帳票は残る


def test_the_home_shows_a_batch_as_one_row_with_the_zip_button(app, client):
    _add_confirmed_document(app, "1.xlsx", value="EQ-001", batch_id="B", order=0)
    _add_confirmed_document(app, "2.xlsx", value="EQ-002", batch_id="B", order=1)
    page = client.get("/").get_data(as_text=True)
    assert "まとめ取り込み（2ファイル）" in page and "/forms/batches/B/download.zip" in page
    # 1件だけ押すとまとまりが壊れることを、押す前に知らせる
    assert "この帳票だけがまとまりから消えます" in page


# ---- 消えたあとの画面・ファイル名・メッセージ ------------------------------------------------------

def test_json_requests_after_the_data_is_gone_get_a_json_error(app, client):
    doc_id, _path = _add_confirmed_document(app, "消えた.xlsx")
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 200
    res = client.post(f"/forms/{doc_id}/draft", json={"values": {}}, headers={"Accept": "application/json"})
    assert res.status_code == 404 and res.mimetype == "application/json"
    assert "残っていません" in res.get_json()["error"]


def test_downloads_have_a_readable_ascii_file_name(app, client):
    import_id = _confirmed_import(app, client)
    disposition = client.get(f"/tables/imports/{import_id}/download.zip").headers["Content-Disposition"]
    # filename* を読めない相手にも、何の zip か分かる名前を渡す（「_2026-09-19.zip」にしない）
    assert "filename=\"records_" in disposition and "filename*=UTF-8''" in disposition


def test_delete_messages_do_not_carry_the_file_name(app, client):
    """flash はブラウザのセッションクッキーに残るので、ファイル名を載せない。"""
    doc_id, _path = _add_confirmed_document(app, "社外秘_取引先A.xlsx")
    res = client.post(f"/forms/{doc_id}/delete", data={"next": "/"})
    assert "社外秘" not in str(res.headers.get("Set-Cookie", ""))
    assert "帳票を削除しました" in client.get("/").get_data(as_text=True)


# ---- AI整形の控え（取り込み単位で消す） ----------------------------------------------------

def _ai_row(conn, template_id: int, import_id: int, row_key: str, cache_key: str) -> None:
    conn.execute("INSERT INTO llm_calls (cache_key, raw_text, created_at) VALUES (?, 'あ', '2026-09-19')",
                 (cache_key,))
    conn.execute("""INSERT INTO ai_items (template_id, import_id, stage_id, row_key, cache_key, status, updated_at)
                    VALUES (?, ?, 'log', ?, ?, 'ok', '2026-09-19')""",
                 (template_id, import_id, row_key, cache_key))


def test_purging_one_import_keeps_the_ai_results_of_another_import(app, client):
    """同じ取り込み設定で作業中の別の取り込みの AI整形結果（課金済み）まで消さない。"""
    import_id = _confirmed_import(app, client)
    with app.app_context():
        template_id = store.get_import(import_id)["template_id"]
        other = store.create_import("作業中.csv", "1" * 64, "tables/other.csv", {"kind": "csv"}, template_id,
                                    store.get_template(template_id)["current_version_id"])
        conn = db.get_db()
        _ai_row(conn, template_id, import_id, "TR-001", "K1")
        _ai_row(conn, template_id, other, "TR-900", "K2")
        conn.commit()

    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200

    with app.app_context():
        conn = db.get_db()
        rows = conn.execute("SELECT row_key, import_id FROM ai_items").fetchall()
        assert [(r["row_key"], r["import_id"]) for r in rows] == [("TR-900", other)]
        assert [r[0] for r in conn.execute("SELECT cache_key FROM llm_calls")] == ["K2"]
        assert store.get_import(other) is not None


def test_ai_results_of_an_import_that_is_gone_are_swept(app, client, tmp_path):
    """消した取り込みを指す AI整形の結果（消したあとに動いていた AI が書いた分など）も残さない。"""
    from app import create_app
    from tests.conftest import make_config

    import_id = _confirmed_import(app, client)
    with app.app_context():
        template_id = store.get_import(import_id)["template_id"]
        conn = db.get_db()
        _ai_row(conn, template_id, 777, "TR-777", "K7")   # もう無い取り込みの分
        conn.commit()
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    with app.app_context():
        conn = db.get_db()
        assert conn.execute("SELECT COUNT(*) FROM ai_items").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0
        # 起動時の片付けでも消える
        _ai_row(conn, template_id, 778, "TR-778", "K8")
        conn.commit()
        config = {"DATABASE": app.config["DATABASE"], "UPLOAD_DIR": app.config["UPLOAD_DIR"],
                  "DATA_DIR": app.config["DATA_DIR"], "TABLES_DIR": app.config["TABLES_DIR"]}
    restarted = create_app(make_config(tmp_path, **config))
    with restarted.app_context():
        conn = db.get_db()
        assert conn.execute("SELECT COUNT(*) FROM ai_items").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0


# ---- 他サイトのページからのダウンロード（＝削除）を断る -----------------------------------------------
# ダウンロードはデータを消すので、GET でも他サイト発なら断る（<img> や別サイトのリンクで消されないように）。

CROSS_SITE_IMG = {"Sec-Fetch-Site": "cross-site", "Sec-Fetch-Mode": "no-cors", "Sec-Fetch-Dest": "image",
                  "Referer": "http://evil.example/"}


def test_a_cross_site_form_download_is_refused_and_keeps_the_data(app, client):
    doc_id, path = _add_confirmed_document(app, "他サイト.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md", headers=CROSS_SITE_IMG)
    assert res.status_code == 403 and "EQ-001" not in res.get_data(as_text=True)
    assert path.exists() and _rows_for(app, "documents", ("document_id",), doc_id)


def test_a_form_download_with_a_foreign_referer_and_no_fetch_headers_is_refused(app, client):
    """Sec-Fetch-Site を付けないブラウザでも、Referer が別サイトなら断る。"""
    doc_id, path = _add_confirmed_document(app, "古いブラウザ.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md", headers={"Referer": "http://evil.example/page"})
    assert res.status_code == 403
    assert path.exists() and _rows_for(app, "documents", ("document_id",), doc_id)


def test_a_cross_site_batch_zip_download_is_refused_and_keeps_the_data(app, client):
    first, path1 = _add_confirmed_document(app, "b1.xlsx", value="EQ-001", batch_id="X", order=0)
    second, path2 = _add_confirmed_document(app, "b2.xlsx", value="EQ-002", batch_id="X", order=1)
    res = client.get("/forms/batches/X/download.zip", headers=CROSS_SITE_IMG)
    assert res.status_code == 403 and res.mimetype != "application/zip"
    assert path1.exists() and path2.exists()
    for doc_id in (first, second):
        assert _rows_for(app, "documents", ("document_id",), doc_id)


def test_a_cross_site_table_zip_download_is_refused_and_keeps_the_data(app, client):
    import_id = _confirmed_import(app, client)
    with app.app_context():
        stored = Path(app.config["UPLOAD_DIR"]) / store.get_import(import_id)["stored_path"]
    res = client.get(f"/tables/imports/{import_id}/download.zip", headers=CROSS_SITE_IMG)
    assert res.status_code == 403 and res.mimetype != "application/zip"
    with app.app_context():
        assert stored.exists() and pipeline.import_dir(import_id).exists()
    assert _rows_for(app, "table_imports", ("import_id", "table_import_id"), import_id)


def test_same_origin_and_typed_downloads_still_work(app, client):
    """アプリ内のクリック（same-origin）とアドレス欄への入力（none）は、これまでどおりダウンロードして消す。"""
    doc_id, path = _add_confirmed_document(app, "同じサイト.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md",
                     headers={"Sec-Fetch-Site": "same-origin", "Referer": f"http://localhost/forms/{doc_id}/done"})
    assert res.status_code == 200 and "EQ-001" in res.get_data(as_text=True)
    assert not path.exists()

    doc_id, path = _add_confirmed_document(app, "手入力.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md", headers={"Sec-Fetch-Site": "none"})
    assert res.status_code == 200 and not path.exists()


def test_other_cross_site_gets_are_still_allowed(app, client):
    """消さない GET（画面の表示）は他サイトからのリンクでも開ける。"""
    assert client.get("/", headers=CROSS_SITE_IMG).status_code == 200


# ---- 一部だけのダウンロード（Range）・削除の順番・取り込み設定の削除 -----------------------------------

def test_a_ranged_form_download_keeps_the_data(app, client):
    """Range 付き（ダウンロードマネージャー・再開）で一部だけ渡したときは消さない。全部渡したときだけ消す。"""
    doc_id, path = _add_confirmed_document(app, "一部だけ.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md", headers={"Range": "bytes=0-9"})
    assert res.status_code in (200, 206)
    if res.status_code == 206:
        assert path.exists()
        with app.app_context():
            assert db.get_document(doc_id) is not None
        res = client.get(f"/forms/{doc_id}/download.md")
        assert res.status_code == 200 and "EQ-001" in res.get_data(as_text=True)
    assert not path.exists()
    with app.app_context():
        assert db.get_document(doc_id) is None


def test_a_ranged_table_download_keeps_the_data(app, client):
    import_id = _confirmed_import(app, client)
    res = client.get(f"/tables/imports/{import_id}/download.zip", headers={"Range": "bytes=0-99"})
    assert res.status_code in (200, 206)
    if res.status_code == 206:
        with app.app_context():
            assert store.get_import(import_id) is not None
        assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    with app.app_context():
        assert store.get_import(import_id) is None


def _failing_delete(*_args, **_kwargs):
    raise RuntimeError("database is locked")


def test_files_stay_when_deleting_the_form_rows_fails(app, client, monkeypatch):
    """DB の削除に失敗したら、ファイルも残す（ファイルの無い帳票が作業中として残らないように）。"""
    doc_id, path = _add_confirmed_document(app, "消せない.xlsx")
    monkeypatch.setattr(purge, "_delete_by_columns", _failing_delete)
    assert client.get(f"/forms/{doc_id}/download.md").status_code == 200
    assert path.exists()
    with app.app_context():
        assert db.get_document(doc_id) is not None


def test_files_stay_when_deleting_the_table_rows_fails(app, client, monkeypatch):
    import_id = _confirmed_import(app, client)
    with app.app_context():
        stored = Path(app.config["UPLOAD_DIR"]) / store.get_import(import_id)["stored_path"]
        folder = pipeline.import_dir(import_id)
    monkeypatch.setattr(purge, "_delete_by_columns", _failing_delete)
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    assert stored.exists() and folder.exists()
    with app.app_context():
        assert store.get_import(import_id) is not None
    # 失敗が解消すれば、もう一度ダウンロードして消せる
    monkeypatch.undo()
    assert client.get(f"/tables/imports/{import_id}/download.zip").status_code == 200
    assert not stored.exists() and not folder.exists()


def test_ai_responses_left_by_a_template_delete_are_swept_at_startup(app, client):
    """取り込み設定を消すと ai_items は消えるが、生の応答（llm_calls）が残る。起動時の片付けで消す。"""
    from app import _cleanup_leftovers

    import_id = _confirmed_import(app, client)
    with app.app_context():
        template_id = store.get_import(import_id)["template_id"]
        conn = db.get_db()
        _ai_row(conn, template_id, import_id, "TR-001", "K")
        conn.commit()
        store.delete_template(template_id)
        assert conn.execute("SELECT COUNT(*) FROM ai_items").fetchone()[0] == 0
    _cleanup_leftovers(app)
    with app.app_context():
        assert db.get_db().execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0


def test_sweep_orphan_ai_removes_responses_nobody_uses(app):
    with app.app_context():
        conn = db.get_db()
        conn.execute("INSERT INTO llm_calls (cache_key, raw_text, created_at) VALUES ('Z', ?, '2026-09-19')",
                     (SECRET,))
        conn.commit()
        assert purge.sweep_orphan_ai(conn) == 1
        assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0
    assert SECRET.encode("utf-8") not in _db_bytes(app)


def test_deleting_a_template_removes_the_raw_ai_responses_at_once(app, client):
    """取り込み設定を消したら、生の応答（llm_calls）も起動を待たずにすぐ消える。"""
    import_id = _confirmed_import(app, client)
    with app.app_context():
        template_id = store.get_import(import_id)["template_id"]
        conn = db.get_db()
        _ai_row(conn, template_id, import_id, "TR-001", "K")
        conn.commit()
        assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 1
        store.delete_template(template_id)
        assert conn.execute("SELECT COUNT(*) FROM llm_calls").fetchone()[0] == 0


def test_a_ranged_download_gets_the_whole_file_and_removes_the_item(app, client):
    """Range 付きでも全体を 200 で返す（一部だけ渡して消すことが無い）。全部渡したので消える。"""
    doc_id, path = _add_confirmed_document(app, "範囲指定.xlsx")
    res = client.get(f"/forms/{doc_id}/download.md", headers={"Range": "bytes=0-9"})
    assert res.status_code == 200 and "EQ-001" in res.get_data(as_text=True)
    assert not path.exists()
    import_id = _confirmed_import(app, client)
    res = client.get(f"/tables/imports/{import_id}/download.zip", headers={"Range": "bytes=0-99"})
    assert res.status_code == 200
    assert "RAG投入用/" in " ".join(zipfile.ZipFile(io.BytesIO(res.data)).namelist())
    with app.app_context():
        assert store.get_import(import_id) is None


def test_purge_after_send_keeps_the_item_for_a_partial_response(app):
    """206（一部だけ）・304 の応答では、送り終えても消さない。"""
    from flask import Response

    called = []
    with app.app_context():
        for status in (206, 304):
            res = purge.purge_after_send(Response(b"abc", status=status), called.append, status)
            list(res.response)
            res.close()
        assert called == []
        res = purge.purge_after_send(Response(b"abc", status=200), called.append, 200)
        list(res.response)
        res.close()
    assert called == [200]


def test_a_folder_that_cannot_be_removed_is_logged(app, client, monkeypatch, caplog):
    """imports/<id>/ を消し切れなかったら、黙って成功扱いにせず記録する。"""
    import logging

    import_id = _confirmed_import(app, client)
    monkeypatch.setattr(purge.shutil, "rmtree", lambda *a, **k: None)
    with app.app_context(), caplog.at_level(logging.WARNING, logger="core.purge"):
        assert purge.import_dir(import_id).exists()
        purge.purge_table_import(import_id)
        assert store.get_import(import_id) is None
    assert f"imports/{import_id}" in caplog.text
