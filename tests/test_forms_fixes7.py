"""帳票取り込みの直し（まとめ取り込みの件数・シートのチェック・見回り・ダウンロード）。

このまわりで見つかった不具合:
  - ②へ戻るとシートのチェックが先頭ファイル分だけになり、読み取り直すと帳票が黙って消える
  - ④が③に並んでいない帳票まで数え、案内の件数と zip の中身が合わない
  - 見回りが「取り込んだ時刻」で切るので、2時間かけて確認すると編集中の帳票が消える
  - 修正中の帳票をリンクから直接ダウンロードすると、直した値が入らないまま消える
"""
import io
import json
import zipfile
from datetime import datetime, timedelta

from core import purge
from models import database as db
from tests.test_forms_batch_review import _copies, _pattern, read_all
from tests.test_forms_flow import activate, add_field, create_type, finish, upload_forms


def _type_fragment(client, ids) -> str:
    return client.get("/forms/type?ids=" + ",".join(str(i) for i in ids)).get_json()["html"]


def _checked_sheets(html: str) -> list[str]:
    """②「読み取るシート」でチェックが入っているシートの名前。"""
    out = []
    for part in html.split('name="sheets"')[1:]:
        tag = part.split(">")[0]
        if " checked" in tag:
            out.append(tag.split('value="')[1].split('"')[0])
    return out


# ---- ② シートのチェックは、読み取ったファイル全部を合わせる -------------------------------------

def test_reopening_the_type_step_keeps_every_sheet_that_was_read(app, client, sample_dir, tmp_path):
    """②へ戻っても、読み取ったシートのチェックが先頭ファイル分に減らない。

    1件分の読み取り結果には、そのファイルに在ったシートしか残らない。先頭のファイルだけを見ると
    ほかのファイルのシートのチェックが外れ、そのまま読み取り直すと大半の帳票が③から消えていた。
    """
    pattern_id = _pattern(client, sample_dir)
    # シートの名前が違う2件（standard は「修理報告書」、shifted は「修理報告書(2)」）
    ids = upload_forms(client, sample_dir / "standard.xlsx", sample_dir / "shifted.xlsx")
    chosen = ["修理報告書", "修理報告書(2)"]
    assert read_all(client, ids, pattern_id, chosen).status_code == 200

    checked = _checked_sheets(_type_fragment(client, ids))
    assert sorted(checked) == sorted(chosen)

    # そのまま読み取り直しても2件とも読める（前は先頭ファイルのシートだけになり1件に減っていた）
    body = read_all(client, ids, pattern_id, checked, acknowledge="on").get_json()
    assert [d["id"] for d in body["docs"]] == ids and not body.get("errors")


def test_a_sheet_the_user_unchecked_stays_unchecked(app, client, sample_dir, tmp_path):
    """手で外したシートは、②へ戻ったときにチェックが戻らない。"""
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 2))

    first = _type_fragment(client, ids)
    assert "参考資料" in first                      # 見本のブックには参考資料シートもある
    assert read_all(client, ids, pattern_id, ["修理報告書"]).status_code == 200

    assert _checked_sheets(_type_fragment(client, ids)) == ["修理報告書"]


# ---- ④ 数えるのは zip に入る帳票だけ -----------------------------------------------------------

def test_the_finish_step_counts_only_the_forms_that_can_reach_the_zip(app, client, sample_dir, tmp_path):
    """読み取れなかったファイルを件数に入れない（案内の件数と zip の中身を合わせる）。"""
    pattern_id = _pattern(client, sample_dir)
    paths = _copies(sample_dir, tmp_path, 2)
    inspection = tmp_path / "点検記録表.xlsx"
    inspection.write_bytes((sample_dir / "inspection.xlsx").read_bytes())
    ids = upload_forms(client, inspection, *paths)

    body = read_all(client, ids, pattern_id, ["修理報告書"]).get_json()
    assert body["errors"] == ["点検記録表.xlsx: 選んだシートがファイルにありません"]

    state = finish(client, ids)
    assert state["total"] == 2                       # 置いたのは3件だが、zip に入るのは2件
    page = state["html"]
    assert "確定済み <strong>0</strong> / 2 件" in page
    assert "残り2件も確定して、まとめてダウンロード（zip）" in page
    # 読み取れなかったファイルは、入らないことを名前を出して知らせる
    assert "まだ読み取れていない1件（zip には入りません）" in page and "点検記録表.xlsx" in page
    assert "サーバーからすべて消えます" not in page and "サーバーに残ります" in page
    # ③に並んでいない帳票には［読み取り結果を見る］を出さない（押しても何も起きないボタンを出さない）
    assert page.count("data-goto-doc") == 2

    for doc_id in [d["id"] for d in body["docs"]]:
        assert client.post(f"/forms/{doc_id}/confirm", json={}).get_json()["ok"] is True
    with app.app_context():
        batch_id = db.get_document(ids[0])["batch_id"]
    res = client.get(f"/forms/batches/{batch_id}/download.zip")
    assert res.status_code == 200
    assert len(zipfile.ZipFile(io.BytesIO(res.data)).namelist()) == 2   # 案内どおりの件数


def test_a_form_that_fails_a_second_read_drops_out_of_the_count(app, client, sample_dir, tmp_path):
    """一度読めた帳票が読み取り直しで読めなくなったら、③から消えるのと一緒に件数からも外す。"""
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, sample_dir / "standard.xlsx", sample_dir / "shifted.xlsx")
    assert read_all(client, ids, pattern_id, ["修理報告書", "修理報告書(2)"]).status_code == 200
    assert finish(client, ids)["total"] == 2

    # 「修理報告書」だけで読み直す → shifted はそのシートが無いので読めない
    body = read_all(client, ids, pattern_id, ["修理報告書"], acknowledge="on").get_json()
    assert [d["id"] for d in body["docs"]] == [ids[0]]
    with app.app_context():
        assert db.get_document(ids[1])["data_json"] is None     # 前の読み取り結果は残さない
    assert finish(client, ids)["total"] == 1


# ---- 見回りは「最後にさわった時刻」で切る -------------------------------------------------------

def _age(app, doc_ids, hours: float, column: str = "created_at") -> None:
    stamp = (datetime.now() - timedelta(hours=hours)).isoformat(timespec="seconds")
    with app.app_context():
        for doc_id in doc_ids:
            db.update_document(doc_id, **{column: stamp})


def test_the_sweeper_keeps_a_form_that_is_still_being_corrected(app, client, sample_dir, tmp_path):
    """途中保存をしている帳票は、取り込みから2時間たっても捨てない。

    帳票には updated_at が無く、取り込んだ時刻で切られていたので、50件を2時間かけて確認すると
    手で直した値ごと消えていた（一覧表は updated_at を持っていて、約束どおり動いていた）。
    """
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))
    read_all(client, ids, pattern_id, ["修理報告書"])
    _age(app, ids, 2.5)                                   # 2時間半前に取り込んだ

    for doc_id in ids:                                    # ずっと手で直している
        assert client.post(f"/forms/{doc_id}/draft",
                           json={"values": {"equipment_name": "直した設備名"}}).status_code == 204
    with app.app_context():
        assert purge.sweep_stale(2) == (0, 0)
        assert all(db.get_document(i) is not None for i in ids)


def test_the_sweeper_keeps_a_batch_together_while_one_form_is_fresh(app, client, sample_dir, tmp_path):
    """まとまりのどれか1件でもさわられていれば、まとまりごと残す（zip が欠けないように）。"""
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 3))
    read_all(client, ids, pattern_id, ["修理報告書"])
    _age(app, ids, 2.5)
    _age(app, ids, 2.5, column="updated_at")

    # いちばん下の1件だけをいま直した（上から順に見ていくと、まだ手が届いていない分がこうなる）
    assert client.post(f"/forms/{ids[2]}/draft", json={"values": {"equipment_name": "直した"}}).status_code == 204
    with app.app_context():
        assert purge.sweep_stale(2) == (0, 0)
        assert all(db.get_document(i) is not None for i in ids)

    # まとまりのどれもさわられなくなれば、これまでどおり捨てる
    _age(app, ids, 2.5, column="updated_at")
    with app.app_context():
        assert purge.sweep_stale(2) == (3, 0)
        assert all(db.get_document(i) is None for i in ids)


# ---- 修正中の帳票は、直した値で渡す -------------------------------------------------------------

def test_a_modified_form_downloads_with_the_corrected_value(app, client, sample_dir):
    """画面の JS を通さずにリンクを開いても、直した値が .md に入る（design.md 3.3）。"""
    pattern_id = create_type(client, sample_dir / "standard.xlsx", "設備修理報告書")
    add_field(client, pattern_id, "修理報告書", "E4", "F4")        # 設備名
    activate(client, pattern_id)
    doc_id, = upload_forms(client, sample_dir / "standard.xlsx")
    read_all(client, [doc_id], pattern_id, ["修理報告書"])

    assert client.post(f"/forms/{doc_id}/confirm", json={}).get_json()["ok"] is True
    assert client.post(f"/forms/{doc_id}/draft",
                       json={"values": {"equipment_name": "直した設備名"}}).status_code == 204
    with app.app_context():
        assert db.get_document(doc_id)["state"] == "modified"

    md = client.get(f"/forms/{doc_id}/download.md").get_data(as_text=True)
    assert "直した設備名" in md


def test_a_modified_form_in_a_batch_zip_carries_the_corrected_value(app, client, sample_dir, tmp_path):
    """まとまりの zip も同じ（中クリックなどで JS を通さずに開いたとき）。"""
    pattern_id = _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 2))
    read_all(client, ids, pattern_id, ["修理報告書"])
    for doc_id in ids:
        assert client.post(f"/forms/{doc_id}/confirm", json={}).get_json()["ok"] is True
    assert client.post(f"/forms/{ids[0]}/draft",
                       json={"values": {"equipment_name": "直した設備名"}}).status_code == 204

    with app.app_context():
        batch_id = db.get_document(ids[0])["batch_id"]
    res = client.get(f"/forms/batches/{batch_id}/download.zip")
    assert res.status_code == 200
    with zipfile.ZipFile(io.BytesIO(res.data)) as zf:
        bodies = [zf.read(n).decode("utf-8") for n in zf.namelist()]
    assert any("直した設備名" in b for b in bodies)


# ---- ② は置かれたブックを1件ずつ開く（50件でメモリを食いつぶさない） -------------------------------

def test_the_type_step_opens_one_workbook_at_a_time(app, client, sample_dir, tmp_path, monkeypatch):
    """置かれたファイル全部のブックを同時にメモリへ広げない（50件でサーバーが落ちないように）。

    1件ぶんの WorkbookInfo はセルごとの控えを持つので大きい。全部を持ち続けると件数に比例して
    増え、50件で数百MB〜数GBになっていた。
    """
    import gc
    import weakref

    from views import forms as forms_view

    seen: list[weakref.ref] = []
    alive_during: list[int] = []
    original = forms_view._load_info

    def watching(doc):
        gc.collect()
        alive_during.append(sum(1 for ref in seen if ref() is not None))
        info = original(doc)
        seen.append(weakref.ref(info))
        return info

    _pattern(client, sample_dir)
    ids = upload_forms(client, *_copies(sample_dir, tmp_path, 4))
    monkeypatch.setattr(forms_view, "_load_info", watching)
    body = client.get("/forms/type?ids=" + ",".join(str(i) for i in ids)).get_json()

    assert len(seen) == 4 and len(body["docs"]) == 4
    # 次の1件を開くとき、前に開いたブックはもう残っていない
    assert alive_during == [0, 0, 0, 0]
    # 画面に出す中身はこれまでどおり（シートは名前ごとに1行にまとまる。4件ぶん並ばない）
    assert body["html"].count('name="sheets"') == 3 and "4件をまとめて読み取る" in body["html"]
