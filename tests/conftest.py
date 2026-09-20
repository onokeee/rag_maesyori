from pathlib import Path

import pytest
from flask.testing import FlaskClient

from app import create_app
from excel.workbook import load_workbook_info
from scripts.make_samples import make_inspection, make_repair_shifted, make_repair_standard, make_repair_table


@pytest.fixture(scope="session")
def sample_dir(tmp_path_factory) -> Path:
    directory = tmp_path_factory.mktemp("samples")
    make_repair_standard(directory / "standard.xlsx")
    make_repair_shifted(directory / "shifted.xlsx")
    make_repair_table(directory / "table.xlsx")
    make_inspection(directory / "inspection.xlsx")
    return directory


@pytest.fixture(scope="session")
def repair_infos(sample_dir):
    return [load_workbook_info(sample_dir / n) for n in ("standard.xlsx", "shifted.xlsx", "table.xlsx")]


def make_config(tmp_path, **extra) -> dict:
    return {
        "TESTING": True,
        "SECRET_KEY": "test-secret",
        "DATABASE": tmp_path / "app.db",
        "UPLOAD_DIR": tmp_path / "uploads",
        "DATA_DIR": tmp_path / "data",
        # env ファイルに本物のAPIキーがあっても、テストが外部のAIに接続しないようにする
        # （AIを使うテストは偽サーバーの URL とキーで上書きする）
        "OPENAI_BASE_URL": "http://127.0.0.1:9/v1",
        "OPENAI_API_KEY": "",
        **extra,
    }


@pytest.fixture
def app(tmp_path):
    return create_app(make_config(tmp_path))


class BufferedClient(FlaskClient):
    """本文を最後まで読んで応答を閉じるテスト用クライアント（本番のサーバと同じ扱い）。

    ダウンロードしたデータを消すのは「本文を送り終えて応答を閉じたとき」なので（core.purge.purge_after_send）、
    応答を閉じないテストクライアントでは消える処理が動かない。buffered=True で毎回閉じる。
    """

    def open(self, *args, **kwargs):
        kwargs.setdefault("buffered", True)
        return super().open(*args, **kwargs)


@pytest.fixture
def client(app):
    app.test_client_class = BufferedClient
    return app.test_client()


# ---- 出来上がったデータを作る（画面と同じ道すじで） ------------------------------------------
# 画面は「帳票取り込み・表の取り込み・帳票登録」の3つだけで、どの段も fetch で進む（2026-09-20 の作り直し）。
# 「確定済みの帳票」「確定済みの一覧表の取り込み」は片付け・ダウンロードのテストで何度も要るので、ここに置く。

EXTRACTION = {
    "pattern": {"id": 1, "name": "設備修理報告書", "version": "v1"},
    "values": {"equipment_id": "EQ-001"},
    "fields": [{"field_name": "equipment_id", "display_name": "設備番号", "data_type": "string", "value": "EQ-001",
                "sheet": "修理報告書", "label_cell": "A4", "value_cell": "B4", "edited": False,
                "required": False}],
    "missing_required": [], "attachments": [], "sheets": ["修理報告書"],
}


def extraction_json(value: str = "EQ-001") -> str:
    import json

    data = json.loads(json.dumps(EXTRACTION))
    data["values"]["equipment_id"] = value
    data["fields"][0]["value"] = value
    return json.dumps(data, ensure_ascii=False)


def add_confirmed_document(app, name: str, *, value: str = "EQ-001", batch_id: str = "",
                           order: int = 0) -> tuple[int, Path]:
    """確定済みの帳票を1件作る（アップロードしたファイルの実体も置く）。"""
    from models import database as db

    with app.app_context():
        stored = f"documents/{name}"
        path = Path(app.config["UPLOAD_DIR"]) / stored
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"dummy-excel")
        doc_id = db.create_document(name, "0" * 64, stored, batch_id=batch_id, batch_order=order)
        db.update_document(doc_id, data_json=extraction_json(value), confirmed_json=extraction_json(value),
                           title=f"{value} 確定")
    return doc_id, path


def confirmed_import(app, client, file_name: str = "トラブル一覧.csv", template_name: str = "トラブル対応一覧") -> int:
    """CSV を置いて確定まで進めた取り込み（zip をダウンロードできる状態）。

    段ごとの fetch は tests/tables_helpers.py（表の取り込みのテスト用ヘルパー）と同じものを使う。
    """
    from tests.tables_helpers import confirmed

    return confirmed(app, client, file_name, template_name)
