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
