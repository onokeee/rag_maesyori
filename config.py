import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent

# 接続設定は aiagent_minimal_rag_tougou と同じく、ドット無しの "env" ファイル（export KEY="..." 形式も可）から読む
load_dotenv(BASE_DIR / "env")


def _split(value: str) -> list[str]:
    return [v.strip() for v in value.replace(",", ";").split(";") if v.strip()]


def _optional_float(name: str) -> float | None:
    raw = os.environ.get(name, "").strip()
    return float(raw) if raw else None


def _optional_int(name: str) -> int | None:
    raw = os.environ.get(name, "").strip()
    return int(raw) if raw else None


class Config:
    # セッション署名鍵は app.py が FLASK_SECRET_KEY または .flask_secret ファイルから設定する
    SECRET_KEY = None
    DATABASE = Path(os.environ.get("DATABASE", BASE_DIR / "instance" / "app.db"))
    UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", BASE_DIR / "uploads"))
    # 画面から保存する設定（model_settings.yaml）の置き場所
    DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data"))
    # 一覧表の行データ・状態ファイルの置き場所（DATA_DIR/tables。create_app で DATA_DIR に合わせて決め直す）
    TABLES_DIR = Path(os.environ.get("TABLES_DIR", DATA_DIR / "tables"))
    # 1回のリクエストの上限（一覧表の大きいCSV/Excelを想定）
    MAX_CONTENT_LENGTH = 200 * 1024 * 1024
    # 帳票（1ファイル＝1件）
    ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
    # 一覧表（Excel/CSV・1行＝1件）
    TABLE_ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".tsv", ".txt"}
    TABLE_MAX_UPLOAD_BYTES = int(os.environ.get("TABLE_MAX_UPLOAD_BYTES", str(200 * 1024 * 1024)))
    # Excel のセル数の上限（これを超えるシートは読み込まない）
    EXCEL_MAX_CELLS = int(os.environ.get("EXCEL_MAX_CELLS", "500000"))
    # 読み込み・下書き・AI整形を同時に動かす本数（列ごと。core/jobs.py）。
    # 社内LANのサーバーで数人が同時に使うので、1本だと誰かの長い読み込みでほかの人が待たされる。
    # 1〜8 に丸められる。列ごとの本数はプロセス内でその列を最初に使ったときに決まり、あとから減らない
    JOB_WORKERS = int(os.environ.get("JOB_WORKERS", "3"))

    # ---- LLM（OpenAI互換API） ----
    OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1").strip().rstrip("/")
    OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
    OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.6-sol").strip()
    OPENAI_MODELS = _split(os.environ.get("OPENAI_MODELS", ""))
    OPENAI_TEMPERATURE = float(os.environ.get("OPENAI_TEMPERATURE", "0"))
    OPENAI_TOP_P = _optional_float("OPENAI_TOP_P")
    OPENAI_MAX_TOKENS = _optional_int("OPENAI_MAX_TOKENS")
    LLM_RATE_LIMIT_RETRIES = int(os.environ.get("LLM_RATE_LIMIT_RETRIES", "3"))
    LLM_RATE_LIMIT_MAX_WAIT = float(os.environ.get("LLM_RATE_LIMIT_MAX_WAIT", "20"))
