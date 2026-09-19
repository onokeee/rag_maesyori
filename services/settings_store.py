"""data/ 以下の設定ファイルの読み書き（スレッドセーフ・一時ファイル経由で置き換え）。

読むときも書くときも同じロックを取る。Windows では、別のスレッドが開いているファイルを os.replace で
置き換えられない（Python の open は FILE_SHARE_DELETE を付けないので WinError 5 になる）。画面を出すたびに
設定を読む（app.py の context_processor）ので、読み書きが重ならないようにする。
ウイルス対策・同期ソフトが少しの間ファイルを掴んでいることもあるので、core.files.remove_upload と同じく
少し待って何度か試す。
"""
from __future__ import annotations

import os
import threading
import time
from pathlib import Path

import yaml
from flask import current_app

from core.files import REMOVE_RETRIES, REMOVE_RETRY_WAIT

_lock = threading.RLock()

# 手で編集して UTF-8 以外で保存された設定ファイルも読めるように試す文字コード（BOM 付き UTF-8 もここで読める）
_FALLBACK_ENCODINGS = ("utf-8-sig", "cp932")


def data_path(name: str) -> Path:
    return Path(current_app.config["DATA_DIR"]) / name


def _retry(action):
    """PermissionError（ほかのプログラムが掴んでいる）なら、少し待って何度か試す。最後の失敗はそのまま出す。"""
    for attempt in range(REMOVE_RETRIES):
        try:
            return action()
        except PermissionError:
            if attempt == REMOVE_RETRIES - 1:
                raise
            time.sleep(REMOVE_RETRY_WAIT * (attempt + 1))


def _decode(raw: bytes) -> str:
    """設定ファイルの中身を文字にする（UTF-16 は BOM で見分け、それ以外は UTF-8 → CP932 の順に試す）。"""
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for encoding in _FALLBACK_ENCODINGS:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("utf-8", raw, 0, 1, "UTF-8・CP932 のどちらでも読めません")


def read_yaml(name: str) -> dict:
    path = data_path(name)
    try:
        with _lock:
            if not path.exists():
                return {}
            raw = _retry(path.read_bytes)
        data = yaml.safe_load(_decode(raw))
    except (OSError, ValueError, yaml.YAMLError) as exc:   # UnicodeDecodeError は ValueError の仲間
        print(f"[settings] {path} を読めませんでした（{exc}）")
        return {}
    return data if isinstance(data, dict) else {}


def write_yaml(name: str, data: dict) -> None:
    """書けなかったとき（ほかのプログラムが掴んだまま・権限が無い）は OSError。"""
    _write(data_path(name), yaml.safe_dump(data, allow_unicode=True, sort_keys=False))


def _write(path: Path, text: str) -> None:
    with _lock:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(path.name + ".tmp")
        try:
            tmp.write_text(text, encoding="utf-8")
            _retry(lambda: os.replace(tmp, path))
        except OSError:
            try:
                tmp.unlink()
            except OSError:
                pass
            raise
        try:
            os.chmod(path, 0o600)  # APIキーを含むことがあるので所有者だけに絞る
        except OSError:
            pass


def lock() -> threading.RLock:
    return _lock
