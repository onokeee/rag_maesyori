"""data/ 以下の設定ファイルの読み書き（スレッドセーフ・一時ファイル経由で置き換え）。"""
from __future__ import annotations

import os
import threading
from pathlib import Path

import yaml
from flask import current_app

_lock = threading.RLock()


def data_path(name: str) -> Path:
    return Path(current_app.config["DATA_DIR"]) / name


def read_yaml(name: str) -> dict:
    path = data_path(name)
    if not path.exists():
        return {}
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except (OSError, yaml.YAMLError) as exc:
        print(f"[settings] {path} を読めませんでした（{exc}）")
        return {}
    return data if isinstance(data, dict) else {}


def write_yaml(name: str, data: dict) -> None:
    _write(data_path(name), yaml.safe_dump(data, allow_unicode=True, sort_keys=False))


def _write(path: Path, text: str) -> None:
    with _lock:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(path.name + ".tmp")
        tmp.write_text(text, encoding="utf-8")
        os.replace(tmp, path)
        try:
            os.chmod(path, 0o600)  # APIキーを含むことがあるので所有者だけに絞る
        except OSError:
            pass


def lock() -> threading.RLock:
    return _lock
