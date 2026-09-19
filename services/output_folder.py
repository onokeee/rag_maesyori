"""保存先フォルダ（作った Markdown をダウンロードせず、選んだフォルダに直接書く。design.md 2.6・3.3）。

ふつうは LightRAG サーバーの INPUT_DIR を選ぶ。LightRAG の [スキャン]（POST /documents/scan）は
INPUT_DIR の直下にあるファイルだけを読み、サブフォルダは読まない（lightrag/api/routers/document_routes.py
1528-1535 行の iter_new_files: os.scandir で直下だけを見て、ファイル以外と対応していない拡張子を飛ばす）。
なので、RAG に入れない管理用のファイル（正規化データ・問題一覧・取込レポート）は、サブフォルダ
_管理用_RAGには入れない/<取り込み名>_<日時>/ に置く。

設定は data/output_settings.yaml に置く（取り込んだデータではなく設定なので残す。design.md 3.3）。
保存は「ダウンロードと同じ中身」を1ファイルずつ書く。書き終えるまでは一時ファイル（拡張子 .tmp。LightRAG の
スキャンは .tmp を読まない）に書き、書けたら os.replace で本来の名前にする。途中で1つでも失敗したら、
この回に書いたファイルを消し（上書きしたファイルは元に戻し）、アプリのデータは消さない。
"""
from __future__ import annotations

import os
import re
import secrets
import shutil
import stat
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from flask import current_app

from core.files import REMOVE_RETRIES, REMOVE_RETRY_WAIT
from services import settings_store

SETTINGS_FILE = "output_settings.yaml"
# 管理用のファイルを入れるサブフォルダ（LightRAG のスキャンはサブフォルダを読まない）
ADMIN_SUBDIR = "_管理用_RAGには入れない"
# LightRAG が取り込み終えたファイルを移すフォルダ（lightrag/constants.py 373行 PARSED_DIR_NAME）
LIGHTRAG_PARSED_DIR = "__parsed__"
CONFLICT_POLICIES = {"stop": "保存を止めて知らせる", "overwrite": "上書きする"}
DEFAULT_POLICY = "stop"

# 同じものを2回押した（ダブルクリック・2つのタブ）ときに、書き込みと削除が重ならないようにする
_save_lock = threading.Lock()


# ---- 設定 ---------------------------------------------------------------------------

def load() -> dict:
    """保存先フォルダの設定。folder が空なら機能はオフ。"""
    data = settings_store.read_yaml(SETTINGS_FILE)
    policy = str(data.get("on_conflict") or DEFAULT_POLICY)
    return {"folder": str(data.get("folder") or "").strip(),
            "save_admin": bool(data.get("save_admin")),
            "on_conflict": policy if policy in CONFLICT_POLICIES else DEFAULT_POLICY}


def configured_folder() -> str:
    """設定された保存先フォルダ（未設定なら ''）。ボタンを出すかどうかに使う。"""
    return load()["folder"]


def clean_folder_text(text) -> str:
    """入力されたパスの前後の空白と、エクスプローラーの「パスのコピー」が付ける " を外す。"""
    s = str(text or "").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        s = s[1:-1].strip()
    return s


def _app_dirs() -> list[Path]:
    """このアプリ自身のデータの置き場所（保存先にすると、起動時の片付けや取り込みの削除で消える）。"""
    cfg = current_app.config
    return [Path(d).resolve() for d in (cfg.get("UPLOAD_DIR"), cfg.get("DATA_DIR"), cfg.get("TABLES_DIR")) if d]


def _inside(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


# Windows のパスの長さの上限（MAX_PATH 260 から終端の1文字を引いた数）。フォルダを作るときは 8.3 形式の名前の分を
# さらに空けておく必要がある（CreateDirectoryW は 248 未満）。レジストリの LongPathsEnabled が 1 なら上限は無い。
MAX_PATH_CHARS = 259
MAX_DIR_CHARS = 247
# 一時ファイル・退避ファイルの名前の長さ（.rag_<16桁>.tmp / .bak）と、確かめる一時ファイルの名前の長さ
_TEMP_NAME_CHARS = len(".rag_") + 16 + len(".tmp")
_PROBE_NAME_CHARS = len(".rag_probe_") + 16 + len(".tmp")
# ふつうの md の名前が入るように、フォルダのあとに残しておきたい文字数（一覧表の名前は分割ヒントで 40 文字ほど長い）
_ROOM_FOR_NAMES = 120
_EXAMPLE_FOLDER = "D:\\LightRAG\\inputs"


def _long_paths_enabled() -> bool:
    if os.name != "nt":
        return True
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem") as key:
            value, _kind = winreg.QueryValueEx(key, "LongPathsEnabled")
        return bool(value)
    except OSError:
        return False


def path_limit() -> int | None:
    """書けるファイルのパスの長さの上限（上限が無ければ None）。"""
    return None if _long_paths_enabled() else MAX_PATH_CHARS


def path_warning(folder) -> str:
    """フォルダのパスが長く、ふつうの名前の md が書けないおそれがあるときの注意（無ければ ''）。"""
    limit = path_limit()
    if limit is None or not folder:
        return ""
    room = limit - len(str(folder)) - 1
    if room >= _ROOM_FOR_NAMES:
        return ""
    return (f"保存先フォルダのパスが長いため（{len(str(folder))}文字）、ファイル名に使えるのは{max(room, 0)}文字までです。"
            f"Windows ではパス全体が{limit}文字を超えるファイルを書けないので、名前の長い md は保存できません。"
            f"ドライブに近いフォルダ（例: {_EXAMPLE_FOLDER}）にすると安全です")


def _probe_write(folder: Path) -> None:
    """一時ファイルを1つ作って消せるか確かめる（書けなければ OSError）。

    tempfile.mkstemp は使わない。Windows でアクセス権（ACL）で書き込みを断られたフォルダだと、mkstemp は
    PermissionError を「名前がぶつかった」とみなして名前を変えながら試し続け、戻ってこない（CPython の
    tempfile._mkstemp_inner。os.access は読み取り専用の属性しか見ないため）。ここでは1回だけ試す。
    """
    path = folder / f".rag_probe_{secrets.token_hex(8)}.tmp"
    fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY | getattr(os, "O_BINARY", 0))
    try:
        os.write(fd, b"probe")
    finally:
        os.close(fd)
        os.unlink(path)


def check_folder(text) -> tuple[Path | None, list[str]]:
    """保存先フォルダとして使えるか確かめる。戻り値: (解決したパス, 使えない理由の一覧)。"""
    raw = clean_folder_text(text)
    if not raw:
        return None, ["保存先フォルダが入力されていません"]
    path = Path(raw)
    if not path.is_absolute():
        return None, ["フォルダはドライブ名から書いてください（例: D:\\LightRAG\\inputs）。相対パスは使えません"]
    try:
        resolved = path.resolve()
    except OSError:
        return None, ["このパスは読めません。書き方を確かめてください"]
    if not resolved.exists():
        return resolved, ["このフォルダがありません。先にエクスプローラーでフォルダを作ってください"]
    if not resolved.is_dir():
        return resolved, ["これはファイルです。フォルダを指定してください"]
    for app_dir in _app_dirs():
        if _inside(resolved, app_dir):
            return resolved, ["このアプリがデータを置くフォルダの中は保存先にできません"
                              "（アプリの片付けで消されることがあります）"]
    limit = path_limit()
    if limit is not None and len(str(resolved)) + 1 + _PROBE_NAME_CHARS > limit:
        return resolved, [f"フォルダのパスが長すぎます（{len(str(resolved))}文字）。Windows ではパス全体が{limit}文字を"
                          f"超えるファイルを書けません。ドライブに近いフォルダ（例: {_EXAMPLE_FOLDER}）を指定してください"]
    try:
        _probe_write(resolved)
    except OSError as exc:
        return resolved, [f"このフォルダにファイルを書き込めません（{exc.strerror or exc.__class__.__name__}）。"
                          "書き込みの権限を確かめてください"]
    return resolved, []


def save_settings(folder_text, save_admin: bool, on_conflict: str) -> dict:
    """設定を確かめてから保存する。使えないときは ValueError（理由を / でつないだ文）。"""
    raw = clean_folder_text(folder_text)
    policy = on_conflict if on_conflict in CONFLICT_POLICIES else DEFAULT_POLICY
    folder = ""
    if raw:
        resolved, errors = check_folder(raw)
        if errors:
            raise ValueError(" / ".join(errors))
        folder = str(resolved)
    data = {"folder": folder, "save_admin": bool(save_admin), "on_conflict": policy}
    settings_store.write_yaml(SETTINGS_FILE, data)
    return data


def folder_report(folder: Path) -> dict:
    """[フォルダを確かめる] の補足（LightRAG の INPUT_DIR らしいか、.md が何件あるか）。"""
    try:
        md_count = sum(1 for p in folder.iterdir() if p.suffix == ".md" and p.is_file())
    except OSError:
        md_count = 0
    return {"parsed_dir": (folder / LIGHTRAG_PARSED_DIR).is_dir(), "md_count": md_count}


# ---- 保存 ---------------------------------------------------------------------------

class SaveConflict(Exception):
    """同じ名前のファイルがある（方針が「止める」）。何も書いていない。"""

    def __init__(self, folder: Path, existing: list[str], parsed: list[str]):
        super().__init__("同じ名前のファイルがあります")
        self.folder, self.existing, self.parsed = folder, existing, parsed


class SaveError(Exception):
    """保存できなかった（フォルダが使えない・書き込みの失敗）。この回に書いたファイルは消してある。

    status: 画面の HTTP ステータス（設定・名前の問題は 400、書き込みの失敗は 500）。
    leftovers: 失敗したあと、消せずに残ってしまったファイル（ウイルス対策・同期ソフトが掴んでいたなど）。
    """

    def __init__(self, message: str, status: int = 400, leftovers: list[Path] | None = None):
        super().__init__(message)
        self.status = status
        self.leftovers = leftovers or []


@dataclass
class SaveResult:
    folder: Path
    md_names: list[str]
    overwritten: list[str] = field(default_factory=list)
    # 上書きの方針のとき、LightRAG で同じ文書とみなされる別のファイル（ヒントだけ違う名前・__parsed__/ の取り込み済み）
    already_parsed: list[str] = field(default_factory=list)
    admin_dir: Path | None = None
    admin_names: list[str] = field(default_factory=list)
    # 上書きの前に退避した元のファイルのうち、消せずに残ったもの（LightRAG は .bak を読まない）
    leftover_backups: list[Path] = field(default_factory=list)


@contextmanager
def saving():
    """保存と、そのあとの削除を1つずつ行う（同じものを2回押しても2回目は「もう無い」になる）。"""
    with _save_lock:
        yield


# LightRAG 1.5.7 が同じ文書とみなす名前（lightrag/parser/routing.py 62行 _PARSER_HINT_RE・1085行
# canonicalize_parser_hinted_basename: 末尾の「.[ヒント]」を外した名前が文書ID（ファイル名の MD5）の元になる。
# __parsed__ に移したファイルは、同名があると「_001」などを足す（lightrag/utils.py 917行 move_file_to_parsed_dir、
# document_routes.py 164行 ARCHIVED_FILE_SUFFIX_RE）。Windows では大文字・小文字の違いも同じファイルになる。
_HINT_RE = re.compile(r"\.\[([^\]]*)\](\.[^.]+)$")
_ARCHIVED_SUFFIX_RE = re.compile(r"_(?:\d{3}|\d{10,})$")


def lightrag_key(name: str, archived: bool = False) -> str:
    """LightRAG で同じ文書になるかを比べるための名前（ヒントを外し、大文字・小文字をそろえる）。"""
    if archived:
        path = Path(name)
        name = f"{_ARCHIVED_SUFFIX_RE.sub('', path.stem)}{path.suffix}"
    return _HINT_RE.sub(r"\2", name).casefold()


def _listing(directory: Path) -> list[str]:
    try:
        return [e.name for e in os.scandir(directory) if e.is_file()]
    except (FileNotFoundError, NotADirectoryError):
        return []


def find_clashes(folder: Path, names: list[str]) -> tuple[list[str], list[str], list[str]]:
    """保存しようとする名前と重なるファイル。

    戻り値: (保存先の直下の同じ名前, 直下にある「ヒントだけ違う」別名, __parsed__ にある同じ文書の名前)
    """
    wanted = {lightrag_key(n) for n in names}
    exact = {n.casefold() for n in names}
    existing, aliases = [], []
    for entry in sorted(_listing(folder)):
        if lightrag_key(entry) not in wanted:
            continue
        (existing if entry.casefold() in exact else aliases).append(entry)
    # __parsed__ の名前の末尾の _001 は、LightRAG が足した番号とも、もとの名前の一部（No._123 など）とも読める。
    # どちらの読み方でも同じ文書になるなら重なりとみなす（LightRAG の削除も __parsed__ では番号を外して比べる:
    # document_routes.py 2115〜2126行 canonicalize_archived_file_variant_basename・2167〜2185行）
    parsed = sorted(e for e in _listing(folder / LIGHTRAG_PARSED_DIR)
                    if lightrag_key(e) in wanted or lightrag_key(e, archived=True) in wanted)
    return existing, aliases, parsed


def _check_name(folder: Path, name: str) -> Path:
    """フォルダの直下に置くファイル名か確かめる（名前は core.naming で作るが、念のため外に出ないことを確かめる）。"""
    if not name or name in (".", "..") or Path(name).name != name or "/" in name or "\\" in name or ":" in name:
        raise SaveError("ファイル名に使えない文字が含まれていたため、保存しませんでした")
    target = folder / name
    if target.resolve().parent != folder.resolve():
        raise SaveError("保存先フォルダの外を指すファイル名だったため、保存しませんでした")
    return target


def _check_length(path: Path, limit: int | None, *, directory: bool = False) -> None:
    """パスが Windows の上限を超えないか確かめる（ファイルは一時ファイル .rag_….tmp の名前の長さも含めて）。"""
    if limit is None:
        return
    if directory:
        length, cap = len(str(path)), min(limit, MAX_DIR_CHARS)
    else:
        length, cap = len(str(path.parent)) + 1 + max(len(path.name), _TEMP_NAME_CHARS), limit
    if length > cap:
        raise SaveError(f"パスが長すぎて Windows で書けないため、保存しませんでした（「{path.name}」、{len(str(path))}文字。"
                        f"上限は{cap}文字）。保存先フォルダをドライブに近いフォルダ（例: {_EXAMPLE_FOLDER}）にするか、"
                        "取り込み名・帳票の種類の名前を短くしてください")


def _check_admin_base(folder: Path) -> None:
    """管理用のサブフォルダが、保存先フォルダの中の普通のフォルダか確かめる（ジャンクション・リンクで外へ出ない）。"""
    base = folder / ADMIN_SUBDIR
    if base.is_symlink() or base.is_junction():
        raise SaveError(f"「{ADMIN_SUBDIR}」がリンク（ジャンクション）になっていて、保存先フォルダの外を指すおそれがあるため、"
                        "保存しませんでした。普通のフォルダにしてください")
    if base.exists() and (not base.is_dir() or base.resolve().parent != folder.resolve()):
        raise SaveError(f"「{ADMIN_SUBDIR}」がフォルダではないか、保存先フォルダの外を指しているため、保存しませんでした")


def _fresh_admin_dir(folder: Path, stem: str) -> Path:
    base = folder / ADMIN_SUBDIR
    name = f"{stem}_{datetime.now():%Y%m%d_%H%M%S}"
    candidate, n = base / name, 2
    while candidate.exists():
        candidate = base / f"{name}_{n}"
        n += 1
    return candidate


def _write_bytes(tmp: Path, data: bytes) -> None:
    """一時ファイルに書いて、ディスクまで書き出す（書き終えてから本来の名前にするので、途中の中身は見えない）。"""
    with open(tmp, "wb") as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())


def write_files(md_files: list[tuple[str, bytes]], admin_files: list[tuple[str, bytes]] | None = None,
                admin_stem: str = "") -> SaveResult:
    """md を保存先フォルダの直下に、管理用のファイルを _管理用_RAGには入れない/<admin_stem>_<日時>/ に書く。

    すべて書けたときだけ SaveResult を返す。同じ名前があって方針が「止める」なら SaveConflict（何も書かない）、
    書き込みに失敗したら、この回に書いたものを消して（上書きしたものは戻して）SaveError。
    """
    settings = load()
    if not settings["folder"]:
        raise SaveError("保存先フォルダが設定されていません。設定 → 保存先フォルダ で決めてください")
    folder, errors = check_folder(settings["folder"])
    if errors:
        raise SaveError("保存先フォルダが使えません: " + " / ".join(errors))
    overwrite = settings["on_conflict"] == "overwrite"

    names = [name for name, _data in md_files]
    targets = [_check_name(folder, name) for name in names]
    if len({lightrag_key(n) for n in names}) != len(names):
        raise SaveError("LightRAG で同じ文書とみなされるファイルが2つ作られるため、保存しませんでした")
    for t in targets:
        if t.exists() and not t.is_file():
            raise SaveError("保存先に同じ名前のフォルダがあるため、保存しませんでした")
    limit = path_limit()
    for t in targets:
        _check_length(t, limit)
    admin_dir = None
    if admin_files:
        _check_admin_base(folder)
        admin_dir = _fresh_admin_dir(folder, admin_stem or "取り込み")
        _check_length(admin_dir, limit, directory=True)
        for name, _data in admin_files:
            _check_length(admin_dir / name, limit)
    existing, aliases, parsed = find_clashes(folder, names)
    if (existing or aliases or parsed) and not overwrite:
        raise SaveConflict(folder, existing + aliases, parsed)

    written: list[Path] = []                    # この回に新しく作ったファイル
    replaced: list[tuple[Path, Path]] = []      # (上書きしたファイル, 元の中身を退避したファイル)
    temps: list[Path] = []
    created_dirs: list[Path] = []
    result = SaveResult(folder=folder, md_names=names)
    current = ""                                # いま書いているもの（失敗したときに名前を出す）
    try:
        for target, (_name, data) in zip(targets, md_files):
            current = target.name
            _place(target, data, written, replaced, temps)
            if any(r[0] == target for r in replaced):
                result.overwritten.append(target.name)
        if admin_dir is not None:
            current = admin_dir.name
            for d in (admin_dir.parent, admin_dir):
                if not d.exists():
                    d.mkdir()
                    created_dirs.append(d)
            if not _inside(admin_dir.resolve(), folder.resolve()):
                raise SaveError("管理用のフォルダが保存先フォルダの外を指しているため、保存しませんでした")
            for name, data in admin_files:
                current = name
                _place(_check_name(admin_dir, name), data, written, replaced, temps)
            result.admin_dir = admin_dir
            result.admin_names = [name for name, _data in admin_files]
    except BaseException as exc:
        leftovers = _rollback(written, replaced, temps, created_dirs)
        if isinstance(exc, (KeyboardInterrupt, SystemExit)):
            raise
        if isinstance(exc, SaveError):
            exc.leftovers = leftovers
            raise
        if isinstance(exc, OSError):
            reason = exc.strerror or exc.__class__.__name__
            if leftovers:
                message = (f"「{current}」を書き込めませんでした（{reason}）。データはこのアプリに残しました。"
                           f"ただし、この回に書いたファイルのうち{len(leftovers)}件を消せませんでした（下の一覧）。"
                           "LightRAG でスキャンする前に、エクスプローラーで消してください")
            else:
                message = (f"「{current}」を書き込めませんでした（{reason}）。"
                           "この回に書いた分は消し、データはこのアプリに残しました")
            raise SaveError(message, status=500, leftovers=leftovers) from exc
        raise
    for _target, backup in replaced:
        if not _remove(backup):
            result.leftover_backups.append(backup)
    result.already_parsed = aliases + [f"{LIGHTRAG_PARSED_DIR}/{n}" for n in parsed]
    return result


def _place(target: Path, data: bytes, written: list, replaced: list, temps: list) -> None:
    """1ファイルを置く。一時ファイルは同じフォルダに、LightRAG のスキャンが読まない拡張子（.tmp）で作る。"""
    tmp = target.with_name(f".rag_{secrets.token_hex(8)}.tmp")
    temps.append(tmp)   # 書き込みの途中で失敗しても消せるよう、作る前に控える
    _write_bytes(tmp, data)
    if target.exists():
        # 上書き: 元の中身を退避してから置き換える（失敗したら戻せるように）
        backup = target.with_name(f".rag_{secrets.token_hex(8)}.bak")
        os.replace(target, backup)
        replaced.append((target, backup))
        os.replace(tmp, target)
    else:
        os.replace(tmp, target)
        written.append(target)
    temps.remove(tmp)


def _rollback(written: list[Path], replaced: list[tuple[Path, Path]], temps: list[Path],
              created_dirs: list[Path]) -> list[Path]:
    """この回に書いたものを消し、上書きしたものを元に戻す。消せずに残ったファイルを返す。"""
    leftovers = [path for path in written if not _remove(path)]
    for target, backup in replaced:
        try:
            if backup.exists():
                os.replace(backup, target)
        except OSError:
            # 戻せなかった: 直下には新しい中身が、そばに元の中身（.bak）が残っている
            leftovers += [target, backup]
    leftovers += [tmp for tmp in temps if not _remove(tmp)]
    for d in reversed(created_dirs):
        shutil.rmtree(d, ignore_errors=True)
    return leftovers


def _remove(path: Path) -> bool:
    """ファイルを消す。消せたら（もともと無ければ）True。

    Windows ではウイルス対策・検索のインデックス・同期ソフトが作ったばかりのファイルを少しの間掴んでいることがあるので、
    core.files.remove_upload と同じく少し待って何度か試す。読み取り専用のファイル（上書きで退避した元のファイル）は、
    読み取り専用を外してから消す。
    """
    for attempt in range(REMOVE_RETRIES):
        try:
            path.unlink()
            return True
        except FileNotFoundError:
            return True
        except OSError:
            try:
                os.chmod(path, stat.S_IWRITE)
            except OSError:
                pass
            time.sleep(REMOVE_RETRY_WAIT * (attempt + 1))
    return not path.exists()
