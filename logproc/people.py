"""記入者の判定。人物一覧（別名・イニシャル）と担当列の名前で照合する。

人物一覧はアプリ内だけで使い、AIには送らない。
位置の形（日付直後の「田中：」、行末の「（田中）」、「田中→佐藤」など）に合うときだけ人名として扱う。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Iterable

from logproc.models import AuthorInfo

DEFAULT_GROUPS = [
    "保全G", "保全課", "保全", "製造課", "製造", "品証", "品質保証", "施設課", "技術", "生技",
    "メーカーFE", "メーカー", "業者", "夜勤者", "班長", "課長",
]

# 辞書にない名前でも「日付 名前 本文」の形で人名とみなす一般的な姓
COMMON_SURNAMES = set("""
佐藤 鈴木 高橋 田中 伊藤 渡辺 渡部 渡邊 山本 中村 小林 加藤 吉田 山田 佐々木 山口 松本 井上 木村 林 斎藤 斉藤
清水 山崎 森 池田 橋本 阿部 石川 山下 中島 石井 小川 前田 岡田 長谷川 藤田 後藤 近藤 村上 遠藤 青木 坂本
福田 太田 西村 藤井 金子 岡本 藤原 中野 三浦 原田 中川 松田 竹内 小野 田村 中山 和田 石田 森田 上田 原
柴田 酒井 工藤 横山 宮崎 宮本 内田 高木 安藤 島田 谷口 大野 高田 丸山 今井 河野 藤本 村田 武田 上野 杉山
増田 小山 大塚 平野 菅原 久保 松井 千葉 岩崎 桜井 木下 野口 松尾 菊地 野村 新井 小西 大西 西田 北村 石原
永井 荒木 本田 久保田 中西 浅野 服部 市川 飯田 片山 小島 水野 岡崎 西川 伊東 五十嵐 松下 吉川 山内 北川
""".split())

# 「〇〇：」の〇〇が人名ではない語
_COLON_STOP = set("""
原因 現象 対応 処置 暫定 恒久 結果 備考 内容 状況 理由 対策 再発防止 補足 注意 注記 予定 回答 連絡 報告 確認
作業 停止 復旧 部品 調査 点検 判断 結論 経過 方針 課題 問題 要因 件名 場所 時間 日時 期間 費用 金額 担当
記入者 対応者 設備 工程 品番 型番 数量 状態 目的 依頼 指示 保留 完了 未 済 注 訂正 追記 水平展開 異常 不具合
""".split())

_KANJI = r"[一-鿿々ヶ]"
TOKEN_RE = re.compile(rf"(?:{_KANJI}{{1,4}}(?:\({_KANJI}{{1,2}}\))?|[A-Za-z][A-Za-z.]{{0,14}}[A-Za-z.]?)")
_INITIALS_RE = re.compile(r"[A-Z]\.[A-Z]\.?")
_LATIN_NAME_RE = re.compile(r"[A-Z][a-z]{2,}")


def key(text: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text or "")).casefold()


@dataclass
class Person:
    name: str
    aliases: list[str] = field(default_factory=list)
    org: str = ""


class PeopleIndex:
    """登録人物・担当列の名前・部署名の索引。"""

    def __init__(self, registered: Iterable[dict | Person] | None = None, column_names: Iterable[str] = (),
                 groups: Iterable[str] | None = None):
        self.persons: list[Person] = []
        self._exact: dict[str, list[Person]] = {}     # 氏名・別名
        self._surname: dict[str, list[Person]] = {}   # 姓
        for item in registered or []:
            p = item if isinstance(item, Person) else Person(
                str(item.get("name", "")).strip(), [str(a) for a in item.get("aliases", []) or []], str(item.get("org", "") or ""))
            if p.name:
                self._add(p)
        for raw in column_names or []:
            for part in re.split(r"[/／、,・]", str(raw or "")):
                part = part.strip()
                if not part or len(part) > 12:
                    continue
                k = key(part)
                if k in self._exact or k in self._surname:
                    continue
                self._add(Person(part))
        self.groups = list(groups) if groups is not None else list(DEFAULT_GROUPS)
        self._group_keys = {key(g): g for g in self.groups}
        # 本文先頭で区切りなしに照合する表記（長い順）
        forms = {unicodedata.normalize("NFKC", f) for f in list(self._exact_forms()) + self.groups}
        self._forms = sorted((f for f in forms if len(f) >= 2), key=len, reverse=True)

    def _add(self, p: Person) -> None:
        self.persons.append(p)
        for form in [p.name, *p.aliases]:
            self._exact.setdefault(key(form), []).append(p)
        parts = unicodedata.normalize("NFKC", p.name).split()
        if len(parts) >= 2:
            self._surname.setdefault(key(parts[0]), []).append(p)

    def _exact_forms(self):
        for p in self.persons:
            yield p.name
            yield from p.aliases
            parts = p.name.split()
            if len(parts) >= 2:
                yield parts[0]
                yield "".join(parts)

    # ---- 照合 ----
    def is_group(self, token: str) -> bool:
        return key(token) in self._group_keys

    def is_known(self, token: str) -> bool:
        k = key(token)
        return k in self._exact or k in self._surname or k in self._group_keys

    def group_at(self, sh: str, p: int, end: int) -> str | None:
        """p から始まる部署名（後ろが空白・「:」・行末のもの）。"""
        for g in sorted(self.groups, key=len, reverse=True):
            gn = unicodedata.normalize("NFKC", g)
            q = p + len(gn)
            if q <= end and sh.startswith(gn, p) and (q == end or sh[q] in " \t:"):
                return gn
        return None

    def known_at(self, sh: str, p: int, end: int) -> str | None:
        """p から区切りなしで始まる登録済みの名前（「4/3佐藤エンコーダ…」用）。"""
        for f in self._forms:
            if p + len(f) <= end and sh.startswith(f, p) and not self.is_group(f):
                return f
        return None

    def names(self) -> list[str]:
        """マスク用の名前一覧（氏名・姓・別名）。"""
        out = set()
        for f in self._exact_forms():
            if len(f) >= 2:
                out.add(f)
        return sorted(out, key=len, reverse=True)

    def accept_colon(self, token: str) -> bool:
        if token in _COLON_STOP:
            return False
        if self.is_known(token) or _INITIALS_RE.fullmatch(token) or _LATIN_NAME_RE.fullmatch(token):
            return True
        base = re.sub(r"\(.*\)$", "", token)
        return bool(re.fullmatch(rf"{_KANJI}{{2,4}}", base)) and base not in _COLON_STOP

    def accept_weak(self, token: str) -> bool:
        if self.is_known(token) or _INITIALS_RE.fullmatch(token):
            return True
        return re.sub(r"\(.*\)$", "", token) in COMMON_SURNAMES

    def resolve(self, raw: str) -> AuthorInfo:
        """表記から AuthorInfo を作る（未登録でも返す）。"""
        k = key(raw)
        if k in self._group_keys:
            return AuthorInfo(raw, self._group_keys[k], False, "")
        cands = self._exact.get(k)
        if cands:
            uniq = _uniq(cands)
            if len(uniq) == 1:
                p = uniq[0]
                note = f"人物一覧の別名「{raw}」から特定した" if key(p.name) != k else ""
                return AuthorInfo(raw, p.name, False, note)
            return AuthorInfo(raw, raw, False, _ambiguous_note(uniq))
        cands = self._surname.get(k)
        if cands:
            uniq = _uniq(cands)
            if len(uniq) == 1:
                return AuthorInfo(raw, uniq[0].name, False, "")
            return AuthorInfo(raw, raw, False, _ambiguous_note(uniq))
        if _INITIALS_RE.fullmatch(unicodedata.normalize("NFKC", raw)):
            return AuthorInfo(raw, None, False, "人物一覧にないイニシャルのため、誰か特定していない")
        return AuthorInfo(raw, raw, False, "")


def _uniq(persons: list[Person]) -> list[Person]:
    seen, out = set(), []
    for p in persons:
        if id(p) not in seen:
            seen.add(id(p))
            out.append(p)
    return out


def _ambiguous_note(persons: list[Person]) -> str:
    names = "／".join(p.name for p in persons)
    return f"人物一覧に{len(persons)}人（{names}）いるため、どちらか特定していない"


def _skip(sh: str, p: int, end: int, chars: str) -> int:
    while p < end and sh[p] in chars:
        p += 1
    return p


def _token_at(sh: str, p: int, end: int) -> str | None:
    m = TOKEN_RE.match(sh[:end], p)
    return m.group(0) if m else None


def detect_head_author(sh: str, p: int, end: int, index: PeopleIndex) -> tuple[AuthorInfo | None, int]:
    """本文先頭（日時の直後）の記入者。戻り値: (記入者, 本文の開始位置)"""
    p = _skip(sh, p, end, " \t:")
    g = index.group_at(sh, p, end)
    if g:
        q = p + len(g)
        q2 = _skip(sh, q, end, " \t")
        t = _token_at(sh, q2, end) if q2 > q else None
        if t and index.accept_weak(t) and not index.is_group(t):
            after = q2 + len(t)
            if after == end or sh[after] in " \t:":
                info = index.resolve(t)
                info.raw = sh[p:after]
                return info, _skip(sh, after, end, " \t:")
        return index.resolve(g), _skip(sh, q, end, " \t:")
    t = _token_at(sh, p, end)
    if t:
        q = p + len(t)
        m = re.compile(r"\s*→\s*").match(sh[:end], q)
        if m:
            t2 = _token_at(sh, m.end(), end)
            if t2 and index.accept_weak(t) and index.accept_weak(t2):
                after = m.end() + len(t2)
                info = index.resolve(t)
                info.raw = sh[p:after]
                info.note = (info.note + "。" if info.note else "") + f"原文は「{sh[p:after]}」（引継ぎまたは連絡の相手は{t2}）"
                return info, _skip(sh, after, end, " \t:")
        r = _skip(sh, q, end, " \t")
        if r < end and sh[r] == ":" and index.accept_colon(t):
            return index.resolve(t), _skip(sh, r + 1, end, " \t")
        if (r > q or r == end) and index.accept_weak(t):
            return index.resolve(t), r
    k = index.known_at(sh, p, end)
    if k:
        return index.resolve(k), p + len(k)
    return None, p


_TAIL_PAREN_RE = re.compile(rf"(?:(\d{{1,2}}/\d{{1,2}})\s*)?({TOKEN_RE.pattern}|[^\s()]{{2,8}})?")
_TAIL_BARE_RE = re.compile(rf"(?<=[\s)）])({_KANJI}{{1,4}})$")


def detect_tail_author(sh: str, start: int, end: int, index: PeopleIndex) -> tuple[AuthorInfo | None, int, tuple[int, int] | None]:
    """末尾の「（田中）」「(4/3 西村)」「…）小西」。戻り値: (記入者, 本文の終了位置, 括弧内の日付の位置)"""
    e = end
    while e > start and sh[e - 1] in " \t。、/→":
        e -= 1
    if e - start >= 3 and sh[e - 1] == ")":
        o = sh.rfind("(", start, e - 1)
        if o > start:
            inner_start = _skip(sh, o + 1, e - 1, " \t")
            inner = sh[inner_start:e - 1].rstrip()
            m = _TAIL_PAREN_RE.fullmatch(inner)
            if m and (m.group(1) or m.group(2)):
                tok = m.group(2)
                if tok and not (index.accept_weak(tok) or index.is_group(tok)):
                    return None, end, None
                info = index.resolve(tok) if tok else None
                dspan = (inner_start + m.start(1), inner_start + m.end(1)) if m.group(1) else None
                return info, o, dspan
    m = _TAIL_BARE_RE.search(sh[start:e])
    if m and index.accept_weak(m.group(1)) and start + m.start(1) > start:
        return index.resolve(m.group(1)), start + m.start(1), None
    return None, end, None


def inherit_authors(authors: list[AuthorInfo | None], skip: set[int] | None = None) -> list[AuthorInfo | None]:
    """記入者の書かれていないセグメントに直前の記入者を（推定）で付ける。"""
    skip = skip or set()
    out: list[AuthorInfo | None] = []
    last: AuthorInfo | None = None
    for i, a in enumerate(authors):
        if i in skip:
            out.append(a)
            continue
        if a is None and last is not None:
            label = last.name or last.raw
            a = AuthorInfo("", label, True, f"原文になく、直前の{label}を引き継いだ")
        if a is not None:
            last = a
        out.append(a)
    return out
