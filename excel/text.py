"""セル値の正規化と型変換。"""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, time

from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel

_EDGE_CHARS = "[]()<>【】〈〉《》「」『』■□◆◇●○・*※:;."
_BRACKETS = {"[": "]", "(": ")", "<": ">", "【": "】", "〈": "〉", "《": "》", "「": "」", "『": "』"}
_CLOSERS = {close: open_ for open_, close in _BRACKETS.items()}
_EDGE_MARKS = "".join(ch for ch in _EDGE_CHARS if ch not in _BRACKETS and ch not in _CLOSERS)
# 見出しの先頭の項番: 「3.」「3)」「3、」「(3)」「D3」（8D の D1〜D8）「A.機構部」。丸数字は section_stripped で別に見る
_SECTION_NO_RE = re.compile(r"^(?:\d{1,2}[.)、](?!\d)|\(\d{1,2}\)|d[1-8](?=\D)|[a-z][.)、](?=[^\x00-\x7f]))")
_TRAILING_PAREN_RE = re.compile(r"\([^()]{1,12}\)$")
_INLINE_SEP = re.compile(r"[:：]")
_DATE_RE = re.compile(r"(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")
# 年の無い日付（「2/12」「2/12 3時17分」「2月12日 14:05」）。年は推測せず、原文のまま残して警告だけ出す。
# 「24/8/25」のような2桁の年は、年の欄が空なのか2桁で書いたのか分からないのでこの形には含めない
_NO_YEAR_RE = re.compile(r"^\s*(\d{1,2})\s*(?:/|月|-|\.)\s*(\d{1,2})\s*日?(?:\s|$|[^\d/\-.])")
# 日付の直後の時刻「2026/2/12 14:05」
_CLOCK_RE = re.compile(r"\s*(\d{1,2}):(\d{2})(?!\d)")
_ERA_RE = re.compile(r"(?:令和|R)\s*(\d{1,2}|元)\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})")
_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")
_DURATION_RE = re.compile(r"(\d+(?:\.\d+)?)\s*時間\s*(\d+(?:\.\d+)?)\s*分")
# 時刻の範囲「09:30-12:45」「9:30～13:55」（作業時間の欄によくある書き方。先頭の 9 を数値として読まない）
# 「12/24 21:53-12/25 11:09」のように各時刻の前に月日が付いた書き方も範囲として見る（先頭の月を数値として読まない）
_TIME_RANGE_RE = re.compile(r"(?:\d{1,2}/\d{1,2}\s*)?(\d{1,2})\s*:\s*(\d{2})\s*[-~〜ー―]\s*"
                            r"(?:\d{1,2}/\d{1,2}\s*)?(\d{1,2})\s*:\s*(\d{2})")
# 時刻の範囲に添えた時間数「（3.2h）」「（工数 8.75h）」「(195分)」
_TIME_AMOUNT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(時間|hours?|hrs?|h|分|mins?)(?![A-Za-z])", re.IGNORECASE)
_SPACES_RE = re.compile(r"[^\S\n]+")
# 見出し末尾の単位: 「作業時間(h)」「停止時間（分）」「金額[円]」
_HEADER_UNIT_RE = re.compile(r"^(.+?)\s*[(\[]\s*([^()\[\]\d]{1,6})\s*[)\]]$")
# 値の末尾の単位: 「1.5時間」「95分」「120 min」
_VALUE_UNIT_RE = re.compile(r"^-?\d+(?:\.\d+)?\s*([^\d\s.,\-]{1,4})$")
# 前後に言葉や括弧書きのある値（「約90分」「595分（9.9h）」）の、最初の数値の直後の単位
_UNIT_AFTER_RE = re.compile(r"\s*([^\d\s.,\-:/~()\[\]{}<>、。・]{1,4})(?![^\d\s.,\-:/~()\[\]{}<>、。・])")
# 会計の書き方の負号「▲50万円」「△0.8」（一覧表の側と同じく負の数として読む）
_MINUS_MARK_RE = re.compile(r"^[△▲]\s*(?=\d)")
# 見出しの括弧書きのうち単位とみなす和文（「発生原因（推定）」「担当（記入）」の括弧書きは単位でない）。
# 英字・記号の単位（h, min, mm, %, ℃）と1文字の和文（分・円・枚・個）はこの一覧に無くても単位とみなす
_JA_UNITS = {"時間", "千円", "万円", "百万円", "人日", "人時", "日間", "ヶ月", "か月", "カ月", "箇所", "ケ所"}
UNIT_ALIASES = {"h": "時間", "hr": "時間", "hrs": "時間", "hour": "時間", "hours": "時間",
                "min": "分", "mins": "分", "sec": "秒", "yen": "円", "¥": "円"}

MAX_LABEL_LENGTH = 20

# チェックボックス表記「■重大　□大　□中」「☑同型機　□類似設備」
CHECKED_MARKS = "■☑☒✓✔"
_CHECKBOX_RE = re.compile(r"([■□☑☐☒✓✔])\s*([^■□☑☐☒✓✔\s]+)")


def normalize_label(text) -> str:
    """ラベル比較用の正規化。全角/半角・空白・前後の記号・大文字小文字の揺れを吸収する。

    例: 「設備№：」「 設備 No. 」「【設備NO】」→ "設備no"
    """
    s = unicodedata.normalize("NFKC", str(text))
    s = re.sub(r"\s+", "", s)
    return _strip_edges(s).lower()


def _strip_edges(s: str) -> str:
    """前後の記号を除く。括弧は対になっているときだけ外す（「停止時間(分)」の「)」だけを消さない）。"""
    prev = None
    while s and s != prev:
        prev = s
        s = s.strip(_EDGE_MARKS)
        if not s:
            break
        close = _BRACKETS.get(s[0])
        if close and s.endswith(close) and len(s) > 1:
            s = s[1:-1]
            continue
        if close and close not in s[1:]:
            s = s[1:]
        if s and s[-1] in _CLOSERS and _CLOSERS[s[-1]] not in s[:-1]:
            s = s[:-1]
    return s


def section_stripped(text) -> str:
    """見出しの先頭の項番を除いた正規化ラベル。「３．暫定対策」「①何が」「(2)原因」「D2 問題の記述」→ 項番なし。

    項番が無い、または除くと文字が残らない（「1.5」など）ときは "" を返す。
    「2号機」のような数字始まりの値は項番とみなさない（区切りの「.」「)」「、」か丸数字が必要）。
    """
    raw = str(text).strip()
    if raw[:1] and "①" <= raw[0] <= "⑳":
        rest = normalize_label(raw[1:])
    else:
        compact = re.sub(r"\s+", "", unicodedata.normalize("NFKC", raw)).lower()
        m = _SECTION_NO_RE.match(compact)
        rest = normalize_label(compact[m.end():]) if m else ""
    return rest if any(ch.isalpha() for ch in rest) else ""


def label_parts(text) -> tuple[str, ...]:
    """「ライン／工程」のように「／」で2〜3個のラベルをまとめた見出しの、各部分の正規化ラベル。

    各部分が2文字以上で文字を含むときだけ分ける（「L/min」は分けない）。
    """
    s = unicodedata.normalize("NFKC", str(text or "")).strip()
    if "\n" in s or "/" not in s:
        return ()
    parts = [normalize_label(p) for p in s.split("/")]
    if not 2 <= len(parts) <= 3 or any(len(p) < 2 or not any(ch.isalpha() for ch in p) for p in parts):
        return ()
    return tuple(parts)


def split_combined_value(text, count: int) -> list[str] | None:
    """「L6／STI-CMP」を count 個に分ける。区切りの数が合わなければ None。"""
    parts = [p.strip() for p in re.split(r"[/／]", str(text or ""))]
    return parts if len(parts) == count and all(parts) else None


# 設備番号らしい記号: 英字と数字を両方含む空白なしの英数字（CMP-108, ROB-821, EQ001）
_CODE_TOKEN = r"(?=[A-Za-z0-9\-_.#]*[A-Za-z])(?=[A-Za-z0-9\-_.#]*\d)[A-Za-z0-9][A-Za-z0-9\-_.#]{2,19}"
_CODE_NAME_PATTERNS = (
    (re.compile(rf"^({_CODE_TOKEN})\s*\((.+)\)$"), 1, 2),        # CVD-202（TEOS CVD 2号機）
    (re.compile(rf"^(.+?)\s*\(({_CODE_TOKEN})\)$"), 2, 1),        # TEOS CVD 2号機（CVD-202）
    (re.compile(rf"^({_CODE_TOKEN})(?:\s*[/:]\s*|\s+)(.+)$"), 1, 2),  # CMP-108　STI-CMP 8号機 / CMP-108：STI-CMP
    (re.compile(rf"^(.+?)(?:\s*[/:]\s*|\s+)({_CODE_TOKEN})$"), 2, 1),  # W-CMP 3号機　CMP-103
)
_CODE_ONLY = re.compile(rf"^{_CODE_TOKEN}$")


def split_code_name(text) -> tuple[str, str] | None:
    """「CMP-108　STI-CMP 8号機」「ROB-821（ウェーハソーター 1号機）」→ (設備番号, 設備名)。分けられなければ None。

    1つのセルに「対象設備」「使用設備」として番号と名前をまとめて書く帳票用。名前の側が番号だけ（「CMP-101 / CMP-102」）なら分けない。
    """
    s = unicodedata.normalize("NFKC", str(text or "")).strip()
    if not s or "\n" in s:
        return None
    for pattern, code_group, name_group in _CODE_NAME_PATTERNS:
        m = pattern.match(s)
        if m:
            code, name = m[code_group].strip(), m[name_group].strip()
            if name and not _CODE_ONLY.match(name) and not all(_CODE_ONLY.match(p) for p in re.split(r"[\s,、/]+", name) if p):
                return code, name
    return None


def paren_stripped(norm: str) -> str:
    """末尾の括弧書きを除いた正規化ラベル。「停止時間(分)」→「停止時間」。括弧が無ければ ""。"""
    m = _TRAILING_PAREN_RE.search(norm)
    if not m or m.start() == 0:
        return ""
    return norm[: m.start()]


_LABEL_SUFFIXES = ("内容", "欄", "日時")


def label_base(norm: str) -> str:
    """表記の揺れを比べるための見出しの基本形。末尾の括弧書きと「内容」「欄」「日時」を除く。変わらなければ ""。

    例: 「発生原因(なぜ起きたか)」→「発生原因」、「応急処置内容」→「応急処置」、「復旧完了日時」→「復旧完了」。
    """
    s = paren_stripped(norm) or norm
    for suffix in _LABEL_SUFFIXES:
        if s.endswith(suffix) and len(s) - len(suffix) >= 2:
            s = s[: -len(suffix)]
            break
    return s if s != norm else ""


def normalize_sheet_name(name) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(name))).lower()


def nfkc_value(text) -> str:
    """値の正規化（Markdown 出力用）。NFKC＋行内の連続空白を1つに畳む。改行と日本語間の空白は残す。

    ラベル比較用の normalize_label とは別物（こちらは空白を消さない）。
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text)).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [_SPACES_RE.sub(" ", line).strip() for line in s.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def normalize_unit(unit: str) -> str:
    u = unicodedata.normalize("NFKC", unit or "").strip()
    return UNIT_ALIASES.get(u.lower(), u)


def split_label_unit(label: str) -> tuple[str, str]:
    """「作業時間(h)」→ ("作業時間", "時間")。単位がなければ (label, "")。"""
    s = unicodedata.normalize("NFKC", str(label or "")).strip()
    m = _HEADER_UNIT_RE.match(s)
    if not m or not m[1].strip() or not _unit_like(m[2].strip()):
        return str(label or "").strip(), ""
    return m[1].strip(), normalize_unit(m[2])


def _unit_like(text: str) -> bool:
    if text.isascii() or text in _JA_UNITS or len(text) == 1:
        return True
    return normalize_unit(text) != text  # 別名の一覧にあるもの


def _number_text(text) -> str:
    """数値として読む前の正規化: NFKC・桁区切りのカンマを除く・「▲」「△」の負号を「-」にする。"""
    s = unicodedata.normalize("NFKC", str(text or "")).replace(",", "").strip()
    return _MINUS_MARK_RE.sub("-", s)


def is_plain_number(text) -> bool:
    """「626」「1,032」「▲5」のような数字だけの値か。"""
    return re.fullmatch(r"-?\d+(?:\.\d+)?", _number_text(text)) is not None


def value_unit(text) -> str:
    """「1.5時間」→ "時間"。数値＋単位の形でなければ ""。"""
    m = _VALUE_UNIT_RE.match(_number_text(text))
    return normalize_unit(m[1]) if m else ""


def written_unit(text) -> str:
    """値に書かれた単位。「1.5時間」に加えて「約90分」「595分（9.9h）」のような前後に言葉のある値からも、
    最初の数値（to_number が読む数値）の直後の単位を返す。単位らしくなければ ""。

    「3時間40分」「09:30-12:45」は to_number が換算するので、ここでは単位を返さない。
    """
    s = _number_text(text)
    whole = value_unit(s)
    if whole or _DURATION_RE.search(s) or _TIME_RANGE_RE.search(s):
        return whole
    first = _NUM_RE.search(s)
    m = _UNIT_AFTER_RE.match(s, first.end()) if first else None
    if not m or not _unit_like(m[1]):
        return ""
    return normalize_unit(m[1])


def numeric_unit(text, unit: str = "") -> str:
    """数値項目の単位。帳票の種類の設定があればそれ、無ければ書かれた値（「390分」「1,032分」「3時間40分」）から。

    どちらからも決まらなければ ""（単位不明）。単位を勝手に決めない（既定値は持たない）。
    """
    u = normalize_unit(unit)
    if u:
        return u
    s = _number_text(text)
    if _DURATION_RE.search(s) or _TIME_RANGE_RE.search(s):
        return "分"  # to_number が「3時間40分」「09:30-12:45」を分に換算する
    return written_unit(s)


def cell_text(value) -> str:
    """セル値を表示用の文字列にする。"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        if value.time() == time(0):
            return value.date().isoformat()
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.strftime("%H:%M" if value.second == 0 else "%H:%M:%S")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).replace("_x000D_", "").replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def pick_checked(text) -> tuple[str | None, str | None] | None:
    """「□重大　■大　□中」→ ("大", None)。チェックボックス表記でなければ None。

    1行に □/■ などの印が2つ以上あるときだけ解釈する。複数チェックは「、」でつなぐ。
    どれにもチェックが無ければ (None, 警告)。
    """
    if not isinstance(text, str) or "\n" in text.strip():
        return None
    options = _CHECKBOX_RE.findall(text)
    if len(options) < 2:
        return None
    picked = [option for mark, option in options if mark in CHECKED_MARKS]
    if not picked:
        return None, "チェックの入った選択肢がありません"
    return "、".join(picked), None


def split_inline(text: str) -> tuple[str, str] | None:
    """「設備番号：EQ-001」のように1セルにラベルと値が入っている場合に分割する。

    戻り値: (正規化済みラベル, 値テキスト)
    """
    m = _INLINE_SEP.search(text)
    if not m:
        return None
    label, value = text[: m.start()], text[m.end():].strip()
    if "\n" in label or not value:
        return None
    label_norm = normalize_label(label)
    if not label_norm or len(label_norm) > MAX_LABEL_LENGTH:
        return None
    return label_norm, value


def to_date(raw, text: str, date1904: bool = False) -> tuple[str | None, str | None]:
    """日付に変換する。戻り値: (ISO形式の日付 or 元テキスト, 警告)

    date1904: ブックが1904年基準（Mac版Excel由来）のときシリアル値の起点を変える。
    """
    if isinstance(raw, datetime):
        # 時刻が入っていれば残す（「発生日時」の 23:08 を黙って落とさない）
        return (raw.date().isoformat() if raw.time() == time(0) else raw.strftime("%Y-%m-%d %H:%M")), None
    if isinstance(raw, date):
        return raw.isoformat(), None
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        low = 20000 - 1462 if date1904 else 20000
        if low < raw < 80000:
            epoch = MAC_EPOCH if date1904 else WINDOWS_EPOCH
            return from_excel(raw, epoch=epoch).date().isoformat(), None

    s = unicodedata.normalize("NFKC", text or "")
    m = _ERA_RE.search(s)
    if m:
        year = 2018 + (1 if m[1] == "元" else int(m[1]))
        parsed = _safe_date(year, int(m[2]), int(m[3]))
        if parsed:
            return parsed, None
    m = _DATE_RE.search(s)
    if m:
        parsed = _safe_date(int(m[1]), int(m[2]), int(m[3]))
        if parsed:
            t = _CLOCK_RE.match(s, m.end())
            if t and int(t[1]) < 24 and int(t[2]) < 60:
                parsed += f" {int(t[1]):02d}:{t[2]}"  # 日付の直後の「14:05」は残す
            return parsed, None
    m = _NO_YEAR_RE.match(s)
    if m and 1 <= int(m[1]) <= 12 and 1 <= int(m[2]) <= 31:
        # 年は補わない（同じ帳票の別の項目やファイル名から推すと、外れたときに誤った日付を Markdown に書くことになる）
        return (text or None), "年が書かれていません。元のファイルを確かめて、2026-02-12 のように年から書いてください"
    return (text or None), "日付として解釈できません"


def _safe_date(y: int, m: int, d: int) -> str | None:
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def to_number(raw, text: str, unit: str = "") -> tuple[int | float | str | None, str | None]:
    """数値に変換する。「2.5時間」のような単位付きは数値部分を取り出して警告を付ける。

    値が「3時間40分」の形なら、unit（「分」「時間」）に換算する（220 / 3.67）。unit が空なら分にする。
    """
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return (int(raw) if float(raw).is_integer() else raw), None

    s = _number_text(text)
    m = _DURATION_RE.search(s)
    if m and normalize_unit(unit) in ("分", "時間", ""):
        # 単位の決まっていない項目では分にする（「3時間40分」を 3 と読まない）
        target = normalize_unit(unit) or "分"
        minutes = float(m[1]) * 60 + float(m[2])
        number = minutes if target == "分" else round(minutes / 60, 2)
        value = int(number) if float(number).is_integer() else number
        return value, f"「{text}」を{target}に換算しました"
    r = _TIME_RANGE_RE.search(s)
    if r:
        return _time_range_number(text, s, r, unit)
    m = _NUM_RE.search(s)
    if not m:
        return (text or None), "数値として読み取れません"
    number = float(m[0])
    value = int(number) if number.is_integer() else number
    warning = None if m[0] == s else f"「{text}」から数値の部分だけを読み取りました"
    return value, warning


def _time_range_number(text: str, s: str, r: re.Match, unit: str) -> tuple[int | float | str | None, str | None]:
    """「09:30-12:45（3.2h）」の形。添えた時間数があればそれを、無ければ範囲の長さを、項目の単位で返す。

    開始時刻（9）を数値として返さない。単位の決まっていない項目では分にする（「3時間40分」と同じ）。
    """
    target = normalize_unit(unit) or "分"
    if target not in ("分", "時間"):
        return (text or None), "時刻の範囲です。数値として読み取れません"
    rest = s[:r.start()] + " " + s[r.end():]
    m = _TIME_AMOUNT_RE.search(rest)
    if not m and "/" in r[0]:
        # 月日の付いた範囲は日をまたぐ日数が分からないので、終了−開始では出さない
        return (text or None), "時刻の範囲です。時間数が書かれていないので数値として読み取れません"
    if m:
        amount, written = float(m[1]), normalize_unit(m[2])
        minutes = amount * 60 if written == "時間" else amount
        note = f"「{text}」の時間数（{m[1]}{written}）を読み取りました"
    else:
        start = int(r[1]) * 60 + int(r[2])
        end = int(r[3]) * 60 + int(r[4])
        minutes = end - start if end >= start else end + 24 * 60 - start  # 日をまたぐ作業
        note = f"「{text}」は時刻の範囲なので、開始から終了までの時間を出しました"
    number = minutes if target == "分" else round(minutes / 60, 2)
    value = int(number) if float(number).is_integer() else number
    return value, f"{note}（{target}）。元のファイルと照らして確かめてください"
