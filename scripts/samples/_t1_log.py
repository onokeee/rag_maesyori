"""T1 トラブル対応一覧の「対応内容」列（複数人が1セルに追記していく対応ログ）を作る。

    from scripts.samples import _t1_log
    text, meta = _t1_log.response_log(inc, prev_inc, extract_date)

アプリの AI整形／ログ分割（logproc/、docs/research/対応内容AI整形設計.md 2章）の主な入力を想定した列。
- エントリは Incident の 発生→連絡→初動→調査→原因→処置→結果（完了日）の順に、実際の日時から作る。
  完了・経過観察は完了日時で締め、経過観察はその後の様子見を追記。保留・対応中は締めずに終わる。
- 件数は重要度で変える（小 2〜4 / 中 3〜6 / 大 4〜9 / 重大 6〜12）。多すぎる分は近いエントリを結合する。
- 書き方は記入者ごと（domain.people() の社員番号から決まる癖）に変える:
  日付（4/1 10:00 / 2024/04/01 10:00 / 4月1日 / R6.4.1 / 【4/2 夜勤】）、同日・翌日・翌週、時刻だけの行、
  記入者（田中： / (田中) / 田中→佐藤 / K.T / 保全G / メーカーFE）、略語（様子見・TEL済・交換済・FE来場・チョコ停）、
  全角数字（管理No・品番・コードの中には入れない）、まれな誤変換。
- セル単位のばらつき: メール転記（-----Original Message-----）、新しい順、見出し型（【現象】【原因】【処置】）、
  「／」詰め・「→」連結・「。」だけでつないだ段落、空欄、「同上」（瞬低などの同時多発の2件目以降）。
乱数は dm.rng("t1log:...") からのみ取るので、何度実行しても同じ文字列になる。
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from functools import lru_cache

from . import domain as dm

# イニシャル用のローマ字頭文字（姓, 名）。社員番号 → (姓の頭文字, 名の頭文字)
_INITIALS = {
    "M10231": ("T", "M"), "M10544": ("N", "K"), "M11302": ("K", "D"), "M11876": ("Y", "S"), "M12410": ("M", "T"),
    "M12655": ("I", "R"), "M13021": ("K", "Y"), "M13388": ("H", "K"), "M13702": ("S", "A"), "M10688": ("S", "K"),
    "M11544": ("Y", "J"), "M12127": ("M", "Y"), "M12893": ("I", "N"), "M13155": ("H", "R"), "M13520": ("A", "M"),
    "M13811": ("I", "R"), "M10902": ("M", "O"), "M12033": ("F", "T"), "M12978": ("O", "S"), "M13640": ("G", "Y"),
    "M11020": ("H", "S"), "M12744": ("M", "M"), "M13902": ("K", "D"), "M11133": ("E", "T"), "M12810": ("A", "Y"),
    "M13950": ("S", "H"), "M11247": ("F", "K"), "M12866": ("N", "A"), "M14011": ("F", "K"), "M11390": ("O", "H"),
    "M12921": ("M", "S"), "M14057": ("O", "H"), "M10870": ("H", "T"), "M12302": ("F", "T"), "M13266": ("N", "K"),
    "M11655": ("O", "M"), "M13477": ("T", "S"), "FE-0107": ("K", "M"), "FE-0213": ("U", "H"), "FE-0320": ("N", "T"),
    "FE-0431": ("M", "T"), "FE-0615": ("K", "T"),
}

# 重要度ごとのエントリ数の範囲
_N_RANGE = {"小": (2, 4), "中": (3, 6), "大": (4, 9), "重大": (6, 12)}

# まれな誤変換・打ち間違い（原文の語, 誤り）
_TYPOS = [("様子見", "様子身"), ("異常なし", "異常ナシ"), ("異常", "以上"), ("確認", "確人"), ("交換", "交感"),
          ("復旧", "複旧"), ("回収", "改修"), ("規定", "既定"), ("清掃", "政争")]

_WEEKDAY_EN = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_MONTH_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
             "November", "December"]
_STEP_NO_RE = re.compile(r"^\s*(?:[\(（]?[0-9０-９]{1,2}[\.．\)）]\s*|[\u2460-\u2473]\s*|・\s*)")


# ---------------------------------------------------------------------------
# 記入者の癖
# ---------------------------------------------------------------------------
@lru_cache(maxsize=None)
def writer_habit(employee_id: str) -> dict:
    """記入者ごとの対応ログの書き癖（日付の書き方・記入者の書き方・略語・時刻だけの行など）。"""
    r = dm.rng(f"t1log:writer:{employee_id}")
    base = dm._writer_habit(employee_id)
    fe = employee_id.startswith("FE-")
    date_style = r.choices(["md", "ymd", "jp", "bracket", "wareki", "rel"], weights=[38, 18, 12, 10, 4, 18])[0]
    author = "fe" if fe else r.choices(["colon", "paren", "space", "initials", "group", "none"],
                                       weights=[34, 22, 13, 10, 11, 10])[0]
    return dict(
        date_style=date_style,
        author=author,
        colon=r.choice(["：", "：", ":", ":", "："]),
        paren_full=r.random() < 0.55,
        initials_order=r.choice(["gf", "fg"]),       # M.T（名.姓）か T.M（姓.名）か
        initials_tail=r.random() < 0.4,
        time_only=r.choice([0.15, 0.35, 0.6, 0.8]),  # 同じ日の2件目以降を時刻だけで書く確率
        rel=r.choice([0.1, 0.25, 0.45, 0.7]) if date_style == "rel" else r.choice([0.0, 0.05, 0.15, 0.3]),
        nodate=r.choice([0.0, 0.0, 0.03, 0.08]),
        time_p=r.uniform(0.45, 0.95),
        abbr=r.choice([0.2, 0.5, 0.8, 0.95]),
        omit_same=r.uniform(0.2, 0.6),               # 直前と同じ記入者なら名前を省く確率
        zen=base["zen"],
        zen_date=base["zen"] >= 0.6 and r.random() < 0.6,
        hkana=base["hkana"],
        typo=base["typo"] * 2 + (0.01 if r.random() < 0.25 else 0.0),
        polite=base["polite"],
        jp_time=r.choice(["colon", "ji"]),           # 4月1日 10:00 か 4月1日 10時
    )


def _group_of(p: dm.Person) -> str:
    if p.department == "設備メーカー":
        return "メーカーFE"
    if "設備保全課" in p.section:
        return "保全G"
    if "製造課" in p.section:
        return "製造"
    if "生産技術" in p.section:
        return "生技"
    if "品質保証" in p.section:
        return "品証"
    return "保全G"


def _initials(p: dm.Person, habit: dict) -> str:
    fam, giv = _INITIALS.get(p.employee_id, ("X", "X"))
    return f"{giv}.{fam}" if habit["initials_order"] == "gf" else f"{fam}.{giv}"


def _maker_short(inc: dm.Incident, fe: dm.Person | None) -> str:
    if fe is not None:
        return dm.MAKER_SHORT.get(fe.section, fe.section)
    return dm.MAKER_SHORT.get(inc.equipment.maker, inc.equipment.maker)


# ---------------------------------------------------------------------------
# エントリ（出来事）の組み立て
# ---------------------------------------------------------------------------
@dataclass
class _Ev:
    at: datetime
    who: dm.Person
    body: str
    kind: str                  # report/first/inv/cause/tel/mail/visit/action/result/follow/hold/ongoing
    short: str = ""            # 「→」連結型で使う短い書き方
    notes: list = field(default_factory=list)   # ※行（日付なしの注記）
    handoff_from: dm.Person | None = None
    mail: dict | None = None


def _sentences(text: str) -> list[str]:
    return [s.strip() + "。" for s in (text or "").split("。") if s.strip()]


def _short_symptom(s: str) -> str:
    s = re.split(r"[。]", s or "")[0]
    return s if len(s) <= 60 else s[:58].rstrip("、，,（(") + "…"


def _steps(action: str) -> list[str]:
    out = []
    for line in (action or "").split("\n"):
        s = _STEP_NO_RE.sub("", line).strip()
        if not s or "継続対応中" in s:
            continue
        out.append(s)
    return out


def _floor5(dt: datetime) -> datetime:
    return dt.replace(minute=dt.minute - dt.minute % 5, second=0, microsecond=0)


def _spread(start: datetime, end: datetime, n: int, lo: float, hi: float, r) -> list[datetime]:
    """start〜end の区間の lo〜hi の範囲に n 個の時刻を昇順で置く。"""
    if n <= 0:
        return []
    span = (end - start).total_seconds()
    pts = sorted(r.uniform(lo, hi) for _ in range(n))
    return [start + timedelta(seconds=span * p) for p in pts]


def _events(inc: dm.Incident, r, extract_dt: datetime) -> list[_Ev]:
    staff = [p for p in inc.assignees if p.department != "設備メーカー"]
    fe = next((p for p in inc.assignees if p.department == "設備メーカー"), None)
    lead = staff[0] if staff else inc.reporter
    second = staff[1] if len(staff) > 1 else None
    rep = inc.reporter
    t_start = inc.response_started_at
    if inc.completed_at is not None:
        t_end = inc.completed_at
    else:
        t_end = inc.occurred_at + timedelta(minutes=max(inc.downtime_min, 30))
        t_end = min(max(t_end, t_start + timedelta(minutes=30)), extract_dt)
    t_end = max(t_end, t_start + timedelta(minutes=5))
    evs: list[_Ev] = []

    # 1. 発生・連絡
    sym = _short_symptom(inc.symptom)
    gap = (inc.reported_at - inc.occurred_at).total_seconds() / 60
    occ = f"発生は{inc.occurred_at.hour}:{inc.occurred_at.minute:02d}頃。" if gap >= 40 and r.random() < 0.6 else ""
    choko = r.choice(["チョコ停が数回続いた後に停止。", "直前からチョコ停多発、その後停止。"]) if inc.detected_by in ("装置アラーム", "オペレーター") and r.random() < 0.07 else ""
    rec = ""
    if inc.recurrence and inc.related_incident_id and r.random() < 0.5:
        rec = f"（{inc.related_incident_id}と同じ現象、再発）"
    if rep in staff or "設備保全課" in rep.section:
        body = f"{choko}{'巡回時に発見。' if inc.detected_by == '保全巡回' else ''}{sym}{rec}。{occ}"
        evs.append(_Ev(inc.reported_at, rep, body, "report", short="停止連絡"))
    elif r.random() < 0.5:
        body = f"{choko}{sym}{rec}。{occ}保全へ連絡"
        evs.append(_Ev(inc.reported_at, rep, body, "report", short="停止連絡"))
    else:
        grp = _group_of(rep)
        body = f"{grp} {dm.surname(rep)}さんより連絡あり。{choko}{sym}{rec}。{occ}"
        evs.append(_Ev(inc.reported_at, lead, body, "report", short="連絡受け"))

    # 2. 初動
    first = re.sub(r"\s*→\s*", "、", inc.first_response or "現場確認")
    evs.append(_Ev(t_start, lead, f"現場着。{first}", "first", short="現場確認"))

    inv = [s for s in _sentences(inc.investigation) if "引き続き調査中" not in s]
    steps = _steps(inc.action)
    parts = list(inc.parts_used)
    long_case = (t_end - t_start) >= timedelta(hours=20)

    # 3. 調査（2〜3文ずつまとめて置く）
    chunks: list[list[str]] = []
    for s in inv:
        if chunks and len(chunks[-1]) < r.choice([1, 2, 2, 3]):
            chunks[-1].append(s)
        else:
            chunks.append([s])
    for at, chunk in zip(_spread(t_start, t_end, len(chunks), 0.05, 0.35, r), chunks):
        who = second if (second is not None and r.random() < 0.35) else lead
        evs.append(_Ev(at, who, "".join(chunk), "inv", short="調査"))

    # 4. 原因
    if inc.cause:
        at = t_start + (t_end - t_start) * r.uniform(0.36, 0.45)
        c = inc.cause.rstrip("。")
        forms = ["原因：{c}", "原因は{c}", "{c}。"]
        if not re.search(r"(推定|判断|可能性|要確認|と思われる|あり)$", c):
            forms.append("{c}と判断")
        evs.append(_Ev(at, lead, r.choice(forms).format(c=c), "cause", short="原因判明"))

    # 5. メーカー（電話・メール・来場）
    if fe is not None:
        mk = _maker_short(inc, fe)
        at_tel = t_start + (t_end - t_start) * r.uniform(0.2, 0.33)
        at_visit = t_start + (t_end - t_start) * r.uniform(0.47, 0.52)
        # 調査文に「○時間後に来場」とあれば、来場時刻をそれに合わせる（初動から数えて、完了より前）
        m = re.search(r"(\d+)時間後に来場", inc.investigation or "")
        if m:
            at_visit = min(t_start + timedelta(hours=int(m.group(1))), t_end - (t_end - t_start) * 0.1)
            at_tel = min(at_tel, t_start + (at_visit - t_start) * 0.3)
        evs.append(_Ev(at_tel, lead, f"{mk}へTEL、状況説明しFE手配", "tel", short="メーカーTEL"))
        if (at_visit.date() > at_tel.date() or (at_visit - at_tel) >= timedelta(hours=3)) and r.random() < 0.45:
            at_mail = at_tel + (at_visit - at_tel) * r.uniform(0.2, 0.6)
            evs.append(_Ev(at_mail, lead, "メーカー回答（メール転記）", "mail", mail=dict(
                fe=fe, maker=fe.section, sent=at_mail, visit=at_visit, cause=inc.cause, eq=inc.equipment.equipment_id,
                sym=re.split(r"[、，]", _short_symptom(inc.symptom))[0][:30], to=lead)))
        if r.random() < 0.6:
            evs.append(_Ev(at_visit, lead, f"来場、{mk} {dm.surname(fe)}様と点検", "visit", short="FE来場"))
        else:
            evs.append(_Ev(at_visit, fe, "現地着、点検開始", "visit", short="FE来場"))

    # 6. 処置（部品は交換の手順に付ける）
    n_act = len(steps)
    if n_act:
        groups: list[list[str]] = []
        for s in steps:
            if groups and len(groups[-1]) < r.choice([1, 2, 2, 3]):
                groups[-1].append(s)
            else:
                groups.append([s])
        hi = 0.95 if inc.completed_at is not None else 0.9
        times = _spread(t_start, t_end, len(groups), 0.5, hi, r)
        for gi, (at, g) in enumerate(zip(times, groups)):
            who = fe if (fe is not None and r.random() < 0.3) else (second if second is not None and gi % 2 == 1 else lead)
            text = "、".join(g)
            if parts and "交換" in text:
                p, n = parts.pop(0)
                text += f"（{p.name} {p.part_no} ×{n}）"
            evs.append(_Ev(at, who, text, "action", short="処置"))
        if parts:
            p, n = parts[0]
            evs[-1].body += f"。使用部品 {p.name}（{p.part_no}）×{n}"

    # 7. 結果・締め
    status = inc.status
    if status in ("完了", "経過観察"):
        tail = r.choice(["完了", "クローズ", "対応終了", "完了。"]) if status == "完了" else r.choice(["様子見", "経過観察とする"])
        res = (inc.result or "復旧").rstrip("。")
        ev = _Ev(t_end, lead, f"{res}。{tail}", "result", short="復旧→完了" if status == "完了" else "復旧 様子見")
        if status == "完了" and inc.severity in ("大", "重大") and inc.prevention and r.random() < 0.55:
            ev.notes.append(f"※再発防止：{inc.prevention.replace(' → ', '、')}")
        if inc.severity == "重大" and inc.horizontal_deployment and r.random() < 0.6:
            ev.notes.append(f"※水平展開：{inc.horizontal_deployment}")
        evs.append(ev)
        if status == "経過観察":
            follow_bodies = r.sample(["再発なし、引き続き様子見", "トレンド確認、異常なし。監視継続", "同アラーム出ていない。もう少し様子見",
                                      "製造に聞き取り、問題なしとのこと。経過観察継続"], 2)
            for k in range(r.choice([1, 1, 2])):
                at = t_end + timedelta(days=r.randint(1 + k * 5, 4 + k * 6), hours=r.randint(0, 6))
                at = at.replace(hour=r.choice([9, 10, 13, 14, 16]), minute=r.choice([0, 10, 30, 45]))
                if at.date() >= extract_dt.date() or at <= t_end:
                    break
                body = follow_bodies[k]
                evs.append(_Ev(at, second or lead, body, "follow", short="様子見"))
    elif status == "保留":
        res = (inc.result or "部品入荷待ち").rstrip("。")
        at = min(t_end + timedelta(days=r.randint(0, 3), minutes=r.randint(10, 300)), extract_dt)
        at = max(at, t_end)
        extra = r.choice(["", "", "。課長承認待ち", "。見積依頼中", "。次回PMで対応予定"])
        evs.append(_Ev(at, lead, f"{res}{extra}", "hold", short="保留"))
    else:
        body = r.choice(["引き続き調査中", "原因調査継続。次シフトへ引継ぎ", "部品手配中、入荷待ち", "ログ回収済、メーカー解析待ち",
                         "暫定で運転再開、継続対応中"])
        evs.append(_Ev(t_end, lead, body, "ongoing", short="調査継続"))

    # 時刻をそろえる（5分単位・昇順・範囲内）
    evs.sort(key=lambda e: (e.at, ["report", "first", "inv", "tel", "cause", "mail", "visit", "action", "result",
                                   "hold", "ongoing", "follow"].index(e.kind)))
    last = None
    for e in evs:
        at = _floor5(e.at)
        if last is not None and at < last:
            at = last
        if e.kind not in ("report",) and at < _floor5(inc.reported_at):
            at = _floor5(inc.reported_at)
        e.at = at
        last = at
    evs[0].at = inc.reported_at.replace(second=0, microsecond=0)
    return evs


def _merge(evs: list[_Ev], target: int, r) -> list[_Ev]:
    """エントリ数が target を超える分を、時間の近い隣同士で結合して減らす（最初と最後、メールは残す）。"""
    evs = list(evs)
    while len(evs) > target:
        best, best_gap = None, None
        for i in range(1, len(evs) - 1):
            a, b = evs[i], evs[i + 1]
            if a.mail or b.mail or b is evs[-1] and len(evs) > 2 and b.kind in ("result", "hold", "ongoing", "follow"):
                continue
            gap = (b.at - a.at).total_seconds() + (0 if a.who == b.who else 3600)
            if best_gap is None or gap < best_gap:
                best, best_gap = i, gap
        if best is None:
            # 残りは中間の最初の2つを結合
            if len(evs) <= 2:
                break
            best = 1 if len(evs) > 3 else 0
        a, b = evs[best], evs[best + 1]
        a.body = a.body.rstrip("。") + "。" + b.body.lstrip("→")
        a.notes += b.notes
        evs.pop(best + 1)
    return evs


# ---------------------------------------------------------------------------
# 描画
# ---------------------------------------------------------------------------
def _zen(s: str) -> str:
    return dm._ZEN_DIGITS_RE.sub(lambda m: dm.to_zenkaku(m.group(0), digits_only=True), s)


def _time_text(at: datetime, habit: dict) -> str:
    if habit["date_style"] == "jp" and habit["jp_time"] == "ji":
        return f"{at.hour}時" if at.minute == 0 else (f"{at.hour}時半" if at.minute == 30 else f"{at.hour}時{at.minute}分")
    return f"{at.hour}:{at.minute:02d}"


def _full_date(at: datetime, style: str, habit: dict, r, meta: Counter) -> str:
    with_time = r.random() < habit["time_p"]
    if style == "bracket":
        shift = "夜勤" if (at.hour >= 20 or at.hour < 8) else "日勤"
        d = at - timedelta(days=1) if at.hour < 8 else at
        meta["date:【M/D 勤務】"] += 1
        return f"【{d.month}/{d.day} {shift}】"
    if style == "ymd":
        meta["date:YYYY/MM/DD"] += 1
        text = f"{at.year}/{at.month:02d}/{at.day:02d}" + (f" {at.hour:02d}:{at.minute:02d}" if with_time else "")
    elif style == "jp":
        meta["date:M月D日"] += 1
        text = f"{at.month}月{at.day}日" + (f" {_time_text(at, habit)}" if with_time else "")
    elif style == "wareki":
        meta["date:R6.4.1"] += 1
        text = f"R{at.year - 2018}.{at.month}.{at.day}" + (f" {at.hour}:{at.minute:02d}" if with_time and r.random() < 0.3 else "")
    else:
        meta["date:M/D"] += 1
        text = f"{at.month}/{at.day}" + (f" {_time_text(at, habit)}" if with_time else "")
    if habit["zen_date"] and r.random() < 0.5:
        text = _zen(text).replace("/", "／").replace(":", "：")
        meta["date:全角"] += 1
    return text


def _when_text(ev: _Ev, prev: _Ev | None, habit: dict, r, meta: Counter, force: bool, style: str | None = None) -> str:
    style = style or habit["date_style"]
    if style == "rel":
        style = "md"
    if prev is not None and not force:
        days = (ev.at.date() - prev.at.date()).days
        if r.random() < habit["nodate"]:
            meta["date:なし"] += 1
            return ""
        if days == 0 and style != "bracket" and r.random() < habit["time_only"]:
            meta["date:時刻のみ"] += 1
            return _time_text(ev.at, habit)
        if r.random() < habit["rel"]:
            wk = ev.at.isocalendar()[1] - prev.at.isocalendar()[1]
            word = None
            if days == 0:
                word = f"同日{ev.at.hour}時" if ev.at.minute == 0 else f"同日 {ev.at.hour}:{ev.at.minute:02d}"
            elif days == 1:
                word = "翌朝" if ev.at.hour < 10 and r.random() < 0.3 else "翌日"
            elif days == 2:
                word = "翌々日"
            elif 2 < days <= 13 and wk == 1 and ev.at.year == prev.at.year:
                word = "翌週"
            elif 3 <= days <= 6:
                word = f"{days}日後"
            if word:
                meta["date:相対(" + re.sub(r"[\d:\s]+|時$", "", word).replace("同日時", "同日") + ")"] += 1
                return word
    return _full_date(ev.at, style, habit, r, meta)


def _author_text(p: dm.Person, habit: dict, r, meta: Counter, inc: dm.Incident) -> tuple[str, str]:
    """(先頭に付ける記入者, 末尾に付ける記入者)。"""
    sur = dm.surname(p)
    c = habit["colon"]
    kind = habit["author"]
    if kind == "fe":
        style = r.choice(["メーカーFE", "メーカーFE", "maker", "fe_sur"])
        meta["author:メーカーFE"] += 1
        if style == "maker":
            return f"{dm.MAKER_SHORT.get(p.section, p.section)}FE {sur}{c}", ""
        if style == "fe_sur":
            return "", f"（FE {sur}）"
        return f"メーカーFE{c}", ""
    if kind == "colon":
        meta["author:田中："] += 1
        return f"{sur}{c}", ""
    if kind == "space":
        meta["author:田中(空白)"] += 1
        return f"{sur} ", ""
    if kind == "paren":
        meta["author:(田中)"] += 1
        return "", (f"（{sur}）" if habit["paren_full"] else f"({sur})")
    if kind == "initials":
        meta["author:K.T"] += 1
        ini = _initials(p, habit)
        return ("", f"({ini})") if habit["initials_tail"] else (f"{ini}{c} ", "")
    if kind == "group":
        meta["author:保全G等"] += 1
        g = _group_of(p)
        return (f"{g} {sur}{c}" if r.random() < 0.6 else f"{g}{c}"), ""
    meta["author:なし"] += 1
    return "", ""


def _typo(s: str, habit: dict, r, meta: Counter) -> str:
    if habit["typo"] and r.random() < habit["typo"]:
        for a, b in r.sample(_TYPOS, len(_TYPOS)):
            if a in s:
                meta["typo"] += 1
                return s.replace(a, b, 1)
    return s


def _body_text(ev: _Ev, habit: dict, r, meta: Counter) -> str:
    s = ev.body
    if r.random() < habit["abbr"]:
        rep = [("経過観察とする", "様子見"), ("経過観察", "様子見"), ("トレンド監視とする", "様子見"), ("へTEL、", "TEL済、"),
               ("電話で状況共有", "TEL済"), ("交換）", "交換済）"), ("来場、", "FE来場、")]
        for a, b in rep:
            if a in s:
                s = s.replace(a, b)
                meta["abbr"] += 1
        s = re.sub(r"交換$", "交換済", s)
    if habit["polite"] and r.random() < 0.5:
        s = dm._polite(s if s.endswith("。") else s + "。")
    if habit["zen"] and r.random() < habit["zen"]:
        z = _zen(s)
        if z != s:
            meta["zen"] += 1
        s = z
    if habit["hkana"] and r.random() < 0.4:
        s = dm.to_hankaku_kana(s)
    return _typo(s, habit, r, meta)


def _mail_block(m: dict, r) -> list[str]:
    sent: datetime = m["sent"] - timedelta(minutes=r.randint(3, 25))   # 転記はメール受信の少し後
    ampm = "AM" if sent.hour < 12 else "PM"
    h12 = sent.hour % 12 or 12
    fe: dm.Person = m["fe"]
    visit: datetime = m["visit"]
    cause = re.sub(r"(と推定|と思われる|の可能性が高い|の可能性あり|（要確認）|。.*$|、ただし.*$)", "", m["cause"] or "").strip()
    guess = f"{cause}の可能性が高いと思われます。" if cause else "いただいたログを解析中です。"
    return [
        "-----Original Message-----",
        f"From: {fe.section} サービス部 {fe.name}",
        f"Sent: {_WEEKDAY_EN[sent.weekday()]}, {_MONTH_EN[sent.month - 1]} {sent.day}, {sent.year} {h12}:{sent.minute:02d} {ampm}",
        f"To: 設備保全課 {m['to'].name}",
        f"Subject: RE: {m['eq']} {m['sym']}の件",
        "",
        f"{m['to'].name.split(' ')[0]}様",
        "",
        "お世話になっております。",
        f"ご連絡いただいた件、{guess}",
        r.choice(["念のため装置ログの追加取得をお願いいたします。", "交換部品は当方で持参いたします。", ""]),
        f"{visit.month}/{visit.day}に弊社FEが伺います。",
        "以上、よろしくお願いいたします。",
        "-----",
    ]


def _render_log(inc: dm.Incident, evs: list[_Ev], r, meta: Counter) -> str:
    joiner = "newline"
    if inc.severity in ("小", "中") and len(evs) <= 5 and not any(e.mail for e in evs):
        joiner = r.choices(["newline", "slash", "arrow", "maru"], weights=[82, 6, 4 if inc.severity == "小" else 0, 8])[0]
    meta[f"joiner:{joiner}"] += 1

    if joiner == "arrow":
        # 「4/1 停止→リセット復帰→4/2再発→…」型（短い言葉だけで連結）
        out, prev_day = [], None
        for e in evs:
            head = f"{e.at.month}/{e.at.day} " if e.at.date() != prev_day else ""
            out.append(head + e.short)
            prev_day = e.at.date()
        meta["date:M/D"] += len({e.at.date() for e in evs})
        return "→".join(out) + f"（{dm.surname(evs[0].who if evs[0].kind != 'report' else evs[-1].who)}）"

    # 新しい順に書くセル（上へ追記）。逆順では「同日」や時刻だけの行が成り立たないので全エントリに日付を書く
    newest = joiner == "newline" and len(evs) >= 3 and not any(e.mail for e in evs) and r.random() < 0.05
    lines: list[str] = []
    prev: _Ev | None = None
    force_next = False
    cell_style = None
    if joiner == "slash":
        cell_style = r.choice(["md", "ymd"])
    for e in evs:
        habit = writer_habit(e.who.employee_id)
        style = cell_style
        if force_next and habit["date_style"] in ("bracket", "rel"):
            style = "md"
        when = _when_text(e, prev, habit, r, meta, force=force_next or newest or joiner == "slash" or prev is None, style=style)
        force_next = False
        head, tail = _author_text(e.who, habit, r, meta, inc)
        if joiner == "slash" and not tail:
            head, tail = "", f"({dm.surname(e.who)})"
        same = prev is not None and prev.who == e.who
        if same and (r.random() < habit["omit_same"] or (when and ":" in when and len(when) <= 5)):
            head, tail = "", ""
            meta["author:省略(直前と同じ)"] += 1
        if e.handoff_from is None and prev is not None and prev.who != e.who and e.kind in ("inv", "action") \
                and prev.who.department != "設備メーカー" and e.who.department != "設備メーカー" and r.random() < 0.12:
            head = f"{dm.surname(prev.who)}→{dm.surname(e.who)} 引継ぎ。"
            tail = ""
            meta["author:田中→佐藤"] += 1
        body = _body_text(e, habit, r, meta)
        if "チョコ停" in body:
            meta["abbr:チョコ停"] += 1
        if joiner == "maru":
            body = body.rstrip("。") + "。"
        sep = "" if (when.endswith("】") or not when) else " "
        line = f"{when}{sep}{head}{body}{tail}".strip()
        if e.mail:
            meta["email"] += 1
            line = "\n".join([line] + _mail_block(e.mail, r))
            force_next = True
        if e.notes and joiner == "newline":
            line = "\n".join([line] + e.notes)
        lines.append(line)
        prev = e

    if joiner == "slash":
        return "／".join(lines)
    if joiner == "maru":
        return "".join(x.replace("\n", "") for x in lines)
    if newest:
        meta["order:新しい順"] += 1
        lines.reverse()
    return "\n".join(lines)


def _header_cell(inc: dm.Incident, r) -> str:
    labels = r.choice([("現象", "原因", "処置", "結果"), ("現象", "原因", "対応", "再発防止"), ("現象", "原因", "処置", "状況")])
    steps = _steps(inc.action)
    act = "、".join(steps) if r.random() < 0.6 else "\n".join(f"・{s}" for s in steps)
    fourth = {"結果": inc.result or "対応中", "再発防止": inc.prevention or "検討中", "状況": inc.status}[labels[3]]
    return "\n".join([
        f"【{labels[0]}】{_short_symptom(inc.symptom)}",
        f"【{labels[1]}】{inc.cause or '調査中'}",
        f"【{labels[2]}】{act or '調査中'}",
        f"【{labels[3]}】{fourth}",
    ])


def response_log(inc: dm.Incident, prev: dm.Incident | None, extract_date: date) -> tuple[str | None, Counter]:
    """1件分の「対応内容」セルの文字列と、README 用の集計カウンタを返す。空欄は None。"""
    r = dm.rng(f"t1log:{inc.incident_id}")
    meta: Counter = Counter()
    # 瞬低などの同時多発で、直前の行と同じ時刻・同じ外部要因なら「同上」
    if (prev is not None and inc.category == "外部要因" and prev.category == "外部要因"
            and prev.occurred_at == inc.occurred_at and r.random() < 0.6):
        meta["cell:同上"] += 1
        return (r.choice(["同上", "同上", f"同上（{prev.incident_id}と同じ処置）"]), meta)
    if inc.severity == "小" and inc.status == "完了" and r.random() < 0.02:
        meta["cell:空欄"] += 1
        return (r.choice([None, None, None, "-"]), meta)
    if r.random() < 0.03:
        meta["cell:見出し型"] += 1
        return (_header_cell(inc, r), meta)

    extract_dt = datetime.combine(extract_date, datetime.min.time()).replace(hour=17)
    evs = _events(inc, r, extract_dt)
    lo, hi = _N_RANGE[inc.severity]
    target = r.randint(lo, hi)
    evs = _merge(evs, max(target, 2), r)
    meta["cell:時系列"] += 1
    if inc.status in ("対応中", "保留"):
        meta["open"] += 1
    meta[f"n:{len(evs)}"] += 1
    return (_render_log(inc, evs, r, meta), meta)


def response_logs(incidents: list[dm.Incident], extract_date: date) -> tuple[dict, Counter]:
    """並び順（発生日時順）の incidents 全件の対応内容と、全体の集計を返す。"""
    logs, total = {}, Counter()
    prev = None
    for inc in incidents:
        text, meta = response_log(inc, prev, extract_date)
        logs[inc.incident_id] = text
        total.update(meta)
        prev = inc
    return logs, total
