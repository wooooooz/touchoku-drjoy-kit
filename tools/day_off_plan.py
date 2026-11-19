# -*- coding: utf-8 -*-
r"""day_off_plan.py — 当直予定から、外す必要のある通常勤務（全日休み・遅出・早帰り）を決める。20261028

管理者 2026-10-26: 勤務間インターバルは通常勤務も含めた連続勤務で見る。当直予定から、外す必要のある通常勤務を自動で決め、
  のちに Dr.JOY の勤務予定へ入れたい。まず決め方（案）を作り、救急部の辛島先生に確かめる予定だったが、
  確かめたいことは全部 管理者が答えた（同日）。夜の当直を続けて組むのは禁止・外した時間は勤務時間上の休み（年休ではない）。
  同日: 初期研（C-1 水準）とそれ以外（B 水準）では決まりが違う → 「明け」「前日休」の決め打ちをやめ、勤務を時間の流れとして
  つないで、水準ごとの上限から計算する形にした。同日 コードのレビューで、夜の当直が続くひと続き・説明・質問を直した。
  ここは決め方と、確かめてもらうための一覧（xlsx）まで。Dr.JOY には書かない（このリポジトリは Dr.JOY に触らない）。

計算:
  勤務 = 通常勤務（通常勤務の日 8:30〜17:15。始まりを遅らせる・終わりを早める・全日休み に変えられる）
       ＋ 夜の当直（17:15〜翌朝。翌日が通常勤務の日なら 8:30、休日なら 9:00 まで）
       ＋ 昼の当直（休日 9:00〜17:15）。数えるのは当直（泊まる「当」）だけで、オンコールは数えない。
       開院記念日など daytime_oncall_only の日の昼は全員オンコールになるので当直に数えない（その日は通常勤務の日）。
  隙間なくつながる勤務をひと続きとし、夜の当直を含むひと続きに水準ごとの上限を当てる（宿日直許可の無い当直として）。
       B（初期研以外）: 続けて 28時間まで、そのあと 18時間休む（＝始業から46時間以内に18時間の連続した休息）
       C-1（初期研）  : 続けて 24時間まで、そのあと 24時間休む
       日直（休日の昼の当直）だけのひと続きは通常の決まり（始業から24時間以内に9時間休む＝続けて15時間まで）で、この案の時刻なら
       いつも満たす（通常勤務 8.75時間・日直 8.25時間）ので当てない。
  夜勤明けの日（平日）は、当直が平日でも休日でも、B は 12:30 まで・C-1 は全日休み（20261027 管理者: シンプルにしたい）。
  夜の当直から翌日（休日）の昼の当直に続くときは、当直に入る日（平日）を B・C-1 とも全日休み（同日）。
  ほかに上限を超えたら、上限を満たすぎりぎりまで通常勤務を外す（20261026 管理者: 決まった半日でなく、時刻をアプリで合わせる）。
       1) 夜の当直のあとの通常勤務を早く終える（早帰り） 2) 当直のあとに通常勤務が無ければ（翌日が休日）、当直の日を遅く始める（遅出）
       3) そのあとの休みが足りなければ、まず当直のあとを早く終え、次に次の勤務の日を遅く始める
       時刻は15分単位で上限を超えない側に切る。残る勤務が1時間未満ならその日は全日休み。
  夜の当直が2つ以上つながるひと続き（夜→通常勤務→夜 など）は、通常勤務を外しても満たせないので外さずに「要相談」。
       夜の当直を続けて組むのは禁止（管理者 20261026）なので、アプリが提出を止める（renderer consecutiveChecks）。ここは念のための受け皿。
  水準は職名で決める（初期研＝C-1、それ以外＝B）。職名が分からない人は B として計算し、理由に書く。

usage: python -X utf8 tools\day_off_plan.py --month 2026-11 [--owner 救急科] [--out xlsx] [--json PATH]
  --json = Dr.JOY の勤務予定に入れる側（手当計算の月次投入）が読む形（20261028）。schema hospital.day-off-plan/1:
    {month, made, offs: [{owner, login, name, title, level, date, kind, start, end, label, why}], notes: [要相談]}
    kind = 全日休み（start/end は null）／遅出／早帰り／遅出・早帰り。start/end は その日の通常勤務の時刻 "H:MM"。
    職員コードは入れない（全職員が読む置き場に置かれても困らないように。入れる側がログインIDで引く）。
"""
import argparse
import collections
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C            # noqa: E402

H = dt.timedelta(hours=1)
STEP = dt.timedelta(minutes=15)                 # 時刻は15分単位（安全側に切る）
MIN_WORK = H                                     # 残りが1時間未満なら全日休み
REG = (dt.time(8, 30), dt.time(17, 15))          # 通常勤務
NIGHT_START = dt.time(17, 15)
NIGHT_END_WORKDAY, NIGHT_END_HOLIDAY = dt.time(8, 30), dt.time(9, 0)
DAY_DUTY = (dt.time(9, 0), dt.time(17, 15))     # 休日の昼の当直
LIMITS = {"B": (28 * H, 18 * H), "C-1": (24 * H, 24 * H)}    # (続けて働ける時間, そのあとの休み)
AKE_END = {"B": dt.time(12, 30), "C-1": None}   # 夜勤明けの日（平日）の終わり。None = 全日休み（20261027 管理者）
OFF, LATE, EARLY, SHORT = "全日休み", "遅出", "早帰り", "遅出・早帰り"
WD = "月火水木金土日"
_DOO = []


def workday(d, holidays):
    """通常勤務の日（土日祝・年末年始でない日。開院記念日は通常勤務の日）。"""
    if not _DOO:
        _DOO.append(C.daytime_oncall_only())
    return not C.is_holiday(d, holidays) or d in _DOO[0]


def level_of(title):
    return "C-1" if C.rank_of(title or "") == "初" else "B"


def _at(d, t):
    return dt.datetime.combine(d, t)


def _floor(x):
    """15分単位で切り下げ（早帰りの時刻）。"""
    m = (x.hour * 60 + x.minute) // 15 * 15
    return x.replace(hour=m // 60, minute=m % 60, second=0, microsecond=0)


def _ceil(x):
    """15分単位で切り上げ（遅出の時刻）。"""
    f = _floor(x)
    return f if f == x.replace(second=0, microsecond=0) and x.second == 0 and x.microsecond == 0 else f + STEP


def hm(t):
    return f"{t.hour}:{t.minute:02d}"


def kind_of(st):
    """通常勤務の (始, 終) か None → 区分（全日休み／遅出／早帰り／遅出・早帰り）か ""（通常どおり）。"""
    if st is None:
        return OFF
    late, early = st[0] > REG[0], st[1] < REG[1]
    return SHORT if late and early else LATE if late else EARLY if early else ""


def label_of(st):
    """一覧に出す文: 全日休み／9:00 から勤務／12:30 まで勤務／9:00〜12:30 勤務。"""
    k = kind_of(st)
    if k == OFF:
        return OFF
    if k == LATE:
        return f"{hm(st[0])} から勤務"
    if k == EARLY:
        return f"{hm(st[1])} まで勤務"
    return f"{hm(st[0])}〜{hm(st[1])} 勤務"


def _blocks(duty_days, reg, holidays):
    """duty_days = {date: {"昼","夜"}}、reg = {date: (始, 終) か None} → 時刻順の勤務 [(始, 終, 種類, 日付)]。
    種類 = "reg"（通常勤務）／"day"（休日の昼の当直）／"night"（夜の当直）。"""
    out = []
    for d, st in reg.items():
        if st:
            out.append((_at(d, st[0]), _at(d, st[1]), "reg", d))
    for d, bands in duty_days.items():
        if "昼" in bands:
            out.append((_at(d, DAY_DUTY[0]), _at(d, DAY_DUTY[1]), "day", d))
        if "夜" in bands:
            nx = d + dt.timedelta(days=1)
            end = NIGHT_END_WORKDAY if workday(nx, holidays) else NIGHT_END_HOLIDAY
            out.append((_at(d, NIGHT_START), _at(nx, end), "night", d))
    return sorted(out)


def _segments(blocks):
    segs = []
    for b in blocks:
        if segs and b[0] <= max(x[1] for x in segs[-1]):
            segs[-1].append(b)
        else:
            segs.append([b])
    return segs


def plan_person(duty_days, level, holidays, span):
    """1人分。duty_days = {date: {"昼","夜"}}、span = (最初の日, 最後の日)（通常勤務を置く範囲）。
    戻り: (reg {date: (始, 終) か None}, 理由 {date: 文}, 要相談 [(date, 文)])。
    20261026 管理者「アプリで制度にあわせられるでしょう」: 決まった半日で外すのをやめ、上限を満たすぎりぎりの時刻
    （15分単位・安全側）まで、始業を遅らせるか終業を早める。残りが1時間未満なら全日休み。"""
    maxc, rest = LIMITS[level]
    d0, d1 = span
    reg = {}
    d = d0
    while d <= d1:
        if workday(d, holidays):
            reg[d] = REG
        d += dt.timedelta(days=1)
    why, notes, stuck = collections.defaultdict(list), [], set()
    lim_c = f"{level}: 続けて{int(maxc / H)}時間まで"
    lim_r = f"{level}: 夜の当直のあと{int(rest / H)}時間休む"

    def note_why(day, text):
        if text not in why[day]:
            why[day].append(text)

    def set_end(day, latest, reason):
        """day の通常勤務を latest（datetime）までに終える。短すぎれば全日休み。変えたら True。"""
        st = reg.get(day)
        if not st:
            return False
        end = min(_floor(latest), _at(day, st[1]))
        new = None if end - _at(day, st[0]) < MIN_WORK else (st[0], end.time())
        if new == st:
            return False
        reg[day] = new
        note_why(day, reason)
        return True

    def set_start(day, earliest, reason):
        """day の通常勤務を earliest（datetime）から始める。短すぎれば全日休み。変えたら True。"""
        st = reg.get(day)
        if not st:
            return False
        start = max(_ceil(earliest), _at(day, st[0]))
        new = None if _at(day, st[1]) - start < MIN_WORK else (start.time(), st[1])
        if new == st:
            return False
        reg[day] = new
        note_why(day, reason)
        return True

    # 20261027 管理者「シンプルにしたい」: 夜勤明けの日（平日）は、当直が平日でも休日でも同じに、B は 12:30 まで・C-1 は全日休み。
    #   平日の夜の当直なら 28時間ちょうど。休日の夜の当直（昼と夜も）は上限より 0.5〜2時間早いが、決まりを一つにそろえる。
    #   どの当直の曜日でも上限を満たす（休日の夜 17:15→12:30 = 19.25時間・休日の昼夜 9:00→12:30 = 27.5時間、翌朝まで20時間休める）。
    #   夜の当直が続く日（禁止）は外さずに要相談に回す（下のループ）。
    one = dt.timedelta(days=1)
    for nd, bands in sorted(duty_days.items()):
        ake = nd + one
        if "夜" in duty_days.get(nd - one, ()) or "夜" in duty_days.get(ake, ()):
            continue
        if "夜" in bands and reg.get(ake):
            end = AKE_END[level]
            reg[ake] = None if end is None else (reg[ake][0], min(reg[ake][1], end))
            note_why(ake, f"夜勤明け（{level}: {'全日休み' if end is None else hm(end) + ' まで'}）")
        # 20261027 管理者: 夜の当直から翌日（休日）の昼の当直に続くときは、当直に入る日（平日）を全日休みに（B の 13:15 から をやめた）。
        #   当直だけで 17:15→翌 17:15 = 24時間なので B・C-1 とも上限内。例 丁野先生 11/12 夜→11/13 昼（12日は休暇の予定とも合う）。
        if "夜" in bands and "昼" in duty_days.get(ake, ()) and reg.get(nd):
            reg[nd] = None
            note_why(nd, "夜の当直から翌日の昼の当直に続く（当直に入る日は全日休み）")

    for _ in range(400):
        segs = _segments(_blocks(duty_days, reg, holidays))
        fixed = False
        for i, seg in enumerate(segs):
            nights = [b for b in seg if b[2] == "night"]
            if not nights:
                continue
            start, end = seg[0][0], max(b[1] for b in seg)
            if len(nights) >= 2:
                # 夜→（通常勤務）→夜 とつながる: 通常勤務を外しても満たせない。外さずに要相談（コードレビュー 20261026）
                key = ("multi", start)
                if key not in stuck:
                    stuck.add(key)
                    ds = "・".join(f"{b[3].month}/{b[3].day}" for b in nights)
                    notes.append((nights[0][3], f"夜の当直が続く（{ds}）。通常勤務を含めて続けて{(end - start) / H:.1f}時間"
                                                f"（{lim_c}）。組み方の見直しか代償休息が要る"))
                continue
            first_duty = next(b for b in seg if b[2] != "reg")
            last_duty = max((b for b in seg if b[2] != "reg"), key=lambda b: b[1])
            after = [b for b in seg if b[2] == "reg" and b[0] >= last_duty[1] - dt.timedelta(minutes=1)]
            before = [b for b in seg if b[2] == "reg" and b[1] <= first_duty[0] + dt.timedelta(minutes=1)]
            if end - start > maxc:
                # まず当直のあとを早く終える。当直のあとに通常勤務が無ければ（翌日が休日）、当直の日を遅く始める
                if after and set_end(after[-1][3], start + maxc, f"夜の当直のあと（{lim_c}）"):
                    fixed = True
                elif before and set_start(before[0][3], end - maxc, f"この日の夜から当直（{lim_c}）"):
                    fixed = True
                else:
                    key = ("max", start)
                    if key not in stuck:
                        stuck.add(key)
                        notes.append((first_duty[3], f"当直だけで続けて{(end - start) / H:.1f}時間（{lim_c}）。通常勤務を外しても満たせない"))
                    continue
                break
            nxt = segs[i + 1] if i + 1 < len(segs) else None
            if nxt and nxt[0][0] - end < rest:
                if after and set_end(after[-1][3], nxt[0][0] - rest, f"夜の当直のあとの休み（{lim_r}）"):
                    fixed = True
                    break
                b = nxt[0]
                if b[2] == "reg" and set_start(b[3], end + rest, f"夜の当直のあとの休み（{lim_r}）"):
                    fixed = True
                    break
                key = ("rest", end)
                if key not in stuck:
                    stuck.add(key)
                    notes.append((b[3], f"夜の当直のあとの休みが{(nxt[0][0] - end) / H:.1f}時間しかない（{lim_r}）。"
                                        "次も当直なので、通常勤務を外しても満たせない"))
        if not fixed:
            break
    return reg, {k: "／".join(v) for k, v in why.items()}, notes


def duty_map(duties, holidays):
    """duties → ({key: {date: {"昼","夜"}}}, {key: 氏名})。オンコール・開院記念日の昼は数えない。"""
    workday(dt.date.today(), holidays)
    doo = _DOO[0]
    by = collections.defaultdict(lambda: collections.defaultdict(set))
    names = {}
    for r in duties:
        if not r.get("date") or r.get("oncall") or r.get("band") not in ("昼", "夜"):
            continue
        d = dt.date.fromisoformat(r["date"])
        if r["band"] == "昼" and d in doo:
            continue                            # 開院記念日の昼は全員オンコール
        nm = r.get("name") or r.get("name_disp") or ""
        if not (r.get("staff_login") or nm):
            continue                            # 誰か分からない行（別々の人を1人にまとめてしまうので数えない）
        key = r.get("staff_login") or ("name:" + nm)
        by[key][d].add(r["band"])
        names[key] = nm or key
    return by, names


def plan(duties, holidays, owner="", titles=None, span=None):
    """1科分の duties → (外す通常勤務の一覧, 要相談の一覧)。titles = {login: 職名}（水準を決める）。
    外す = [{owner, login, name, title, level, date, kind(全日休み/遅出/早帰り/遅出・早帰り), label(9:00 から勤務 など), why}]。"""
    titles = titles or {}
    by, names = duty_map(duties, holidays)
    offs, notes = [], []
    for key, days in by.items():
        lo, hi = min(days) - dt.timedelta(days=2), max(days) + dt.timedelta(days=3)
        if span:
            lo, hi = min(lo, span[0]), max(hi, span[1])
        title = titles.get(key, "")
        lv = level_of(title)
        unknown = "（職名が分からないので B として計算）" if not title else ""
        reg, why, ns = plan_person(days, lv, holidays, (lo, hi))
        for d, st in sorted(reg.items()):
            if kind_of(st):
                offs.append({"owner": owner, "login": key, "name": names[key], "title": title, "level": lv,
                             "date": d, "kind": kind_of(st), "label": label_of(st), "time": st, "why": why.get(d, "") + unknown})
        for d, w in ns:
            notes.append({"owner": owner, "login": key, "name": names[key], "title": title, "level": lv,
                          "date": d, "why": w + unknown})
    return offs, notes


def month_data(y, m, owner, log=print):
    """その月と前後の月の当直（月の境目のため）と、名簿の職名 {login: 職名}。読めた月を log に出す。"""
    import master_store as MST
    duties, titles, read = [], {}, []
    for yy, mm in (((y, m - 1) if m > 1 else (y - 1, 12)), (y, m), ((y, m + 1) if m < 12 else (y + 1, 1))):
        import shozoku as SZ
        if (yy, mm) < tuple(SZ.REAL_FROM):      # 20261028: 10月までは試験用フォルダの作り物の割当（本物は DTX14）→ 読まない
            continue
        mp = MST.master_path(os.path.join(C.PROD_ROOT, f"R{yy - 2018}年{mm}月"), yy, mm)
        if not os.path.isfile(mp):              # 9・10月は「R8年10月(動作確認用テストフォルダ)」の名前（管理者だけ）
            import glob
            alt = [p for p in glob.glob(os.path.join(C.PROD_ROOT, f"R{yy - 2018}年{mm}月(*)"))
                   if os.path.isfile(MST.master_path(p, yy, mm))]
            if not alt:
                continue
            mp = MST.master_path(alt[0], yy, mm)
        ent = (MST.load(mp, yy, mm).get("depts") or {}).get(owner) or {}
        duties += ent.get("duties") or []
        read.append(f"{yy}-{mm:02d}")
        for p in ent.get("roster") or []:
            if p.get("login") and p.get("title"):
                titles[p["login"]] = p["title"]
    if any(not titles.get(r.get("staff_login")) for r in duties if r.get("staff_login")):
        try:                                     # 名簿に職名が無い人は Dr.JOY の名簿から
            people, _e, _c = C.load_staffs(C.newest_staffs())
            for p in people:
                if p.get("title"):
                    titles.setdefault(p["login"], p["title"])
        except Exception:
            pass
    log(f"  {owner}: 読んだ月 {', '.join(read) or 'なし'}（前後の月が無ければ、月の境目の当直は見えない）")
    return duties, titles


def rules_rows(y, m):
    """「決め方」シートの行 [(文, 太字, 大きさ)]。20261026 管理者: 初期研（C-1）とそれ以外（B）を分けて書く。
    同日: 決まった半日で外すのをやめ、上限を満たすぎりぎりの時刻（15分単位）で「何時から／何時まで勤務」と出す。
    同日: 確かめたいことは全部 管理者が答えた（辛島先生に聞かなくてよい）→「決めたこと」に置き換えた。"""
    t = 11
    return [
        (f"当直予定から、通常勤務をどこまで外すかを自動で決める決め方　R{y - 2018}年{m}月分で計算", True, 14),
        ("", False, t),
        ("対象は救急科と小児科。数えるのは当直（泊まる「当」）だけで、オンコールは数えない。", False, t),
        ("当直は宿日直許可の無い当直（勤務時間）として数える。", False, t),
        ("", False, t),
        ("■ 勤務の時刻と外し方　※初期研もそれ以外も同じ", True, 12),
        ("通常勤務: 平日と開院記念日の 8:30〜17:15。", False, t),
        ("夜の当直: 17:15〜翌朝。翌日が平日なら 8:30、休日なら 9:00 まで（初期研も同じ）。　休日の昼の当直（日直）: 9:00〜17:15。", False, t),
        ("隙間なくつながる勤務は、ひと続きの勤務として数える（例: 通常勤務 8:30〜17:15 → 夜の当直 → 翌 8:30 = 24時間）。", False, t),
        ("外し方: 夜勤明けの日（平日）は、当直が平日でも休日でも、初期研以外は 12:30 まで、初期研は全日休み。", False, t),
        ("　夜の当直から翌日（休日）の昼の当直に続くときは、当直に入る日（平日）を全日休み（初期研も同じ）。", False, t),
        ("　ほかに上限を超えるときは、上限を満たすぎりぎりまで、当直に入る日などの出勤を遅らせる（遅出）。時刻は15分単位で、上限を超えない側に寄せる。", False, t),
        ("　残る勤務が1時間未満になるときは、その日を全日休みにする。", False, t),
        ("外した時間は、年次休暇ではなく勤務時間上の休み。平日オフの日は、Dr.JOY では「勤務しない日」として入れる。", False, t),
        ("夜の当直を含まないとき（日直だけなど）は、通常の決まり（始業から24時間以内に9時間休む＝続けて15時間まで）。この時刻ならいつも満たすので、何も外さない。", False, t),
        ("夜の当直を続けて組む（夜 → 翌日の夜）のは禁止。通常勤務を外しても休みが取れないため。当直勤務表アプリで提出を止める。", False, t),
        ("", False, t),
        ("■ 初期研以外（B 水準）", True, 12),
        ("上限: 夜の当直を含めて続けて 28時間まで。そのあと 18時間休む（＝始業から46時間以内に18時間の連続した休息）。", False, t),
        ("外し方の順:", False, t),
        ("　① 夜勤明けの日（平日）は 12:30 まで勤務（早帰り）。当直が平日でも休日でも同じ。", False, t),
        ("　② 夜の当直から翌日（休日）の昼の当直に続くときは、当直に入る日（平日）を全日休み。", False, t),
        ("　③ それでも 18時間休めなければ、次の勤務の日を遅出にする。", False, t),
        ("こうなる（6〜11月の当直予定で出てくる型）:", False, t),
        ("　・平日の夜の当直で、翌日も平日（例: 水曜の夜）: 木曜は 12:30 まで勤務。水曜 8:30 から木曜 12:30 まで続けて 28時間。"
         "木曜 12:30 から金曜 8:30 まで 20時間休める。", False, t),
        ("　・休日の夜の当直のあと、翌日が平日（例: 日曜の夜の当直のあとの月曜）: 月曜も 12:30 まで勤務。日曜 17:15 から 19.25時間"
         "（日曜の昼と夜の当直なら 9:00 から 27.5時間）で 28時間以内、月曜 12:30 から火曜 8:30 まで 20時間休める。"
         "上限だけなら 14:30（昼と夜なら 13:00）まで働けるが、決まりを簡単にするため平日の夜の当直のあとにそろえた（2026-10-27）。", False, t),
        ("　・平日の夜の当直のあと、翌日（休日）の昼の当直に続く（例: 金曜の夜 → 土曜の昼）: 金曜は全日休み。"
         "金曜 17:15 から土曜 17:15 まで続けて 24時間。上限だけなら金曜 13:15 から働けるが、決まりを簡単にするため全日休みにした（2026-10-27）。", False, t),
        ("　　例: 救急科 丁野先生 11/12 夜 → 11/13 昼。12日は全日休み（休暇の予定とも合う）。", False, t),
        ("　・平日の夜の当直で、翌日が休日（例: 金曜の夜。土曜 9:00 まで）: 何も外さない。金曜 8:30 から土曜 9:00 まで 24.5時間で 28時間以内。", False, t),
        ("　・（6〜11月には無いが起こりうる）休日の夜 → 翌日（休日）の昼の当直 → 次の日が平日（例: 土曜の夜 → 日曜の昼）: 月曜は 11:15 から勤務。"
         "日曜 17:15 から 18時間休むため。", False, t),
        ("", False, t),
        ("■ 初期研（C-1 水準）", True, 12),
        ("上限: 夜の当直を含めて続けて 24時間まで。そのあと 24時間休む。代償休息で補うことは原則できない。", False, t),
        ("外し方の順: ① 夜勤明けの日（平日）は全日休み。当直が平日でも休日でも同じ。② 夜の当直から翌日（休日）の昼の当直に続くときは、当直に入る日（平日）も全日休み。"
         "③ ほかに上限を超えるときは、上限を満たすぎりぎりまで、当直に入る日や次の勤務の日の出勤を遅らせる（遅出）。", False, t),
        ("こうなる（6〜11月の当直予定で出てくる型）:", False, t),
        ("　・平日の夜の当直で、翌日も平日（例: 水曜の夜）: 木曜は全日休み。水曜 8:30 から木曜 8:30 で 24時間ちょうどなので、"
         "木曜は少しも働けない。そのあと金曜 8:30 まで 24時間休む。", False, t),
        ("　・休日の夜の当直（昼と夜の当直も同じ）で、翌日が平日（例: 日曜の夜）: 月曜は全日休み。月曜 8:30 に当直を終えてから火曜 8:30 まで 24時間休むため。", False, t),
        ("　・平日の夜の当直で、翌日が休日（例: 金曜の夜。土曜 9:00 まで）: 金曜は 9:00 から勤務。8:30 からだと土曜 9:00 まで 24.5時間で 30分超えるため。", False, t),
        ("　・祝日の前の日（平日）の夜の当直（例: 11/2 月曜の夜 → 11/3 祝日）: 11/2 は 9:00 から勤務（上と同じ）。"
         "祝日の翌日（11/4）も 9:00 から勤務。11/3 の 9:00 に当直を終えてから 24時間休むため。", False, t),
        ("　・土曜の昼と夜の当直（土曜 9:00 → 日曜 9:00）のあと: 月曜は 9:00 から勤務。日曜 9:00 から 24時間休むため。", False, t),
        ("", False, t),
        ("■ 両方に共通すること", True, 12),
        ("月末の夜の当直の翌日（翌月1日）の分は、翌月分の一覧に出る（この表には出ない）。", False, t),
        ("水準は職名で決める（初期研＝C-1、それ以外＝B。Dr.JOY の特例水準の指定と同じ）。職名が分からない人は B として計算し、理由に書く。", False, t),
        ("", False, t),
        ("■ 決めたこと（管理者 2026-10-26）", True, 12),
        ("・救急科・小児科の当直は宿日直許可の無い当直。救急・小児の医師の水準は B、初期研は C-1。", False, t),
        ("・上限は B＝続けて28時間・そのあと18時間、C-1＝続けて24時間・そのあと24時間。24時間ちょうどは「まで」の内側。", False, t),
        ("・夜の当直の終わりは、翌日が平日なら 8:30、休日なら 9:00（初期研も同じ）。", False, t),
        ("・外した時間は勤務時間上の休み（年次休暇ではない）。", False, t),
        ("・夜の当直を続けて組むのは禁止。", False, t),
        ("・夜勤明けの日（平日）は、当直が平日でも休日でも、初期研以外は 12:30 まで、初期研は全日休み（2026-10-27）。", False, t),
        ("・夜の当直から翌日（休日）の昼の当直に続くときは、当直に入る日（平日）を全日休み（2026-10-27）。", False, t),
        ("・平日オフの日は、Dr.JOY では「勤務しない日」として入れる。", False, t),
        ("・オンコールは数えない。", False, t),
        ("・Dr.JOY の勤務パターンには、午前休みや 9:00 からの遅出は今は無い。", False, t),
        ("", False, t),
        ("このあと: Dr.JOY の勤務予定に入れる仕組みを作る。6〜11月に出てくる時刻は、B が「12:30 まで」、"
         "C-1 が「9:00 から」と、全日休みだけ（起こりうるものに B の「11:15 から」）。"
         "Dr.JOY にこれらの勤務パターンを足すか、日ごとに時刻を入れられれば、そのまま入る。", False, t),
    ]


def write_xlsx(path, y, m, results):
    """results = [(owner, offs, notes, duties)] → 確かめてもらうための xlsx（決め方・一覧・科ごとのカレンダー）。"""
    import calendar
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    F = "Yu Gothic"
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    head = PatternFill("solid", fgColor="D9E1F2")
    fill = {EARLY: PatternFill("solid", fgColor="FFF2CC"), LATE: PatternFill("solid", fgColor="E2EFDA"),
            SHORT: PatternFill("solid", fgColor="DDEBF7"), OFF: PatternFill("solid", fgColor="F8CBAD")}
    wb = Workbook()
    ws = wb.active
    ws.title = "決め方"
    for i, (t, b, sz) in enumerate(rules_rows(y, m), 1):
        c = ws.cell(i, 1, t)
        c.font = Font(name=F, bold=b, size=sz)
        c.alignment = Alignment(wrap_text=True, vertical="top")     # 長い文は折り返す（20261026: 例を書き足して長くなった）
    ws.column_dimensions["A"].width = 110

    ws2 = wb.create_sheet("外す通常勤務の一覧")
    hdr = ["科", "氏名", "職名", "水準", "日付", "曜日", "外す", "理由"]
    for j, h in enumerate(hdr, 1):
        c = ws2.cell(1, j, h)
        c.font, c.fill, c.border = Font(name=F, bold=True), head, box
        c.alignment = Alignment(horizontal="center")
    r = 2
    for owner, offs, notes, _duties in results:
        items = [(o["date"], o["name"], o, False) for o in offs] + [(n["date"], n["name"], n, True) for n in notes]
        for _d, _n, o, is_note in sorted(items, key=lambda x: (x[0], x[1])):
            vals = [owner, o["name"], o["title"], o["level"], o["date"], WD[o["date"].weekday()],
                    "要相談" if is_note else o["label"], o["why"]]
            for j, v in enumerate(vals, 1):
                c = ws2.cell(r, j, v)
                c.font, c.border = Font(name=F, color="C00000" if (is_note and j == 7) else None), box
                if j == 5:
                    c.number_format = "m/d"
                if j == 7 and not is_note:
                    c.fill = fill[o["kind"]]
            r += 1
    for j, w in enumerate([12, 16, 14, 6, 8, 6, 16, 80], 1):
        ws2.column_dimensions[get_column_letter(j)].width = w
    ws2.freeze_panes = "A2"

    nd = calendar.monthrange(y, m)[1]
    hol = C.load_holidays()
    for owner, offs, notes, duties in results:
        ws3 = wb.create_sheet(f"{owner}（カレンダー）"[:31])
        ws3.cell(1, 1, f"{owner}　R{y - 2018}年{m}月　9:00〜＝その時刻から勤務（遅出）　〜12:30＝その時刻まで勤務（早帰り）　全休＝全日休み　"
                       "！＝要相談　灰字＝当直（昼・夜）　灰色の列は土日祝").font = Font(name=F, bold=True, size=12)
        ws3.cell(3, 1, "氏名（水準）").font = Font(name=F, bold=True)
        for d in range(1, nd + 1):
            day = dt.date(y, m, d)
            for rr, v in ((2, d), (3, WD[day.weekday()])):
                c = ws3.cell(rr, d + 1, v)
                c.font, c.border = Font(name=F, bold=True), box
                c.alignment = Alignment(horizontal="center")
                if not workday(day, hol):
                    c.fill = PatternFill("solid", fgColor="DDDDDD")
            ws3.column_dimensions[get_column_letter(d + 1)].width = 6
        by, names = duty_map(duties, hol)
        inm = [o for o in offs if (o["date"].year, o["date"].month) == (y, m)]
        nim = [n for n in notes if (n["date"].year, n["date"].month) == (y, m)]
        lv = {o["login"]: o["level"] for o in inm + nim}
        people = sorted(lv, key=lambda k: (lv[k] != "B", names.get(k, k)))
        for i, key in enumerate(people):
            rr = 4 + i
            c = ws3.cell(rr, 1, f"{names.get(key, key)}（{lv[key]}）")
            c.font, c.border = Font(name=F), box
            offd = {o["date"]: o for o in inm if o["login"] == key}
            noted = {n["date"] for n in nim if n["login"] == key}
            for d in range(1, nd + 1):
                day = dt.date(y, m, d)
                c = ws3.cell(rr, d + 1)
                c.border = box
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if not workday(day, hol):
                    c.fill = PatternFill("solid", fgColor="EEEEEE")
                duty = "".join(b for b in ("昼", "夜") if b in by.get(key, {}).get(day, ()))
                parts = []
                if day in offd:
                    o = offd[day]
                    parts.append("全休" if o["kind"] == OFF else
                                 (f"{hm(o['time'][0])}〜" if o["kind"] == LATE else
                                  f"〜{hm(o['time'][1])}" if o["kind"] == EARLY else f"{hm(o['time'][0])}〜{hm(o['time'][1])}"))
                    c.fill = fill[o["kind"]]
                if day in noted:
                    parts.append("！")
                if duty:
                    parts.append(duty)
                c.value = "\n".join(parts) if parts else None
                c.font = Font(name=F, bold=bool(day in offd or day in noted), size=9,
                              color="C00000" if day in noted else ("808080" if not (day in offd) else None))
            ws3.row_dimensions[rr].height = 26
        ws3.column_dimensions["A"].width = 18
        ws3.freeze_panes = "B4"
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM")
    ap.add_argument("--owner", action="append")
    ap.add_argument("--out")
    ap.add_argument("--json", help="Dr.JOY に入れる側が読む JSON の書き出し先")
    a = ap.parse_args()
    y, m = (int(x) for x in a.month.split("-"))
    hol = C.load_holidays()
    results = []
    for owner in a.owner or C.consecutive_warn_owners():
        duties, titles = month_data(y, m, owner)
        offs, notes = plan(duties, hol, owner, titles)
        offs = [o for o in offs if (o["date"].year, o["date"].month) == (y, m)]
        notes = [n for n in notes if (n["date"].year, n["date"].month) == (y, m)]
        c = collections.Counter((o["level"], o["kind"]) for o in offs)
        print(f"{owner}: " + "・".join(f"{lv} {k} {n}" for (lv, k), n in sorted(c.items())) + f"・要相談 {len(notes)}")
        for n in notes:
            print(f"   要相談: {n['name']}（{n['level']}） {n['date']:%m/%d} {n['why']}")
        results.append((owner, offs, notes, duties))
    if a.out:
        write_xlsx(a.out, y, m, results)
        print("書いた:", a.out)
    if a.json:
        write_json(a.json, y, m, results)
        print("書いた:", a.json)


def to_json(y, m, results):
    """results = [(owner, offs, notes, duties)] → schema hospital.day-off-plan/1 の dict。"""
    def one(o, with_time=True):
        d = {k: o.get(k, "") for k in ("owner", "login", "name", "title", "level")}
        d["date"] = o["date"].isoformat()
        if with_time:
            st = o.get("time")
            d.update(kind=o["kind"], start=hm(st[0]) if st else None, end=hm(st[1]) if st else None, label=o.get("label", ""))
        d["why"] = o.get("why", "")
        return d
    return {"schema": "hospital.day-off-plan/1", "month": f"{y:04d}-{m:02d}",
            "made": dt.datetime.now().isoformat(timespec="seconds"),
            "offs": [one(o) for _ow, offs, _n, _d in results for o in offs],
            "notes": [one(n, False) for _ow, _o, notes, _d in results for n in notes]}


def write_json(path, y, m, results):
    import json
    import master_store as MST
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = MST.tmp_name(path)
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(to_json(y, m, results), f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)
    return path


if __name__ == "__main__":
    main()
