# -*- coding: utf-8 -*-
r"""shozoku.py — Dr.JOY の所属（スタッフ所属履歴）と、所属を決める提出単位の決まり。20261031

管理者 2026-09-27 決定:
  アプリの左の名簿（提出単位 owner ごと・月ごと）は、その科の Dr.JOY 所属医師の申告。
  所属は今日の staff_*.csv（名簿\ の最新）の所属で見る。
  ⚠ 同日の訂正（管理者）: スタッフ所属履歴は正本ではない。初期研の1年分のローテーションを一括で入れるための道具で、
    所属の変更・配置のほとんどは統合ファイルからのマクロ（ときどき事務担当・管理者が手で入れる）。
    だから所属履歴で対象月の1日の所属を引くのは、職位が初期研（contract.rank_of == "初"）の人だけ。
    それ以外の人は今日の staff CSV の所属だけを使い、所属履歴の未来の組（消し残しの「他院」など）は見ない。
  ⚠ 20261028 管理者: 初期研・外科・内科全部はローテーションの予定を事務担当が department_history.csv で
    日付付きで入れている（ほかの科の異動も入っている＝11/1 救急部→放射線科・11/16 他院→脳神経外科など）。
    そこで初期研以外も「今日より後に始まる組」（予定）だけは見る。今日以前に始まった組は今日の staff CSV が正。
    ただし「他院」の組は見ない（2026-09-06 決定: 他院は初期研だけ。専攻医は院外研修中も各科の所属のまま）。
    初期研以外で他院にいる人・他院の予定がある人は ingai_violations() で拾い、事務担当へのメールで知らせる。

所属履歴の置き場 = 本番の 管理ツール\所属履歴\所属履歴_YYYY.xlsx（毎日 3:30 の名簿更新で上書き。11〜12月は翌年の分も置く。
20261009 に 名簿\ から移した＝名簿\ は全職員が読めるが、所属履歴には医師以外の職員コード・ログインID・所属の移り変わりが入っている）。
  シート1枚。見出し: 職員コード, ログインID, 氏名, 所属（まとめ・使わない）, 職業, 所属1, 開始日1, 所属2, 開始日2, …
  開始日は "2026/05/28" の文字列（Excel の日付・数値も受ける）。所属は "/外科(消化管)/"（2つなら "/循環器内科/集中治療部/"）。

初期研の日付 D の所属 = D の年のファイルで 開始日 <= D の最後の組の所属。
  その年に無ければ前の年のファイル。どちらにも無ければ今日の staff CSV の所属
  （D より前に予定された変更が無い＝今日の所属のままのはず）。
ファイルが無い・読めないときは落ちずに staff CSV へ倒す（problems に理由を残す）。
職位は staff CSV の職名で見る（dept_on に title を渡せばそれで見る）。

所属を決める提出単位（defining owners）= 活きている提出単位から EXCLUDE_OWNERS を除いたもの。
  owner → Dr.JOY の所属 = data/dept_map.json の map（{"dept":[...],"rank":[...]} の形は dept だけを使う）。
  ある人がその owner に「合う」= その人の所属の区切り（/…/）のどれかに、owner のパターンのどれかが含まれる。
当直をしない部署（NON_DUTY_DEPTS）の人を科から外す提案はしない（別の見出しに出す）。
初期研（contract.rank_of == "初"）と名簿外（unreg: / unregistered）は、所属の照合から外す。
"""
from __future__ import annotations

import datetime as dt
import glob
import hashlib
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

HISTORY_DIR = os.path.join(C.PROD_ADMIN, "所属履歴")   # テストでは差し替える。20261009: 名簿\ から移した（全職員が読めたため）
STAFFS = None                               # 今日の所属を読む staff CSV（None=最新の staff_*.csv。テストでは差し替える）
TODAY = None                                # 「今日」（None=本当の今日。テストでは日付を固定する。20261028）
HISTORY_GLOB = "所属履歴_*.xlsx"
EXCLUDE_OWNERS = ("集中治療室", "内科", "管理当直", "内科C 初期研")
NON_DUTY_DEPTS = ("リハビリテーション部", "病理診断科", "医療安全", "研修管理室", "所属未分類")
EXCLUDE_RANKS = ("初",)
# 本番の最初の月。これより前の月（R8.10 など）は予行で、「前の月」として区分の引き継ぎ・比べに使わない
# （管理者 20260927: 10月フォルダは試しに使う。dtx14_daily_sync.SYNC_FROM と同じ月）
REAL_FROM = (2026, 11)
_SEG_SPLIT = re.compile(r"[/／、]")


def history_path(dir_, year):
    return os.path.join(dir_, f"所属履歴_{year}.xlsx")


def parse_date(v):
    """開始日のセル → date（読めなければ None）。文字列 2026/05/28・2026-05-28、datetime/date、Excel の通し番号。"""
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 1 <= v < 200000:
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        return None
    m = re.match(r"\s*(\d{4})[/\-.年](\d{1,2})[/\-.月](\d{1,2})", str(v))
    if not m:
        return None
    try:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def read_history_file(path):
    """所属履歴の xlsx 1本 → {ログインID(小文字): [(開始日, 所属, 並び順)]}。見出しの名前で列を探す。"""
    import warnings
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")          # 「Workbook contains no default style」
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        head = None
        for r in rows:
            if r and any(str(c or "").strip() == "ログインID" for c in r):
                head = [str(c or "").strip() for c in r]
                break
        if head is None:
            raise ValueError("見出し（ログインID）が見つからない")
        i_login = head.index("ログインID")
        pairs = []                                # (所属の列, 開始日の列)
        for j, h in enumerate(head):
            m = re.fullmatch(r"所属(\d+)", h)
            if m and f"開始日{m.group(1)}" in head:
                pairs.append((j, head.index(f"開始日{m.group(1)}")))
        out = {}
        for r in rows:
            login = str(r[i_login] or "").strip().lower() if i_login < len(r) else ""
            if not login:
                continue
            lst = out.setdefault(login, [])
            for n, (jd, js) in enumerate(pairs):
                dept = str(r[jd] or "").strip() if jd < len(r) else ""
                start = parse_date(r[js]) if js < len(r) else None
                if dept and start:
                    lst.append((start, dept, n))
        return out
    finally:
        wb.close()


class History:
    """所属履歴（年ごとに遅延読み込み）＋今日の staff CSV。dept_on(login, date) で日付の所属を引く。"""

    def __init__(self, dir_=None, staffs=None, log=None, today=None):
        self.dir = dir_ or HISTORY_DIR
        self.today = today or TODAY or dt.date.today()     # これより後に始まる組＝予定（初期研以外はこれだけ見る。20261028）
        self.log = log or (lambda s: None)
        self.problems = []                        # 読めなかったものの説明
        self.files = []                           # 読めた所属履歴 [(ファイル名, 更新日時)]
        self._years = {}
        self.people = {}                          # ログインID(小文字) → staff CSV の1人
        self.staffs_name = ""
        self._load_staffs(staffs)

    def _load_staffs(self, staffs):
        people = None
        if isinstance(staffs, (list, tuple)):
            people = list(staffs)
        else:
            path = staffs or STAFFS or C.newest_staffs()
            try:
                people = C.load_staffs(path)[0]
                self.staffs_name = os.path.basename(path)
            except Exception as e:                # 読めなくても落ちない（所属履歴だけで引ける人は引く）
                self.problems.append(f"名簿（staff CSV）が読めない: {type(e).__name__}: {e}")
                people = []
        for p in people:
            lg = str(p.get("login") or "").strip().lower()
            if lg:
                self.people[lg] = p

    def year_pairs(self, year):
        if year not in self._years:
            path = history_path(self.dir, year)
            data = {}
            try:
                if os.path.isfile(path):
                    data = read_history_file(path)
                    self.files.append((os.path.basename(path),
                                       dt.datetime.fromtimestamp(os.path.getmtime(path))))
            except Exception as e:
                self.problems.append(f"{os.path.basename(path)} が読めない: {type(e).__name__}: {e}")
                self.log(f"  ⚠ {os.path.basename(path)} が読めない（{e}）→ 今日の所属で見る")
                data = {}
            self._years[year] = data
        return self._years[year]

    def today_dept(self, login):
        p = self.people.get(str(login or "").strip().lower())
        return (p or {}).get("dept") or ""

    def _title(self, k, title):
        return title if title is not None else ((self.people.get(k) or {}).get("title") or "")

    def _pairs(self, k, date, title):
        """date までに始まる、見てよい組 [(開始日, 所属, 番号)]。初期研は全部（D の年と前の年のファイル）、
        ほかの人は今日より後に始まる組（日付付きの予定）だけ。"""
        if C.rank_of(title) == "初":
            for y in (date.year, date.year - 1):
                cands = [x for x in self.year_pairs(y).get(k, []) if x[0] <= date]
                if cands:
                    return cands
            return []
        out = []
        for y in range(self.today.year, date.year + 1):
            out += [x for x in self.year_pairs(y).get(k, []) if self.today < x[0] <= date and not is_ingai(x[1])]
        return out

    def ingai_violations(self, until):
        """初期研以外の医師で、今日の所属が他院の人と、until までに他院に移る予定の組がある人
        （2026-09-06 決定: 他院は初期研だけ）。戻り: [(氏名, 職名, 説明)] かな順。"""
        out = []
        for k, p in sorted(self.people.items(), key=lambda kv: (kv[1].get("kana") or "", kv[0])):
            if not p.get("is_doctor", True) or C.rank_of(p.get("title")) == "初":
                continue
            what = []
            if is_ingai(p.get("dept")):
                what.append("今の所属が他院")
            for y in range(self.today.year, until.year + 1):
                for s_, dep, _n in sorted(self.year_pairs(y).get(k, [])):
                    if self.today < s_ <= until and is_ingai(dep):
                        what.append(f"{s_.month}/{s_.day} から他院の予定")
            if what:
                out.append((p.get("name") or k, p.get("title") or "", "・".join(what)))
        return out

    def dept_on(self, login, date, title=None):
        """その日の Dr.JOY の所属（無ければ ""）。初期研は所属履歴（20260927）。
        ほかの人は今日の staff CSV の所属に、今日より後に始まる所属履歴の組（日付付きの予定）を重ねる（20261028）。
        title = 職名（省略=staff CSV の職名）。"""
        k = str(login or "").strip().lower()
        if not k:
            return ""
        cands = self._pairs(k, date, self._title(k, title))
        if cands:
            return max(cands, key=lambda x: (x[0], x[2]))[1]
        return self.today_dept(k)

    def depts_during(self, login, start, end, title=None):
        """start〜end に就く所属（start の所属＋その間に始まる組）。月の途中の異動を見るのに使う（20261028）。"""
        k = str(login or "").strip().lower()
        if not k:
            return []
        t = self._title(k, title)
        out = [self.dept_on(k, start, t)]
        for s_, dep, _n in sorted(self._pairs(k, end, t)):
            if start < s_ <= end and dep not in out:
                out.append(dep)
        return [d for d in out if d]


def load_history(dir_=None, staffs=None, log=None, today=None):
    """所属履歴を読む準備をして History を返す（ファイルは要るときに読む）。
    staffs = staff CSV のパス（省略=最新の staff_*.csv）か、contract.load_staffs の people のリスト。"""
    return History(dir_, staffs, log, today)


def dept_on(login, date, history, title=None):
    """History.dept_on の関数版（所属履歴を見るのは初期研だけ）。"""
    return history.dept_on(login, date, title)


def is_ingai(dept):
    """所属が他院か（区切りのどれか）。"""
    return "他院" in str(dept or "")


def segments(dept):
    """'/循環器内科/集中治療部/' → ['循環器内科', '集中治療部']"""
    return [s.strip() for s in _SEG_SPLIT.split(str(dept or "")) if s.strip()]


def dept_display(dept):
    """人に見せる所属（区切りを「、」に）。"""
    return "、".join(segments(dept))


def matches(dept, patterns):
    """所属の区切りのどれかに、パターンのどれかが含まれるか。"""
    segs = segments(dept)
    return any(p and p in s for s in segs for p in patterns)


def chosen_seg(mem, patterns):
    """集合当直（パターンが2つ以上）の「貴科リスト」の行で、庶務が選んだ「Dr.JOY所属」（行の seg）。無い・パターンに無ければ None。
    20261031 管理者（コードレビュー③）: 内科 消化器・呼吸器などは診療科ではないので、医師ごとにどちらの所属かを持たせる。"""
    s = str((mem or {}).get("seg") or "").strip()
    return s if len(patterns) >= 2 and s in patterns else None


def member_seg(mem, dept, patterns):
    """集合当直の行の「Dr.JOY所属」= 選んだもの（seg）→ 無ければ Dr.JOY の所属で当たるパターン → 無ければ None。
    パターンが1つ以下（ふつうの科）は None（見出しは提出単位のまま）。"""
    if len(patterns) < 2:
        return None
    return chosen_seg(mem, patterns) or next((p for p in patterns if matches(dept, [p])), None)


def is_non_duty(dept):
    """当直をしない部署（リハビリテーション部・病理診断科・医療安全部・研修管理室・所属未分類）か。"""
    return matches(dept, NON_DUTY_DEPTS)


def load_dept_map():
    with open(os.path.join(C.DATA, "dept_map.json"), encoding="utf-8-sig") as f:
        return json.load(f)["map"]


def owner_patterns(owner, dept_map):
    """owner → Dr.JOY の所属のパターン（dict の形は dept だけ）。"""
    return C.dept_map_entry(dept_map.get(owner, []))[0]


def defining_owners(frames=None, dept_map=None):
    """所属を決める提出単位（枠マスタの並び・retired と EXCLUDE_OWNERS を除く・パターンが無いものも除く）。"""
    frames = frames if frames is not None else C.load_frames()
    dept_map = dept_map if dept_map is not None else load_dept_map()
    return [o for o in C.owners_of(frames)
            if o not in EXCLUDE_OWNERS and owner_patterns(o, dept_map)]


def is_unregistered(member):
    return bool(member.get("unregistered")) or str(member.get("login") or "").startswith("unreg:")


def is_excluded(member, title=None):
    """所属の照合から外す人か（名簿外・初期研）。title を渡せばそれで職位を見る（無ければ member の title）。"""
    if is_unregistered(member):
        return True
    return C.rank_of(title if title is not None else member.get("title")) in EXCLUDE_RANKS


def config_block(frames=None, dept_map=None):
    """config/current.json の "shozoku"（アプリも同じ決まりで見られるように）。"""
    return {"owners": defining_owners(frames, dept_map),
            "exclude_owners": list(EXCLUDE_OWNERS),
            "non_duty_depts": list(NON_DUTY_DEPTS),
            "exclude_ranks": list(EXCLUDE_RANKS),
            "real_from": "%04d-%02d" % REAL_FROM}


def history_files(dir_=None):
    d = dir_ or HISTORY_DIR
    try:
        return sorted(f for f in glob.glob(os.path.join(d, HISTORY_GLOB))
                      if re.fullmatch(r"所属履歴_\d{4}\.xlsx", os.path.basename(f)))
    except OSError:
        return []


def content_hash(dir_=None):
    """所属履歴（全年分）の中身のハッシュ。xlsx はバイト列が毎回変わりうるのでセルの値で取る。無ければ ""。
    読めないファイルはバイト列で代える（変わったとみなされる側に倒す）。"""
    files = history_files(dir_)
    if not files:
        return ""
    import warnings
    import openpyxl
    h = hashlib.sha256()
    for f in files:
        h.update(os.path.basename(f).encode("utf-8"))
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            try:
                for r in wb.worksheets[0].iter_rows(values_only=True):
                    h.update(repr(tuple("" if c is None else str(c) for c in r)).encode("utf-8"))
            finally:
                wb.close()
        except Exception:
            with open(f, "rb") as fh:
                h.update(fh.read())
    return h.hexdigest()
