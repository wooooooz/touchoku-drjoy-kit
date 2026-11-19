# -*- coding: utf-8 -*-
"""contract.py  20261015
Touchoku_Plan のデータ契約エンジン（手当計算非依存の自己完結版）。

  - 原子レコード = 1人 × 1日 × 1バンド（昼/夜） × 1枠（設計案§1.1）
  - 20260831 ユーザー決定: バンドは常に昼/夜の2区分（「昼夜」「全日」バンドは廃止）。
    旧形式の行（昼夜列が「昼夜」「全日」または空欄）は取込時に昼+夜へ自動展開する（後方互換）
  - 人物キー = Dr.JOYログインID（staffs.csvのログインID列）
  - 氏名解決 = exact → NFC → canon(異体字21組) の3段。自由入力は受けない
  - 提出（dept単位のJSON / 入力規則付きExcelテンプレ / CSV）の読取と検証
  - 原子レコード → DTX14互換の人別3シート（人別当直日/人別オンコール日/人別ファースト）

依存: openpyxl のみ（祝日は data/holidays.csv を読む。jpholidayは使わない）。
"""
from __future__ import annotations

import csv
import datetime as dt
import io
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if getattr(sys, "frozen", False):          # PyInstaller の exe: data/ は同梱先（_MEIPASS）から読む（20260904）
    ROOT = getattr(sys, "_MEIPASS", ROOT)

# 本番共有（ユーザー決定 2026-09-09。月ごとにフォルダを分ける＝10月と11月が同時に活きる期間があるため）
#   当直勤務表\_本体\ 原本exe ／ 名簿\ staff_YYYYMMDD.csv＋院外研修期間表 ／ 科別ツール\ ／ 管理ツール\ ／
#   R8年N月\ 当直勤務表入力_R8年N月.exe・config\・提出\・集約\・テンプレ\・HTML（tools/month_setup.py が作る）
PROD_ROOT = r"\\fileserver\share$\00_DrJOY\当直勤務表"
PROD_MASTER = os.path.join(PROD_ROOT, "_本体")
PROD_ROSTER = os.path.join(PROD_ROOT, "名簿")
PROD_TOOLS = os.path.join(PROD_ROOT, "科別ツール")
PROD_ADMIN = os.path.join(PROD_ROOT, "管理ツール")
# 職員コードの入った名簿は管理者だけの 管理ツール\名簿\ に置き、全職員が読む 名簿\ の写しは職員コードの列を空にする
# （20261015 管理者: 職員コードは人事給与システムの職員番号と同じ11桁。統合ファイルでは暗号にしている値）。
# 管理側の処理はこちらを先に読み、読めない PC（科別ツールを使う各科）は 名簿\ を読む。
PROD_ROSTER_FULL = os.path.join(PROD_ADMIN, "名簿")


def resolve_data(frozen=None, root=None, admin_dir=None):
    r"""data\ の場所（20260926）。exe（PyInstaller）では、本番の 管理ツール\data\ に frames.json があればそれを優先し、
    無ければ（届かない・読めない＝管理者でない）同梱先（_MEIPASS）。toban.csv・frames.json・holidays.csv を直しても、
    管理ツールを作り直さずに deploy_prod --what admin で 管理ツール\data\ へ置けば①に届く。exe でなければリポジトリの data\。"""
    frozen = getattr(sys, "frozen", False) if frozen is None else frozen
    root = ROOT if root is None else root
    if frozen:
        prod = os.path.join(PROD_ADMIN if admin_dir is None else admin_dir, "data")
        try:
            if os.path.isfile(os.path.join(prod, "frames.json")):
                return prod
        except OSError:
            pass
    return os.path.join(root, "data")


DATA = resolve_data()


def month_dir_name(year, month):
    return f"R{year - 2018}年{month}月"


def prod_month_dir(year, month):
    return os.path.join(PROD_ROOT, month_dir_name(year, month))


def parse_month_dir(name):
    """'R8年11月' → (2026, 11)。違う名前なら None"""
    m = re.fullmatch(r"R(\d{1,2})年(\d{1,2})月", name)
    if not m:
        return None
    return 2018 + int(m.group(1)), int(m.group(2))


def find_kikan():
    r"""院外研修期間表: 名簿\ にあればそれ（最新）、無ければ従来の Desktop のファイル、無ければ None"""
    import glob
    cands = sorted(glob.glob(os.path.join(PROD_ROSTER, "*期間表*.xlsx")), key=os.path.getmtime) if os.path.isdir(PROD_ROSTER) else []
    if cands:
        return cands[-1]
    # 20261015: 管理者のデスクトップのフォルダは「名簿たち」から「ソースリスト」に名前を変えた（狙われやすい語を避ける）
    p = os.path.join(os.path.expanduser("~"), "Desktop", "ソースリスト", "令和８年度専門研修医他院期間表.xlsx")
    return p if os.path.isfile(p) else None

P_DUTY = "人別当直日"
P_ONCALL = "人別オンコール日"
P_FIRST = "人別ファースト"

# ---------------------------------------------------------------- 氏名の正規化
# dtx14_names._ITAIJI と同一表（vendored 2026-08-30。正本は drjoy_import/dtx14_names.py）
_ITAIJI = str.maketrans({
    "髙": "高", "﨑": "崎", "桒": "桑", "濵": "浜", "栁": "柳", "德": "徳",
    "眞": "真", "淺": "浅", "靑": "青", "賴": "頼", "曻": "昇", "硏": "研",
    "瀨": "瀬", "圡": "土", "祐": "祐", "凰": "凰",
})


def nfc(s: str) -> str:
    """CJK互換漢字を正規字へ（隆 U+F9DC→U+9686 の実例）。"""
    return unicodedata.normalize("NFC", str(s or ""))


def canon(name: str) -> str:
    return str(name or "").replace(" ", "").replace("　", "").translate(_ITAIJI)


def namekey(s: str) -> str:
    return str(s or "").replace(" ", "").replace("　", "")


# ---------------------------------------------------------------- 祝日・名簿・枠

def load_holidays(path=None):
    """data/holidays.csv → {date: name}。土日は含まない（is_holidayが曜日を見る）。"""
    path = path or os.path.join(DATA, "holidays.csv")
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("date"):
                out[dt.date.fromisoformat(row["date"].strip())] = (row.get("name") or "").strip()
    return out


def is_holiday(d: dt.date, holidays) -> bool:
    return d.weekday() >= 5 or d in holidays


def daytime_oncall_only(path=None):
    """data/holidays.csv の daytime_oncall_only=1 の日 → {date}。
    休診日だが宿日直手当が出ない日（例 開院記念日・管理者 20260922）。昼のファーストは立つが全員オンコール。
    アプリは当直で出されてもよく、集約と DTX14 書き出しの側で昼を全部オンコールに変える（apply_daytime_oncall）。"""
    path = path or os.path.join(DATA, "holidays.csv")
    out = set()
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("date") and str(row.get("daytime_oncall_only") or "").strip() in ("1", "true", "True"):
                out.add(dt.date.fromisoformat(row["date"].strip()))
    return out


def load_toban(path=None):
    """data/toban.csv（tools/toban_extract.py が DTX14 の「当番当直」シートから作る）→ [(date, band, dept)]。
    当番当直 = その日・そのバンドの当直の要（管理者 20260922）。dept は DTX14 ファースト表の科名（dept_dtx14）。"""
    path = path or os.path.join(DATA, "toban.csv")
    out = []
    if not os.path.isfile(path):
        return out
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("date") and row.get("band") and row.get("dept"):
                out.append((row["date"].strip(), row["band"].strip(), row["dept"].strip()))
    return out


def apply_daytime_oncall(records, days=None):
    """昼のレコードを、daytime_oncall_only の日だけ全部オンコールにする。戻り: 変えた件数。"""
    days = daytime_oncall_only() if days is None else days
    if not days:
        return 0
    iso = {d.isoformat() for d in days}
    n = 0
    for r in records:
        if r.get("band") == "昼" and r.get("date") in iso and not r.get("oncall"):
            r["oncall"] = True
            n += 1
    return n


def newest_staffs(extra_dirs=()):
    """名簿の既定（20260909 ユーザー指示）:
      1) 当直アプリ用に絞った staff_YYYYMMDD.csv（職業=医師。make_staff_csv.py が作る）を
         extra_dirs → 本番の 名簿\（UNC）→ OneDrive の旧 config → Downloads の順に探し、最新（更新日時）を使う
      2) 無ければ Dr.JOY の生のエクスポート staffs*.csv（500行以上）の最新
      3) それも無ければ従来の固定パス"""
    import glob
    home = os.path.expanduser("~")
    onedrive_cfg = os.path.join(home, "OneDrive - Contoso", "(C) 当直勤務表アプリ", "当直勤務表", "config")
    dirs = list(extra_dirs) + [PROD_ROSTER, onedrive_cfg, os.path.join(home, "Downloads")]

    def pick(pattern, min_rows, exact=None, only=None):
        cands = []
        for d in (only or dirs):
            try:
                files = glob.glob(os.path.join(d, pattern))
            except OSError:
                continue
            for f in files:
                if exact and not re.fullmatch(exact, os.path.basename(f)):
                    continue
                try:
                    with open(f, "rb") as fh:
                        n = sum(1 for _ in fh)
                except OSError:
                    continue
                if n >= min_rows:
                    cands.append((os.path.getmtime(f), f))
        return max(cands)[1] if cands else None

    # 20261015: 職員コードの入った 管理ツール\名簿\ を先に（同じ日に 名簿\ にも置くので、更新日時の比べだけだと
    #   職員コードが空の写しを選ぶことがある）。読めない PC では glob が空になって次へ進む
    return (pick("staff_????????.csv", 200, exact=r"staff_\d{8}\.csv", only=list(extra_dirs) + [PROD_ROSTER_FULL])
            or pick("staff_????????.csv", 200, exact=r"staff_\d{8}\.csv") or pick("staffs*.csv", 500)
            or os.path.join(home, "Downloads", "staffs (65).csv"))


_CODE_MAP = {}


def staff_code_map(path=None):
    """ログインID → 職員コード（管理側の処理だけが使う。20261015: 全職員が読む config・予定表・名簿\\ からは職員コードを
    外したので、人別シートの H列と交代記録の from_code/to_code はここから引く）。読めなければ空の dict。"""
    p = path or newest_staffs()
    if p in _CODE_MAP:
        return _CODE_MAP[p]
    m = {}
    try:
        people, _e, _c = load_staffs(p)
        m = {x["login"]: x["code"] for x in people if x.get("login") and x.get("code")}
    except (OSError, ValueError, UnicodeDecodeError):
        pass
    _CODE_MAP[p] = m
    return m


def code_of(login, path=None):
    """ログインID の職員コード（無ければ ""）。"""
    return staff_code_map(path).get(login or "", "") if login else ""


def load_staffs(path):
    """staffs.csv(CP932) → (people[list], exact{namekey:[i]}, canon_map{canonkey:[i]})。
    people[i] = {login, code, name('姓 名'), kana, dept, title, job, is_doctor}"""
    people = []
    exact = defaultdict(list)
    canon_map = defaultdict(list)
    with open(path, encoding="cp932", newline="") as f:
        for row in csv.DictReader(f):
            sei = (row.get("姓") or "").strip()
            mei = (row.get("名") or "").strip()
            if not sei:
                continue
            job = (row.get("職業") or "").strip()
            p = {
                "login": (row.get("ログインID") or "").strip(),
                "code": (row.get("職員コード") or "").strip(),
                "name": f"{sei} {mei}".strip(),
                "kana": f"{(row.get('せい') or '').strip()} {(row.get('めい') or '').strip()}".strip(),
                "dept": (row.get("所属") or "").strip(),
                "title": (row.get("職名") or "").strip(),
                "job": job,
                "is_doctor": job in ("医師", "歯科医師"),
            }
            i = len(people)
            people.append(p)
            k = namekey(sei + mei)
            exact[k].append(i)
            canon_map[canon(nfc(k))].append(i)
    return people, exact, canon_map


def resolve_name(fullname, people, exact, canon_map):
    """氏名 → (person|None, how)。how=exact/nfc/canon/ambiguous/none。"""
    k = namekey(fullname)
    for key, how in ((k, "exact"), (nfc(k), "nfc")):
        hits = exact.get(key, [])
        if len(hits) == 1:
            return people[hits[0]], how
        if len(hits) > 1:
            return None, "ambiguous"
    hits = canon_map.get(canon(nfc(k)), [])
    if len(hits) == 1:
        return people[hits[0]], "canon"
    if len(hits) > 1:
        return None, "ambiguous"
    return None, "none"


def load_frames(path=None):
    path = path or os.path.join(DATA, "frames.json")
    with open(path, encoding="utf-8-sig") as f:
        return json.load(f)["frames"]


def consecutive_warn_owners(path=None):
    """連続当直の警告を出す提出単位（frames.json の consecutive_warn_owners。20260914 管理者決定: 救急科と小児科だけ）。
    無ければ空＝どの科にも出さない。"""
    path = path or os.path.join(DATA, "frames.json")
    with open(path, encoding="utf-8-sig") as f:
        return list(json.load(f).get("consecutive_warn_owners") or [])


def is_active(frame):
    """入力させる枠か。retired=true は消滅した科の枠（20260913 管理者決定: 内科 旧科）。
    アプリの科セレクタ・当直表xlsx・未提出リストからは外すが、取込と互換出力では
    今までどおり解釈する（過去月の現物にレコードが残っているため）。"""
    return not frame.get("retired")


def active_frames(frames):
    return [f for f in frames if is_active(f)]


def owners_of(frames):
    """提出単位の一覧（枠マスタの並び順・retired は除く）。"""
    out = []
    for f in active_frames(frames):
        o = f.get("owner", f["dept"])
        if o not in out:
            out.append(o)
    return out


def rank_of(title):
    """職名 → 職位の印（教/研/専/初、その他は空）。当直表の印・名簿の並び・dept_map の職位条件で共用。"""
    t = str(title or "")
    if "初期研" in t or "初期研修" in t:
        return "初"
    if "専攻医" in t or "専門研修" in t:
        return "専"
    if "大学院生" in t:
        return "研"
    if "医師" in t or "病院長" in t or "指定医" in t:
        return "教"
    return ""


# 勤務の区分（管理者 20260927）: 名簿の各人が Dr.JOY の週パターンをどうするか。
#   full  = フルタイム（平日は毎日当院で日中勤務。産休・育休・病休中も含む）→ 週パターン「通常」
#   part  = パートタイム（当直・外来など一部だけ来る）→ 週パターン「勤務なし」
#   ingai = 院外研修専従（院外研修中で当院の仕事はしない。科の先生のまま）→ 週パターン「勤務なし」
WORK_TYPES = ("full", "part", "ingai")
WORK_LABEL = {"full": "フルタイム", "part": "パートタイム", "ingai": "院外研修専従"}


def work_of(member):
    """名簿の1人 → 勤務の区分（full/part/ingai）。work があればそれ。
    無い旧い名簿は in_house で読む（in_house が False = 旧「当直のみ」→ part、それ以外 → full）。"""
    w = (member or {}).get("work")
    if w in WORK_TYPES:
        return w
    return "part" if (member or {}).get("in_house") is False else "full"


def dept_map_entry(entry):
    """dept_map の1項目 → (所属の部分一致パターン, 職位の条件)。
    旧形式 = パターンの配列（職位を問わない）。
    20260913: {"dept": [...], "rank": ["初"]} の形も受ける
    （内科C 初期研 = 血液膠原病・神経に所属する初期研だけ。管理者）。"""
    if isinstance(entry, dict):
        return list(entry.get("dept") or []), list(entry.get("rank") or [])
    return list(entry or []), []


def matches_dept_map(entry, person):
    """その人が owner の初期候補に入るか（医師であることは呼び出し側で見る）。"""
    pats, ranks = dept_map_entry(entry)
    dept = person.get("dept") or ""
    if not any(x and x in dept for x in pats):
        return False
    return not ranks or rank_of(person.get("title")) in ranks


def expand_bands(shape, d: dt.date, holidays):
    """枠のshapeと日付からバンドを決める（Module9/6/18の全日展開と同じ規則）。"""
    if shape == "全日":
        return ["昼", "夜"] if is_holiday(d, holidays) else ["夜"]
    return [shape]


def band_stands(frame, band, d: dt.date, holidays, oncall_only=None) -> bool:
    """その日にその枠のそのバンドが立つか（当直表xlsx の灰色のマスと同じ規則。duty_sheet.Col.stands_on）。
    昼は土日祝だけ。skip_daytime_oncall_only の枠（管理当直）は開院記念日などの昼も立たない（管理者 20260923）。
    夜・全日・昼夜はいつも立つ（全日・昼夜の昼は expand_bands 側で土日祝だけに絞る）。"""
    if band != "昼":
        return True
    if frame.get("skip_daytime_oncall_only"):
        days = daytime_oncall_only() if oncall_only is None else oncall_only
        if d in days:
            return False
    return is_holiday(d, holidays)


_WD_JP = "月火水木金土日"


def md_label(d: dt.date) -> str:
    """理由の文に出す日付（11月3日（火））。"""
    return f"{d.month}月{d.day}日（{_WD_JP[d.weekday()]}）"


ROLE_LABEL = {"first": "（ファースト）", "second": "（セカンド以下）"}


def frame_label(f, frames_own):
    """テンプレの「枠」列に書くラベル。同じ科名でファースト/セカンドの両枠があるときだけ
    （ファースト）/（セカンド以下）を付けて区別する（新テンプレは1行=1枠×1日なので、
    ラベルが同じだと行の見分けも取込の枠解決もできない）。"""
    dup = any(x is not f and namekey(x["dept"]) == namekey(f["dept"])
              and x["role"] != f["role"] for x in frames_own)
    return f["dept"] + (ROLE_LABEL.get(f["role"], "") if dup else "")


# ---------------------------------------------------------------- 提出の検証・取込

def make_record(person, d: dt.date, band, frame, oncall, name_disp=None):
    return {
        "schema": "hospital.duty/1",
        "staff_login": person["login"],
        "staff_code": person["code"],
        "name": person["name"],
        "name_disp": name_disp if name_disp is not None else namekey(person["name"]),
        "date": d.isoformat(),
        "band": band,
        # 互換出力のA列は現物の当直表A列の記載そのまま（末尾空白等も保持=全行一致のため）
        "frame_dept": frame.get("dept_dtx14", frame["dept"]),
        "frame_id": frame["frame_id"],
        "role": frame["role"],
        "oncall": bool(oncall),
        "note": "",
    }


_TRUE = {"1", "true", "TRUE", "○", "呼", "オンコール", "on", "yes", "はい"}

# 新テンプレ（2026-08-30ユーザーフィードバック）の複数人列: 当直1〜3 / オンコール1〜6
_MULTI_DUTY = re.compile(r"^当直\s*(\d+)$")
_MULTI_ONCALL = re.compile(r"^オンコール\s*(\d+)$")
_BASE_KEYS = ("日付", "曜日", "枠", "昼夜", "備考")


def expand_person_rows(rows):
    """新レイアウト（当直1〜/オンコール1〜の複数人列）の行を、1人=1行の旧形
    （氏名/オンコール列）へ展開する。旧レイアウトの行はそのまま通す（後方互換）。
    空欄の人列は「その人数は居ない」の意味で正常＝何も出さない。"""
    out = []
    for r in rows:
        duty_keys = sorted((k for k in r if isinstance(k, str) and _MULTI_DUTY.match(k)),
                           key=lambda k: int(_MULTI_DUTY.match(k).group(1)))
        on_keys = sorted((k for k in r if isinstance(k, str) and _MULTI_ONCALL.match(k)),
                         key=lambda k: int(_MULTI_ONCALL.match(k).group(1)))
        if not duty_keys and not on_keys:
            out.append(r)          # 旧レイアウト（氏名/オンコール1列ずつ）
            continue
        base = {k: r[k] for k in _BASE_KEYS if k in r}
        base["_line"] = r.get("_line")
        for keys, mark in ((duty_keys, ""), (on_keys, "○")):
            for k in keys:
                v = str(r.get(k) or "").strip()
                if v:
                    out.append({**base, "氏名": v, "オンコール": mark})
    return out


def parse_date_any(v, year, month):
    """テンプレ/CSVの日付セル: ISO・yyyy/mm/dd・日番号・datetime を受ける。"""
    if isinstance(v, (dt.datetime, dt.date)):
        d = v.date() if isinstance(v, dt.datetime) else v
    else:
        s = str(v or "").strip()
        if not s:
            return None
        s = s.replace("年", "/").replace("月", "/").replace("日", "")
        m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
        if m:
            d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        elif re.match(r"^\d{1,2}[-/]\d{1,2}$", s):
            mm, dd = re.split(r"[-/]", s)
            d = dt.date(year, int(mm), int(dd))
        elif s.isdigit() and 1 <= int(s) <= 31:
            d = dt.date(year, month, int(s))
        else:
            return None
    if d.year != year or d.month != month:
        return "out_of_month"
    return d


def import_rows(rows, owner, year, month, frames, holidays,
                people, exact, canon_map,
                roster_label="職員名簿", not_found_hint="（名簿にありません）", warnings=None):
    """テンプレ.xlsx/CSVの行 → (records, rejects)。
    rows = [{'日付':…, '枠':…, '昼夜':…, '氏名':…, 'オンコール':…, '_line': n}]
    検証: 枠の実在(owner配下)・日付が対象月・氏名が名簿と一致(NFC/異体字込み)・
          バンドがshapeと整合・昼が立たない日の昼・capacity超過・同一人同枠重複。黙って捨てない。
    warnings: リストを渡すと、capacity 超過の「当」を弾かずに残し、ここへ
              {line, date, band, frame_id, name, why} を積む（当直表xlsx の取込。20260926:
              連名の当直もあるので取込を止めない。アプリが取り込んだあとに『オ』へ切り替えてもらう）。
              None（既定）は従来どおり、その枠日の当直を落として rejects に積む（集約など）。"""
    own = [f for f in frames if f.get("owner", f["dept"]) == owner or owner == "管理者"]
    by_id = {f["frame_id"]: f for f in own}

    def resolve_frame(fkey, band_hint):
        """枠列の値 → (frame|None, err|None)。frame_id / 科名 / 科名＋（ファースト）等を受け、
        同名複数は昼夜列（空欄=全日）で絞る。"""
        if fkey in by_id:
            return by_id[fkey], None
        k = namekey(fkey)
        cand = [f for f in own
                if namekey(f["dept"]) == k or namekey(frame_label(f, own)) == k]
        if not cand:
            if not fkey and len(own) == 1:
                return own[0], None
            return None, f"枠『{fkey}』がこの科の枠にありません"
        if len(cand) == 1:
            return cand[0], None
        c2 = ([f for f in cand if f["shape"] == band_hint] if band_hint
              else [f for f in cand if f["shape"] == "全日"])
        if band_hint and not c2:
            c2 = [f for f in cand if f["shape"] == "全日"]
        if len(c2) == 1:
            return c2[0], None
        return None, (f"枠『{fkey}』が複数あります"
                      "（枠名に（ファースト）/（セカンド以下）を付け、昼/夜の列も確認してください）")

    records, rejects = [], []
    rec_lines = []            # records と同じ並びの元の行番号（重複・定員超過の理由に行を付けるため）
    oncall_only = []          # 昼を置かない日（daytime_oncall_only）。要るときだけ1回読む

    def rej(r, why):
        rejects.append({"line": r.get("_line"), "row": {k: str(v) for k, v in r.items() if k != "_line"}, "why": why})

    def n_same(nm):
        """同姓同名の人数（resolve_name が ambiguous を返したときの説明用）。"""
        k = namekey(nm)
        for hits in (exact.get(k), exact.get(nfc(k)), canon_map.get(canon(nfc(k)))):
            if hits and len(hits) > 1:
                return len(set(hits))
        return 2

    for r in rows:
        name = str(r.get("氏名") or "").strip()
        datev = r.get("日付")
        if not name and (datev in (None, "") or str(datev).strip() == ""):
            continue  # 完全空行
        if not name:
            continue  # 名前が空の行はテンプレの未記入行（正常）
        d = parse_date_any(datev, year, month)
        if d is None:
            rej(r, "日付が読めません")
            continue
        if d == "out_of_month":
            rej(r, f"対象月({year}/{month})の日付ではありません")
            continue
        # 枠の解決: frame_id → 枠名(＋ファースト/セカンド以下) → owner配下が1枠だけなら省略可
        fkey = str(r.get("枠") or "").strip()
        frame, err = resolve_frame(fkey, str(r.get("昼夜") or "").strip())
        if frame is None:
            rej(r, err)
            continue
        # 氏名の解決（自由入力の穴に戻らない: 名簿と一致した人だけ受理）
        person, how = resolve_name(name, people, exact, canon_map)
        if person is None:
            if how == "ambiguous":        # 20260926: 事務の方に分かる言い方で（どうすればよいかまで）
                rej(r, f"氏名『{name}』は{roster_label}に同姓同名が{n_same(name)}人います。ドロップダウンから選んでください")
            else:
                rej(r, f"氏名『{name}』が{roster_label}と一致しません" + not_found_hint)
            continue
        # バンド（昼/夜のみ。旧形式の「昼夜」「全日」は空欄と同じ=枠の既定へ自動展開）
        band = str(r.get("昼夜") or "").strip()
        if band in ("昼夜", "全日"):
            band = ""
        if band and band not in ("昼", "夜"):
            rej(r, f"昼夜『{band}』は「昼」か「夜」で書いてください（空欄=枠の既定）")
            continue
        oncall = str(r.get("オンコール") or "").strip() in _TRUE
        if band:
            bands = [band]
            if frame["shape"] != "全日" and band != frame["shape"]:
                rej(r, f"この枠は「{frame['shape']}」の枠です（{band}は入れられません）")
                continue
            if frame["shape"] == "全日" and band == "昼" and not is_holiday(d, holidays):
                rej(r, "全日枠の昼は土日祝だけです（平日は夜のみ）")
                continue
        else:
            bands = expand_bands(frame["shape"], d, holidays)
        # 昼が立たない日の昼（平日の昼枠・管理当直の開院記念日の昼。20260926）はどの経路でも弾く
        if "昼" in bands:
            if not oncall_only:
                oncall_only.append(daytime_oncall_only())
            if not band_stands(frame, "昼", d, holidays, oncall_only[0]):
                rej(r, f"{md_label(d)} 昼: この日はこの枠が立ちません（灰色のマス）。『{name}』を消してください")
                continue
        for b in bands:
            records.append(make_record(person, d, b, frame, oncall))
            rec_lines.append(r.get("_line"))

    # 重複（同一人・同日・同バンド・同枠）。1件だけ採用するが、書き間違いなので弾きとして知らせる
    seen = Counter((x["staff_login"], x["date"], x["band"], x["frame_id"]) for x in records)
    dup = {k for k, c in seen.items() if c > 1}
    if dup:
        uniq, uniq_lines, dropped, told = [], [], set(), set()
        for x, ln in zip(records, rec_lines):
            k = (x["staff_login"], x["date"], x["band"], x["frame_id"])
            if k in dup and k in dropped:
                t = (ln, x["staff_login"], x["date"])       # 昼夜1レーンの昼と夜で同じ文を2回出さない
                if t not in told:
                    told.add(t)
                    d0 = dt.date.fromisoformat(x["date"])
                    rejects.append({"line": ln, "row": {"氏名": x["name"], "日付": x["date"], "枠": x["frame_dept"], "昼夜": x["band"]},
                                    "why": f"{md_label(d0)} {x['band']}: 同じ人（{namekey_display(x['name'])}）が"
                                           "同じ枠に2回書かれています。どちらか1つを消してください"})
                continue
            dropped.add(k)
            uniq.append(x)
            uniq_lines.append(ln)
        records, rec_lines = uniq, uniq_lines

    # capacity 超過（当直=泊まる人だけ数える。オンコールは複数人が正常なので上限を掛けない）
    groups = defaultdict(list)
    for x, ln in zip(records, rec_lines):
        if not x["oncall"]:
            groups[(x["frame_id"], x["date"], x["band"])].append((x, ln))
    over = set()
    fmap = {f["frame_id"]: f for f in frames}
    for (fid, d0, b0), lst in groups.items():
        cap = fmap[fid].get("capacity", 1)
        c = len(lst)
        if c <= cap:
            continue
        if warnings is not None:          # 当直表xlsx の取込: 残して知らせる（20260926）
            for x, ln in lst[cap:]:
                warnings.append({"line": ln, "date": d0, "band": b0, "frame_id": fid,
                                 "name": namekey_display(x["name"]),
                                 "why": f"{md_label(dt.date.fromisoformat(d0))} {b0}: 当直が定員（{cap}名）を超えています（{c}名）。"
                                        "連名でなければ取り込んだあとに『オ』に切り替えてください"})
            continue
        over.add((fid, d0, b0))
        rejects.append({"line": None, "row": {"枠": fmap[fid]["dept"], "日付": d0, "昼夜": b0},
                        "why": f"1枠の当直人数({cap}名)を超えています（{c}名）。全員残す場合は枠マスタのcapacityを直してください"})
    records = [x for x in records if x["oncall"] or (x["frame_id"], x["date"], x["band"]) not in over]
    return records, rejects


def read_import_file(path, year, month):
    """テンプレ.xlsx か CSV(UTF-8 BOM/CP932) → rows（import_rowsの入力形）。"""
    if path.lower().endswith(".xlsx"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["当直入力"] if "当直入力" in wb.sheetnames else wb[wb.sheetnames[0]]
        rows = []
        header = None
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            vals = ["" if v is None else v for v in row]
            if header is None:
                if any("日付" in str(v) for v in vals):
                    header = [str(v).strip() for v in vals]
                continue
            rows.append({**{header[j]: vals[j] for j in range(min(len(header), len(vals)))}, "_line": i})
        wb.close()
        return expand_person_rows(rows)
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    rows = []
    rd = csv.DictReader(io.StringIO(text))
    for i, r in enumerate(rd, start=2):
        r = {(k or "").strip(): (v or "") for k, v in r.items()}
        r["_line"] = i
        rows.append(r)
    return expand_person_rows(rows)


# ---------------------------------------------------------------- 提出JSON

def load_submission_json(path):
    with open(path, encoding="utf-8-sig") as f:
        sub = json.load(f)
    if sub.get("schema") != "hospital.duty-submission/1":
        raise ValueError(f"schemaが違います: {sub.get('schema')}")
    return sub


def validate_submission(sub, frames, holidays, people, exact, canon_map):
    """提出JSONの原子レコードも取込と同じ検証を通す（経路によらず必ず通す）。
    戻り: (records, problems)。"""
    fmap = {f["frame_id"]: f for f in frames}
    records, problems = [], []
    for r in sub.get("duties", []):
        fid = r.get("frame_id")
        f = fmap.get(fid)
        if f is None:
            # frame_id無し（旧形式）→ dept+role+行で当てる
            cand = [x for x in frames
                    if namekey(x["dept"]) == namekey(r.get("frame_dept"))
                    and x["role"] == r.get("role")]
            f = cand[0] if len(cand) == 1 else None
        if f is None:
            problems.append({"row": r, "why": f"枠が特定できません: {fid}/{r.get('frame_dept')}"})
            continue
        try:
            d = dt.date.fromisoformat(r["date"])
        except Exception:
            problems.append({"row": r, "why": "日付が不正です"})
            continue
        if (d.year, d.month) != (sub["year"], sub["month"]):
            problems.append({"row": r, "why": "対象月外の日付です"})
            continue
        if r.get("band") not in ("昼", "夜"):
            problems.append({"row": r, "why": f"band不正: {r.get('band')}"})
            continue
        login = (r.get("staff_login") or "").strip()
        person = None
        if login:
            hits = [p for p in people if p["login"] == login]
            person = hits[0] if hits else None
        if person is None:
            person, how = resolve_name(r.get("name") or "", people, exact, canon_map)
        provisional = False
        if person is None and (login.startswith("unreg:") or r.get("unregistered")):
            # Dr.JOY未登録の手入力医師（20260902ユーザー決定）。氏名で名簿に結べなければ
            # 暫定の人物として通す（互換出力は氏名で書くので下流は動く。集約レポートに一覧を出す）
            nm = namekey_display(r.get("name") or "")
            if not nm:
                problems.append({"row": r, "why": "名簿外医師の氏名が空です"})
                continue
            person = {"login": login or "unreg:" + canon(nfc(nm)), "code": "", "name": nm}
            provisional = True
        if person is None:
            # 20260926: 事務の方に分かる言い方で（ログインIDは出さない）
            problems.append({"row": r, "why": f"Dr.JOY の医師一覧にいない人です（退職・ロック・仮登録の解除など）: "
                                             f"{namekey_display(r.get('name') or '') or login}"})
            continue
        rec = make_record(person, d, r["band"], f, r.get("oncall"), r.get("name_disp"))
        rec["note"] = r.get("note", "")
        if provisional:
            rec["unregistered"] = True
        records.append(rec)
    return records, problems


def namekey_display(s):
    """手入力氏名の表示形: 全角空白→半角、連続空白→1つ、前後の空白を落とす（「姓 名」を保つ）。"""
    return re.sub(r"\s+", " ", str(s or "").replace("\u3000", " ")).strip()


# ---------------------------------------------------------------- 互換3シート出力

def rebuild_rows(records, frames):
    """原子レコード → 人別3シート相当 {sheet: [(A,B,C,D,E)]}。
    B=当直表セルの記載(name_disp・オンコールは括弧つき) / E=フルネーム(空白なし)。"""
    fmap = {f["frame_id"]: f for f in frames}
    out = {P_DUTY: [], P_ONCALL: [], P_FIRST: []}
    for r in records:
        e = namekey(r["name"]) or r.get("name_dtx14", "")
        disp = r.get("name_disp") or namekey(r["name"])
        b = f"({disp})" if r["oncall"] else disp
        row = (r["frame_dept"], b, r["date"], r["band"], e)
        out[P_ONCALL if r["oncall"] else P_DUTY].append(row)
        f = fmap.get(r.get("frame_id"))
        first_ok = (f["first_report"] if f else
                    (r.get("role") == "first" and "放射線科" not in r["frame_dept"]
                     and "管理当直" not in r["frame_dept"]))
        if first_ok:
            out[P_FIRST].append(row)
    return out


HEADER = ["科名", "氏名", "当直日", "勤務形態", "フルネーム"]


def write_compat_xlsx(records, frames, path, year, month):
    """DTX14互換の人別3シート入りxlsxを書く（下流仕様: A科名/B氏名/C日付/D昼夜/Eフルネーム）。
    C列はdatetimeで書く（DTX_DrJOY/interval_v2が日付として読むため）。"""
    import openpyxl
    rows = rebuild_rows(records, frames)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet in (P_DUTY, P_ONCALL, P_FIRST):
        ws = wb.create_sheet(sheet)
        ws.append(HEADER)
        data = sorted(rows[sheet], key=lambda t: (t[2], t[0], t[3], t[1]))
        for a, b, c, d, e in data:
            ws.append([a, b, dt.datetime.fromisoformat(c), d, e])
        for cell in ws["C"]:
            cell.number_format = "yyyy/mm/dd"
    ws = wb.create_sheet("生成情報")
    ws.append(["生成", dt.datetime.now().isoformat(timespec="seconds")])
    ws.append(["対象月", f"{year}/{month}"])
    ws.append(["レコード数", len(records)])
    ws.append(["注意", "このブックはTouchoku_Planの集約器が生成した互換出力です。DTX14本体ではありません。"])
    wb.save(path)
    return {k: len(v) for k, v in rows.items()}
