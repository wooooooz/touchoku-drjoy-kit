#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
時間外勤務統計 (DTX_18) Python版 — GUI + ロジック単一ファイル
20260605（当直まるごと加算）/ 20260605（管理当直控除・救急小児判定を前方一致へ修正）
/ 20260605（Excel出力廃止＝配布pptxのみ生成・build_deck分割／詳細ログ）
/ 20260605（D=救急小児の早上がり控除[当直加算を上限・データ欠落期間除外・翌朝退勤1900補正]／給与連携dimension対応reset_dimensions／監査base_dd修正）
/ 20260606（B=勤務間インターバル不足[専攻医/初期研・通常9h/救急小児当直後 専攻医18h・初期研24h・自己研鑽等は休息]）
/ 20260607（C=当直回数・オンコール回数・連続勤務日数・休日数[勤務日=出勤打刻 or 当直/休日=暦日−勤務日]）
/ 20260608（残業の平日/休日区分を日付判定化＝空列依存で休日時間外が全て平日になる不具合を修正）
/ 20260608（連続勤務/休日を申告ベース化＝勤務日を「時間外/自己研鑽/当直の申告ある日」に変更しビーコン汚染を排除／ビーコン置きっぱなし疑いを scan_beacon[置きっぱ日数＝暦日カバレッジ≥22h(前日跨ぎ合算)且つ申告なしが4日+・全員]で別途検出）
/ 20260609（給与連携から「出退勤あり・時間外/研鑽の申告ゼロ」医師を scan_no_claim で抽出し科別・職位順スライドを追加）
/ 20260623（ビーコン常時受信期に対応＝置きっ放しを scan_beacon_v2[給与連携直読み・Dr.JOY退勤自動クローズ印(翌始業締め32:30等/予定未入力は翌23:59締め47:59)をカバー暦日数で集計]に刷新。旧 scan_beacon[22hカバレッジ]は常時受信下で失明のため置換）
/ 20260623（scan_beacon_v2: 退勤打刻種別＝ビーコンの日のみ置きっ放しと計上。アプリ/WEB/打刻機/編集の手入力退勤は病院公認として除外＝手入力で出退勤を記録する医師をランキングから外す）
/ 20260623（scan_beacon_v2: 置きっ放しを24:00判定に刷新＝各日の24時に置きっ放し状態(自動クローズが深夜をまたぐ)なら+1・連日は分割記録をmerge橋渡し。ビーコン稼動日数(正常退勤)を追加。非常勤医師は対象外。スライドは置きっぱ/ビーコン稼動/手入力の3日数列）
/ 20260901（時間外スロットの長さを cell_datetime 差で算出＝24時間以上の申告がちょうど24h短く計上される不具合を修正。
  旧実装は try_time_hours[0〜24h未満]の差を使っており、例 07:22→翌08:30(25:08) が 1:08 になっていた。
  あわせて早上がり控除(E)＝代償休息と見なす仕様・管理当直窓clip(C)＝実装は意図どおりである旨をコメントで明記）
/ 20261016（働き方デッキをDr.JOY公式値にそろえる＝管理者決定。load_work_record_ot/apply_drjoy_overtime＝Dr.JOY「勤務実績」exportの残業時間(25%)+法定休日(35%)で時間外(p.total/per_dept_hours)を上書き。load_drjoy_daisho＝Dr.JOY「勤務予定」exportの代償休息「全期間必要時間」(残高)を読む。どちらもmake_pptx_boxplotが呼ぶ）
/ 20261017（管理者決定: 勤務実績CSVが無い/月違いならInputErrorで停止＝従来値に黙って戻らない[開発用 HATARAKI_ALLOW_LEGACY=1]。代償休息は(M+1)月度の勤務予定exportの「月末まで必要時間」＝その月に発生し翌月末までに付与が必要な1ヶ月分を読む[残高は参考]。人別に total_old/drjoy_ot を残す）
/ 20261107（apply_drjoy_overtime: 人別シートに居ないがDr.JOYに時間外がある医師[締めより後の承認]を足す＝管理者GO）
/ 20261107（勤務実績の「出勤日数」を p.drjoy_days に残す＝P2の休日数用）
/ 20261104（apply_drjoy_overtime: Dr.JOYの時間外を申請ベースの科別時間の比で各科に割り振る＝移動者が「A→B」の別グループにならず、経た全科の箱ひげに入る。aggregate_dept_fte: 科別の重み付き箱ひげ用に wsamples＝(1か月換算h, 在籍割合, 在籍日数, 実時間) を足す＝管理者決定）

VBA版 (vba/DTX_18_Module1_current.bas 0.3) からの移植。担当者には PyInstaller --onefile で単体exe配布。
- 入力: 月フォルダ内の `*_当直情報*.xlsm`（人別シート）と `*給与連携項目一覧*.xlsx`（日別所属）
- 出力: 同フォルダに `時間外勤務統計_R{N}{M}月.xlsx`（自動オープン）
- 集計: 科別=日割りFTE平均、個人別=科遷移ラベル、職位正規化、箱ひげ図（Stacked-Bar trick）
- Excelを起動しない（openpyxlのみ）→ 担当者の他作業に干渉なし

検証メモ(20260604): VBA 0.3 と個人別合計が完全一致(不一致0/292名)。要点2つ —
  ① 厳密整数“分”で累積し min_to_h1() で 0.1h に銀行丸め（VBA Round準拠／FP境界ぶれ排除）。
     round(浮動小数,1) だと 43.85→43.9 等で VBA とズレるため。
  ② 実績ゼロ(時間外0h)の人は全統計から除外（ユーザー指示20260604）。VBA0.3は0残業者も掲載
     していたが、Python版は意図的に除外＝対象人数がVBAより少ない(例 292→244)。
  ※ 遷移ラベルの末尾「→初期研修医(仮)」は仕様通り除外（本Py=正。本番VBA0.3出力は
     未再実行のため(仮)残存＝古い。VBAソースの JoinDictKeysSkipKari は本Pyと同じ意図）。
"""
from __future__ import annotations

import os
import sys
import re
import glob
import queue
import threading
import traceback
import statistics
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from typing import Optional

try:
    import jpholiday as _jph     # 日本の国民の祝日判定（救急/小児の当直まるごと加算で使用）
except Exception:
    _jph = None

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import ColorChoice
from openpyxl.drawing.colors import ColorChoice as DrawColorChoice  # alias


def _ws_full(ws):
    """Dr.JOY直出力のxlsxは <dimension> が壊れていることがあり、read_onlyだと A1:A1 と
    誤読して1列しか読めない。reset_dimensions() で実データ全体を読めるようにする（速度は維持）。
    通常のシートでは無害。失敗は無視。20260605"""
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    return ws


# ============================================================
# 定数
# ============================================================
HDR_ROW = 4
DATA_START = 5
MAX_SLOTS = 10
EXCLUDE_DEPTS = {"初期研修医(仮)", "他院", "所属未分類"}
BOX_EXCLUDE_DEPTS = {"初期研修医(仮)", "所属未分類"}
KARI = "初期研修医(仮)"
POSITION_ORDER = ["医師", "大学院生", "専門研修医", "初期研修医"]

# ビーコン置きっぱなし疑いの判定（20260608）: 各暦日が在席区間（出勤〜退勤・前日からの跨ぎ含む）で
# BEACON_COVERAGE_H 時間以上"覆われている"のに、その日に時間外/自己研鑽の申告が無い日＝「置きっぱ日数」。
# 終日いた事になっているのに勤務を一切申告していない＝ビーコン放置の矛盾。置きっぱ日数が BEACON_MIN_DAYS
# 日以上で疑い。※1行ごとの在席時間ではなくカバレッジで見る理由＝ビーコンの連続在席は深夜0時で
# 「退勤23:59/出勤00:00」に分割記録され、その日が15h等に見えて漏れるため（前日跨ぎを合算して補正）。
# 本物の当直は1日だけ(翌日へ続かない)＝カバレッジ≈15hで非該当。1〜2日は申告漏れの可能性→閾値で除外。
BEACON_COVERAGE_H = 22.0
BEACON_MIN_DAYS = 4

# ビーコン置きっぱなし v2（常時受信期・自動クローズ印方式）20260623
# ビーコン受信機が院内に行き渡り検知が途切れなくなった結果、置きっ放しは「連日24h在席」
# ではなく Dr.JOY の退勤自動クローズ固定値に化け、旧 scan_beacon(22hカバレッジ)は失明(検出0)した。
# v2 は Dr.JOY が付ける自動クローズ印を直接読む（始業時刻に非依存の構造判定。詳細は scan_beacon_v2）。
BEACONV2_MIN_DAYS = 3                      # 置きっ放しカバー暦日がこれ以上で「疑い」（注意1-2/警告3-7/重度8+）
BEACONV2_HANDOFF_TOL = 60                  # 退勤の翌日時刻と翌日出勤打刻の一致許容（秒）。自動クローズは同一の予定始業値で一致
BEACONV2_EOD_MIN = 47 * 3600 + 58 * 60    # ②予定未入力印の下限＝翌23:58（47:58/47:59）。1回のクローズは翌23:59が上限
# 打刻種別＝ビーコン の退勤だけを置きっ放しとして数える。本人がアプリ/WEB/打刻機で手入力 or 編集した退勤は
# 病院公認の記録なので除外（ビーコン置きっぱでも自分で出退勤を入力している医師がいる）。20260623
BEACON_PUNCH_TYPE = "ビーコン"
# 非常勤医師は勤務形態が異なるため置きっ放し表から除外（職名に この語を含む人を除外）。20260623
BEACON_EXCLUDE_JOB = "非常勤"

# 色（VBA版と同じ）
CLR_HEADER_BG = "2B4C7E"
CLR_HEADER_FG = "FFFFFF"
CLR_HOT_BG = "FFEBEB"
CLR_HOT_FG = "C0504D"
CLR_ZEBRA = "F5F8FC"
CLR_BORDER = "C0C0C0"
CLR_PRIMARY = "2B4C7E"
CLR_WARN_HIGH = "FFE0E0"
CLR_WARN_MID = "FFF4E0"

# ============================================================
# データクラス
# ============================================================
@dataclass
class PersonRow:
    dept_label: str             # 個人別表示用（遷移ラベル、(仮)除く）
    fallback_dept: str          # 元A1（仮含む）
    position: str               # 正規化後職位
    name: str                   # 氏名（元のまま、表示用）
    name_norm: str              # 正規化氏名（マッチ用）
    total: float
    teate: float                # 診療手当あり
    other: float                # その他
    weekday_ot: float
    holiday_ot: float
    duty: float = 0.0           # 当直分（救急/小児のみ。当直時間帯の研鑽を時間外化）
    per_dept_hours: dict[str, float] = field(default_factory=dict)
    iv_target: bool = False     # 勤務間インターバル判定の対象(専攻医/初期研)か
    iv_viol: int = 0            # インターバル不足の回数
    iv_short: float = 0.0       # 不足時間合計(h)
    iv_nopunch: int = 0         # 打刻なしで勤務がある日数(判定の信頼性低下=要確認)
    # C: 当直/オンコール回数・連続勤務日数・休日数（20260607）
    n_duty: int = 0             # 当直回数（当直列が非空の日数）
    n_oncall: int = 0           # オンコール回数（オンコール列が非空の日数）
    n_workday: int = 0          # 勤務日数（時間外 or 自己研鑽 or 当直の申告がある日。20260608 申告ベース）
    n_off: int = 0              # 休日数（暦日 − 勤務日数）
    max_streak: int = 0         # 連続勤務日数（勤務日が途切れず続いた最大日数）

@dataclass
class BeaconRow:
    """ビーコン置きっぱなし疑い。全人別シートを対象（申告ゼロも含む）。
    判定＝在席(退勤−出勤)≥BEACON_PRESENCE_H なのに その日に時間外/自己研鑽の申告が無い日（置きっぱ日数）。"""
    name: str
    dept: str
    position: str
    beacon_days: int    # 置きっぱ日数（カバレッジ≥BEACON_COVERAGE_H 且つ その日の申告が無い日数）
    n_placed: int       # 終日在席（カバレッジ≥BEACON_COVERAGE_H）の総日数（参考）
    n_apply: int        # 時間外/自己研鑽の申告がある日数（少ない＝在席だけで悪質）

@dataclass
class BeaconRowV2:
    """ビーコン置きっぱなし疑い（常時受信期・自動クローズ印方式）。20260623。
    給与連携を直読みし、Dr.JOY が退勤を自動クローズした印（翌日始業 or 翌23:59）からカウント。"""
    name: str
    dept: str
    position: str
    leave_days: int     # 置きっ放し日数＝各日の24:00時点で置きっ放し（ビーコン自動クローズが深夜をまたぐ）の日数
    marker_rows: int    # ビーコン自動クローズ行数（参考）
    kind: str           # 'sched'（予定有・翌始業締め） / 'mixed' / 'no-sched'（予定未入力・翌23:59締めのみ）
    n_apply: int        # 時間外/自己研鑽の申告がある日数（少ない＝在席だけで悪質）
    beacon_days: int = 0  # ビーコン稼動日数＝正常にビーコンで退勤を記録した日数（自動クローズでない）
    manual_days: int = 0  # 退勤を手入力/編集で記録した日数（種別≠ビーコン・月内。病院公認＝置きっぱ非該当）

@dataclass
class DeptStat:
    name: str
    hc: int          # 実人数（その科で残業あり）
    fte: float       # 換算人数 Σ(在籍日/月暦日)
    total: float     # その科でやった残業の総時間
    avg: float       # total / fte
    samples: list[float] = field(default_factory=list)  # 箱ひげ用（個人の該当科時間）
    # 20261104 重み付き箱ひげ用: (1か月換算h, 重み=その月の在籍日のうちこの科にいた割合, この科の在籍日数, この科での実時間h)
    wsamples: list[tuple] = field(default_factory=list)

@dataclass
class PosStat:
    name: str
    n: int
    avg: float
    median: float
    max: float
    min: float
    over80: int

@dataclass
class WarnEntry:
    severity: str   # "高"/"中"/"低"
    category: str
    proc: str
    key: str
    dept: str = ""
    item: str = ""
    reason: str = ""
    sheet: str = ""
    row: int = 0


# ============================================================
# テキスト・日付ユーティリティ
# ============================================================
def norm_name(s) -> str:
    if s is None:
        return ""
    t = str(s)
    t = t.replace("　", "").replace(" ", "")
    t = t.replace("髙", "高").replace("﨑", "崎")
    return t.strip()

def first_seg(s) -> str:
    """/循環器内科/集中治療部/ -> 循環器内科。スラッシュ無しはそのまま"""
    if s is None:
        return ""
    t = str(s)
    if t.startswith("/"): t = t[1:]
    if t.endswith("/"): t = t[:-1]
    parts = t.split("/")
    return parts[0] if parts else ""

_DATE_RE = re.compile(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})")

def date_ord(v) -> int:
    """日付っぽい値 -> 日単位ordinal(int)。失敗0
    対応: date/datetime, Excelシリアル(数値), '2026/5/1', '2026-05-01', '2026年05月01日' 等"""
    if v is None:
        return 0
    try:
        if hasattr(v, "toordinal"):
            return int(v.toordinal())
        if isinstance(v, (int, float)):
            # Excelシリアル(1900基準): 25569 = 1970-01-01。ここでは年月日文字列が主なので数値は稀
            return int(v)
        s = str(v).strip()
        if not s:
            return 0
        m = _DATE_RE.search(s)
        if m:
            import datetime as _dt
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return _dt.date(y, mo, d).toordinal()
    except Exception:
        pass
    return 0

def try_time_hours(v) -> Optional[float]:
    """時刻(Variant) -> 0以上24未満の時間(時)。失敗None"""
    if v is None:
        return None
    try:
        if hasattr(v, "hour"):
            return v.hour + v.minute / 60.0 + v.second / 3600.0
        if isinstance(v, (int, float)):
            f = float(v)
            return (f - int(f)) * 24.0
        s = str(v).strip()
        if not s:
            return None
        # "9:30" 等
        if ":" in s:
            parts = s.split(":")
            h = int(parts[0]); m = int(parts[1]) if len(parts) > 1 else 0
            ss = int(parts[2]) if len(parts) > 2 else 0
            return h + m / 60.0 + ss / 3600.0
        f = float(s)
        return (f - int(f)) * 24.0
    except Exception:
        return None

def min_to_h1(total_min) -> float:
    """整数分 -> 0.1h単位の時間。半端(.X5=3分境界)は銀行丸め(half-to-even)でVBA Round準拠。
    Dr.JOY時刻は分単位なので分は厳密整数。round(浮動小数)のFP境界ぶれを排除する。"""
    m = int(round(total_min))
    q, r = divmod(m, 6)            # 0.1h = 6分。r=余り分(0..5)
    if r < 3:
        t = q
    elif r > 3:
        t = q + 1
    else:                          # r==3: ちょうど半分 → 偶数側へ
        t = q if (q % 2 == 0) else q + 1
    return t / 10.0

# ============================================================
# 当直時間帯ヘルパー（救急/小児の「当直まるごと加算」用。手当計算の Excel Module4 と同一定義）
# ============================================================
DUTY_TYPES = ("昼", "夜", "休", "昼夜")

def cell_datetime(v, base_date):
    """セル値を datetime に。datetime はそのまま／time は base_date と合成／数値・文字列は時刻化して合成。失敗 None。"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, time):
        return datetime.combine(base_date, v)
    th = try_time_hours(v)
    if th is None:
        return None
    h = int(th); mi = int(round((th - h) * 60))
    if mi >= 60:
        h += 1; mi -= 60
    return datetime(base_date.year, base_date.month, base_date.day) + timedelta(hours=h, minutes=mi)


def time_on_date(v, base_date):
    """出勤/退勤などの『時刻』を base_date に合成した datetime に。空・失敗は None。
    重要(20260605 C1対策): Dr.JOYは『翌朝の退勤』を Excel時刻シリアル≥1日 で持ち、openpyxlは
    datetime(1900,1,1+n,H,M) で読む（=翌n日のH:M）。単に時刻成分だけ採るとke<=ks補正も効かず、
    24時間勤務日(例 出勤8:00→翌8:10)を『日中不在』と誤判定して誤控除する。よって 1900系の
    datetime は日数オフセットを base_date に足す。実日付つき(年>1900)はそのまま時刻採用。"""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    if isinstance(v, datetime):
        if v.year <= 1900:
            off = (v.date() - date(1899, 12, 31)).days   # 1900-01-01→+1日, 1900-01-02→+2日…
            if off < 0:
                off = 0
            return datetime.combine(base_date + timedelta(days=off), v.time())
        return datetime.combine(base_date, v.time())
    if isinstance(v, time):
        return datetime.combine(base_date, v)
    th = try_time_hours(v)
    if th is None:
        return None
    h = int(th); mi = int(round((th - h) * 60))
    if mi >= 60:
        h += 1; mi -= 60
    return datetime(base_date.year, base_date.month, base_date.day) + timedelta(hours=h, minutes=mi)

def cell_date(v):
    """セル値を date に。失敗 None。"""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    o = date_ord(v)
    return date.fromordinal(o) if o > 0 else None

def is_jp_holiday(d) -> bool:
    """日本の国民の祝日か（土日は含めない）。jpholiday未導入ならFalse。"""
    if _jph is None:
        return False
    try:
        return bool(_jph.is_holiday(d))
    except Exception:
        return False

def is_holiday_date(d, holiday_cell=None) -> bool:
    """休日か＝土日 or 国民の祝日 or 休日列が非空。
    ※ソースの曜日/休日列は当直行で空のことがあるので、日付から土日・祝日を判定する。"""
    if d.weekday() >= 5:           # 土(5)・日(6)
        return True
    if holiday_cell is not None and str(holiday_cell).strip() != "":
        return True
    return is_jp_holiday(d)

def next_is_holiday(next_date, next_row, c_holiday) -> bool:
    """翌日が休日か。Module4 CheckNextDayInArray 準拠＋国民の祝日も考慮。"""
    hcell = None
    if next_row is not None and c_holiday > 0 and c_holiday - 1 < len(next_row):
        hcell = next_row[c_holiday - 1]
    return is_holiday_date(next_date, hcell)

def duty_window(d, duty_type, next_holiday):
    """当直時間帯 (start_dt, end_dt) を返す。d は date。Module4 ProcessDutyInArray 準拠。"""
    base = datetime(d.year, d.month, d.day)
    nxt = base + timedelta(days=1)
    if duty_type == "昼":
        return base + timedelta(hours=9), base + timedelta(hours=17, minutes=15)
    if duty_type in ("夜", "休"):
        s = base + timedelta(hours=17, minutes=15)
        e = (nxt + timedelta(hours=9)) if next_holiday else (nxt + timedelta(hours=8, minutes=30))
        return s, e
    if duty_type == "昼夜":
        s = base + timedelta(hours=9)
        e = (nxt + timedelta(hours=9)) if next_holiday else (nxt + timedelta(hours=8, minutes=30))
        return s, e
    return None, None

def overlap_minutes(s_dt, e_dt, win_s, win_e) -> int:
    """[s_dt,e_dt] と [win_s,win_e] の重なり分（整数分）。"""
    lo = max(s_dt, win_s); hi = min(e_dt, win_e)
    if hi <= lo:
        return 0
    return int(round((hi - lo).total_seconds() / 60.0))


def is_er_peds(dept) -> bool:
    """救急部/小児科か（当直まるごと加算＝宿日直許可なしの対象）。
    前方一致で判定。部分一致だと「外科(肝胆膵・乳腺)」等が "小児" を含んで誤ヒットするため。"""
    d = str(dept).strip() if dept is not None else ""
    return d.startswith("救急") or d.startswith("小児")


# ============================================================
# 勤務間インターバル用 区間演算（20260606）
# ============================================================
def _merge_intervals(ivs, tol_min=15):
    """[(start_dt,end_dt)] を時刻順にマージ。tol_min分以内で接する区間は連結。"""
    if not ivs:
        return []
    ivs = sorted(ivs)
    out = [list(ivs[0])]
    for s, e in ivs[1:]:
        if s <= out[-1][1] + timedelta(minutes=tol_min):
            if e > out[-1][1]:
                out[-1][1] = e
        else:
            out.append([s, e])
    return [(s, e) for s, e in out]


def _subtract_intervals(work, rest):
    """work区間から rest区間を引く（休息を労働から除外）。"""
    if not rest:
        return list(work)
    out = []
    for ws, we in work:
        segs = [(ws, we)]
        for rs, re in rest:
            nxt = []
            for s, e in segs:
                if re <= s or rs >= e:
                    nxt.append((s, e))
                else:
                    if s < rs:
                        nxt.append((s, rs))
                    if re < e:
                        nxt.append((re, e))
            segs = nxt
        out.extend(segs)
    return out


def normalize_position(pos) -> str:
    """VBA NormalizePosition と一致"""
    if pos is None:
        return ""
    p = str(pos).strip()
    if p.endswith("専攻医") or p.startswith("専攻医"):
        return "専門研修医"
    if p.startswith("大学院生"):
        return "大学院生"
    if p == "精神保健指定医":
        return "医師"
    if "初期研" in p or "初期研修医" in p:
        return "初期研修医"
    if p == "歯科医師":
        return "医師"
    return p

def position_color(pos: str) -> Optional[str]:
    if pos == "初期研修医": return "008000"
    if pos == "専門研修医": return "0000C0"
    return None

def label_excluding_kari(per_dept_hours: dict[str, float], fallback: str) -> str:
    """挿入順で (仮)を除いて "→" 連結。仮以外がなければ仮を残す。空ならfallback。"""
    if not per_dept_hours:
        return fallback
    non_kari = [d for d in per_dept_hours.keys() if d != KARI]
    if non_kari:
        return "→".join(non_kari)
    return "→".join(per_dept_hours.keys())


# ============================================================
# 入力検出
# ============================================================
def find_latest(folder: str, pattern: str) -> Optional[str]:
    """フォルダ内の最新ファイルを返す（~$ロックファイル除外）"""
    paths = []
    for p in glob.glob(os.path.join(folder, pattern)):
        name = os.path.basename(p)
        if name.startswith("~$"):
            continue
        paths.append(p)
    if not paths:
        return None
    paths.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return paths[0]

def _hhmm_to_h(s) -> float:
    """'265:10' 形式 → 時間(float)。空・不正は 0。"""
    s = str(s).strip() if s is not None else ""
    if ":" in s:
        try:
            h, m = s.split(":")
            return int(h) + int(m) / 60.0
        except Exception:
            return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


class InputError(Exception):
    """Dr.JOY の export が無い・月が違う・中身がおかしいときに投げる。
    20261017 管理者決定: 勤務実績CSVが無い月はデッキ生成を止める（従来値に黙って戻らない）。"""


def _allow_legacy() -> bool:
    """開発・確認用だけ。環境変数 HATARAKI_ALLOW_LEGACY=1 のとき、export が無くても従来値で進む（配布デッキには使わない）。"""
    return os.environ.get("HATARAKI_ALLOW_LEGACY", "") == "1"


def report_ym(month_label: str) -> tuple[int, int]:
    """'R8年9月' → (2026, 9)。令和N年 = 2018+N年。"""
    m = re.search(r"R(\d+)年(\d+)月", str(month_label))
    if not m:
        raise InputError(f"対象月の表記を読めません: {month_label!r}")
    return 2018 + int(m.group(1)), int(m.group(2))


def _next_ym(y: int, m: int) -> tuple[int, int]:
    return (y + 1, 1) if m == 12 else (y, m + 1)


def _pick_export(folder: str, pattern: str, regexes: list, what: str, log=print):
    """folder の pattern に合うファイルのうち、ファイル名が regexes のどれかに合う最新(mtime)を返す。
    月が違うファイルは無視する（find_latest は月を見ないので、月違いを黙って読む事故を防ぐ）。無ければ None。"""
    hits = []
    for pth in glob.glob(os.path.join(folder, pattern)):
        nm = os.path.basename(pth)
        if nm.startswith("~$") or "バックアップ" in nm or nm.startswith("rejected-"):
            continue
        if any(re.search(rx, nm) for rx in regexes):
            hits.append(pth)
    if not hits:
        return None
    hits.sort(key=lambda q: os.path.getmtime(q), reverse=True)
    if len(hits) > 1:
        log(f"      ※{what}が{len(hits)}本あります。最新を使います: {os.path.basename(hits[0])}")
    return hits[0]


def load_work_record_ot(folder: str, month_label=None, log=print, meta=None):
    """Dr.JOY「勤務実績」export(勤務_*月度*.csv・Shift_JIS)を読み、
    人別の時間外労働時間＝残業時間(25%)＋法定休日時間(35%) を返す。
    20261016: 働き方デッキの時間外を Dr.JOY の公式値にそろえる（管理者決定）。
    20261017: 対象月と同じ月のファイル(YYYYMM月度)だけを読む。無ければ InputError（停止）。
    戻り値: (dict{canon(氏名): 時間h}, 読んだファイル名)
    """
    import csv as _csv
    import dtx14_names as _N
    if month_label:
        y, m = report_ym(month_label)
        path = _pick_export(folder, "勤務_*月度*.csv", [rf"{y}{m:02d}月度"], "勤務実績CSV", log=log)
        if not path:
            msg = (f"勤務実績CSV（{y}{m:02d}月度）が {folder} にありません。"
                   "Dr.JOY データダウンロード→帳票種別「勤務実績」・対象月・職業「医師」・本登録・Shift_JIS・HH:MM で出して置いてください。")
            if _allow_legacy():
                log(f"      ⚠【開発用】{msg} 従来値で進みます。")
                return {}, None
            raise InputError(msg)
    else:
        path = find_latest(folder, "勤務_*月度*.csv")
        if not path:
            return {}, None
    rows = list(_csv.reader(open(path, encoding="cp932", errors="replace")))
    if len(rows) < 3:
        raise InputError(f"勤務実績CSVが空です: {os.path.basename(path)}")
    hdr = rows[1]  # 行0=カテゴリ, 行1=項目名, 行2〜=データ

    def _col(*keys):
        for i, h in enumerate(hdr):
            hs = str(h)
            if all(k in hs for k in keys):
                return i
        return None
    c_name = _col("氏名")
    c_zan = _col("残業時間", "25")      # 残業時間（25％）＝法定時間外
    c_hol = _col("法定休日時間", "35")  # 法定休日時間（35％）＝休日労働
    c_job = _col("職名")
    c_dep = next((i for i, h in enumerate(hdr) if str(h).strip() == "所属"), None)
    c_days = next((i for i, h in enumerate(hdr) if str(h).strip() == "出勤日数"), None)
    if c_name is None or c_zan is None or c_hol is None:
        raise InputError(f"勤務実績CSVに必須の列がありません（氏名・残業時間（25％）・法定休日時間（35％））: {os.path.basename(path)}。"
                         "Dr.JOYの出力様式が変わった可能性があります。")
    out = {}
    for r in rows[2:]:
        if len(r) <= c_zan:
            continue
        nm = str(r[c_name]).strip()
        if not nm:
            continue
        h = _hhmm_to_h(r[c_zan])
        if len(r) > c_hol:
            h += _hhmm_to_h(r[c_hol])
        key = _N.canon(nm.replace("　", " ")).replace(" ", "")
        out[key] = h
        if meta is not None:   # 20261107 氏名・職名・所属（人別シートに無い人を足すとき用）・出勤日数（P2の休日数）
            try:
                _days = int(str(r[c_days]).strip()) if c_days is not None and len(r) > c_days else None
            except ValueError:
                _days = None
            meta[key] = (nm, str(r[c_job]).strip() if c_job is not None and len(r) > c_job else "",
                         str(r[c_dep]).strip() if c_dep is not None and len(r) > c_dep else "", _days)
    return out, os.path.basename(path)


_OK_POS = ("医師", "大学院生", "専門研修医", "初期研修医")


def _add_drjoy_only(persons, ot, meta, day_count, log=print):
    """人別シートに居ないが Dr.JOY に時間外がある医師を persons に足す（20261107）。足した人数を返す。"""
    import dtx14_names as _N
    have = {_N.canon(str(p.name).replace("　", " ")).replace(" ", "") for p in persons}
    added = []
    for key, v in ot.items():
        if key in have or v < 1.0 or key not in meta:
            continue
        nm, job, dep_raw = meta[key][:3]
        pos = normalize_position(job)
        if "非常勤" in job or pos not in _OK_POS:
            continue
        depts = [first_seg(x) or x.strip() for x in str(dep_raw).replace(",", "、").split("、") if x.strip()]
        depts = [d for d in dict.fromkeys(depts) if d and "院外" not in d]
        if not depts:
            continue
        nn = norm_name(nm)
        days = {d: (day_count or {}).get((nn, d), 0) for d in depts}
        if sum(days.values()) > 0:
            w = {d: n for d, n in days.items() if n > 0}
        else:
            w = {d: 1 for d in depts}
        tot = round(v, 1)
        ws = sum(w.values())
        per = {d: round(tot * n / ws, 1) for d, n in w.items()}
        label = "→".join(per) if len(per) > 1 else next(iter(per))
        p = PersonRow(dept_label=label, fallback_dept=label, position=pos, name=nm, name_norm=nn,
                      total=tot, teate=0.0, other=0.0, weekday_ot=0.0, holiday_ot=0.0, per_dept_hours=per)
        p.total_old = 0.0
        p.drjoy_ot = True
        p.drjoy_added = True
        p.drjoy_days = meta[key][3]
        persons.append(p)
        added.append(f"{nm}({pos[:2]}・{tot:.0f}h)")
    if added:
        log(f"      人別シートに無いがDr.JOYに時間外がある医師を足した（締めより後の承認など）: {len(added)}名 "
            + "、".join(added[:15]) + (" …" if len(added) > 15 else ""))
    return len(added)


def apply_drjoy_overtime(persons, folder, log=print, month_label=None, day_count=None, add_missing=True):
    """persons の時間外(p.total・p.per_dept_hours)を Dr.JOY「勤務実績」の
    時間外労働時間(残業25%+法定休日35%)で上書きする（インプレース）。
    per_dept_hours は、Dr.JOYの合計を申請ベースの科別時間の比で各科に割り振る（20261104。
    Dr.JOY値は科別に出ないため。従来どおり移動者は経た全科に入る。1科だけの人はその科に全量）。
    20261016 管理者 GO。20261017: 勤務実績が無い／月違い／ほぼ一致しないときは InputError で停止。
    上書き前の値を p.total_old、Dr.JOYの値で置いたかを p.drjoy_ot に残す（新旧対比と脚注用）。
    Dr.JOY に実績が無い人（退職等でアカウントが消えた人）は旧の申請値のまま残す（管理者決定 2026-10-17: R8.6〜8月）。
    20261107 管理者 GO: 人別シート（締めの時点の給与連携から作る）に居ないのに Dr.JOY に時間外が1時間以上ある医師
    （締めより後に承認された人。例 R8.7月 白銀先生88件を8/5〜8/17に承認）を persons に足す（add_missing）。
    科は day_count（給与連携の在籍日数）の比で割り、無ければ勤務実績の「所属」を等分。非常勤・他院だけの人は足さない。
    足した人は p.drjoy_added=True。
    戻り値 (上書き数, 未マッチ名リスト, ファイル名)。"""
    import dtx14_names as _N
    for p in persons:
        p.total_old = p.total
        p.drjoy_ot = False
        p.drjoy_days = None
    meta = {}
    ot, src = load_work_record_ot(folder, month_label=month_label, log=log, meta=meta)
    if not ot:
        log("      ⚠勤務実績CSVが無いため、時間外はDr.JOY値に差し替えません（従来値のまま・開発用）")
        return 0, [], src
    n_ok = 0
    miss = []
    for p in persons:
        k = _N.canon(str(p.name).replace("　", " ")).replace(" ", "")
        if k in ot:
            v = ot[k]
            p.total = round(v, 1)
            # 20261104: split the Dr.JOY total over the departments the person worked in,
            # in proportion to the old (application-based) hours, so movers stay in each
            # department's box plot instead of forming an "A→B" group of their own.
            old = {d: h for d, h in (p.per_dept_hours or {}).items() if h and h > 0}
            old_sum = sum(old.values())
            if len(old) > 1 and old_sum > 0:
                p.per_dept_hours = {d: round(p.total * h / old_sum, 1) for d, h in old.items()}
            else:
                dep = next(iter(old), None) or (p.dept_label or p.fallback_dept or "不明").strip() or "不明"
                p.per_dept_hours = {dep: p.total}
            p.drjoy_ot = True
            p.drjoy_days = meta.get(k, (None,) * 4)[3]   # 20261107 Dr.JOY 勤務実績の出勤日数（P2の休日数）
            n_ok += 1
        else:
            miss.append(p.name)
    if add_missing:
        _add_drjoy_only(persons, ot, meta, day_count, log)
    log(f"      時間外をDr.JOY勤務実績で上書き: {n_ok}名 一致 / 未マッチ {len(miss)}名（{src}）")
    if miss:
        log(f"        未マッチ（旧の申請値のまま）: {', '.join(miss[:15])}" + (" …" if len(miss) > 15 else ""))
    if persons and n_ok < 0.5 * len(persons):
        raise InputError(f"勤務実績CSVと人別シートの氏名が半分も一致しません（{n_ok}/{len(persons)}名）。"
                         f"ファイルの月・職業（医師）を確かめてください: {src}")
    return n_ok, miss, src


def load_drjoy_daisho(folder: str, month_label=None, log=print, path=None):
    """Dr.JOY「勤務予定」export(1ファイル複数シート xlsx)から代償休息必要時間を読む。
    20261102（締切の書き方） 管理者決定: 載せるのは「その月（M月）に発生し、翌月末までに付与が必要な1ヶ月分」。
       これは (M+1)月度の勤務予定exportの「代償休息 月末まで必要時間」。取るのは申請締切(翌月の第2営業日13時。5日以降なら第1営業日)・
       手当計算(14時)と同じ時点。残高「全期間必要時間」は脚注に併記する。
       ⚠M月度のexportを読んではいけない（M月度の「月末まで」は全員0になる）。
    month_label を渡すと (M+1)月度のファイルだけを読み、無ければ InputError。
    職業が「医師」のシートだけを数える。
    戻り値: {"rows":[{"name","dept","comp_h","bal_h"}...降順], "total","n","bal_total","bal_n","src","export_ym"} or None
    """
    if path:
        # 20261104 読むファイルを呼び出し側が決めたとき（代償休息の消化ページ＝その月の初めの残高）
        export_ym = None
    elif month_label:
        y, m = report_ym(month_label)
        ey, em = _next_ym(y, m)
        path = _pick_export(folder, "*勤務予定*.xlsx", [rf"{ey}年{em:02d}月度"], "勤務予定export", log=log)
        if not path:
            msg = (f"勤務予定export（{ey}年{em:02d}月度）が {folder} にありません。"
                   f"R{y - 2018}年{m}月分の代償休息は、翌月（{ey}年{em}月度）のexportの「月末まで必要時間」を読みます。"
                   "Dr.JOY データダウンロード→帳票種別「勤務予定」・1ファイル複数シート・対象月＝翌月度・職業「医師」で、"
                   "申請締切・承認の後に出して置いてください。")
            if _allow_legacy():
                log(f"      ⚠【開発用】{msg} 従来のinterval_v2で進みます。")
                return None
            raise InputError(msg)
        export_ym = (ey, em)
    else:
        path = find_latest(folder, "*勤務予定*.xlsx")
        if not path:
            return None
        export_ym = None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise InputError(f"勤務予定exportを開けません: {os.path.basename(path)} / {e}")

    def _h(s):
        s = str(s).strip() if s is not None else ""
        if ":" in s:
            try:
                hh, mm = s.split(":")
                return int(hh) + int(mm) / 60.0
            except Exception:
                return 0.0
        return 0.0

    def _after(row, label_pred):
        for j, c in enumerate(row):
            if c and label_pred(str(c).strip()) and j + 2 < len(row):
                return True, row[j + 2]
        return False, None
    rows = []
    saw_label = False
    try:
        for sn in wb.sheetnames:
            ws = wb[sn]
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            rr = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rr.append(row)
                if i >= 5:
                    break
            if not rr:
                continue
            nm = ""
            for j, c in enumerate(rr[0]):
                if c and str(c).strip() == "氏名" and j + 1 < len(rr[0]):
                    nm = str(rr[0][j + 1]).strip() if rr[0][j + 1] else ""
            if not nm and len(rr[0]) > 9 and rr[0][9]:
                nm = str(rr[0][9]).strip()
            dep = job = ""
            if len(rr) > 1:
                for j, c in enumerate(rr[1]):
                    if c and str(c).strip() == "所属" and j + 1 < len(rr[1]):
                        dep = str(rr[1][j + 1]).strip() if rr[1][j + 1] else ""
                    if c and str(c).strip() == "職業" and j + 1 < len(rr[1]):
                        job = str(rr[1][j + 1]).strip() if rr[1][j + 1] else ""
                if not dep and len(rr[1]) > 1 and rr[1][1]:
                    dep = str(rr[1][1]).strip()
            if (month_label or export_ym is None and path) and job and job != "医師":
                continue          # 看護師等は数えない（職業の欄がある export のみ）
            comp = bal = 0.0
            for row in rr:
                ok, v = _after(row, lambda s: "月末まで必要時間" in s)
                if ok:
                    saw_label = True
                    comp = _h(v)
                ok, v = _after(row, lambda s: "全期間必要時間" in s)
                if ok:
                    bal = _h(v)
            if (comp > 0.001 or bal > 0.001) and nm:
                rows.append({"name": nm, "dept": dep, "comp_h": round(comp, 2), "bal_h": round(bal, 2)})
    finally:
        wb.close()
    if not saw_label:
        raise InputError(f"勤務予定exportに「代償休息 月末まで必要時間」の欄がありません: {os.path.basename(path)}。"
                         "Dr.JOYの出力様式が変わった可能性があります。")
    mon = [o for o in rows if o["comp_h"] > 0.001]
    mon.sort(key=lambda o: -o["comp_h"])
    if month_label and not mon:
        raise InputError(f"勤務予定export（{os.path.basename(path)}）で「月末まで必要時間」が全員0時間です。"
                         "月の取り違え（対象月度を出している）か、Dr.JOYの代償休息の設定（開始月）が変わった可能性があります。")
    return {"rows": mon, "total": round(sum(o["comp_h"] for o in mon), 1), "n": len(mon),
            "bal_total": round(sum(o["bal_h"] for o in rows), 1), "bal_n": sum(1 for o in rows if o["bal_h"] > 0.001),
            "src": os.path.basename(path), "export_ym": export_ym,
            "all_rows": sorted(rows, key=lambda o: -o["bal_h"])}


DAY_H = 7.75   # 20261104 所定の1日＝8:30〜17:15から昼休み1時間を除いた7時間45分
_ARCHIVE = chr(92) * 2 + "fileserver" + chr(92) + "share$" + chr(92) + "00_DrJOY" + chr(92) + "働き方レポート" + chr(92) + "archive"


def _leave_hours(seg: str) -> float:
    """「代償休息 | 09:00~13:16」「代償休息 | 午後休 (13:00~17:15)」「代償休息」（時間なし＝終日）の時間。
    昼休み 12:00〜13:00 にかかる分は引く。20261104"""
    import re as _re
    m = _re.search(r"(\d{1,2}):(\d{2})\s*[~〜]\s*(\d{1,2}):(\d{2})", seg)
    if not m:
        return DAY_H
    a = int(m.group(1)) * 60 + int(m.group(2))
    b = int(m.group(3)) * 60 + int(m.group(4))
    lunch = max(0, min(b, 13 * 60) - max(a, 12 * 60))
    return max(0, b - a - lunch) / 60.0


def load_daisho_taken(pay_xlsx: str) -> dict:
    """給与連携項目一覧の「所定時間、休暇等」から、代償休息の休暇申請（承認済み）を人ごとに集める。
    戻り値 {canon氏名: {"name","times","hours","days":[(日付文字列, 時間)]}}。20261104"""
    import dtx14_names as _N
    out = {}
    wb = openpyxl.load_workbook(pay_xlsx, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        it = ws.iter_rows(values_only=True)
        h = [str(c).strip() if c is not None else "" for c in next(it)]
        ci = next((i for i, c in enumerate(h) if c == "所定時間、休暇等"), None)
        ni = next((i for i, c in enumerate(h) if c == "氏名"), None)
        di = next((i for i, c in enumerate(h) if c in ("日付", "勤務日")), None)
        if ci is None or ni is None:
            raise InputError(f"給与連携項目一覧に「所定時間、休暇等」「氏名」の列がありません: {os.path.basename(pay_xlsx)}")
        for r in it:
            v = r[ci] if ci < len(r) else None
            if not v or "代償休息" not in str(v):
                continue
            for seg in str(v).split(","):
                if "[休暇申請]代償休息" not in seg:
                    continue
                hh = _leave_hours(seg.split("代償休息", 1)[1])
                nm = str(r[ni]).strip()
                k = _N.canon(nm.replace("　", " ")).replace(" ", "")
                o = out.setdefault(k, {"name": nm, "times": 0, "hours": 0.0, "days": []})
                o["times"] += 1
                o["hours"] += hh
                o["days"].append((str(r[di]) if di is not None else "", hh))
    finally:
        wb.close()
    return out


def load_daisho_usage(folder: str, month_label: str, pay_xlsx: str, log=print):
    """代償休息の消化（20261104 管理者指示）: その月の初めに残高（全期間必要時間）が1日（7:45）以上あった医師が、
    その月に代償休息をどれだけ取ったか。
    月初めの残高＝その月度（M月度）の勤務予定export のうち、その月の申請締切（tools/deadline.py）以後に取った最も早いもの。
    作業フォルダか、アーカイブの前月フォルダ（手当計算が翌月度として保存）から探す。見つからなければ None（ページを出さない）。"""
    import glob, datetime as _dt, re as _re
    import dtx14_names as _N
    y, m = report_ym(month_label)
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    pat = f"{y}年{m:02d}月度_勤務予定_*.xlsx"
    cands = glob.glob(os.path.join(folder, pat)) + glob.glob(os.path.join(_ARCHIVE, f"{py}-{pm:02d}", pat))
    cands = [c for c in cands if not os.path.basename(c).startswith("~$")]
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "tools"))
        import deadline as _D
        dl, _ = _D.deadline(y, m)
        dl_dt = _dt.datetime(dl.year, dl.month, dl.day, 13, 0)
    except Exception:
        dl_dt = _dt.datetime(y, m, 1)
    cands = [c for c in cands if _dt.datetime.fromtimestamp(os.path.getmtime(c)) >= dl_dt]
    if not cands:
        log(f"      ⚠代償休息の消化: {y}年{m:02d}月度の勤務予定export（{dl_dt:%m/%d %H:%M}以後）が無いので、ページを出しません")
        return None
    path = min(cands, key=os.path.getmtime)
    start = load_drjoy_daisho(folder, log=log, path=path)
    taken = load_daisho_taken(pay_xlsx)
    rows = []
    used = set()
    for o in start["all_rows"]:
        k = _N.canon(o["name"].replace("　", " ")).replace(" ", "")
        if o["bal_h"] + 1e-6 < DAY_H:
            continue
        t = taken.get(k)
        used.add(k)
        rows.append({"name": o["name"], "dept": o["dept"], "bal_h": o["bal_h"],
                     "times": t["times"] if t else 0, "taken_h": round(t["hours"], 2) if t else 0.0})
    others = [v for k, v in taken.items() if k not in used]
    acq = _dt.datetime.fromtimestamp(os.path.getmtime(path))
    log(f"      PP: 代償休息の消化＝{os.path.basename(path)}（{acq:%m/%d %H:%M}取得）で残高1日以上 {len(rows)}名、"
        f"うち取得 {sum(1 for r in rows if r['times'])}名・{sum(r['taken_h'] for r in rows):.1f}時間。"
        f"ほかに残高1日未満で取得 {len(others)}名")
    return {"rows": rows, "others": others, "src": os.path.basename(path), "acq": acq,
            "all_taken": taken}


def detect_month(src_xlsm: str) -> tuple[str, int]:
    """人別シートの最初のデータ行(A列)の日付から R{N}年{M}月 と 暦日数 を取得"""
    wb = openpyxl.load_workbook(src_xlsm, read_only=True, data_only=True)
    try:
        for name in wb.sheetnames:
            if name == "名簿":
                continue
            ws = _ws_full(wb[name])
            a2 = ws.cell(2, 1).value
            if a2 is None or str(a2).strip() == "":
                continue
            v = ws.cell(DATA_START, 1).value
            try:
                if hasattr(v, "year"):
                    y, m = v.year, v.month
                else:
                    mm = _DATE_RE.search(str(v))
                    if not mm:
                        continue
                    y, m = int(mm.group(1)), int(mm.group(2))
                # 暦日数
                next_m = date(y, m + 1, 1) if m < 12 else date(y + 1, 1, 1)
                last = (next_m - date(y, m, 1)).days
                return f"R{y-2018}年{m}月", last
            except Exception:
                continue
        return "", 30
    finally:
        wb.close()


# ============================================================
# 日別所属マップ
# ============================================================
def build_dept_by_date(pay_xlsx: Optional[str]) -> tuple[dict, dict, list[WarnEntry]]:
    """
    給与連携項目一覧.xlsx から:
      dept_by_date[(name_norm, dateord)] = 第1階層科
      day_count[(name_norm, dept)] = 在籍暦日数
    返り値: (dept_by_date, day_count, warns)
    """
    warns: list[WarnEntry] = []
    if pay_xlsx is None:
        warns.append(WarnEntry("中", "給与連携なし", "build_dept_by_date", "(ソース)", item="*給与連携項目一覧*.xlsx",
                               reason="日割り集計不可。最終所属(A1)で代替"))
        return {}, {}, warns
    wb = openpyxl.load_workbook(pay_xlsx, read_only=True, data_only=True)
    try:
        ws = _ws_full(wb[wb.sheetnames[0]])
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return {}, {}, warns
        hdr = rows[0]
        idx = {c: j for j, c in enumerate(hdr) if c}
        cDept = idx.get("所属名履歴（階層含む）") if "所属名履歴（階層含む）" in idx else idx.get("所属")
        cName = idx.get("氏名")
        cDate = idx.get("日付")
        if cDept is None or cName is None or cDate is None:
            warns.append(WarnEntry("中", "給与連携列欠落", "build_dept_by_date", "(ソース)",
                                   item=os.path.basename(pay_xlsx),
                                   reason="所属/氏名/日付列が見つからない。日割り集計不可"))
            return {}, {}, warns
        dept_by_date: dict[tuple[str, int], str] = {}
        day_count: dict[tuple[str, str], int] = {}
        for r in rows[1:]:
            if not r:
                continue
            nm = norm_name(r[cName] if cName < len(r) else None)
            if not nm:
                continue
            dp_raw = r[cDept] if cDept < len(r) else None
            if dp_raw is None:
                continue
            dp = first_seg(dp_raw)
            if not dp:
                continue
            ord_ = date_ord(r[cDate] if cDate < len(r) else None)
            if ord_ <= 0:
                continue
            k1 = (nm, ord_)
            if k1 not in dept_by_date:
                dept_by_date[k1] = dp
            k2 = (nm, dp)
            day_count[k2] = day_count.get(k2, 0) + 1
        return dept_by_date, day_count, warns
    finally:
        wb.close()


# ============================================================
# 管理当直マップ（DTX14）
# ============================================================
def build_bucho_windows(folder: str) -> tuple[dict, Optional[WarnEntry]]:
    """DTX14『人別当直日』から管理当直の (氏名norm -> [当直窓(start,end)]) を構築。
    管理当直(教授/准教授のみ)は宿日直許可あり＝救急/小児でも当直まるごと加算の対象外。
    判別: A列='管理当直' の行（E列=フルネーム, C列=当直日, D列=昼/夜）。
    DTX14 や『人別当直日』が無ければ空dict＋要確認警告を返す（停止しない）。"""
    az14 = find_latest(folder, "*DTX14*.xlsm")
    if not az14:
        return {}, WarnEntry("中", "管理当直控除不可", "build_bucho_windows", "(ソース)",
                             item="*DTX14*.xlsm",
                             reason="DTX14が見つからず、救急/小児の管理当直を当直加算から除外できない")
    wins: dict[str, list[tuple[datetime, datetime]]] = {}
    wb = openpyxl.load_workbook(az14, read_only=True, data_only=True)
    try:
        if "人別当直日" not in wb.sheetnames:
            return {}, WarnEntry("中", "管理当直控除不可", "build_bucho_windows", "(ソース)",
                                 item=os.path.basename(az14),
                                 reason="DTX14に『人別当直日』シートが無く、管理当直を判別できない")
        ws = _ws_full(wb["人別当直日"])
        for row in ws.iter_rows(values_only=True):
            a = row[0] if len(row) > 0 else None
            if a is None or str(a).strip() != "管理当直":
                continue
            nm = norm_name(row[4] if len(row) > 4 else None)            # E列=フルネーム
            dd = cell_date(row[2] if len(row) > 2 else None)            # C列=当直日
            dtype = str(row[3]).strip() if (len(row) > 3 and row[3] is not None) else ""  # D列=昼/夜
            if not nm or dd is None or dtype == "":
                continue
            nhol = is_holiday_date(dd + timedelta(days=1))              # 翌日休日は日付から算出（DTX14に翌行文脈なし）
            win_s, win_e = duty_window(dd, dtype, nhol)
            if win_s is None:
                continue
            wins.setdefault(nm, []).append((win_s, win_e))
        return wins, None
    finally:
        wb.close()


# ============================================================
# 人別シート走査（残業集計＋日別科振り分け）
# ============================================================
def collect_person_rows(src_xlsm: str,
                        dept_by_date: dict,
                        bucho_windows: dict = None,
                        cal_days: int = 0,
                        log=lambda s: None) -> tuple[list[PersonRow], list[WarnEntry]]:
    """全人別シートを走査して PersonRow リストを返す。"""
    warns: list[WarnEntry] = []
    out: list[PersonRow] = []
    wb = openpyxl.load_workbook(src_xlsm, read_only=True, data_only=True)
    try:
        names = [n for n in wb.sheetnames if n != "名簿"]
        total = len(names)
        el_total_min = 0    # 救急/小児 早上がり控除（実効=duty上限後）の合計（ログ用）
        el_count = 0        # 控除が発生した人数
        el_big = []         # 控除が大きい人（要確認用・可視化）
        for i, sname in enumerate(names, 1):
            ws = _ws_full(wb[sname])
            # read_onlyでは ws.cell() が遅いので行を一括取得（タプル, 0-indexed）
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < HDR_ROW + 1:           # ヘッダー(4行)＋最低1データ行
                continue

            def gv(rr1, cc1):                      # 1-indexed安全アクセス
                row = rows[rr1 - 1] if 0 <= rr1 - 1 < len(rows) else ()
                return row[cc1 - 1] if 0 <= cc1 - 1 < len(row) else None

            pName = gv(2, 1)
            if pName is None or str(pName).strip() == "":
                continue
            pName = str(pName).strip()
            deptRaw = gv(1, 1)
            deptRaw = str(deptRaw).strip() if deptRaw is not None else ""
            position = normalize_position(gv(1, 3))
            # 除外科
            if deptRaw in EXCLUDE_DEPTS:
                if deptRaw == KARI:
                    warns.append(WarnEntry("高", "仮所属(Dr.JOY修正要)", "collect_person_rows", pName,
                                           dept=deptRaw, item="科(A1)",
                                           reason="所属未確定(仮)のまま勤務データに混入。Dr.JOYで本所属に修正要",
                                           sheet=sname))
                else:
                    warns.append(WarnEntry("中", "除外科", "collect_person_rows", pName,
                                           dept=deptRaw, item="科(A1)",
                                           reason="集計対象外の所属。意図通りか確認",
                                           sheet=sname))
                continue
            # ヘッダー走査（4行目）
            hdr = list(rows[HDR_ROW - 1])
            def col_of(name):
                try:
                    return hdr.index(name) + 1
                except ValueError:
                    return 0
            cDateCol = col_of("日付")
            if cDateCol == 0:
                warns.append(WarnEntry("中", "列欠落", "collect_person_rows", pName, dept=deptRaw,
                                       item="日付列", reason="ヘッダー行(4)に日付列なし", sheet=sname))
                continue
            cWeekday = col_of("曜日")
            cHoliday = col_of("休日")
            cReason, cStart, cEnd = [0] * (MAX_SLOTS + 1), [0] * (MAX_SLOTS + 1), [0] * (MAX_SLOTS + 1)
            for s in range(1, MAX_SLOTS + 1):
                cReason[s] = col_of(f"時間外理由{s}")
                cStart[s] = col_of(f"時間外開始{s}")
                cEnd[s] = col_of(f"時間外終了{s}")
            # 勤務間インターバル用: 研鑽(種類/開始/終了)・出退勤・当直の列（20260606）
            cKenT, cKenS, cKenE = [0] * (MAX_SLOTS + 1), [0] * (MAX_SLOTS + 1), [0] * (MAX_SLOTS + 1)
            for s in range(1, MAX_SLOTS + 1):
                cKenT[s] = col_of(f"研鑽申請種類名{s}")
                cKenS[s] = col_of(f"自己研鑽開始{s}")
                cKenE[s] = col_of(f"自己研鑽終了{s}")
            cInIv = col_of("出勤")
            cOutIv = col_of("退勤")
            cDutyIv = col_of("当直")
            # データ範囲: 列A(index0)の最終非空行まで（VBA Cells(.,1).End(xlUp) 同等）
            lastRow = 0
            for r in range(DATA_START, len(rows) + 1):
                a = gv(r, 1)
                if a is not None and str(a).strip() != "":
                    lastRow = r
            # 全行集計（厳密整数分で累積：FP境界ぶれ排除＝VBA Round準拠）
            taMin = oMin = wMin = hMin = 0
            per_dept_min: dict[str, int] = {}
            nm_n = norm_name(pName)
            # 管理当直(宿日直許可あり)の窓: 救急/小児の人だけ「当直加算/窓内申告時間外」の対象外にする
            bucho_wins = (bucho_windows or {}).get(nm_n) if is_er_peds(deptRaw) else None
            for r in range(DATA_START, lastRow + 1):
                row = rows[r - 1]
                dv = row[cDateCol - 1] if cDateCol - 1 < len(row) else None
                # 休日判定（ソースの曜日/休日列は空のことが多いので、日付から土日・国民の祝日も判定。
                #   20260608: 旧実装は空列依存で休日時間外が全て平日に誤分類されていた不具合を修正。
                #   休日列が非空ならそれも尊重＝is_holiday_dateに集約。まるごと加算側と同一判定で整合）
                _ddRow = cell_date(dv)
                _holCell = row[cHoliday - 1] if (cHoliday > 0 and cHoliday - 1 < len(row)) else None
                isHol = is_holiday_date(_ddRow, _holCell) if _ddRow is not None else False
                # 時間外スロット（分単位＝整数で集計）
                slot_min = 0
                for s in range(1, MAX_SLOTS + 1):
                    if cReason[s] == 0 or cStart[s] == 0 or cEnd[s] == 0:
                        break
                    reason = row[cReason[s] - 1] if cReason[s] - 1 < len(row) else None
                    if reason is None or str(reason).strip() == "":
                        break
                    sv = row[cStart[s] - 1] if cStart[s] - 1 < len(row) else None
                    ev = row[cEnd[s] - 1] if cEnd[s] - 1 < len(row) else None
                    # 20260901【24時間以上の申告が24h短く計上されるバグの修正】
                    #   try_time_hours() は仕様どおり「0以上24未満の時間」を返す＝datetime を渡しても
                    #   時刻部分しか見ない。旧実装はこれで開始・終了の差を取っていたため、
                    #   例）黒潮浩一 8/9 『診療（手当あり） 07:22 → 翌08:30』(実25:08) が
                    #   8.5-7.37=1.13h ＝ 1:08 として計上されていた（差が正なので +1440 補正も効かない）。
                    #   cell_datetime() は datetime をそのまま返す＝日付を保つので、そこから長さを取る。
                    #   ※ try_time_hours() 自体の仕様は変えない（cell_datetime/time_on_date が依存）。
                    rd = cell_date(dv)
                    ks = cell_datetime(sv, rd) if rd is not None else None
                    ke = cell_datetime(ev, rd) if rd is not None else None
                    if ks is not None and ke is not None:
                        if ke < ks:                       # 日をまたぐ（時刻のみ値のとき）
                            ke = ke + timedelta(days=1)
                        dur_min = int(round((ke - ks).total_seconds() / 60.0))
                    else:
                        # フォールバック（日付が読めない／時刻化できない）＝従来どおり時刻差で算出。
                        # 黙って落とさず、24時間未満の申告はこれまでと同じ結果になる。
                        sh = try_time_hours(sv); eh = try_time_hours(ev)
                        if sh is None or eh is None:
                            continue
                        dur_min = int(round((eh - sh) * 60))   # 厳密分
                        if dur_min < 0: dur_min += 1440
                    if dur_min <= 0: continue
                    # 管理当直窓に重なる申告時間外は除外（非ER当直と同じ＝当直中の時間外は計上しない）
                    #   ※ R8.7月・R8.8月はどちらも発火0件だが、これは「救急/小児 かつ 管理当直」を
                    #     兼ねる医師が実際にいなかったためで、実装は意図どおり（ユーザー確認 20260901）。
                    #     デッドコードではないので消さないこと。
                    if bucho_wins and ks is not None and ke is not None:
                        clip = 0
                        for (bw_s, bw_e) in bucho_wins:
                            clip += overlap_minutes(ks, ke, bw_s, bw_e)
                        if clip > 0:
                            dur_min = max(0, dur_min - clip)
                            if dur_min == 0:
                                continue
                    slot_min += dur_min
                    if "手当あり" in str(reason):
                        taMin += dur_min
                    else:
                        oMin += dur_min
                if slot_min > 0:
                    if isHol:
                        hMin += slot_min
                    else:
                        wMin += slot_min
                    # 日別科振り分け
                    rowOrd = date_ord(dv)
                    rowDept = dept_by_date.get((nm_n, rowOrd), deptRaw) if dept_by_date else deptRaw
                    if rowDept:
                        per_dept_min[rowDept] = per_dept_min.get(rowDept, 0) + slot_min
            # --- 救急/小児: 当直まるごと加算（宿日直許可なし＝当直窓全体が労働時間。20260605） ---
            # 加算分 = 当直窓 −（平日9:00-17:15の通常勤務時間帯）−（窓内に既計上の時間外）。
            # 平日昼の通常勤務帯には当直は存在しない想定。万一あれば除外。夜・早朝・休日は全部加算。
            duty_min = 0
            if is_er_peds(deptRaw):
                cDuty = col_of("当直")
                if cDuty > 0:
                    for r in range(DATA_START, lastRow + 1):
                        row = rows[r - 1]
                        dty = row[cDuty - 1] if cDuty - 1 < len(row) else None
                        dty = str(dty).strip() if dty is not None else ""
                        if dty not in DUTY_TYPES:
                            continue
                        dv = row[cDateCol - 1] if cDateCol - 1 < len(row) else None
                        dd = cell_date(dv)
                        if dd is None:
                            continue
                        # 管理当直日は宿日直許可あり＝当直まるごと加算しない（窓加算をスキップ）
                        if bucho_wins and any(bw_s.date() == dd for (bw_s, _bw_e) in bucho_wins):
                            continue
                        nrow = rows[r] if r < len(rows) else None     # 翌行(1-indexed r → 0-indexed rows[r])
                        nhol = next_is_holiday(dd + timedelta(days=1), nrow, cHoliday)
                        win_s, win_e = duty_window(dd, dty, nhol)
                        if win_s is None:
                            continue
                        # 当直日の休日判定（土日・国民の祝日・休日列）
                        isHolD = is_holiday_date(dd, row[cHoliday - 1] if (cHoliday > 0 and cHoliday - 1 < len(row)) else None)
                        # 平日は通常勤務時間帯(〜17:15)を窓から除外。休日は窓全体が労働。
                        eff_s = win_s
                        if not isHolD:
                            reg_end = datetime(dd.year, dd.month, dd.day, 17, 15)
                            if eff_s < reg_end:
                                eff_s = reg_end
                        eff_e = win_e
                        if eff_e <= eff_s:
                            continue
                        window_min = int((eff_e - eff_s).total_seconds() / 60)
                        # 窓内に既計上の時間外(当日行＋翌行)を差し引く（二重計上防止）
                        scan_rows = [row]
                        if dty != "昼" and nrow is not None:
                            ndd = cell_date(nrow[cDateCol - 1] if cDateCol - 1 < len(nrow) else None)
                            if ndd == dd + timedelta(days=1):
                                scan_rows.append(nrow)
                        ot_in = 0
                        for rr in scan_rows:
                            # 翌行(nrow)の時刻には翌日を基準日に（時刻のみ値が来た場合の保険。
                            # 実データの時間外は実日付つきdatetimeなのでcell_datetimeは基準日を無視＝現結果は不変）
                            base_dd = dd if rr is row else (dd + timedelta(days=1))
                            for s in range(1, MAX_SLOTS + 1):
                                if cReason[s] == 0 or cStart[s] == 0 or cEnd[s] == 0:
                                    break
                                rsn = rr[cReason[s] - 1] if cReason[s] - 1 < len(rr) else None
                                if rsn is None or str(rsn).strip() == "":
                                    break
                                ks = cell_datetime(rr[cStart[s] - 1] if cStart[s] - 1 < len(rr) else None, base_dd)
                                ke = cell_datetime(rr[cEnd[s] - 1] if cEnd[s] - 1 < len(rr) else None, base_dd)
                                if ks is None or ke is None:
                                    continue
                                if ke < ks:
                                    ke = ke + timedelta(days=1)
                                ot_in += overlap_minutes(ks, ke, eff_s, eff_e)
                        add = window_min - ot_in
                        if add <= 0:
                            continue
                        duty_min += add
                        rowOrd = date_ord(dv)
                        rowDept = dept_by_date.get((nm_n, rowOrd), deptRaw) if dept_by_date else deptRaw
                        if rowDept:
                            per_dept_min[rowDept] = per_dept_min.get(rowDept, 0) + add
                        if isHolD:
                            hMin += add
                        else:
                            wMin += add

            # --- 救急/小児: 当直前後の早上がり控除 20260605（監査C1/C2対応で改良） ---
            # 在席=出勤〜退勤。平日(土日祝・休日列以外)の9:00-17:15のうち在席に含まれない分を控除候補。
            #  C1: 翌朝退勤(Excel時刻シリアル≥1日)は time_on_date が翌日扱い→24h勤務日を誤控除しない。
            #  C2: ①終日打刻なしの平日は休みとして控除するが、打刻のある初日〜最終日の外（月途中開始や
            #       データ欠落）は控除しない。②実控除は当直まるごと加算(duty_min)を上限（当直前後の休みで
            #       当直の過大計上を補正する趣旨。夜間のみ勤務者を0にしない）。
            #  出退勤打刻列が無い旧形式ソースでは控除なし（col_of=0）。
            el_min = 0
            if is_er_peds(deptRaw):
                cIn = col_of("出勤"); cOut = col_of("退勤")
                REG_MIN = (17 * 60 + 15) - (9 * 60)   # 9:00-17:15 = 495分
                if cIn > 0 and cOut > 0:
                    # 打刻のある勤務日の範囲（この外＝在籍前/データ欠落とみなし控除しない）
                    first_pd = last_pd = None
                    for r in range(DATA_START, lastRow + 1):
                        row = rows[r - 1]
                        iv = row[cIn - 1] if cIn - 1 < len(row) else None
                        ov = row[cOut - 1] if cOut - 1 < len(row) else None
                        if (iv is not None and str(iv).strip() != "") or (ov is not None and str(ov).strip() != ""):
                            pdd = cell_date(row[cDateCol - 1] if cDateCol - 1 < len(row) else None)
                            if pdd is not None:
                                first_pd = pdd if first_pd is None else min(first_pd, pdd)
                                last_pd = pdd if last_pd is None else max(last_pd, pdd)
                    for r in range(DATA_START, lastRow + 1):
                        row = rows[r - 1]
                        dvv = row[cDateCol - 1] if cDateCol - 1 < len(row) else None
                        dd = cell_date(dvv)
                        if dd is None:
                            continue
                        holcell = row[cHoliday - 1] if (cHoliday > 0 and cHoliday - 1 < len(row)) else None
                        if is_holiday_date(dd, holcell):
                            continue   # 土日・祝日・休日列＝所定労働なし→控除対象外
                        iv = row[cIn - 1] if cIn - 1 < len(row) else None
                        ov = row[cOut - 1] if cOut - 1 < len(row) else None
                        iv_empty = (iv is None or str(iv).strip() == "")
                        ov_empty = (ov is None or str(ov).strip() == "")
                        if iv_empty and ov_empty:
                            # 終日打刻なし＝休み。ただし打刻期間の外＝データ欠落とみなし控除しない（C2）
                            if first_pd is not None and first_pd <= dd <= last_pd:
                                el_min += REG_MIN
                            continue
                        ks = time_on_date(iv, dd); ke = time_on_date(ov, dd)
                        if ks is None or ke is None:
                            continue              # 片方だけ等の異常は控除しない（安全側）
                        if ke <= ks:
                            ke = ke + timedelta(days=1)   # 念のための補正（time値で退勤<出勤の場合）
                        reg_s = datetime(dd.year, dd.month, dd.day, 9, 0)
                        reg_e = datetime(dd.year, dd.month, dd.day, 17, 15)
                        present = overlap_minutes(ks, ke, reg_s, reg_e)
                        el_min += max(0, REG_MIN - present)

            gross_min = taMin + oMin + duty_min
            # 実績ゼロ(活動なし)の人は全統計から除外（早上がり控除の前で判定。20260604）
            if gross_min == 0:
                continue
            # 早上がり控除は当直まるごと加算(duty_min)を上限に総時間外から差引（C2）。per_dept も同額削る。
            #   ※ この控除は「当直明けに早く上がった＝代償休息をとった」と見なすもの（ユーザー確認 20260901）。
            #     規程に明文の根拠があるわけではないが、仕様として正しい。
            #     「根拠不明の控除」と誤認して消さないこと。
            el_eff = el_min if el_min < duty_min else duty_min
            total_min = max(0, gross_min - el_eff)
            if el_eff > 0:
                el_total_min += el_eff
                el_count += 1
                if el_min >= 80 * 60:   # 控除候補が大きい人を可視化（要確認）
                    el_big.append((pName, deptRaw, el_min / 60.0, el_eff / 60.0))
                rem = el_eff
                for _dept in sorted(per_dept_min, key=lambda d: -per_dept_min[d]):
                    take = min(rem, per_dept_min[_dept])
                    per_dept_min[_dept] -= take
                    rem -= take
                    if rem <= 0:
                        break
            # 科別集計用に厳密時間(分/60)へ。挿入順は per_dept_min を維持
            per_dept = {dept: m / 60.0 for dept, m in per_dept_min.items()}
            label = label_excluding_kari(per_dept, deptRaw)

            # --- 勤務間インターバル不足（専攻医・初期研のみ。20260606 医師の働き方改革） ---
            # 勤務=出退勤の在席∪救急小児の当直窓∪時間外∪研鑽「通常当直業務」。休息=研鑽の通常当直業務以外＋空き。
            # 連続する勤務クラスタ(隙間<9h=休息未満は連結)の間の隙間がインターバル。基準は直前クラスタが
            # 救急/小児当直(宿日直許可なし)を含むか＋職位で 24h(初期研)/18h(専攻医)、それ以外9h。
            iv_target = position in ("専門研修医", "初期研修医")
            iv_viol = 0; iv_short_min = 0; iv_nopunch = 0
            if iv_target:
                work_raw = []; rest_raw = []; er_windows = []
                for r in range(DATA_START, lastRow + 1):
                    row = rows[r - 1]
                    dvv = row[cDateCol - 1] if cDateCol - 1 < len(row) else None
                    dd = cell_date(dvv)
                    if dd is None:
                        continue
                    day_work = False; day_punch = False
                    if cInIv > 0 and cOutIv > 0:
                        a = time_on_date(row[cInIv - 1] if cInIv - 1 < len(row) else None, dd)
                        b = time_on_date(row[cOutIv - 1] if cOutIv - 1 < len(row) else None, dd)
                        if a is not None and b is not None:
                            if b <= a:
                                b = b + timedelta(days=1)
                            work_raw.append((a, b)); day_punch = True
                    for s in range(1, MAX_SLOTS + 1):
                        if cReason[s] == 0 or cStart[s] == 0 or cEnd[s] == 0:
                            break
                        rsn = row[cReason[s] - 1] if cReason[s] - 1 < len(row) else None
                        if rsn is None or str(rsn).strip() == "":
                            break
                        a = cell_datetime(row[cStart[s] - 1] if cStart[s] - 1 < len(row) else None, dd)
                        b = cell_datetime(row[cEnd[s] - 1] if cEnd[s] - 1 < len(row) else None, dd)
                        if a is not None and b is not None:
                            if b < a:
                                b = b + timedelta(days=1)
                            work_raw.append((a, b)); day_work = True
                    for s in range(1, MAX_SLOTS + 1):
                        if cKenT[s] == 0 or cKenS[s] == 0 or cKenE[s] == 0:
                            break
                        kt = row[cKenT[s] - 1] if cKenT[s] - 1 < len(row) else None
                        if kt is None or str(kt).strip() == "":
                            break
                        a = cell_datetime(row[cKenS[s] - 1] if cKenS[s] - 1 < len(row) else None, dd)
                        b = cell_datetime(row[cKenE[s] - 1] if cKenE[s] - 1 < len(row) else None, dd)
                        if a is None or b is None:
                            continue
                        if b < a:
                            b = b + timedelta(days=1)
                        if str(kt).strip() == "通常当直業務":
                            work_raw.append((a, b)); day_work = True
                        else:
                            rest_raw.append((a, b))   # 自己研鑽/待機/休憩睡眠/移動＝休息
                    if is_er_peds(deptRaw) and cDutyIv > 0:
                        dty = row[cDutyIv - 1] if cDutyIv - 1 < len(row) else None
                        dty = str(dty).strip() if dty is not None else ""
                        if dty in DUTY_TYPES:
                            nhol = is_holiday_date(dd + timedelta(days=1))
                            w1, w2 = duty_window(dd, dty, nhol)
                            if w1 is not None:
                                work_raw.append((w1, w2)); er_windows.append((w1, w2)); day_work = True
                    if day_work and not day_punch:
                        iv_nopunch += 1
                # phys=元の在席/当直窓(未カット)の連続スパン。carved=phys−休息(自己研鑽等)。
                # carvedの隙間が phys スパンの内側＝当直中の仮眠等(休息扱いだが勤務の区切りではない)→同一勤務に連結。
                # phys の外側の隙間＝帰宅した休息＝インターバル。これで通常日の<9h不足を拾い、仮眠の誤検出を避ける。
                phys = _merge_intervals(work_raw)
                carved = _merge_intervals(_subtract_intervals(phys, _merge_intervals(rest_raw)))

                def _within_phys(a, b):
                    return any(ps <= a and b <= pe for (ps, pe) in phys)

                episodes = []   # [start, end, has_er]
                for (ss, ee) in carved:
                    haser = any(not (we <= ss or ws >= ee) for (ws, we) in er_windows)
                    if episodes and _within_phys(episodes[-1][1], ss):
                        if ee > episodes[-1][1]:
                            episodes[-1][1] = ee
                        episodes[-1][2] = episodes[-1][2] or haser
                    else:
                        episodes.append([ss, ee, haser])
                req_er = 24.0 if position == "初期研修医" else 18.0
                for ci in range(len(episodes) - 1):
                    gap = (episodes[ci + 1][0] - episodes[ci][1]).total_seconds() / 3600.0
                    req = req_er if episodes[ci][2] else 9.0
                    if gap < req:
                        iv_viol += 1
                        iv_short_min += int(round((req - gap) * 60))

            # --- C: 当直回数・オンコール回数・連続勤務日数・休日数（20260607） ---
            # 勤務日 = 出勤打刻あり or 当直あり（オンコール単独は休息扱い・ユーザー合意）。
            # 休日数 = 暦日(cal_days) − 勤務日数。連続勤務日数 = 勤務日が連続した最大日数。
            cOncall = col_of("オンコール")
            n_duty = n_oncall = 0
            work_dates = set()
            # 勤務日 = 申告のある日（時間外 or 自己研鑽 or 当直）。出勤打刻はビーコン放置で汚染される
            #   ため勤務日判定に使わない（20260608 申告ベースへ変更）。当直は信頼できる勤務シグナル。
            #   ビーコン置きっぱなし疑いは別途 scan_beacon（在席ベース・全員対象）で出す。
            for r in range(DATA_START, lastRow + 1):
                row = rows[r - 1]
                dd = cell_date(row[cDateCol - 1] if cDateCol - 1 < len(row) else None)
                if dd is None:
                    continue
                dty = row[cDutyIv - 1] if (cDutyIv > 0 and cDutyIv - 1 < len(row)) else None
                has_duty = dty is not None and str(dty).strip() != ""
                oc = row[cOncall - 1] if (cOncall > 0 and cOncall - 1 < len(row)) else None
                has_oncall = oc is not None and str(oc).strip() != ""
                if has_duty:
                    n_duty += 1
                if has_oncall:
                    n_oncall += 1
                has_apply = has_duty
                if not has_apply:
                    for s in range(1, MAX_SLOTS + 1):
                        cs = cStart[s]
                        if cs > 0 and cs - 1 < len(row) and row[cs - 1] is not None and str(row[cs - 1]).strip() != "":
                            has_apply = True
                            break
                if not has_apply:
                    for s in range(1, MAX_SLOTS + 1):
                        cs = cKenS[s]
                        if cs > 0 and cs - 1 < len(row) and row[cs - 1] is not None and str(row[cs - 1]).strip() != "":
                            has_apply = True
                            break
                if has_apply:
                    work_dates.add(dd)
            n_workday = len(work_dates)
            n_off = max(0, cal_days - n_workday) if cal_days else 0
            max_streak = 0
            if work_dates:
                sd = sorted(work_dates)
                run = max_streak = 1
                for k in range(1, len(sd)):
                    run = run + 1 if (sd[k] - sd[k - 1]).days == 1 else 1
                    if run > max_streak:
                        max_streak = run

            out.append(PersonRow(
                dept_label=label, fallback_dept=deptRaw, position=position,
                name=pName, name_norm=nm_n,
                total=min_to_h1(total_min), teate=min_to_h1(taMin), other=min_to_h1(oMin),
                weekday_ot=min_to_h1(wMin), holiday_ot=min_to_h1(hMin),
                duty=min_to_h1(duty_min),
                per_dept_hours=per_dept,
                iv_target=iv_target, iv_viol=iv_viol,
                iv_short=min_to_h1(iv_short_min), iv_nopunch=iv_nopunch,
                n_duty=n_duty, n_oncall=n_oncall, n_workday=n_workday,
                n_off=n_off, max_streak=max_streak,
            ))
            if i % 15 == 0 or i == total:
                log(f"      走査 {i}/{total} 件（{i * 100 // total}%）／ 集計対象 {len(out)}名")
        if el_count > 0:
            log(f"      → 早上がり控除（救急/小児・平日日中の不在／当直加算が上限）: "
                f"{el_count}名 ／ 合計 {el_total_min / 60.0:.1f} h")
            if el_big:
                log(f"      ⚠ 控除候補が大きい人（要確認・日中打刻が少ない/夜間勤務の可能性）: {len(el_big)}名")
                for nm, dp, raw_h, eff_h in sorted(el_big, key=lambda x: -x[2])[:8]:
                    log(f"          {nm}（{dp}）: 控除候補 {raw_h:.0f}h → 実控除 {eff_h:.0f}h（当直加算上限）")
        iv_v = [p for p in out if p.iv_target and p.iv_viol > 0]
        iv_np = [p for p in out if p.iv_target and p.iv_nopunch > 0]
        if iv_v:
            mx = max(p.iv_short for p in iv_v)
            log(f"      → 勤務間インターバル不足（専攻医/初期研）: {len(iv_v)}名 ／ 最大不足 {mx:.1f} h"
                + (f" ／ ⚠打刻なし勤務日あり {len(iv_np)}名（判定精度↓）" if iv_np else ""))
        return out, warns
    finally:
        wb.close()


def scan_beacon(src_xlsm: str, log=lambda s: None) -> list[BeaconRow]:
    """全人別シートを走査し、「置きっぱ日数」＝在席(退勤−出勤)≥BEACON_PRESENCE_H なのに その日に
    時間外/自己研鑽の申告が無い日数 が BEACON_MIN_DAYS 日以上の人を「ビーコン置きっぱなし疑い」として返す。
    集計対象(残業実績あり)に限らず全員が対象＝申告ゼロの放置も拾う。終日いた事になっているのに勤務の
    申告が無い＝矛盾。本物の24h当直は申告があるので除外される（1〜2日は申告漏れの可能性→閾値で除外）。"""
    out: list[BeaconRow] = []
    wb = openpyxl.load_workbook(src_xlsm, read_only=True, data_only=True)
    try:
        for sname in wb.sheetnames:
            if sname == "名簿":
                continue
            ws = _ws_full(wb[sname])
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < HDR_ROW + 1:
                continue

            def gv(rr1, cc1):
                row = rows[rr1 - 1] if 0 <= rr1 - 1 < len(rows) else ()
                return row[cc1 - 1] if 0 <= cc1 - 1 < len(row) else None

            pName = gv(2, 1)
            if pName is None or str(pName).strip() == "":
                continue
            pName = str(pName).strip()
            deptRaw = gv(1, 1)
            deptRaw = str(deptRaw).strip() if deptRaw is not None else ""
            position = normalize_position(gv(1, 3))
            hdr = list(rows[HDR_ROW - 1])

            def col_of(name):
                try:
                    return hdr.index(name) + 1
                except ValueError:
                    return 0

            cDateCol = col_of("日付"); cIn = col_of("出勤"); cOut = col_of("退勤")
            if cDateCol == 0 or cIn == 0 or cOut == 0:
                continue
            cAS = [col_of(f"時間外開始{s}") for s in range(1, MAX_SLOTS + 1)]
            cKS = [col_of(f"自己研鑽開始{s}") for s in range(1, MAX_SLOTS + 1)]

            ivs = []                # 在席区間（出勤〜退勤・翌日跨ぎ補正済）
            day_dates = []          # シートにある暦日
            apply_dates = set()     # 時間外/自己研鑽の申告がある日
            n_apply = 0
            for r in range(DATA_START, len(rows) + 1):
                row = rows[r - 1]
                dd = cell_date(row[cDateCol - 1] if cDateCol - 1 < len(row) else None)
                if dd is None:
                    continue
                day_dates.append(dd)
                ap = False
                for cs in cAS + cKS:
                    if cs > 0 and cs - 1 < len(row) and row[cs - 1] is not None and str(row[cs - 1]).strip() != "":
                        ap = True
                        break
                if ap:
                    n_apply += 1
                    apply_dates.add(dd)
                iv = row[cIn - 1] if cIn - 1 < len(row) else None
                ov = row[cOut - 1] if cOut - 1 < len(row) else None
                if iv not in (None, "") and ov not in (None, ""):
                    ks = time_on_date(iv, dd); ke = time_on_date(ov, dd)
                    if ks is not None and ke is not None:
                        if ke < ks:
                            ke = ke + timedelta(days=1)
                        ivs.append((ks, ke))
            # 在席区間を結合（深夜0時の分割記録＝小ギャップを橋渡し）し、各暦日のカバレッジを算出
            merged = _merge_intervals(ivs, tol_min=30) if ivs else []
            beacon_days = 0; n_placed = 0
            for dd in day_dates:
                d0 = datetime(dd.year, dd.month, dd.day)
                d1 = d0 + timedelta(days=1)
                cov_min = 0
                for (s, e) in merged:
                    cov_min += overlap_minutes(s, e, d0, d1)
                if cov_min / 60.0 >= BEACON_COVERAGE_H:
                    n_placed += 1
                    if dd not in apply_dates:    # 終日在席なのに勤務の申告が無い＝置きっぱ
                        beacon_days += 1
            if beacon_days >= BEACON_MIN_DAYS:
                out.append(BeaconRow(name=pName, dept=deptRaw, position=position,
                                     beacon_days=beacon_days, n_placed=n_placed, n_apply=n_apply))
    finally:
        wb.close()
    out.sort(key=lambda b: (-b.beacon_days, b.n_apply))
    if out:
        nms = "、".join(f"{b.name}({b.dept}・置きっぱ{b.beacon_days}日/申告{b.n_apply}日)" for b in out)
        log(f"      ⚠ ビーコン置きっぱなし疑い {len(out)}名"
            f"（終日在席[カバレッジ≥{BEACON_COVERAGE_H:.0f}h]且つ申告なしが{BEACON_MIN_DAYS}日+）: {nms}")
    return out


def _dept_leaf(s: str) -> str:
    """所属名履歴（階層含む・"/科名/" 等）から末端の科名を取り出す。"""
    s = str(s or "").strip()
    for sep in (" > ", " ＞ ", "＞", ">", "/", "／"):
        if sep in s:
            parts = [x.strip() for x in s.split(sep) if x.strip()]
            if parts:
                return parts[-1]
    return s


def scan_beacon_v2(pay_xlsx: str, log=lambda s: None) -> list[BeaconRowV2]:
    """給与連携項目一覧を直読みし、Dr.JOY の退勤自動クローズ印からビーコン置きっ放しを検出する。
    （20260623。常時受信期の正式版。旧 scan_beacon〈在席カバレッジ22h〉を置換）

    背景: ビーコン受信機が院内に行き渡り検知が途切れなくなった結果、置きっ放しは
    「連日24h在席」ではなく Dr.JOY の退勤自動クローズ固定値に化け、22hカバレッジ方式は失明した。

    判定（医師ごとの予定始業時刻に非依存の構造判定。本物の当直明けは実退室時刻でバラつき非該当）:
      退勤が翌日にまたがり、かつ
        ① 退勤の翌日時刻 ＝ 翌日の出勤打刻時刻（始業アンカー。値は8:30/9:00/13:30…と可変） → sched 印
        ② 翌日に出勤打刻が無く 退勤＝翌23:59（予定未入力でアンカー無し・翌23:59が上限）        → no-sched 印
    数えるのは行数でなく「カバー暦日数」（②は当日＋翌日の2暦日ぶん）。退勤は openpyxl で timedelta、
    翌日跨ぎは timedelta(days=1,…)＝24h以上で読める（read_only でも翌日成分は保持される）。"""
    if not pay_xlsx:
        return []
    wb = openpyxl.load_workbook(pay_xlsx, read_only=True, data_only=True)
    try:
        ws = _ws_full(wb.worksheets[0])
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            return []
        hdr = [str(x).strip() if x is not None else "" for x in rows[0]]

        def col(n):
            try:
                return hdr.index(n)
            except ValueError:
                return -1

        def col_like(*needles):
            """ヘッダの表記ゆれ（全角/半角括弧など）に強い緩い列検索。"""
            for i, h in enumerate(hdr):
                if all(nd in h for nd in needles):
                    return i
            return -1

        cName = col("氏名"); cCode = col("職員コード"); cJob = col("職名")
        cDept = col("所属名履歴（階層含む）"); cDate = col("日付")
        cIn = col("出勤"); cOut = col("退勤")
        cToutType = col_like("打刻種別", "退勤")    # '打刻種別（退勤)'＝括弧が半角の表記ゆれあり→緩く検索
        cReason = [col(f"時間外理由{k}") for k in range(1, MAX_SLOTS + 1)]
        cKen = [col(f"研鑽申請種類名{k}") for k in range(1, MAX_SLOTS + 1)]
        if cName < 0 or cDate < 0 or cIn < 0 or cOut < 0:
            log("      ⚠ scan_beacon_v2: 必須列（氏名/日付/出勤/退勤）が見つからず中断")
            return []
        if cToutType < 0:
            log("      ⚠ scan_beacon_v2: 「打刻種別（退勤）」列なし→手入力除外なしで集計（旧エクスポート互換）")

        def ne(v):
            return v is not None and str(v).strip() != ""

        def secs(v):
            return int(round(v.total_seconds())) if isinstance(v, timedelta) else None

        # 1) 人ごとに (日付→出勤秒)/(日付→退勤秒)/申告日/メタ を集める
        persons: dict = {}
        for r in rows[1:]:
            nm = r[cName] if cName < len(r) else None
            if not ne(nm):
                continue
            dd = cell_date(r[cDate]) if cDate < len(r) else None
            if dd is None:
                continue
            code = r[cCode] if (0 <= cCode < len(r)) else None
            key = (str(code).strip() if ne(code) else "") + "|" + str(nm).strip()
            p = persons.get(key)
            if p is None:
                p = persons[key] = {"name": str(nm).strip(), "job": "", "dept": "",
                                    "in_by_date": {}, "out_by_date": {}, "outtype_by_date": {}, "apply": set()}
            if 0 <= cJob < len(r) and ne(r[cJob]):
                p["job"] = str(r[cJob]).strip()
            if 0 <= cDept < len(r) and ne(r[cDept]):
                p["dept"] = str(r[cDept]).strip()
            iv = secs(r[cIn]) if cIn < len(r) else None
            ov = secs(r[cOut]) if cOut < len(r) else None
            if iv is not None:
                p["in_by_date"][dd] = iv
            if ov is not None:
                p["out_by_date"][dd] = ov
            if 0 <= cToutType < len(r) and ne(r[cToutType]):
                p["outtype_by_date"][dd] = str(r[cToutType]).strip()
            if any(0 <= c < len(r) and ne(r[c]) for c in cReason + cKen):
                p["apply"].add(dd)

        # 2) 人ごとに判定。置きっぱ日数＝各日の24:00時点で置きっ放し状態か（自動クローズの連続検知が
        #    深夜をまたぐ）で数える。連日置きっ放しは退勤23:59/翌出勤00:00 に分割記録されるので
        #    _merge_intervals(tol30) で橋渡しし、間の深夜も「またいでいる」と判定（前後の状況を考慮）。
        out: list[BeaconRowV2] = []
        have_type = cToutType >= 0
        for p in persons.values():
            if BEACON_EXCLUDE_JOB in p["job"]:              # 非常勤医師は置きっ放し表から除外
                continue
            ivs = []                                        # 自動クローズ(ビーコン)セッション区間
            marker_rows = 0; beacon_days = 0; has_sched = False; has_nosched = False
            for dd, ov in p["out_by_date"].items():
                otype = p["outtype_by_date"].get(dd)
                is_auto = False
                if ov >= 24 * 3600:                          # 退勤が翌日にまたがる
                    tod = ov - 24 * 3600                     # 退勤の「翌日の時刻」（秒）
                    nxt = p["in_by_date"].get(dd + timedelta(days=1))
                    is_sched = nxt is not None and abs(tod - nxt) <= BEACONV2_HANDOFF_TOL
                    is_nosched = nxt is None and ov >= BEACONV2_EOD_MIN
                    # 退勤種別＝ビーコンの自動クローズだけ置きっ放し。手入力/編集は病院公認→除外。
                    if (is_sched or is_nosched) and ((not have_type) or otype == BEACON_PUNCH_TYPE):
                        is_auto = True
                        marker_rows += 1
                        has_sched = has_sched or is_sched
                        has_nosched = has_nosched or is_nosched
                        iv = p["in_by_date"].get(dd)
                        s = datetime(dd.year, dd.month, dd.day) + timedelta(seconds=(iv if iv is not None else 0))
                        e = datetime(dd.year, dd.month, dd.day) + timedelta(seconds=ov)
                        ivs.append((s, e))
                if not is_auto and ((not have_type) or otype == BEACON_PUNCH_TYPE):
                    beacon_days += 1                         # 正常にビーコンで退勤を記録した日（稼動）
            # 24:00判定：自動クローズ区間をmergeし、各日の24:00(=翌0:00)が区間内なら置きっ放し+1
            merged = _merge_intervals(ivs, tol_min=30) if ivs else []
            leave_days = 0
            for (s, e) in merged:
                m = datetime(s.year, s.month, s.day) + timedelta(days=1)   # s当日の24:00
                while m < e:
                    if m > s:
                        leave_days += 1
                    m += timedelta(days=1)
            if leave_days <= 0:
                continue
            kind = "mixed" if (has_sched and has_nosched) else ("no-sched" if has_nosched else "sched")
            # 手入力日数＝退勤をアプリ/WEB/打刻機/編集で記録した日数（種別≠ビーコン・月内。病院公認）。
            manual_days = sum(1 for t in p["outtype_by_date"].values() if t and t != BEACON_PUNCH_TYPE)
            out.append(BeaconRowV2(
                name=p["name"],
                dept=_dept_leaf(p["dept"]) or "(所属不明)",
                position=normalize_position(p["job"]),
                leave_days=leave_days,
                beacon_days=beacon_days,
                marker_rows=marker_rows,
                kind=kind,
                n_apply=len(p["apply"]),
                manual_days=manual_days,
            ))
    finally:
        wb.close()
    out = [b for b in out if b.leave_days >= BEACONV2_MIN_DAYS]
    out.sort(key=lambda b: (-b.leave_days, b.n_apply))
    if out:
        n_ns = sum(1 for b in out if b.kind == "no-sched")
        nms = "、".join(f"{b.name}({b.dept}・置きっぱ{b.leave_days}日/申請{b.n_apply}日"
                       + ("・予定未入力" if b.kind == "no-sched" else "") + ")" for b in out[:12])
        log(f"      ⚠ ビーコン置きっぱなし疑い {len(out)}名"
            f"（自動クローズ印・カバー暦日≥{BEACONV2_MIN_DAYS}日／うち予定未入力{n_ns}名）: {nms}"
            + ("…ほか" if len(out) > 12 else ""))
    return out


def scan_no_claim(pay_xlsx: str, log=lambda s: None) -> list[dict]:
    """給与連携項目一覧から「出退勤(出勤 or 退勤)はあるが、時間外/自己研鑽の申告が一切ない」医師を返す。
    手当計算の Excelはこの層を当直情報シート化しない(実績ゼロ除外)ので給与連携を直接読む。各 dict: name/dept/position。
    （20260609）"""
    if not pay_xlsx:
        return []
    out: list[dict] = []
    wb = openpyxl.load_workbook(pay_xlsx, read_only=True, data_only=True)
    try:
        ws = _ws_full(wb.worksheets[0])
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            return []
        hdr = [str(x).strip() if x is not None else "" for x in rows[0]]

        def col(n):
            try:
                return hdr.index(n)
            except ValueError:
                return -1

        cName = col("氏名"); cCode = col("職員コード"); cJob = col("職名")
        cDept = col("所属名履歴（階層含む）")
        cReason = [col(f"時間外理由{k}") for k in range(1, 11)]
        cKen = [col(f"研鑽申請種類名{k}") for k in range(1, 11)]
        cIn = col("出勤"); cOut = col("退勤")
        if cName < 0:
            return []

        def ne(v):
            return v is not None and str(v).strip() != ""

        persons = {}
        for r in rows[1:]:
            nm = r[cName] if cName < len(r) else None
            if not ne(nm):
                continue
            code = r[cCode] if (0 <= cCode < len(r)) else None
            key = (str(code).strip() if ne(code) else "") + "|" + str(nm).strip()
            p = persons.setdefault(key, {"name": str(nm).strip(), "job": "", "dept": "",
                                         "apply": False, "punch": False})
            if 0 <= cJob < len(r) and ne(r[cJob]):
                p["job"] = str(r[cJob]).strip()
            if 0 <= cDept < len(r) and ne(r[cDept]):
                p["dept"] = str(r[cDept]).strip()
            if not p["apply"] and any(0 <= c < len(r) and ne(r[c]) for c in cReason + cKen):
                p["apply"] = True
            if not p["punch"]:
                iv = r[cIn] if 0 <= cIn < len(r) else None
                ov = r[cOut] if 0 <= cOut < len(r) else None
                if ne(iv) or ne(ov):
                    p["punch"] = True
        for p in persons.values():
            if p["punch"] and not p["apply"]:
                out.append({"name": p["name"],
                            "dept": _dept_leaf(p["dept"]) or "(所属不明)",
                            "position": normalize_position(p["job"])})
    finally:
        wb.close()
    if out:
        log(f"      → 出退勤あり・申告ゼロ: {len(out)}名（科別・職位順スライドに掲載）")
    return out


# ============================================================
# 集計
# ============================================================
def aggregate_dept_fte(persons: list[PersonRow],
                       day_count: dict,
                       cal_days: int) -> list[DeptStat]:
    """科ごとに 実人数/FTE/総時間/平均/サンプル(箱ひげ用)"""
    bucket: dict[str, DeptStat] = {}
    for p in persons:
        # 20261104 weight = share of this person's days spent in each department (sums to 1 per person).
        # Falls back to the share of hours when the payroll export has no day counts for these departments.
        pdays = {dd: (day_count.get((p.name_norm, dd), 0) if day_count else 0) for dd in p.per_dept_hours}
        sum_days = sum(pdays.values())
        sum_h = sum(v for v in p.per_dept_hours.values() if v > 0)
        ndept = len(p.per_dept_hours) or 1
        days_ok = sum_days > 0 and all(pdays[dd] > 0 for dd, hh in p.per_dept_hours.items() if hh > 0)
        for dept, h in p.per_dept_hours.items():
            if days_ok:
                w = pdays[dept] / sum_days
            elif sum_h > 0:
                w = max(h, 0) / sum_h
            else:
                w = 1.0 / ndept
            if w > 0:
                if dept not in bucket:
                    bucket[dept] = DeptStat(name=dept, hc=0, fte=0.0, total=0.0, avg=0.0, samples=[])
                bucket[dept].wsamples.append((h / w, w, pdays[dept], h))
        for dept, h in p.per_dept_hours.items():
            if dept not in bucket:
                bucket[dept] = DeptStat(name=dept, hc=0, fte=0.0, total=0.0, avg=0.0, samples=[])
            d = bucket[dept]
            d.hc += 1
            d.total += h
            d.samples.append(h)
            # FTE
            days = day_count.get((p.name_norm, dept), 0) if day_count else 0
            if days > 0 and cal_days > 0:
                fte_add = min(days / cal_days, 1.0)
            else:
                fte_add = 1.0
            d.fte += fte_add
    # 平均
    for d in bucket.values():
        d.avg = (d.total / d.fte) if d.fte > 0 else 0.0
    # 平均降順
    result = sorted(bucket.values(), key=lambda x: -x.avg)
    return result


def aggregate_position(persons: list[PersonRow]) -> list[PosStat]:
    """職位別を固定順で（人数/平均/中央値/最大/最小/80h超え）"""
    bucket: dict[str, list[float]] = {}
    for p in persons:
        bucket.setdefault(p.position, []).append(p.total)
    out: list[PosStat] = []
    for name in POSITION_ORDER:
        vals = bucket.get(name)
        if not vals:
            continue
        out.append(PosStat(
            name=name, n=len(vals),
            avg=round(sum(vals) / len(vals), 1),
            median=round(statistics.median(vals), 1),
            max=round(max(vals), 1),
            min=round(min(vals), 1),
            over80=sum(1 for v in vals if v >= 80),
        ))
    return out


def box_stats(samples: list[float]) -> Optional[dict]:
    """箱ひげ用 min/Q1/median/Q3/max"""
    n = len(samples)
    if n < 1:
        return None
    s = sorted(samples)
    mn, mx = s[0], s[-1]
    med = statistics.median(s)
    if n >= 4:
        q = statistics.quantiles(s, n=4, method="inclusive")
        q1, q3 = q[0], q[2]
    elif n >= 2:
        q1, q3 = s[0], s[-1]
    else:
        q1 = q3 = med
    return dict(min=mn, q1=q1, median=med, q3=q3, max=mx, n=n)


# ============================================================
# 出力ヘルパー
# ============================================================
_thin = Side(border_style="thin", color=CLR_BORDER)
_box = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

def style_header(ws, rng: str):
    for row in ws[rng]:
        for c in row:
            c.font = Font(bold=True, color=CLR_HEADER_FG, size=10)
            c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
            c.alignment = Alignment(horizontal="center", vertical="center")

def apply_border(ws, rng: str):
    for row in ws[rng]:
        for c in row:
            c.border = _box

def autofit(ws, cols: list[str], min_w: int = 8):
    for col in cols:
        max_len = min_w
        for cell in ws[col]:
            if cell.value is None: continue
            v = str(cell.value)
            # 半角換算（簡易）
            ln = sum(2 if ord(ch) > 127 else 1 for ch in v)
            max_len = max(max_len, ln + 2)
        ws.column_dimensions[col].width = min(max_len, 40)


# ============================================================
# シート: 個人別
# ============================================================
def write_person(wb: Workbook, persons: list[PersonRow]):
    ws = wb.create_sheet("個人別")
    ws.append(["科", "職位", "氏名", "時間外合計(h)", "診療手当あり(h)", "その他(h)", "当直分(h)", "平日時間外(h)", "休日時間外(h)"])
    style_header(ws, "A1:I1")
    rows = sorted(persons, key=lambda p: -p.total)
    for i, p in enumerate(rows, start=2):
        ws.cell(i, 1, p.dept_label)
        ws.cell(i, 2, p.position)
        ws.cell(i, 3, p.name)
        ws.cell(i, 4, p.total)
        ws.cell(i, 5, p.teate)
        ws.cell(i, 6, p.other)
        ws.cell(i, 7, p.duty)
        ws.cell(i, 8, p.weekday_ot)
        ws.cell(i, 9, p.holiday_ot)
        # 80h超えハイライト
        if p.total >= 80:
            for c in range(1, 10):
                ws.cell(i, c).fill = PatternFill("solid", fgColor=CLR_HOT_BG)
            ws.cell(i, 4).font = Font(bold=True, color=CLR_HOT_FG)
        elif (i % 2) == 0:
            for c in range(1, 10):
                ws.cell(i, c).fill = PatternFill("solid", fgColor=CLR_ZEBRA)
        # 職位フォント色
        clr = position_color(p.position)
        if clr:
            for c in range(1, 4):
                cell = ws.cell(i, c)
                if cell.font.color is None or not cell.font.bold:
                    cell.font = Font(color=clr, bold=cell.font.bold or False)
    # 書式
    nrows = len(rows) + 1
    for r in range(2, nrows + 1):
        for c in range(4, 10):
            ws.cell(r, c).number_format = "0.0"
    apply_border(ws, f"A1:I{nrows}")
    autofit(ws, ["A", "B", "C", "D", "E", "F", "G", "H", "I"])
    ws.freeze_panes = "A2"


# ============================================================
# シート: 科別（FTE平均横棒 + 箱ひげ）
# ============================================================
def write_dept(wb: Workbook, dept_stats: list[DeptStat]):
    ws = wb.create_sheet("科別")
    ws.append(["科", "実人数", "換算人数", "平均(h)"])
    style_header(ws, "A1:D1")
    shown = [d for d in dept_stats if d.name not in BOX_EXCLUDE_DEPTS]
    if not shown:
        return
    for i, d in enumerate(shown, start=2):
        ws.cell(i, 1, d.name)
        ws.cell(i, 2, d.hc)
        ws.cell(i, 3, round(d.fte, 2))
        ws.cell(i, 4, min_to_h1(d.avg * 60))
        if (i % 2) == 0:
            for c in range(1, 5):
                ws.cell(i, c).fill = PatternFill("solid", fgColor=CLR_ZEBRA)
    nrows = len(shown) + 1
    for r in range(2, nrows + 1):
        ws.cell(r, 3).number_format = "0.00"
        ws.cell(r, 4).number_format = "0.0"
    apply_border(ws, f"A1:D{nrows}")
    autofit(ws, ["A", "B", "C", "D"])

    # チャート用ラベル隠しシート
    wsL = wb.create_sheet("_chart_labels")
    wsL.append(["dept_label", "avg"])
    for i, d in enumerate(shown, start=2):
        wsL.cell(i, 1, f"{d.name} (n={d.hc})")
        wsL.cell(i, 2, min_to_h1(d.avg * 60))
    wsL.sheet_state = "hidden"

    # FTE平均横棒
    chart = BarChart()
    chart.type = "bar"
    chart.style = 2
    chart.title = "科別 平均時間外勤務時間（FTE換算）"
    chart.y_axis.title = None
    chart.x_axis.title = "時間 (h)"
    chart.legend = None
    data_ref = Reference(wsL, min_col=2, min_row=1, max_row=nrows)
    cats_ref = Reference(wsL, min_col=1, min_row=2, max_row=nrows)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.dLbls = DataLabelList(showVal=True)
    chart.width = 22
    chart.height = max(8, 0.55 * len(shown) + 4)
    ws.add_chart(chart, f"F1")

    # 箱ひげ（Stacked-Bar trick）
    write_dept_box(wb, ws, shown)


def write_dept_box(wb: Workbook, ws_dept, shown: list[DeptStat]):
    """箱ひげをStacked-Bar trickで作る。data: [label, base=min, q1-min, median-q1, q3-median, max-q3]"""
    wsB = wb.create_sheet("_box_data")
    wsB.append(["label", "base", "w_low", "q1_med", "med_q3", "w_high"])
    valid = []
    for d in shown:
        bx = box_stats(d.samples)
        if bx is None or bx["n"] < 3:
            continue
        valid.append((d, bx))
    if not valid:
        wsB.sheet_state = "hidden"
        return
    for i, (d, bx) in enumerate(valid, start=2):
        wsB.cell(i, 1, f"{d.name} (n={bx['n']})")
        wsB.cell(i, 2, bx["min"])
        wsB.cell(i, 3, bx["q1"] - bx["min"])
        wsB.cell(i, 4, bx["median"] - bx["q1"])
        wsB.cell(i, 5, bx["q3"] - bx["median"])
        wsB.cell(i, 6, bx["max"] - bx["q3"])
    wsB.sheet_state = "hidden"
    n = len(valid)
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "科別 時間外勤務 箱ひげ図（移動者は経た全科でカウント）"
    chart.x_axis.title = "時間 (h)"
    chart.legend = None
    cats_ref = Reference(wsB, min_col=1, min_row=2, max_row=n + 1)
    # 系列ごとに追加（base/whisker low/box low/box high/whisker high）
    for col, fill, transparent in [
        (2, None, True),    # base (透明)
        (3, "999999", False),  # whisker low（細い灰色）
        (4, "2B4C7E", False),  # box low（塗りつぶし）
        (5, "5078B0", False),  # box high
        (6, "999999", False),  # whisker high
    ]:
        ref = Reference(wsB, min_col=col, min_row=1, max_row=n + 1)
        chart.add_data(ref, titles_from_data=True)
    # 系列のスタイル: 1番目を透明に
    if chart.series:
        s0 = chart.series[0]
        s0.graphicalProperties = GraphicalProperties(noFill=True)
        s0.graphicalProperties.line.noFill = True
    chart.set_categories(cats_ref)
    chart.width = 22
    chart.height = max(8, 0.55 * n + 4)
    # 平均グラフの下に配置
    ws_dept.add_chart(chart, f"F{int(0.55 * len(shown) + 6) + 1 + 18}")


# ============================================================
# シート: 職位別
# ============================================================
def write_position(wb: Workbook, pos_stats: list[PosStat]):
    ws = wb.create_sheet("職位別")
    ws.append(["職位", "人数", "平均(h)", "中央値(h)", "最大(h)", "最小(h)", "80h超え"])
    style_header(ws, "A1:G1")
    for i, p in enumerate(pos_stats, start=2):
        ws.cell(i, 1, p.name)
        ws.cell(i, 2, p.n)
        ws.cell(i, 3, p.avg)
        ws.cell(i, 4, p.median)
        ws.cell(i, 5, p.max)
        ws.cell(i, 6, p.min)
        ws.cell(i, 7, p.over80)
        if (i % 2) == 0:
            for c in range(1, 8):
                ws.cell(i, c).fill = PatternFill("solid", fgColor=CLR_ZEBRA)
    nrows = len(pos_stats) + 1
    for r in range(2, nrows + 1):
        for c in range(3, 7):
            ws.cell(r, c).number_format = "0.0"
    apply_border(ws, f"A1:G{nrows}")
    autofit(ws, ["A", "B", "C", "D", "E", "F", "G"])

    # チャートラベル
    wsL_name = "_position_labels"
    if wsL_name in wb.sheetnames:
        wsL = wb[wsL_name]
    else:
        wsL = wb.create_sheet(wsL_name)
    wsL.append(["pos_label", "avg"])
    for i, p in enumerate(pos_stats, start=2):
        wsL.cell(i, 1, f"{p.name} (n={p.n})")
        wsL.cell(i, 2, p.avg)
    wsL.sheet_state = "hidden"

    chart = BarChart()
    chart.type = "col"
    chart.style = 2
    chart.title = "職位別 平均時間外勤務時間"
    chart.y_axis.title = "時間 (h)"
    chart.legend = None
    data_ref = Reference(wsL, min_col=2, min_row=1, max_row=nrows)
    cats_ref = Reference(wsL, min_col=1, min_row=2, max_row=nrows)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.dLbls = DataLabelList(showVal=True)
    chart.width = 18
    chart.height = 10
    ws.add_chart(chart, f"I1")


# ============================================================
# シート: 概要
# ============================================================
def write_overview(wb: Workbook, persons: list[PersonRow], month_label: str, cal_days: int):
    ws = wb.create_sheet("概要")
    # タイトル
    ws.merge_cells("A1:H1")
    ws["A1"] = f"時間外勤務統計  {month_label}"
    ws["A1"].font = Font(size=18, bold=True, color=CLR_PRIMARY)
    ws.merge_cells("A2:H2")
    ws["A2"] = "当院 医師・歯科医師"
    ws["A2"].font = Font(size=10, color="787878")

    n = len(persons)
    if n == 0:
        return
    totals = [p.total for p in persons]
    avg = sum(totals) / n
    med = statistics.median(totals)
    mx = max(totals)
    over80 = sum(1 for v in totals if v >= 80)
    over100 = sum(1 for v in totals if v >= 100)
    over155 = sum(1 for v in totals if v >= 155)
    weeks = cal_days / 7.0
    weekly_avg = avg / weeks if weeks > 0 else 0

    def card(row, col, title, value, color=CLR_PRIMARY):
        ws.cell(row, col, title).font = Font(size=9, color="787878")
        ws.cell(row + 1, col, value).font = Font(size=14, bold=True, color=color)

    card(4, 1, "対象人数", f"{n} 名")
    card(4, 3, "月あたり時間外", f"{avg:.1f} h")
    card(4, 5, "週あたり時間外", f"{weekly_avg:.1f} h")
    card(4, 7, "中央値", f"{med:.1f} h")
    card(4, 9, "最大", f"{mx:.1f} h")
    clr_warn = "C0504D"
    card(7, 1, "80h超え", f"{over80} 名", clr_warn if over80 > 0 else CLR_PRIMARY)
    card(7, 3, "100h超え", f"{over100} 名", clr_warn if over100 > 0 else CLR_PRIMARY)
    card(7, 5, "155h超え", f"{over155} 名", clr_warn if over155 > 0 else CLR_PRIMARY)

    # 列幅
    for col, w in [("A", 12), ("B", 14), ("C", 12), ("D", 14), ("E", 12), ("F", 14), ("G", 12), ("H", 14), ("I", 12), ("J", 14)]:
        ws.column_dimensions[col].width = w

    # ヒストグラム（20h刻み）
    bin_w = 20
    bins = int(mx / bin_w) + 1 if mx > 0 else 1
    counts = [0] * bins
    for v in totals:
        idx = min(int(v / bin_w), bins - 1)
        counts[idx] += 1
    wsH = wb.create_sheet("_hist")
    wsH.append(["range", "count"])
    for b in range(bins):
        wsH.cell(b + 2, 1, f"{b*bin_w}-{(b+1)*bin_w}")
        wsH.cell(b + 2, 2, counts[b])
    wsH.sheet_state = "hidden"
    chart = BarChart()
    chart.type = "col"
    chart.title = "時間外勤務 分布 (時間帯別人数)"
    chart.y_axis.title = "人数"
    chart.legend = None
    data_ref = Reference(wsH, min_col=2, min_row=1, max_row=bins + 1)
    cats_ref = Reference(wsH, min_col=1, min_row=2, max_row=bins + 1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.dLbls = DataLabelList(showVal=True)
    # 80h以上を赤に
    if chart.series:
        s = chart.series[0]
        for b in range(bins):
            if b * bin_w >= 80:
                dp = DataPoint(idx=b)
                dp.graphicalProperties = GraphicalProperties(solidFill=CLR_HOT_FG)
                s.dPt.append(dp)
    chart.width = 20
    chart.height = 10
    ws.add_chart(chart, "A12")

    # Top10
    top_row = 32
    ws.cell(top_row, 1, "時間外上位10名").font = Font(size=12, bold=True, color=CLR_PRIMARY)
    top_row += 1
    ws.cell(top_row, 1, "順位"); ws.cell(top_row, 2, "科"); ws.cell(top_row, 3, "氏名")
    ws.cell(top_row, 4, "職位"); ws.cell(top_row, 5, "時間外(h)")
    style_header(ws, f"A{top_row}:E{top_row}")
    rows = sorted(persons, key=lambda p: -p.total)[:10]
    for i, p in enumerate(rows, start=1):
        r = top_row + i
        ws.cell(r, 1, i); ws.cell(r, 2, p.dept_label); ws.cell(r, 3, p.name)
        ws.cell(r, 4, p.position); ws.cell(r, 5, p.total)
        ws.cell(r, 5).number_format = "0.0"
        if p.total >= 80:
            ws.cell(r, 5).font = Font(bold=True, color=CLR_HOT_FG)
        if i % 2 == 0:
            for c in range(1, 6):
                ws.cell(r, c).fill = PatternFill("solid", fgColor=CLR_ZEBRA)
    apply_border(ws, f"A{top_row}:E{top_row + len(rows)}")


# ============================================================
# シート: 要確認
# ============================================================
def write_warn(wb: Workbook, warns: list[WarnEntry]):
    if not warns:
        return
    ws = wb.create_sheet("要確認")
    ws.append(["重要度", "区分", "処理", "氏名/キー", "科名", "対象項目", "原因候補", "元シート", "行"])
    style_header(ws, "A1:I1")
    for i, w in enumerate(warns, start=2):
        ws.cell(i, 1, w.severity); ws.cell(i, 2, w.category); ws.cell(i, 3, w.proc)
        ws.cell(i, 4, w.key); ws.cell(i, 5, w.dept); ws.cell(i, 6, w.item)
        ws.cell(i, 7, w.reason); ws.cell(i, 8, w.sheet)
        if w.row > 0: ws.cell(i, 9, w.row)
        fill = CLR_WARN_HIGH if w.severity == "高" else (CLR_WARN_MID if w.severity == "中" else None)
        if fill:
            for c in range(1, 10):
                ws.cell(i, c).fill = PatternFill("solid", fgColor=fill)
    n = len(warns) + 1
    apply_border(ws, f"A1:I{n}")
    autofit(ws, ["A", "B", "C", "D", "E", "F", "G", "H", "I"])
    ws.auto_filter.ref = f"A1:I{n}"


# ============================================================
# メイン実行（フォルダを受け取って出力xlsxパスを返す）
# ============================================================
def run(folder: str, log=lambda s: None) -> str:
    t0 = datetime.now()
    el = lambda: (datetime.now() - t0).total_seconds()
    log(f"対象フォルダ: {folder}")
    src = find_latest(folder, "*_当直情報*.xlsm")
    if not src:
        # フォールバック
        src = find_latest(folder, "出退勤入力情報_DrJOY*.xlsm")
    if not src:
        raise FileNotFoundError("出退勤入力情報_DrJOY_*_当直情報.xlsm が見つかりません")
    log(f"  勤務データ : {os.path.basename(src)}")
    pay = find_latest(folder, "*給与連携項目一覧*.xlsx")
    if pay:
        log(f"  所属データ : {os.path.basename(pay)}")
    else:
        log("  所属データ : 見つからない（日割り集計なし／A1フォールバック）")
    az14 = find_latest(folder, "*DTX14*.xlsm")
    log(f"  当直データ : {os.path.basename(az14) if az14 else '見つからない（管理当直の控除なし）'}")

    month_label, cal_days = detect_month(src)
    log(f"  対象月　　 : {month_label}（{cal_days}日）")
    log("────────────────────────────")

    log("[1/4] 所属マップを構築中（勤務日ごとの科）...")
    dept_by_date, day_count, warns_map = build_dept_by_date(pay)
    log(f"      → (氏名×日付) {len(dept_by_date):,} 件 ／ (氏名×科) {len(day_count):,} 件")

    log("[2/4] 管理当直マップを構築中（DTX14『人別当直日』）...")
    bucho_windows, warn_bucho = build_bucho_windows(folder)
    if warn_bucho is not None:
        log(f"      → 取得不可（{warn_bucho.reason}）→「要確認」に記録")
    else:
        log(f"      → 管理当直 {sum(len(v) for v in bucho_windows.values())} 件 ／ {len(bucho_windows)} 名"
            f"（救急/小児は当直加算から除外）")

    log(f"[3/4] 人別シートを集計中...　[{el():.0f}秒経過]")
    persons, warns_collect = collect_person_rows(src, dept_by_date, bucho_windows=bucho_windows,
                                                 cal_days=cal_days, log=log)
    # 20261016 時間外を Dr.JOY「勤務実績」の時間外労働時間(残業25%+法定休日35%)に差し替え（管理者GO）
    apply_drjoy_overtime(persons, folder, log=log, month_label=month_label, day_count=day_count)
    n80 = sum(1 for p in persons if p.total > 80)
    n100 = sum(1 for p in persons if p.total > 100)
    n155 = sum(1 for p in persons if p.total > 155)
    avg = (sum(p.total for p in persons) / len(persons)) if persons else 0.0
    log(f"      → 集計対象 {len(persons)} 名（実績ゼロを除外）／ 平均 {avg:.1f} h/月")
    log(f"      → 80h超 {n80}名 ／ 100h超 {n100}名 ／ 155h超 {n155}名")
    if persons:
        top = max(persons, key=lambda p: p.total)
        log(f"      → 最多: {top.name}（{top.dept_label}） {top.total} h")
    nwarn = len(warns_map) + len(warns_collect) + (1 if warn_bucho is not None else 0)
    log(f"      → 「要確認」 {nwarn} 件")

    log("      ビーコン置きっぱなし疑いを走査中（給与連携・自動クローズ印方式 v2）...")
    beacon_rows = scan_beacon_v2(pay, log=log)
    no_claim_rows = scan_no_claim(pay, log=log)

    log(f"[4/4] 配布PowerPoint を作成中（集計＋グラフ描画）...　[{el():.0f}秒経過]")
    dept_stats = aggregate_dept_fte(persons, day_count, cal_days)
    log(f"      → 科別 {len(dept_stats)}科 を集計")
    import make_pptx_boxplot
    out_path = make_pptx_boxplot.build_deck(folder, month_label, persons, dept_stats,
                                            beacon_rows=beacon_rows, no_claim_rows=no_claim_rows, log=log)
    log(f"      → 配布PowerPoint 完了: {os.path.basename(out_path)}　[{el():.0f}秒経過]")

    log("────────────────────────────")
    log(f"✅ すべて完了しました（合計 {el():.0f} 秒）")
    return out_path


# ============================================================
# GUI (tkinter)
# ============================================================
def gui_main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("時間外勤務統計 (DTX_18 Python版)")
    root.geometry("680x500")

    # デフォルトフォルダ: exeまたは.pyのあるディレクトリ
    default_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()

    frm_top = ttk.Frame(root, padding=10)
    frm_top.pack(fill="x")
    ttk.Label(frm_top, text="対象フォルダ:").pack(side="left")
    var_folder = tk.StringVar(value=default_dir)
    ent = ttk.Entry(frm_top, textvariable=var_folder, width=60)
    ent.pack(side="left", padx=4, fill="x", expand=True)
    def browse():
        d = filedialog.askdirectory(initialdir=var_folder.get() or default_dir, title="対象フォルダを選択")
        if d:
            var_folder.set(d)
    ttk.Button(frm_top, text="参照...", command=browse).pack(side="left")

    frm_mid = ttk.Frame(root, padding=(10, 0))
    frm_mid.pack(fill="x")
    btn_run = ttk.Button(frm_mid, text="統計を作成")
    btn_run.pack(side="left")
    pb = ttk.Progressbar(frm_mid, mode="indeterminate", length=260)
    pb.pack(side="left", padx=8)

    frm_log = ttk.LabelFrame(root, text="ログ", padding=8)
    frm_log.pack(fill="both", expand=True, padx=10, pady=10)
    log_widget = scrolledtext.ScrolledText(frm_log, height=18, state="disabled", wrap="word")
    log_widget.pack(fill="both", expand=True)

    q: queue.Queue = queue.Queue()

    def log_to_widget(msg: str):
        log_widget.configure(state="normal")
        log_widget.insert("end", msg + "\n")
        log_widget.see("end")
        log_widget.configure(state="disabled")

    def drain():
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == "log":
                    log_to_widget(payload)
                elif kind == "done":
                    pb.stop()
                    btn_run.config(state="normal")
                    out_path = payload
                    log_to_widget(f"=== 完了: {out_path} ===")
                    try:
                        os.startfile(out_path)
                    except Exception as e:
                        log_to_widget(f"ファイルを開けません: {e}")
                    messagebox.showinfo("完了", f"配布PowerPoint を作成しました\n{out_path}")
                elif kind == "error":
                    pb.stop()
                    btn_run.config(state="normal")
                    log_to_widget(payload)
                    messagebox.showerror("エラー", payload.splitlines()[0] if payload else "不明なエラー")
        except queue.Empty:
            pass
        root.after(150, drain)

    def worker(folder: str):
        try:
            out = run(folder, log=lambda s: q.put(("log", s)))
            q.put(("done", out))
        except Exception:
            q.put(("error", traceback.format_exc()))

    def on_run():
        folder = var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("エラー", "対象フォルダを選んでください")
            return
        log_to_widget("--------")
        btn_run.config(state="disabled")
        pb.start(12)
        threading.Thread(target=worker, args=(folder,), daemon=True).start()

    btn_run.config(command=on_run)
    root.after(150, drain)

    # ダブルクリック起動時の自動開始: exe(または.py)のあるフォルダに勤務データがあれば
    # 起動直後に作成を始める（別フォルダを集計したい場合は「参照...」で選び直して再実行）。
    def _autostart():
        folder = var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            return
        if find_latest(folder, "*_当直情報*.xlsm") or find_latest(folder, "出退勤入力情報_DrJOY*.xlsm"):
            log_to_widget(f"フォルダ「{folder}」を自動で集計します（停止する場合はウィンドウを閉じてください）")
            on_run()
        else:
            log_to_widget("このフォルダに勤務データが無いため自動開始しません。「参照...」でフォルダを選び「統計を作成」を押してください。")
    root.after(500, _autostart)

    root.mainloop()


# ============================================================
# エントリポイント
# ============================================================
if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] not in ("--gui",):
        # CLIモード: python dtx18_stats.py <folder>
        folder = sys.argv[1]
        out = run(folder, log=print)
        print(f"OUT: {out}")
    else:
        gui_main()
