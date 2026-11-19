# -*- coding: utf-8 -*-
"""Before the monthly Dr.JOY bulk import (drjoy-touchoku-monthly, the 29th): is the month confirmed, and has the
duty-roster app rewritten DTX14 after the confirmation?  20260926

  python -X utf8 tools/dtx14_confirm_check.py --month 2026-11 [--dtx14 <path>] [--now 2026-10-29T17:40]

Months before R8.11 (事務担当の手作業の DTX14 の月) are out of scope: exit 0 with 「対象外」, and the old way applies
(管理者が決裁を確かめて GO). The app never presses ③確定 nor stamps AF1 for them (20260926).

Exit codes
  0  out of scope (before R8.11), or confirmed, and DTX14's 「17時更新」 stamp (AF1 of 当直表 ファースト) is at/after the confirmation → import
  3  not confirmed yet → do not import (look again tomorrow)
  4  confirmed, but DTX14 was not rewritten after the confirmation → do not import (the 17:00 run will do it)
  2  cannot tell (DTX14 missing / two candidates / stamp unreadable / 確定.json unreadable) → do not import, tell 管理者

Rules (Touchoku_Plan, 当直勤務表アプリ):
  * Confirmation = 当直勤務表\\R8年N月\\集約\\確定.json の confirmed（{at, by, revs} or null）. Only when there is no
    確定.json (months confirmed before 2026-09-25) fall back to master\\当直予定表_R8年N月.json の confirmed.
    If they disagree, 確定.json wins (master_store.effective_confirmed).
  * The app stamps 『R8.9.22』『17時更新』 into AF1:AG1 whenever it writes DTX14 (the hour it actually wrote).
    Treat the stamp as <date> <hour>:00 and require it to be at/after confirmed.at (a stamp in the same hour but
    before the confirmation is not accepted: safe side, the next day's run passes).
  * DTX14 = 当直勤務表_正本\\R8.N月 当直勤務表\\★R8.N月用 DTX14*.xlsm (same name rule as the app: after DTX14 only
    digits and dots; blanks are ignored).
"""
import argparse
import datetime as dt
import glob
import json
import os
import re
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _C():
    # collect_month_sources (CLERK_DIR, hydrate) re-wraps stdout when imported; import it only when needed
    import collect_month_sources as C
    return C

DUTY_ROOT = r"\\fileserver\share$\00_DrJOY\当直勤務表"
SHEET = "当直表 ファースト"
APP_FROM = (2026, 11)   # first month whose DTX14 the duty-roster app writes (SYNC_FROM; same as drjoy_kotai_sync.MIN_MONTH)


def out_of_scope(y, m):
    """True for months before R8.11 (事務担当の手作業の DTX14): no ③確定, no AF1 stamp; 管理者 checks and gives GO."""
    return (y, m) < APP_FROM


def reiwa(y):
    return y - 2018


def is_master_name(name, y, m):
    core = name.replace(" ", "").replace("　", "")
    return re.fullmatch(rf"★R{reiwa(y)}\.{m}月用DTX14[0-9.]*\.xlsm", core) is not None


def dtx14_candidates(y, m):
    d = os.path.join(_C().CLERK_DIR, "当直勤務表_正本", f"R{reiwa(y)}.{m}月 当直勤務表")
    return [os.path.join(d, f) for f in (os.listdir(d) if os.path.isdir(d) else []) if is_master_name(f, y, m)]


def confirmation(y, m, root=DUTY_ROOT):
    """(confirmed dict or None, where it came from). Raises ValueError when 確定.json exists but cannot be read."""
    md = os.path.join(root, f"R{reiwa(y)}年{m}月")
    truth = os.path.join(md, "集約", "確定.json")
    if os.path.exists(truth):
        try:
            with open(truth, encoding="utf-8") as f:
                t = json.load(f)
        except Exception as e:
            raise ValueError(f"確定.json が読めない: {e}") from e
        if not isinstance(t, dict) or "confirmed" not in t:
            raise ValueError("確定.json の形が違う（confirmed が無い）")
        return t.get("confirmed") or None, "集約\\確定.json"
    master = os.path.join(md, "master", f"当直予定表_R{reiwa(y)}年{m}月.json")
    if os.path.exists(master):
        with open(master, encoding="utf-8") as f:
            return (json.load(f).get("confirmed") or None), "master（確定.json が無い月）"
    return None, "（確定.json もマスターも無い）"


def parse_stamp(v):
    """『R8.9.22\\n17時更新』 → datetime(2026, 9, 22, 17). None when it does not look like the stamp."""
    if v is None:
        return None
    s = str(v).replace("\r", "\n")
    mt = re.search(r"R(\d+)\.(\d+)\.(\d+)\D+?(\d{1,2})時更新", s)
    if not mt:
        return None
    n, mo, d, h = map(int, mt.groups())
    try:
        return dt.datetime(2018 + n, mo, d, h)
    except ValueError:
        return None


def read_stamp(path):
    from openpyxl import load_workbook
    _C().hydrate(path)                   # OneDrive on-demand: read it once so the real file is local
    tmp = os.path.join(tempfile.gettempdir(), "_dtx14_stamp_check.xlsm")
    shutil.copy2(path, tmp)
    try:
        wb = load_workbook(tmp, read_only=True, data_only=True, keep_links=False)
        try:
            if SHEET not in wb.sheetnames:
                return None, f"シート「{SHEET}」が無い"
            v = wb[SHEET]["AF1"].value
            return parse_stamp(v), repr(v)
        finally:
            wb.close()
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def judge(confirmed, stamp):
    """(exit code, message) from the confirmation dict and the AF1 stamp datetime."""
    if not confirmed:
        return 3, "未確定（決裁待ち）。投入しない。明日もう一度見る"
    try:
        at = dt.datetime.fromisoformat(str(confirmed.get("at")))
    except ValueError:
        return 2, f"確定の時刻が読めない: {confirmed.get('at')!r}"
    if stamp is None:
        return 2, "DTX14 の AF1（17時更新）が読めない。管理者に知らせる"
    if stamp < at:
        return 4, (f"DTX14 は確定（{at:%m/%d %H:%M}）のあとに書き直されていない（AF1＝{stamp:%m/%d} {stamp:%H}時更新）。"
                   "投入しない。17時の反映のあと、明日もう一度見る")
    return 0, f"確定 {at:%Y/%m/%d %H:%M} ／ DTX14 の更新 {stamp:%Y/%m/%d} {stamp:%H}時 → 投入してよい"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM（投入する月）")
    ap.add_argument("--dtx14", default=None)
    a = ap.parse_args()
    y, m = map(int, a.month.split("-"))
    if out_of_scope(y, m):
        print(f"→ 対象外（R{reiwa(y)}.{m}月分は事務担当の手作業の DTX14 の月）。"
              "このチェックは使わず、R8.10月分までの決裁の確かめ方（管理者の GO）に従う")
        return 0
    try:
        conf, src = confirmation(y, m)
    except ValueError as e:
        print(f"★{e}")
        return 2
    print(f"確定: {conf.get('at') + ' ' + str(conf.get('by') or '') if conf else '未確定'}（{src}）")
    if not conf:
        print("→", judge(None, None)[1])
        return 3
    if a.dtx14:
        path = a.dtx14
    else:
        c = dtx14_candidates(y, m)
        if len(c) != 1:
            print(f"★DTX14 の正本が {len(c)} 本: {[os.path.basename(x) for x in c]}。管理者に知らせる")
            return 2
        path = c[0]
    stamp, raw = read_stamp(path)
    print(f"DTX14: {os.path.basename(path)}  AF1={raw}")
    code, msg = judge(conf, stamp)
    print(("→ " if code == 0 else "★") + msg)
    return code


if __name__ == "__main__":
    sys.exit(main())
