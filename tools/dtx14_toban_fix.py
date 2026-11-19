# -*- coding: utf-8 -*-
"""dtx14_toban_fix.py — DTX14 の「当番当直」シートの1日ぶんを直す（休日→平日など）。20261008

例: 2026-12-28 は当番当直シートで「休日-昼／休日-夜」になっていたが通常の開院日（管理者 20260922）。
    昼の行を消し、夜の区分を「平日-夜」にする。
    python -X utf8 tools\\dtx14_toban_fix.py --file <DTX14.xlsm> --date 2026-12-28 --weekday
    --weekday … その日を平日にする（休日-昼 の行を削除、休日-夜 → 平日-夜）
    --holiday … その日を休日にする（平日-夜 → 休日-夜。昼の行は --day-dept 科名 で足す）
    --add-holidays data/holidays.csv [--until 2027-01-03] … 「休日判定用」シートに未登録の休日（開院記念日・年末年始など）を足す
    --add-toban data/toban.csv --year 2027 … 「当番当直」シートの最後に、その年の当番当直を足す（その年の行が1つでもあれば断る。20261008）

Excel COM でブックを開いて（マクロは走らせない）値だけ直し、上書き保存する。書く前に同じフォルダの
`_上書き履歴\\` へ「<名前>_<日付>当番当直修正前.xlsm」を写す（17時の反映・④の取り込みと同じ置き場。
以前の置き場「待避」は、20260929 に「【不要？】待避」へ改名して使わなくなった）。SharePoint 同期フォルダの原本を直すときに使う。
"""
import argparse
import datetime as dt
import os
import shutil

SHEET = "当番当直"
BACKUP_DIR = "_上書き履歴"   # same place as dtx14_daily_sync / dtx14_manual_import


def fix(path, date, weekday=True, day_dept="", log=print):
    import win32com.client
    path = os.path.abspath(path)
    bak_dir = os.path.join(os.path.dirname(path), BACKUP_DIR)
    os.makedirs(bak_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    bak = os.path.join(bak_dir, f"{stem}_{dt.date.today():%Y%m%d}当番当直修正前{ext}")
    if not os.path.exists(bak):
        shutil.copy2(path, bak)
        log(f"退避: {bak}")
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
        ws = wb.Worksheets(SHEET)
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row          # xlUp
        hits = []
        for r in range(2, last + 1):
            v = ws.Cells(r, 1).Value
            d = v.date() if hasattr(v, "date") else None
            if d == date:
                hits.append((r, str(ws.Cells(r, 3).Value or ""), str(ws.Cells(r, 4).Value or "")))
        log(f"{date}: {[(r, k, d) for r, k, d in hits]}")
        changed = 0
        if weekday:
            for r, kubun, dept in sorted(hits, reverse=True):
                if kubun.endswith("昼"):
                    ws.Rows(r).Delete()
                    log(f"  行{r} 『{kubun} {dept}』を削除")
                    changed += 1
                elif kubun == "休日-夜":
                    ws.Cells(r, 3).Value = "平日-夜"
                    log(f"  行{r} 休日-夜 → 平日-夜（{dept}）")
                    changed += 1
        else:
            for r, kubun, dept in hits:
                if kubun == "平日-夜":
                    ws.Cells(r, 3).Value = "休日-夜"
                    log(f"  行{r} 平日-夜 → 休日-夜（{dept}）")
                    changed += 1
            if day_dept and not any(k.endswith("昼") for _r, k, _d in hits):
                r = min(r for r, _k, _d in hits) if hits else last + 1
                ws.Rows(r).Insert(-4121)
                ws.Cells(r, 1).Value = float((date - dt.date(1899, 12, 30)).days)
                ws.Cells(r, 1).NumberFormat = ws.Cells(r + 1, 1).NumberFormat
                ws.Cells(r, 2).Value = ws.Cells(r + 1, 2).Value
                ws.Cells(r, 3).Value = "休日-昼"
                ws.Cells(r, 4).Value = day_dept
                log(f"  行{r} に 休日-昼 {day_dept} を追加")
                changed += 1
        if changed:
            wb.Save()
        wb.Close(SaveChanges=False)
        log(f"{'保存' if changed else '変更なし'}: {path}")
        return changed
    finally:
        xl.Quit()


WD_JP = "月火水木金土日"


def add_holidays(path, csv_path, until=None, log=print):
    """DTX14 の「休日判定用」シート（A=日付／B=曜日1文字／C=名称）に、data/holidays.csv のうち
    シートに無い日（最後の日付より後）を足す。土日に当たる祝日・年末年始も入れる（シートの従来の書き方。
    例 2026/5/3（日）憲法記念日）。戻り: 足した件数。"""
    import csv as _csv
    import win32com.client
    path = os.path.abspath(path)
    bak_dir = os.path.join(os.path.dirname(path), BACKUP_DIR)
    os.makedirs(bak_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    bak = os.path.join(bak_dir, f"{stem}_{dt.date.today():%Y%m%d}休日判定用追加前{ext}")
    if not os.path.exists(bak):
        shutil.copy2(path, bak)
        log(f"退避: {bak}")
    hol = []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in _csv.DictReader(f):
            if row.get("date"):
                d = dt.date.fromisoformat(row["date"].strip())
                if until is None or d <= until:
                    hol.append((d, (row.get("name") or "").strip()))
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
        ws = wb.Worksheets("休日判定用")
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        have = set()
        for r in range(1, last + 1):
            v = ws.Cells(r, 1).Value
            if hasattr(v, "date"):
                have.add(v.date())
        last_d = max(have) if have else dt.date(1900, 1, 1)
        fmt = ws.Cells(last, 1).NumberFormat
        added = 0
        for d, name in sorted(hol):
            if d in have or d <= last_d:
                continue
            r = last + 1 + added
            ws.Cells(r, 1).Value = float((d - dt.date(1899, 12, 30)).days)
            ws.Cells(r, 1).NumberFormat = fmt
            ws.Cells(r, 2).Value = WD_JP[d.weekday()]
            ws.Cells(r, 3).Value = name
            log(f"  行{r}: {d} {WD_JP[d.weekday()]} {name}")
            added += 1
        if added:
            wb.Save()
        wb.Close(SaveChanges=False)
        log(f"{'保存' if added else '変更なし'}（{added} 件）: {path}")
        return added
    finally:
        xl.Quit()


def add_toban(path, csv_path, year, log=print):
    """「当番当直」シート（A=日付／B=曜日「月曜」／C=区分 休日-昼・休日-夜・平日-夜／D=担当科）の最後に、
    csv（date,band,dept＝data/toban.csv の形）のうち year 年の行を足す。区分の休日は土日＋data/holidays.csv（contract.is_holiday）。
    シートにその年の行が1つでもあれば何もしない（二重に足さない）。書式は今の最後の行を写す。戻り: 足した件数。"""
    import csv as _csv
    import sys
    import win32com.client
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import contract as C
    hol = C.load_holidays()
    rows = []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in _csv.DictReader(f):
            d = dt.date.fromisoformat(row["date"].strip())
            if d.year != year:
                continue
            band, dept = row["band"].strip(), row["dept"].strip()
            kubun = "休日-昼" if band == "昼" else ("休日-夜" if C.is_holiday(d, hol) else "平日-夜")
            if band == "昼" and not C.is_holiday(d, hol):
                raise SystemExit(f"{d} は平日なのに昼の当番当直がある（{dept}）")
            rows.append((d, kubun, dept))
    if not rows:
        raise SystemExit(f"{csv_path} に {year} 年の行がない")
    path = os.path.abspath(path)
    bak_dir = os.path.join(os.path.dirname(path), BACKUP_DIR)
    os.makedirs(bak_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    bak = os.path.join(bak_dir, f"{stem}_{dt.date.today():%Y%m%d}当番当直{year}追加前{ext}")
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
        ws = wb.Worksheets(SHEET)
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row          # xlUp
        for r in range(2, last + 1):
            v = ws.Cells(r, 1).Value
            if hasattr(v, "year") and v.year == year:
                wb.Close(SaveChanges=False)
                log(f"変更なし: {year} 年の行がもうある（行{r}）: {path}")
                return 0
        if not os.path.exists(bak):
            shutil.copy2(path, bak)
            log(f"退避: {bak}")
        first, end = last + 1, last + len(rows)
        ws.Range(ws.Cells(last, 1), ws.Cells(last, 4)).Copy()
        ws.Range(ws.Cells(first, 1), ws.Cells(end, 4)).PasteSpecial(-4122)   # xlPasteFormats
        xl.CutCopyMode = False
        vals = [[float((d - dt.date(1899, 12, 30)).days), WD_JP[d.weekday()] + "曜", k, dept] for d, k, dept in rows]
        ws.Range(ws.Cells(first, 1), ws.Cells(end, 4)).Value = vals
        wb.Save()
        wb.Close(SaveChanges=False)
        log(f"保存（{len(rows)} 件・行{first}〜{end}）: {path}")
        return len(rows)
    finally:
        xl.Quit()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--date")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--weekday", action="store_true")
    g.add_argument("--holiday", action="store_true")
    ap.add_argument("--day-dept", default="")
    ap.add_argument("--add-holidays", metavar="CSV", help="休日判定用シートに holidays.csv の未登録の日を足す")
    ap.add_argument("--until", default="", help="--add-holidays で足す上限の日（例 2027-01-03。既定=次の1月3日）")
    ap.add_argument("--add-toban", metavar="CSV", help="当番当直シートの最後に、--year の年の当番当直を足す")
    ap.add_argument("--year", type=int)
    a = ap.parse_args()
    if a.add_toban:
        if not a.year:
            ap.error("--add-toban には --year が要ります")
        add_toban(a.file, a.add_toban, a.year)
        return
    if a.add_holidays:
        if a.until:
            until = dt.date.fromisoformat(a.until)
        else:
            t = dt.date.today()
            until = dt.date(t.year + 1, 1, 3)
        add_holidays(a.file, a.add_holidays, until)
        return
    if not (a.date and (a.weekday or a.holiday)):
        ap.error("--date と --weekday/--holiday、または --add-holidays を指定してください")
    fix(a.file, dt.date.fromisoformat(a.date), weekday=a.weekday, day_dept=a.day_dept)


if __name__ == "__main__":
    main()
