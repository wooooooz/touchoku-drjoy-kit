# -*- coding: utf-8 -*-
"""dtx14_split_rows.py — DTX14 の雛形で「全日」の1行しか無い診療科の行を 昼/夜 の2行に割る。20260922

管理者 2026-09-22 決定（8/31 の「全日の行を廃止」の具体化）:
  雛形は全枠が 昼＋夜 の2行。月ごとのタイプは書き込み側（dtx14_fill）がデータから決め、
  全日タイプなら上の行を「全日」にして下の行を隠す。だから雛形そのものは常に2行でよい。

R8.10 の雛形（★R8.10月用 DTX14.9.5.12.xlsm）では「当直表 ファースト」は既に全科が 昼/夜 の2行で、
「当直表 セカンド以下」に 全日の1行の科が 22（消滅した 内科 旧科 を含む）ある。
このスクリプトは、枠マスタ data/frames.json にある診療科の行だけを対象にし、
既に2行ある科・看護部・管理当直・薬剤部などは触らない。

やること（Excel COM・マクロは走らせない・値と書式だけ）:
  1) 対象の行のすぐ下に1行挿入し、上の行の書式と高さを写す（Rows.Copy → Insert）
  2) A列を上下で結合（既存の2行の科と同じ見た目）。C列は 上=「昼」・下=「夜」
  3) 日付セル（D列〜月末）を空にする（--keep-data で上の行の中身を残す）
  4) --roll-month YYYY-MM で A1（両シートの D3/D4 が =A1 で参照する月初シリアル）を差し替え、
     行5以降の日付セルを全部空にする（新しい月の雛形は空）

usage: python -X utf8 tools\\dtx14_split_rows.py --in <雛形.xlsm> --out <出力.xlsm>
                                               [--roll-month 2026-11] [--keep-data]
"""
import argparse
import calendar
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402
import dtx14_extract as X     # noqa: E402

XL_MACRO_ENABLED = 52
COL_DAY1 = 4
CELL_MONTH = (1, 1)           # A1 = 月初のシリアル（D3/D4 が =A1+0 で参照）
BLOCK_END = {X.SHEET_FIRST: 64, X.SHEET_SECOND: 40}   # 走査の上限（挿入で伸びる分は余裕をみる）
LABELS = {"昼", "夜", "全日", "深夜", "日勤", "準夜", "遅出"}   # C列の区分（この行だけが当直の行）


def target_rows(template):
    """{シート: [(行, 科名)]} 全日の1行しか無い診療科（枠マスタにある科）の行。"""
    frames = C.load_frames()
    out = {}
    for sheet in (X.SHEET_FIRST, X.SHEET_SECOND):
        rows, _ = X.load_sheet_grid(template, sheet)
        loc = X.locate_rows(rows, sheet)
        seen, lst = set(), []
        for f in frames:
            if f.get("sheet") != sheet or f.get("retired"):   # 消滅した科（内科 旧科）の隠し行は触らない
                continue
            dept = f.get("dept_dtx14") or f.get("dept") or ""
            k = X.namekey(X.nfc(dept))
            if k in seen:
                continue
            seen.add(k)
            _rows, two = X.dept_band_rows(dept, loc)
            if two or not _rows:
                continue
            r = _rows["昼"]
            if loc.get((k, "夜")) == r + 1:      # 前回の書き込みで畳んだ状態（上=全日・下=夜）は既に2行
                continue
            lst.append((r, dept))
        out[sheet] = sorted(lst)
    return out


def split(src, out, roll=None, keep_data=False, log=print):
    import win32com.client
    targets = target_rows(src)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(src), UpdateLinks=0, ReadOnly=True)
        xl.Calculation = -4135         # xlCalculationManual（ブックを開いてからでないと設定できない。保存前に自動へ戻す）
        summary = {}
        for sheet in (X.SHEET_FIRST, X.SHEET_SECOND):
            ws = wb.Worksheets(sheet)
            cur = ws.Cells(*CELL_MONTH).Value
            cur = cur.date() if hasattr(cur, "date") else dt.date(1899, 12, 30) + dt.timedelta(days=int(cur or 0))
            ndays = calendar.monthrange(cur.year, cur.month)[1]
            lastcol = COL_DAY1 + 31 - 1                     # 日付セルは最大31列ぶんを空にする
            done = []
            shapes0 = {sh.Name for sh in ws.Shapes}          # 行の Copy→Insert は行に載っている図形も複製するので、後で消す
            for r, dept in sorted(targets.get(sheet, []), reverse=True):   # 下から割れば行番号がずれない
                if str(ws.Cells(r, 1).Value or "").strip() != dept.strip():
                    raise SystemExit(f"{sheet} 行{r}: A列が『{ws.Cells(r, 1).Value}』で『{dept}』ではありません")
                ws.Rows(r).Copy()
                ws.Rows(r + 1).Insert(-4121)                # xlDown（挿入行は上の行の書式・高さを引き継ぐ）
                xl.CutCopyMode = False
                ws.Rows(r + 1).RowHeight = ws.Rows(r).RowHeight
                ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, lastcol)).ClearContents()
                ws.Cells(r + 1, 1).Value = None
                ws.Range(ws.Cells(r, 1), ws.Cells(r + 1, 1)).Merge()
                ws.Cells(r, 3).Value = "昼"
                ws.Cells(r + 1, 3).Value = "夜"
                ws.Rows(r + 1).Hidden = False
                if not keep_data:
                    ws.Range(ws.Cells(r, COL_DAY1), ws.Cells(r, lastcol)).ClearContents()
                done.append((r, dept))
            dup = [sh for sh in ws.Shapes if sh.Name not in shapes0]
            for sh in dup:                                   # 複製された図形（例: 外科 心臓血管 の電話番号の枠）
                sh.Delete()
            if dup:
                log(f"  {sheet}: 行の複製で増えた図形 {len(dup)} 個を消した")
            # ⚠ 複数行にまたがる装飾の図形（セカンド以下の 内科 血液〜内科C 初期研 を括る L 字）は Excel の
            #   「セルに合わせて移動やサイズ変更」でも伸びない。Top/Height を行に合わせて付け直しても PDF では
            #   1〜2行ずれて描かれた（20260922 実測）ので、ここでは触らず事務担当に手で直してもらう。
            summary[sheet] = sorted(done)
            if roll:
                y, m = roll
                ws.Cells(*CELL_MONTH).Value = float((dt.date(y, m, 1) - dt.date(1899, 12, 30)).days)
                if not keep_data:                            # 当直の行（C列に区分がある行）だけ空にする。医事の欄などは触らない
                    for r in range(5, BLOCK_END[sheet] + 31):
                        if str(ws.Cells(r, 3).Value or "").strip() in LABELS:
                            ws.Range(ws.Cells(r, COL_DAY1), ws.Cells(r, lastcol)).ClearContents()
        xl.Calculation = -4105          # xlCalculationAutomatic
        wb.SaveAs(os.path.abspath(out), FileFormat=XL_MACRO_ENABLED)
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    for sheet, lst in summary.items():
        log(f"{sheet}: {len(lst)} 科を 昼/夜 の2行に割った"
            + (f"（{', '.join(d for _, d in lst)}）" if lst else ""))
    if roll:
        log(f"月を {roll[0]}年{roll[1]}月 にした（A1 のシリアル。D3/D4 は =A1 で追従）")
    log(f"出力: {out}")
    return summary


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--roll-month", help="YYYY-MM。月初シリアルを差し替え、日付セルを全部空にする")
    ap.add_argument("--keep-data", action="store_true", help="上の行の中身（全日として書かれたもの）を残す")
    a = ap.parse_args()
    roll = None
    if a.roll_month:
        y, m = a.roll_month.split("-")
        roll = (int(y), int(m))
    split(a.in_, a.out, roll, a.keep_data)


if __name__ == "__main__":
    main()
