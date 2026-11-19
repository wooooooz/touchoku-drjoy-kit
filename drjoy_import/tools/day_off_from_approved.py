# -*- coding: utf-8 -*-
r"""day_off_from_approved.py — 救急科・小児科の外す日中勤務の一覧を、交代記録台帳の approved（承認済みの割当）から作る。20261028

当直勤務表アプリの tools\day_off_plan.py（決め方の正本）は、各月のマスターの duties から一覧を作る。
ところがマスターは「各科が出した今の予定」で、決裁後の交代申請は承認を待たずに入る（承認待ちも含む。
2026-10-28）。Dr.JOY に入れる当直（DTX14・交代の自動登録）は台帳の approved
（reflected_at の付いた、DTX14 に書けた割当）なので、一覧も同じ割当から作らないと食い違う。
→ 決め方（day_off_plan.plan）はそのまま借り、duties だけ台帳の approved に差し替える。Dr.JOY には触らない。

月の境目のため、前後の月も読む（day_off_plan.month_data と同じ。shozoku.REAL_FROM＝2026-11 より前の月は試験用の作り物なので読まない）。月ごとに:
  台帳（R8年N月\集約\交代記録.json）が決裁済み（confirmed_at あり）で approved があれば approved。
  無ければマスター（決裁前の月・交代が1件も無い月）。どちらを使ったかを出力に書く。
職名（水準 B／C-1 を決める）は day_off_plan.month_data と同じくマスターの名簿から取る。

使い方:
  python -X utf8 tools\day_off_from_approved.py --month 2026-11 --json <出力.json>
出力は hospital.day-off-plan/1（day_off_plan.py --json と同じ形）。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

TA_TOOLS = r"C:\tools\touchoku\tools"
sys.path.insert(0, TA_TOOLS)
import contract as C            # noqa: E402
import day_off_plan as D        # noqa: E402
import master_store as MST      # noqa: E402

DUTY_ROOT = r"\\fileserver\share$\00_DrJOY\当直勤務表"


def month_folder(yy, mm):
    """R8年N月 (or the test-folder name for 9・10月, as day_off_plan.month_data does)."""
    p = os.path.join(C.PROD_ROOT, f"R{yy - 2018}年{mm}月")
    if os.path.isfile(MST.master_path(p, yy, mm)):
        return p
    import glob
    alt = [x for x in glob.glob(os.path.join(C.PROD_ROOT, f"R{yy - 2018}年{mm}月(*)"))
           if os.path.isfile(MST.master_path(x, yy, mm))]
    return alt[0] if alt else None


def approved_duties(folder, owner):
    """台帳の approved → duties（day_off_plan が読む形）。決裁前・approved が無ければ None。"""
    p = os.path.join(folder, "集約", "交代記録.json")
    try:
        lg = json.load(open(p, encoding="utf-8-sig"))
    except (OSError, ValueError):
        return None
    if not lg.get("confirmed_at") or not isinstance(lg.get("approved"), dict):
        return None
    out = []
    for r in lg["approved"].get(owner) or []:
        date, band, fid, login, name, oc = r[:6]
        out.append({"date": date, "band": band, "frame_id": fid, "staff_login": login, "name": name,
                    "oncall": bool(oc), "staff_code": r[6] if len(r) > 6 else ""})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM")
    ap.add_argument("--json", required=True)
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    y, m = (int(x) for x in a.month.split("-"))
    hol = C.load_holidays()
    months = (((y, m - 1) if m > 1 else (y - 1, 12)), (y, m), ((y, m + 1) if m < 12 else (y + 1, 1)))
    results = []
    for owner in C.consecutive_warn_owners():
        duties, used = [], []
        _md, titles = D.month_data(y, m, owner, log=lambda *_a: None)      # titles from the masters' rosters
        import shozoku as SZ
        for yy, mm in months:
            if (yy, mm) < tuple(SZ.REAL_FROM):
                # 20261028: 10月までは試験用フォルダの作り物の割当（本物は事務担当の DTX14）。月の境目の計算に混ぜない
                #   （day_off_plan.month_data と同じ。2026-10-28）
                continue
            folder = month_folder(yy, mm)
            if not folder:
                continue
            ap_ = approved_duties(folder, owner)
            if ap_ is not None:
                duties += ap_
                used.append(f"{yy}-{mm:02d}=approved")
            else:
                ent = (MST.load(MST.master_path(folder, yy, mm), yy, mm).get("depts") or {}).get(owner) or {}
                duties += ent.get("duties") or []
                used.append(f"{yy}-{mm:02d}=master")
        offs, notes = D.plan(duties, hol, owner, titles)
        offs = [o for o in offs if (o["date"].year, o["date"].month) == (y, m)]
        notes = [n for n in notes if (n["date"].year, n["date"].month) == (y, m)]
        print(f"{owner}: 読んだ割当 {', '.join(used) or 'なし'} → 外す日中勤務 {len(offs)}件・要相談 {len(notes)}件")
        results.append((owner, offs, notes, duties))
    D.write_json(a.json, y, m, results)
    print("書いた:", a.json)


if __name__ == "__main__":
    main()
