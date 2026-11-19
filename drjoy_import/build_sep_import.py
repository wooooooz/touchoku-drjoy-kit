# -*- coding: utf-8 -*-
"""翌月の当直をDr.JOY「勤務予定」インポート用xlsxに作る（毎月の一括投入の STEP 4）。20261028

★2026-08-27 の作り直しの理由（投入直前のレビューで判明）
  ① オンコール(略称が括弧つき)を当直として投入していた … interval_v2.load_duty で除外
  ② 休日の全日当直で「昼」だけが残り泊まりが落ちていた … 同 load_duty で 昼+夜→昼夜 に合成
  ③ 秋川三郎・癸田四郎の職員コードが入れ替わっていた
     … 勤務予定エクスポートが職員コード修正(8/27 07:48)より前の 01:11 取得だったため。
       21:25 に取り直した版を使う。

★20261028 引数で月とファイルを渡す形にした（以前は毎月この中のパスを書き換えていた）。
  救急科・小児科の当直の前後に外す日中勤務は、当直勤務表アプリの day_off_plan を正にする
  （管理者 2026-10-28）。--month を渡すと day_off_plan.py を走らせて一覧を作り、生成器に渡す。

使い方:
  python -X utf8 build_sep_import.py --month 2026-11 --dtx14 <DTX14の複製.xlsm>
      --plan <2026年11月度_勤務予定_*.xlsx> --staffs <staffs (N).csv> --out-dir <作業フォルダ>
  --no-day-off … 一覧を使わず、以前の明けの削り（B=12:00・C-1=明けでオフ）にする（比べるとき用）
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, r"C:\tools\touchoku\drjoy_import")

import dtx14_names
import dtx14_to_drjoy_schedule as G


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM（投入する月）")
    ap.add_argument("--dtx14", required=True)
    ap.add_argument("--plan", required=True, help="その月の勤務予定エクスポート（at0029・WorkingPlan）")
    ap.add_argument("--plan-next", help="翌月の勤務予定エクスポート（月末の当直の明けを書くのに使う。無ければ全日休みだけ書く）")
    ap.add_argument("--staffs", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--no-day-off", action="store_true")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    for p in (a.dtx14, a.plan, a.staffs):
        if not os.path.exists(p):
            raise SystemExit(f"見つかりません: {p}")
    os.makedirs(a.out_dir, exist_ok=True)
    y, m = (int(x) for x in a.month.split("-"))
    tag = f"R{y - 2018}年{m}月"
    out = os.path.join(a.out_dir, f"DrJOY勤務予定_当直_{tag}{'_明け旧方式' if a.no_day_off else ''}.xlsx")

    print("DTX14  :", os.path.basename(a.dtx14))
    print("勤務予定:", os.path.basename(a.plan))
    print("staffs :", os.path.basename(a.staffs))

    day_off = None
    if not a.no_day_off:
        day_off, notes = [], []
        nxt = f"{y + (m == 12)}-{m % 12 + 1:02d}"
        for mon in (a.month, nxt):      # 翌月の一覧は、月末の当直の明け（翌月1日など）のため
            js = G.make_day_off(mon, os.path.join(a.out_dir, f"day_off_plan_{mon}.json"))
            offs, nts = G.load_day_off(js, a.staffs)
            day_off += offs
            notes += nts if mon == a.month else []
            print(f"外す日中勤務（当直勤務表アプリ・{mon}）: {len(offs)}件 / 職員コードが引けない "
                  f"{sum(1 for o in offs if not o['code'])}件" + (f" / 要相談 {len(nts)}件" if mon == a.month else ""))
        for n in notes:
            print(f"  ★要相談: {n}")
    print()

    resolve = dtx14_names.build_resolver(a.dtx14, staffs_csv=a.staffs)
    G.build(a.out_dir, out, dtx14=a.dtx14, resolve=resolve, plan_xlsx=a.plan, staffs_csv=a.staffs,
            day_off=day_off, plan_next_xlsx=a.plan_next)
    print(f"\n→ {out}")


if __name__ == "__main__":
    main()
