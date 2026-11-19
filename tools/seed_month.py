# -*- coding: utf-8 -*-
"""seed_month.py — 現物 DTX14 の月を共有マスターへ流し込み、科ごとの当直表xlsxを書き出す。20261015

目視検証用（ユーザー指示 2026-09-13）: 確定済みの月（R8.9）を新形式で見る。
  1. dtx14_extract で現物 → 原子レコード（selftest T1 と同じ読み口）
  2. 枠の owner ごとに分けて master_store.put_dept（roster は出現した人＝常勤扱い）
  3. duty_sheet.write_book で科ごとの当直表xlsx を <月フォルダ>\\当直表\\ に書く

usage: python -X utf8 tools\\seed_month.py --year 2026 --month 9 [--xlsm PATH] [--staffs PATH] [--month-dir DIR] [--no-master]
  --month-dir 省略 = 本番の月フォルダ（contract.prod_month_dir）。無ければ先に month_setup で作る。
"""
import argparse
import datetime as dt
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C            # noqa: E402
import dtx14_extract as X       # noqa: E402
import duty_sheet as D          # noqa: E402
import master_store as MS       # noqa: E402

DEFAULT_XLSM = r"C:\tools\touchoku\sample\DTX14_sample.xlsm"


def safe(owner):
    return owner.replace("/", "・").replace(" ", "").replace("　", "")


def members_of(records, staff_by_login):
    """原子レコードに出てきた人 → duty_sheet のメンバー（selftest.members_of と同じ）"""
    out, seen = [], set()
    for r in records:
        lg = r.get("staff_login") or ""
        if not lg or lg in seen:
            continue
        seen.add(lg)
        p = staff_by_login.get(lg) or {}
        name = p.get("name") or r.get("name") or ""
        sei, _, mei = C.namekey_display(name).partition(" ")
        out.append({"login": lg, "code": p.get("code", ""), "sei": sei, "mei": mei,
                    "name": name, "kana": p.get("kana", ""), "title": p.get("title", ""),
                    "in_house": True, "unregistered": False})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--month", type=int, required=True)
    ap.add_argument("--xlsm", default=DEFAULT_XLSM)
    ap.add_argument("--staffs", default=None)
    ap.add_argument("--month-dir", default=None)
    ap.add_argument("--no-master", action="store_true", help="マスターには書かず xlsx だけ出す")
    a = ap.parse_args()
    year, month = a.year, a.month
    staffs = a.staffs or C.newest_staffs()
    md = a.month_dir or C.prod_month_dir(year, month)
    if not os.path.isdir(md):
        raise SystemExit(f"月フォルダがありません（先に month_setup で作る）: {md}")
    if not os.path.isfile(a.xlsm):
        raise SystemExit(f"DTX14 の複製がありません: {a.xlsm}")

    frames = C.load_frames()
    holidays = C.load_holidays()
    people, _, _ = C.load_staffs(staffs)
    staff_by_login = {p["login"]: p for p in people}
    records, unresolved = X.extract(a.xlsm, staffs, year, month)
    print(f"extract: {len(records)}件（ログインID未解決 {len(unresolved)}件）  名簿={os.path.basename(staffs)}")

    fmap = {f["frame_id"]: f for f in frames}
    by_owner = defaultdict(list)
    for r in records:
        f = fmap.get(r["frame_id"])
        if f and not C.is_active(f):
            continue        # 消滅した科の枠（retired）は流し込まない・当直表も作らない（20260914）
        owner = f.get("owner", f["dept"]) if f else r["frame_dept"]
        by_owner[owner].append(r)

    mpath = MS.master_path(md, year, month)
    out_dir = os.path.join(md, "当直表")
    os.makedirs(out_dir, exist_ok=True)
    over = []
    for owner, recs in sorted(by_owner.items()):
        own = [f for f in C.active_frames(frames) if f.get("owner", f["dept"]) == owner]
        members = members_of(recs, staff_by_login)
        rev = None
        if not a.no_master:
            roster = [{"login": m["login"], "code": "", "name": m["name"],   # 20261015: 予定表に職員コードを置かない
                       "title": m["title"], "in_house": True} for m in members]
            _, ent = MS.put_dept(mpath, owner, duties=recs, roster=roster,
                                 by="seed_month(DTX14現物)", year=year, month=month)
            rev = (ent or {}).get("rev")
        xp = os.path.join(out_dir, f"当直表_R{year - 2018}年{month}月_{safe(owner)}.xlsx")
        info = D.write_book(xp, year=year, month=month, owner=owner, frames_own=own,
                            holidays=holidays, members=members, duties=recs,
                            source="master" if rev else "seed", master_rev=rev,
                            asof=dt.datetime.now())
        if info.get("overflow"):
            over.append(f"{owner}:{info['overflow']}人")
        print(f"  {owner}: {len(recs)}件 {len(members)}名 → {os.path.basename(xp)}")
    print(f"{len(by_owner)}科 → {out_dir}")
    if not a.no_master:
        m = MS.load(mpath, year, month)
        print(f"マスター: {len(m['depts'])}科  {mpath}")
    if over:
        print("⚠ レーンにあふれた人（備考行へ）: " + " / ".join(over))


if __name__ == "__main__":
    main()
