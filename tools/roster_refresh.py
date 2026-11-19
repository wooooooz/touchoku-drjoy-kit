# -*- coding: utf-8 -*-
r"""roster_refresh.py — 名簿の更新を一括で（毎日3:30の名簿更新のタスク用）。20261028

Dr.JOY のエクスポート（ブラウザ操作）は先に drjoy/Export-DrJoyStaff.ps1 で済ませる（drjoy/tasks/Run-TouchokuRosterDaily.ps1）。ここは残り:
  1. Downloads の最新 staffs*.csv（500行以上）→ staff_YYYYMMDD.csv（職業=医師）を
     本番の 名簿\ と 科別ツール\{救急科,形成外科} に置く（古い staff_YYYYMMDD.csv は消す）
  2. 名簿か Dr.JOY の所属履歴（名簿\所属履歴_YYYY.xlsx。中身のハッシュ＝.shozoku_sha）が前回から変わっていれば、活きている月フォルダ（今月以降・確定していない）の
     config/HTML/テンプレを新しい名簿で作り直す（month_setup.create。exe と提出物は触らない）
  月フォルダを新しく作るのは毎日15時の反映（dtx14_daily_sync → new_month。来月と次の次の月）。ここでは作らない（無ければ結果に書くだけ）。
結果は標準出力に1行ずつ。失敗は非0で終了（呼び出し側がメールする）。

usage: python -X utf8 tools\roster_refresh.py [--force]
"""
import argparse
import datetime as dt
import glob
import hashlib
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C            # noqa: E402
import make_staff_csv as MS     # noqa: E402
import month_setup as MO        # noqa: E402
import shozoku as SZ            # noqa: E402

# 20260915 管理者指示: 科別ツールは共有の 名簿\ を直接読むので、exe の隣への写しはやめた（同じ日付の写しが既定に勝っていた）
TOOL_DIRS = []


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="名簿が変わっていなくても月フォルダを作り直す")
    a = ap.parse_args()
    if not os.path.isdir(C.PROD_ROOT):
        raise SystemExit(f"本番共有に届きません（職場ネットワークが要ります）: {C.PROD_ROOT}")
    os.makedirs(C.PROD_ROSTER, exist_ok=True)
    os.makedirs(C.PROD_ROSTER_FULL, exist_ok=True)
    src = MS.newest_raw()
    if not src:
        raise SystemExit("Downloads に Dr.JOY のエクスポート（staffs*.csv・500行以上）がありません")
    # 1. staff_YYYYMMDD.csv（20261014: 職員コードの入った版は 管理ツール\名簿\ だけ。全職員が読む 名簿\ は職員コードを空に）
    dirs = [C.PROD_ROSTER_FULL, C.PROD_ROSTER] + [d for d in TOOL_DIRS if os.path.isdir(d)]
    stamp, n_all, n_kept, dropped, paths = MS.make(src, dirs, blank_code_dirs=dirs[1:])
    print(f"名簿: {os.path.basename(src)}（{n_all}行）→ staff_{stamp}.csv 医師 {n_kept} 名 → {len(paths)} か所"
          "（職員コードは 管理ツール\\名簿\\ だけ）")
    for d in dirs:
        for f in glob.glob(os.path.join(d, "staff_*.csv")):
            if re.fullmatch(r"staff_\d{8}\.csv", os.path.basename(f)) and os.path.basename(f) != f"staff_{stamp}.csv":
                os.remove(f)
                print(f"削除: {f}")
    # 1b. 貴科の医師リストを直せる人（部門管理者）の一覧 名簿\roster_editors.json（20261028 管理者）。
    #     作れなくても名簿の更新は止めない（前の日の一覧が残る。アプリは読めなければ誰にも × と区分の切り替えをさせない）
    try:
        import roster_editors as RE
        ed = RE.build(src)
        RE.write(ed, os.path.join(C.PROD_ROSTER, RE.FILE_NAME))
        n = {lv: sum(1 for e in ed["editors"].values() if e["level"] == lv) for lv in ("全体", "所属", "追加")}
        print(f"部門管理者: 全体管理者 {n['全体']}・所属管理者 {n['所属']}・追加の作成者 {n['追加']} → 名簿\\{RE.FILE_NAME}")
    except (Exception, SystemExit) as e:
        print(f"注意: 部門管理者の一覧を作れなかった（{type(e).__name__}: {e}）→ 前の一覧のまま")
    marker = os.path.join(C.PROD_ROSTER, ".roster_sha")
    cur = sha(paths[0])
    prev = open(marker, encoding="utf-8").read().strip() if os.path.isfile(marker) else ""
    changed = cur != prev
    # 20260927: 所属履歴が変わっても作り直す（config の roster[].dept は対象月の1日の所属＝所属履歴から引く）
    smarker = os.path.join(C.PROD_ROSTER, ".shozoku_sha")
    try:
        scur = SZ.content_hash(C.PROD_ROSTER)
    except Exception as e:                  # 読めなくても名簿の更新は止めない（変わったとみなす）
        print(f"注意: 所属履歴のハッシュが取れない（{type(e).__name__}: {e}）")
        scur = "?"
    sprev = open(smarker, encoding="utf-8").read().strip() if os.path.isfile(smarker) else ""
    if scur != sprev:
        print("所属履歴が前回から変わった" if sprev else "所属履歴のハッシュを初めて記録する")
        changed = True
    # 2. 活きている月フォルダを作り直す
    today = dt.date.today()
    months = MO.list_months()
    active = [(y, m) for y, m, closed in months if not closed and (y, m) >= (today.year, today.month)]
    tgt = MO.default_target(today)
    if tgt not in [(y, m) for y, m, _ in months]:
        print(f"注意: 対象月 {tgt[0]}/{tgt[1]} の月フォルダがまだありません（毎日15時の反映が自動で作ります。"
              "作れなかったときは管理者へ警報メールが届きます。急ぐときは管理ツールの①で作れます）")
    if not changed and not a.force:
        print(f"名簿・所属履歴は前回と同じ。月フォルダはそのまま（活きている月: {', '.join(f'{y}/{m}' for y, m in active) or 'なし'}）")
        return
    for y, m in active:
        MO.create(y, m, log=lambda t: print("  " + t))
        print(f"作り直し: {C.month_dir_name(y, m)}")
    # 確定済みの月（今月以降）は、名簿（config の roster）だけ作り直す（20261028 管理者。これまでは確定した日の名簿で止まっていた）
    failed = []
    for y, m, closed in months:
        if not closed or (y, m) < (today.year, today.month):
            continue
        try:
            if not MO.refresh_roster_only(y, m, log=lambda t: print("  " + t)):
                print(f"名簿は変わらない（確定済み）: {C.month_dir_name(y, m)}")
        except (Exception, SystemExit) as e:
            failed.append(f"{C.month_dir_name(y, m)}（{type(e).__name__}: {e}）")
    if failed:
        raise SystemExit("確定済みの月の名簿を更新できませんでした（次の回にもう一度やります）: " + " / ".join(failed))
    with open(marker, "w", encoding="utf-8") as f:
        f.write(cur)
    if scur != "?":
        with open(smarker, "w", encoding="utf-8") as f:
            f.write(scur)
    if not active:
        print("活きている月フォルダがありません")


if __name__ == "__main__":
    main()
