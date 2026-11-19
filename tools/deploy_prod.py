# -*- coding: utf-8 -*-
r"""deploy_prod.py — 本番共有（UNC）へ配置する。20261014

本番 = contract.PROD_ROOT（\\fileserver\share$\00_DrJOY\当直勤務表）。月ごとのフォルダは
tools/month_setup.py（管理ツール）が作るので、ここが配るのは月に依らないものだけ:
  master … app\release\当直勤務表入力(Win).exe と dist\touchoku_xlsx.exe（xlsx ヘルパー）→ _本体\
           （月フォルダはここから複写される。使用中なら退避→配置、--wait-minutes）
  tools  … dist\救急元表ツール.exe / 形成外科当直表ツール.exe と _お読みください → 科別ツール\
  admin  … dist\当直勤務表管理ツール.exe と _お読みください → 管理ツール\、リポジトリの data\*.json・*.csv → 管理ツール\data\
           （20260926: 管理ツールの exe は 管理ツール\data\ に frames.json があればそちらを読む＝contract.resolve_data。
             toban.csv などを直したら --what admin で置けば、exe を作り直さずに①へ届く）
  docs   … README/TODO/docs の md・docx → ドキュメント\
  readme … 本番ルートの _お読みください.txt
usage: python -X utf8 tools\deploy_prod.py [--what all|master|months|tools|admin|docs|readme ...] [--wait-minutes 70]
"""
import argparse
import hashlib
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

PROD = C.PROD_ROOT
DIST_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "docs", "配布")


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def data_files(root=None):
    r"""管理ツール\data\ へ置くもの = リポジトリの data\ の *.json・*.csv（frames.json を最後に＝
    管理ツールは frames.json があるかで 管理ツール\data\ を使うか決めるので、ほかが揃ってから置く）。"""
    d = os.path.join(root or C.ROOT, "data")
    names = sorted(n for n in os.listdir(d) if n.lower().endswith((".json", ".csv")))
    names.sort(key=lambda n: n == "frames.json")
    return [os.path.join(d, n) for n in names]


WAIT_MINUTES = 0   # --wait-minutes: exe が使用中（誰かが開いている＝SMBロック）のとき、閉じるまで待って再試行する上限


def put(src, dst):
    """コピーしてハッシュで検証。使用中（PermissionError）なら リネーム退避→新規配置 を試し、
    それも無理なら WAIT_MINUTES の範囲で5分ごとに再試行する（開いている人が閉じるのを待つ。
    60分の自動終了は 20260925 に廃止した。月フォルダのアプリは起動係なので、共有の exe はふつう使用中にならない）。"""
    import time, datetime as dt
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.isfile(dst) and sha(src) == sha(dst):
        print(f"  同じ: {os.path.relpath(dst, PROD)}")
        return False
    deadline = time.time() + WAIT_MINUTES * 60
    while True:
        try:
            shutil.copy2(src, dst)
            break
        except PermissionError:
            # 1) リネーム退避 + 新規配置（管理当直の deploy-win-exe.ps1 と同じ順）
            try:
                # 20261014: 「….exe.旧_日時」のように拡張子の後ろに足す名前は、EDR がランサムウェアの疑い
                # （Possible Ransomware Encrypted File Creation via Python Script）として通報する。拡張子は元のまま1つにする
                stem, ext = os.path.splitext(dst)
                aside = f"{stem}_旧_{dt.datetime.now():%Y%m%d-%H%M}{ext}"
                os.rename(dst, aside)
                shutil.copy2(src, dst)
                print(f"  退避: {os.path.basename(aside)}（使用中だったので旧版を改名して残した。閉じられたら消してよい）")
                break
            except PermissionError:
                pass
            if time.time() >= deadline:
                raise SystemExit(f"使用中で差し替えできません: {dst}"
                                 "  誰かがこのファイルを開いています（アプリ・ツールを開いたままの人）。"
                                 "閉じてもらってから再実行するか、--wait-minutes 30 で待ちながら再試行してください")
            print(f"  使用中: {os.path.relpath(dst, PROD)} … 5分後に再試行（残り {int((deadline - time.time()) / 60)} 分）")
            time.sleep(300)
    if sha(src) != sha(dst):
        raise SystemExit(f"配置の検証に失敗: {dst}")
    print(f"  配置: {os.path.relpath(dst, PROD)}  {os.path.getsize(dst):,} bytes")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--what", nargs="*", default=["all"])
    ap.add_argument("--wait-minutes", type=int, default=0, help="exe が使用中のとき閉じるまで待つ上限（分）。例 30")
    a = ap.parse_args()
    global WAIT_MINUTES
    WAIT_MINUTES = a.wait_minutes
    what = set(a.what)
    if "all" in what:
        what = {"master", "months", "tools", "admin", "docs", "readme"}
    if not os.path.isdir(os.path.dirname(PROD)):
        raise SystemExit(f"本番共有に届きません: {PROD}")
    for d in (PROD, C.PROD_MASTER, C.PROD_ROSTER, C.PROD_TOOLS, C.PROD_ADMIN, os.path.join(PROD, "ドキュメント")):
        os.makedirs(d, exist_ok=True)
    if "master" in what:
        print("■ 本体の原本（_本体）")
        put(os.path.join(C.ROOT, "app", "release", "当直勤務表入力(Win).exe"), os.path.join(C.PROD_MASTER, "当直勤務表入力(Win).exe"))
        # 20260925: 起動係＋展開済みアプリの zip＋版の名前（month_setup.put_app が月フォルダへ写す。起動が約19秒→約3秒）
        for n in ("touchoku_app.zip", "touchoku_app.ver", "touchoku_launcher.exe"):   # zip を先・版と起動係をあとに
            p = os.path.join(C.ROOT, "dist", n)
            if os.path.isfile(p):
                put(p, os.path.join(C.PROD_MASTER, n))
            else:
                print(f"  ⚠ {n} がありません（tools\\build_launcher.bat で作る）")
        helper = os.path.join(C.ROOT, "dist", "touchoku_xlsx.exe")          # 20260912: xlsx ヘルパー（month_setup が月フォルダへ複写）
        if os.path.isfile(helper):
            put(helper, os.path.join(C.PROD_MASTER, "touchoku_xlsx.exe"))
        else:
            print(f"  ⚠ ヘルパーがありません（build_kyukyu_exe.bat で作る）: {helper}")
    if "months" in what:
        # 20260925: 活きている月フォルダ（R8年N月）へ、_本体\ の起動係・アプリの zip・版を配る
        import month_setup as MS
        print("■ 月フォルダのアプリ")
        for name in sorted(os.listdir(PROD)):
            ym = C.parse_month_dir(name)
            if not ym or not os.path.isdir(os.path.join(PROD, name, "config")):
                continue
            MS.put_app(os.path.join(PROD, name), ym[0] - 2018, ym[1], log=lambda s: print("  " + name + ": " + s.strip()))
        # 20260928: 試験用フォルダ（R8年9月・10月(動作確認用テストフォルダ)）にも同じ版を配る（起動係の名前は中にあるものに合わせる）
        import set_acl as SA
        for name in sorted(os.listdir(PROD)):
            ym = SA.parse_test_dir(name)
            md = os.path.join(PROD, name)
            if not ym or not os.path.isdir(os.path.join(md, "config")):
                continue
            exes = [n for n in os.listdir(md) if n.startswith("当直勤務表入力") and n.endswith(".exe")]
            MS.put_app(md, ym[0] - 2018, ym[1], exe_name=exes[0] if exes else None,
                       log=lambda s, name=name: print("  " + name + ": " + s.strip()))
    if "tools" in what:
        print("■ 科別ツール")
        for dept, exe in (("救急科", "救急元表ツール.exe"), ("形成外科", "形成外科当直表ツール.exe")):
            put(os.path.join(C.ROOT, "dist", exe), os.path.join(C.PROD_TOOLS, dept, exe))
            rd = os.path.join(DIST_DIR, f"_お読みください_{dept}.txt")
            if os.path.isfile(rd):
                put(rd, os.path.join(C.PROD_TOOLS, dept, "_お読みください.txt"))
    if "admin" in what:
        print("■ 管理ツール")
        put(os.path.join(C.ROOT, "dist", "当直勤務表管理ツール.exe"), os.path.join(C.PROD_ADMIN, "当直勤務表管理ツール.exe"))
        rd = os.path.join(DIST_DIR, "_お読みください_管理ツール.txt")
        if os.path.isfile(rd):
            put(rd, os.path.join(C.PROD_ADMIN, "_お読みください.txt"))
        # 20260926: 枠マスタ・祝日・当番当直・科マップ（管理ツールの exe が同梱の data より優先して読む）
        for src in data_files():
            put(src, os.path.join(C.PROD_ADMIN, "data", os.path.basename(src)))
    if "docs" in what:
        print("■ ドキュメント")
        for name in ("README.md", "TODO.md"):
            put(os.path.join(C.ROOT, name), os.path.join(PROD, "ドキュメント", name))
        for name in os.listdir(os.path.join(C.ROOT, "docs")):
            if name.endswith(".md") or name.endswith(".docx"):
                put(os.path.join(C.ROOT, "docs", name), os.path.join(PROD, "ドキュメント", name))
    if "readme" in what:
        print("■ お読みください")
        rd = os.path.join(DIST_DIR, "_お読みください_本番.txt")
        if os.path.isfile(rd):
            put(rd, os.path.join(PROD, "_お読みください.txt"))
    print(f"完了: {PROD}")


if __name__ == "__main__":
    main()
