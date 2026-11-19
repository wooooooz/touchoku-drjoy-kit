# -*- coding: utf-8 -*-
r"""strip_codes.py — 全職員が読む月フォルダのファイルから職員コードを消す。20261015

管理者 2026-10-15: 職員コード（人事給与システムの職員番号と同じ11桁。統合ファイルでは暗号にしている値）が、Domain Users の読める
  名簿\staff_*.csv・各月の config\current.json・master\当直予定表_*.json（と master\_履歴\）に平文で入っていた。
  → 名簿は roster_refresh が 管理ツール\名簿\（管理者だけ）と 名簿\（職員コードの列は空）に分けて置く。
    config は make_month が職員コードを書かない。予定表はアプリが config の空の職員コードを写すので、新しくは入らない。
  ここは、すでに入っている分（と、開いたままの古いアプリが書き戻した分）を消す。15時の反映が毎日呼ぶ（入っていなければ何もしない）。
消すもの: config の roster[].code、予定表の depts.*.roster[].code・depts.*.duties[].staff_code・depts.*.baseline[].staff_code。
  管理側の処理（人別シートの H列・交代記録の from_code/to_code）は contract.code_of で 管理ツール\名簿\ から引く。
書き方: その場で書き換える（一時ファイルからの差し替えにしない）。交代記録は予定表と _履歴\ の複製の「持ち主」（NTFS）と
  更新時刻で、誰がいつ保存したかを割り出すので、持ち主を変えず、更新時刻も元に戻す。今の予定表はロックを取ってから書く。
対象は名前の決まった月フォルダ（R8年11月 など）だけ。動作確認用の試験フォルダ（管理者だけ）は触らない。

usage: python -X utf8 tools\strip_codes.py [--dry-run]
"""
import argparse
import glob
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C            # noqa: E402
import master_store as MST      # noqa: E402


def strip_obj(obj):
    """config か予定表の dict から職員コードを消す（その場で）。戻り: 消した数。"""
    n = 0

    def roster(lst):
        nonlocal n
        for p in lst or []:
            if isinstance(p, dict) and p.get("code"):
                p["code"] = ""
                n += 1

    def duties(lst):
        nonlocal n
        for r in lst or []:
            if isinstance(r, dict) and r.get("staff_code"):
                r["staff_code"] = ""
                n += 1

    roster(obj.get("roster"))
    for ent in (obj.get("depts") or {}).values():
        if isinstance(ent, dict):
            roster(ent.get("roster"))
            duties(ent.get("duties"))
            duties(ent.get("baseline"))
    return n


def strip_file(path, dry=False, use_lock=False):
    """1つのファイルから職員コードを消す。持ち主と更新時刻は変えない。戻り: 消した数。"""
    def work():
        with open(path, encoding="utf-8-sig") as f:
            raw = f.read()
        obj = json.loads(raw)
        n = strip_obj(obj)
        if n and not dry:
            st = os.stat(path)
            bom = raw.startswith("﻿")
            data = json.dumps(obj, ensure_ascii=False)
            with open(path, "r+", encoding="utf-8-sig" if bom else "utf-8", newline="\n") as f:
                f.seek(0)
                f.write(data)
                f.truncate()
            os.utime(path, ns=(st.st_atime_ns, st.st_mtime_ns))
        return n

    if use_lock:
        with MST.lock(path):
            return work()
    return work()


def month_dirs(root=None):
    """名前の決まった月フォルダ（R8年11月 など）。試験フォルダ（…(動作確認用テストフォルダ)）は外す。"""
    return sorted(d for d in glob.glob(os.path.join(root or C.PROD_ROOT, "R*年*月"))
                  if os.path.isdir(d) and re.fullmatch(r"R\d+年\d+月", os.path.basename(d)))


def strip_all(dry=False, log=print, root=None):
    """全部の月フォルダの config・予定表・_履歴 から職員コードを消す。戻り: (消した数, 失敗したファイル)。"""
    total, bad = 0, []
    base = root or C.PROD_ROOT

    def rel(p):
        try:
            return os.path.relpath(p, base)
        except ValueError:
            return p
    for md in month_dirs(root):
        files = [(p, False) for p in glob.glob(os.path.join(md, "config", "*.json"))]
        files += [(p, True) for p in glob.glob(os.path.join(md, "master", "*.json"))]
        files += [(p, False) for p in glob.glob(os.path.join(md, "master", MST.HISTORY, "*.json"))]
        for p, lk in files:
            try:
                n = strip_file(p, dry=dry, use_lock=lk)
            except Exception as e:                      # 1つの失敗で残りを止めない
                bad.append(f"{rel(p)}（{type(e).__name__}: {e}）")
                continue
            if n:
                total += n
                log(f"  職員コードを消した{'（dry-run）' if dry else ''}: {rel(p)} {n} 件")
    return total, bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    total, bad = strip_all(a.dry_run)
    print(f"合計 {total} 件{'（dry-run・書いていない）' if a.dry_run else ''}")
    for b in bad:
        print(f"✗ {b}")
    if bad:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
