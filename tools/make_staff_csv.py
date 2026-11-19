# -*- coding: utf-8 -*-
"""make_staff_csv.py — Dr.JOY のスタッフ一覧CSV → 当直アプリ用の名簿 staff_YYYYMMDD.csv。20261015

ユーザー指示（2026-09-09）:
  ・名前は staff_YYYYMMDD.csv（日付 = 元CSVの更新日）
  ・中身は 職業「医師」だけ。アカウント状態は「有効」と「仮登録」（エクスポートはロック・削除を出さない設定なので実質この2つ。
    仮登録は当初は除く案だったがユーザー指示で入れることになった）
列は読む側が使う列だけ（20260923 第2陣。それまでは元CSVの28列を全部書いていた）。
  本番の 名簿\\ は Domain Users 全員が読めるので、メールアドレス・性別・入職日・雇用形態・NameID などは載せない。
  読む側はすべて contract.load_staffs を通る（make_month・aggregate・dtx14_extract/fill・ingai_report・kotai_record・
  seed_month・bucho_roster・kyukyu/keisei の変換と GUI）。load_staffs が読む列 = COLUMNS。
  見出しの名前・並び（元CSVの順）・CP932 は元のまま。

usage:
  python -X utf8 tools\\make_staff_csv.py [--src "staffs (81).csv"] [--out DIR ...]
    --src 省略 = Downloads の最新 staffs*.csv（500行以上）
    --out 省略 = 本番共有の 名簿\\ フォルダ（無ければ Downloads）。複数指定可
"""
import argparse
import csv
import datetime as dt
import glob
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

DEFAULT_OUT = C.PROD_ROSTER   # 本番共有の 名簿\（20260909）
JOBS = ("医師",)
STATES = ("有効", "仮登録")   # 20260909 ユーザー指示: 仮登録も入れる（エクスポート既定はロック・削除を出さない）
# 書き出す列（20260923 第2陣）= contract.load_staffs が読む列だけ。読む列を増やすときはここにも足すこと
COLUMNS = ("職員コード", "ログインID", "姓", "名", "せい", "めい", "所属", "職業", "職名")


def newest_raw():
    cands = []
    for f in glob.glob(os.path.join(os.path.expanduser("~"), "Downloads", "staffs*.csv")):
        with open(f, "rb") as fh:
            if sum(1 for _ in fh) >= 500:
                cands.append((os.path.getmtime(f), f))
    return max(cands)[1] if cands else None


def make(src, out_dirs, blank_code_dirs=()):
    """blank_code_dirs: 職員コードの列を空にして書くフォルダ（全職員が読む 名簿\\。20261015 管理者: 職員コードは
    人事給与システムの職員番号と同じ11桁なので、管理者だけの 管理ツール\\名簿\\ にだけ置く。列は残す＝読む側の形を変えない）。"""
    blank = {os.path.normcase(os.path.abspath(d)) for d in blank_code_dirs}
    stamp = dt.datetime.fromtimestamp(os.path.getmtime(src)).strftime("%Y%m%d")
    with open(src, encoding="cp932", newline="") as f:
        rd = csv.DictReader(f)
        rows = list(rd)
        src_fields = rd.fieldnames or []
    lack = [c for c in COLUMNS if c not in src_fields]
    if lack:                         # 読む側が要る列が無い名簿を置くと、どのツールも黙って空の名簿を読む
        raise SystemExit(f"元CSVに要る列がありません: {'・'.join(lack)}（{src}）")
    fields = [c for c in src_fields if c in COLUMNS]     # 元CSVの並びのまま、要る列だけ
    kept = [r for r in rows if (r.get("職業") or "").strip() in JOBS and (r.get("アカウント状態") or "").strip() in STATES]
    dropped = {}
    for r in rows:
        if r in kept:
            continue
        key = f"{(r.get('職業') or '').strip()}/{(r.get('アカウント状態') or '').strip()}"
        dropped[key] = dropped.get(key, 0) + 1
    outs = []
    for d in out_dirs:
        os.makedirs(d, exist_ok=True)
        p = os.path.join(d, f"staff_{stamp}.csv")
        rows_out = kept
        if os.path.normcase(os.path.abspath(d)) in blank:
            rows_out = [{**r, "職員コード": ""} for r in kept]
        with open(p, "w", encoding="cp932", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows_out)
        outs.append(p)
    return stamp, len(rows), len(kept), dropped, outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=None)
    ap.add_argument("--out", nargs="*", default=None)
    a = ap.parse_args()
    src = a.src or newest_raw()
    if not src or not os.path.isfile(src):
        raise SystemExit("元の staffs*.csv が見つかりません（Downloads に Dr.JOY のエクスポートを置いてください）")
    outs = a.out or [DEFAULT_OUT if os.path.isdir(os.path.dirname(DEFAULT_OUT)) else os.path.join(os.path.expanduser("~"), "Downloads")]
    stamp, n_all, n_kept, dropped, paths = make(src, outs, blank_code_dirs=[C.PROD_ROSTER])
    print(f"元: {src}  {n_all} 行")
    print(f"残した: {n_kept} 行（職業={'/'.join(JOBS)} かつ アカウント状態={'/'.join(STATES)}）")
    print("除いた: " + "、".join(f"{k}×{v}" for k, v in sorted(dropped.items())))
    for p in paths:
        print(f"出力: {p}")


if __name__ == "__main__":
    main()
