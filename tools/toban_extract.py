# -*- coding: utf-8 -*-
"""toban_extract.py — DTX14 の「当番当直」シート → data/toban.csv。20260922

当番当直（管理者 2026-09-22）: 日ごと・昼/夜ごとに担当の診療科が決まっている役割。年末まで DTX14 の
「当番当直」シート（列 = 日付／曜日／区分（休日-昼・休日-夜・平日-夜）／担当科）に書いてある。
本表ではグレー地に白文字。アプリではその科のファーストの欄に赤字で「当番当直」と出し、
オンコールでは提出できないようにする（当直が必須）。

担当科の名前はファースト表の A 列の科名（dept_dtx14）と同じ表記（事務担当が置換済み）。

usage: python -X utf8 tools\\toban_extract.py --in <DTX14.xlsm> [--out data\\toban.csv]
出力: date,band,dept（UTF-8 BOM・CRLF）。読む側 = contract.load_toban / make_month.build_config
"""
import argparse
import csv
import datetime as dt
import os
import sys
import warnings

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402

SHEET = "当番当直"
BAND = {"休日-昼": "昼", "休日-夜": "夜", "平日-夜": "夜", "平日-昼": "昼"}


def read_toban(xlsm):
    import openpyxl
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"シート『{SHEET}』がありません: {xlsm}")
    ws = wb[SHEET]
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        r = list(r) + [None] * 4
        d, kubun, dept = r[0], r[2], r[3]
        if not isinstance(d, (dt.date, dt.datetime)) or not kubun or not dept:
            continue
        band = BAND.get(str(kubun).strip())
        if not band:
            print(f"  ⚠ 区分が読めません: {d} {kubun!r}")
            continue
        d = d.date() if isinstance(d, dt.datetime) else d
        out.append((d.isoformat(), band, str(dept).strip()))
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_", required=True)
    ap.add_argument("--out", default=os.path.join(C.DATA, "toban.csv"))
    a = ap.parse_args()
    rows = read_toban(a.in_)
    frames = C.load_frames()
    known = {C.namekey(f.get("dept_dtx14") or f["dept"]) for f in frames}
    unknown = sorted({dept for _, _, dept in rows if C.namekey(dept) not in known})
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(["date", "band", "dept"])
        w.writerows(rows)
    print(f"当番当直 {len(rows)} 件（{rows[0][0]}〜{rows[-1][0]}）→ {a.out}")
    if unknown:
        print(f"  ⚠ 枠マスタに無い科名: {unknown}（frames.json の dept_dtx14 と表記が違う。アプリで赤字が出ない）")


if __name__ == "__main__":
    main()
