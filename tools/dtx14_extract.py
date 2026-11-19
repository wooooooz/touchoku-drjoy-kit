# -*- coding: utf-8 -*-
"""dtx14_extract.py  20260922
本物のDTX14（xlsm）から原子レコードを起こす読み口（drjoy_import/tools/dtx14_contract.py の
extract 部分を Touchoku_Plan 向けに vendored したもの。Phase 0 で往復一致を実証済みのコード）。

用途:
  1) セルフテスト: 現物 → 原子レコード → 科別提出JSONに分解 → 集約器 → 互換xlsx → 現物と突合
  2) 並行運用中の月次比較: アプリの出力と現物DTX14の人別3シートの差分確認
  3) 従来どおりのExcel提出の受け皿（提出が無い科をDTX14現物から補うとき）

現物は読み取り専用（openpyxl read_only・COM不使用）。

20260922: 行の割り出しを**見出し（A列の科名・C列の昼/夜/全日）**に一本化した（locate_rows）。
事務担当の雛形が全日の1行から昼/夜の2行に変わると以降の行が全部ずれるので、
枠マスタ frames.json の固定 row は見出しで引けなかったときの落とし先にだけ使う。
"""
from __future__ import annotations

import datetime as dt
import json
import re
import zipfile
from collections import Counter

import openpyxl

from contract import (is_holiday, load_frames, load_holidays, load_staffs, namekey, nfc,
                      resolve_name)

SHEET_FIRST = "当直表 ファースト"
SHEET_SECOND = "当直表 セカンド以下"
P_DUTY = "人別当直日"
P_ONCALL = "人別オンコール日"
P_FIRST = "人別ファースト"
RANGES_FIRST = [(5, 45), (63, 64)]   # 診療科当直 + 管理当直
RANGES_SECOND = [(5, 38)]


def load_sheet_grid(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    z = zipfile.ZipFile(path)
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml)
    rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    name2file = {n: "xl/" + rid2t[r] for n, r in sheets}
    x = z.read(name2file[sheet]).decode("utf-8")
    z.close()
    merged_a = {}
    for ref in re.findall(r'<mergeCell ref="([^"]+)"/>', x):
        m = re.match(r"^A(\d+):[A-Z]+(\d+)$", ref)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            for r in range(a, b + 1):
                merged_a[r] = a
    return rows, merged_a


def cellv(rows, r, c):
    if r - 1 < len(rows):
        row = rows[r - 1]
        if c - 1 < len(row):
            return row[c - 1]
    return None


# ------------------------------------------------------------------ 行の見出し索引
BANDS = ("昼", "夜")


def locate_rows(ws_rows, sheet_name=None, start=5, end=None):
    """A列の科名とC列の見出しだけで行を割り出す。戻り: {(科名のnamekey, 見出し): 行}。
    見出しは "昼" / "夜" / それ以外は全部 "全日"（空欄・「全日」・看護部の「深夜」等）。

    現物の2行組は上の行にだけ科名があるので、A列が空の行は直前の科の続きとして扱う。
    A列もC列も空の行で科の続きは切れる（現物の45行目・セカンド以下の7/33行目）。
    sheet_name は呼び出し側の見通しのために受けるだけで、判定には使わない。"""
    end = end or len(ws_rows)
    out, cur = {}, None
    for r in range(start, end + 1):
        a = cellv(ws_rows, r, 1)
        c = cellv(ws_rows, r, 3)
        a = "" if a is None else str(a).strip()
        c = "" if c is None else str(c).strip()
        if a:
            cur = a
        elif not c:
            cur = None
            continue
        if not cur:
            continue
        out.setdefault((namekey(nfc(cur)), c if c in BANDS else "全日"), r)
    return out


def dept_band_rows(dept_dtx14, loc):
    """科名 → ({"昼": 行, "夜": 行}, 2行かどうか)。
    昼と夜の行が揃っていればそれぞれ、1行しか無ければ（全日の行）両方のバンドを同じ行に向ける。
    20260922 の「全日タイプ」（C列を全日に戻して夜の行を隠した状態）もここで1行として拾える。"""
    k = namekey(nfc(dept_dtx14 or ""))
    day, night = loc.get((k, "昼")), loc.get((k, "夜"))
    if day and night:
        return {"昼": day, "夜": night}, True
    r = loc.get((k, "全日")) or day or night
    return ({"昼": r, "夜": r}, False) if r else ({}, False)


def frame_band_rows(frame, loc):
    """枠 → {band: 行}。見出しで引けなければ枠マスタの固定 row に落ちる（後方互換）。"""
    rows, _two = dept_band_rows(frame.get("dept_dtx14") or frame.get("dept"), loc)
    bands = BANDS if frame.get("shape") == "全日" else (frame.get("shape"),)
    fixed = int(frame.get("row") or 0)
    out = {}
    for b in bands:
        r = rows.get(b) or fixed
        if b and r:
            out[b] = int(r)
    return out


def frame_id_index(frames, loc, sheet):
    """{(行, band): frame_id}。枠マスタに無い行（管理当直・看護部など）は入らない。"""
    out = {}
    for f in frames:
        if f.get("sheet") != sheet:
            continue
        for b, r in frame_band_rows(f, loc).items():
            out.setdefault((r, b), f["frame_id"])
    return out


def shift_ranges(base, frames, loc, sheet):
    """行が増えた分だけ既定の走査範囲を広げる（ブロックごとに、そのブロックの枠の最大のずれ）。"""
    out = []
    for (a, b) in base:
        delta = 0
        for f in frames:
            if f.get("sheet") != sheet:
                continue
            fixed = int(f.get("row") or 0)
            if not (a <= fixed <= b):
                continue
            for r in frame_band_rows(f, loc).values():
                delta = max(delta, r - fixed)
        out.append((a, b + delta))
    return out


def last_date_col(rows):
    r3 = rows[2]
    return max(i + 1 for i, v in enumerate(r3) if v not in (None, ""))


def vba_trim(s):
    return s.strip(" ")


def wrapped_parens(s):
    return len(s) >= 1 and s[:1] in "(（" and s[-1:] in ")）"


def simulate_extract(rows, merged_a, ranges, mode, holidays):
    out = []
    lastcol = last_date_col(rows)
    for (r1, r2) in ranges:
        for i in range(r1, r2 + 1):
            ar = merged_a.get(i, i)
            dept = cellv(rows, ar, 1)
            dept = "" if dept is None else str(dept)
            if mode == "first" and ("放射線科" in dept or "管理当直" in dept):
                continue
            wt = cellv(rows, i, 3)
            wt = "" if wt is None else str(wt)
            for j in range(4, lastcol + 1):
                v = cellv(rows, i, j)
                if v is None or len(str(v)) == 0:
                    continue
                d = cellv(rows, 3, j)
                if isinstance(d, dt.datetime):
                    d = d.date()
                for line in str(v).split("\n"):
                    if mode == "duty":
                        t = vba_trim(line)
                        if not t or wrapped_parens(t):
                            continue
                    elif mode == "oncall":
                        t = vba_trim(line)
                        if not wrapped_parens(t):
                            continue
                    else:
                        t = line
                    if wt == "全日":
                        if is_holiday(d, holidays):
                            out.append((dept, t, d, "昼", i))
                            out.append((dept, t, d, "夜", i))
                        else:
                            out.append((dept, t, d, "夜", i))
                    else:
                        out.append((dept, t, d, wt, i))
    return out


def post_chuya(rows4):
    out = []
    for dept, b, d, band, src in rows4:
        b2 = str(b).replace("：", ":")
        if len(b2) == 0 or ":" not in b2:
            out.append((dept, b, d, band, src))
            continue
        shift = b2[1] if b2.startswith("(") else b2[:1]
        if (shift == "昼" and band == "夜") or (shift == "夜" and band == "昼"):
            continue
        if b2.startswith("("):
            b2 = b2.replace("(昼:", "(", 1).replace("(夜:", "(", 1)
        else:
            b2 = b2.replace("昼:", "", 1).replace("夜:", "", 1)
        out.append((dept, b2, d, band, src))
    return out


def post_shoki(rows4):
    return [r for r in rows4 if "初期研未定" not in str(r[0]) + str(r[1])]


def post_paren_half(rows4):
    return [(a, str(b).replace("（", "(").replace("）", ")"), d, band, src)
            for a, b, d, band, src in rows4]


def read_actual(path, sheet):
    """現物の人別シート → [(A,B,dateISO,D,E)]（検証用）。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        r = list(r) + [None] * 5
        a, b, c, d, e = r[0], r[1], r[2], r[3], r[4]
        if all(v in (None, "") for v in (a, b, c, d, e)):
            continue
        if isinstance(c, dt.datetime):
            c = c.date().isoformat()
        elif isinstance(c, dt.date):
            c = c.isoformat()
        else:
            c = str(c)
        out.append((str(a or ""), str(b or ""), c, str(d or ""), str(e or "")))
    wb.close()
    return out


def extract(xlsm, staffs_path, year, month, holidays=None):
    """現物DTX14 → 原子レコードのリスト（frame_id = f-r{n}/s-r{n}）。
    E列（現物の人別シート）を氏名の教師にしてログインIDへ解決する。"""
    holidays = holidays or load_holidays()
    people, exact, canon_map = load_staffs(staffs_path)

    def emap(sheet):
        m = {}
        for a, b, c, d, e in read_actual(xlsm, sheet):
            m.setdefault((a, b, c, d), []).append(e)
        return m

    em_duty, em_on = emap(P_DUTY), emap(P_ONCALL)

    rows1, ma1 = load_sheet_grid(xlsm, SHEET_FIRST)
    rows2, ma2 = load_sheet_grid(xlsm, SHEET_SECOND)

    # 行は見出しで割り出す（全日の1行 / 昼・夜の2行 のどちらの雛形でも同じ frame_id になる）
    frames = load_frames()
    loc1 = locate_rows(rows1, SHEET_FIRST)
    loc2 = locate_rows(rows2, SHEET_SECOND)
    fid1 = frame_id_index(frames, loc1, SHEET_FIRST)
    fid2 = frame_id_index(frames, loc2, SHEET_SECOND)
    rng1 = shift_ranges(RANGES_FIRST, frames, loc1, SHEET_FIRST)
    rng2 = shift_ranges(RANGES_SECOND, frames, loc2, SHEET_SECOND)

    def run(rows, ma, ranges, mode):
        return post_paren_half(post_shoki(post_chuya(
            simulate_extract(rows, ma, ranges, mode, holidays))))

    records = []
    unresolved = []

    def mk(rows4, em, oncall_flag, sheet_role, fidx):
        pre = "f-r" if sheet_role == "first" else "s-r"
        for dept, b, d, band, src in rows4:
            k = (dept, b, d.isoformat(), band)
            evs = em.get(k) or []
            e = evs.pop(0) if evs else ""
            core = b[1:-1] if oncall_flag and wrapped_parens(b) else b
            person = how = None
            if e:
                person, how = resolve_name(e, people, exact, canon_map)
            if person is None:
                unresolved.append((dept, b, d.isoformat(), band, e))
            records.append({
                "schema": "hospital.duty/1",
                "staff_login": person["login"] if person else None,
                "staff_code": person["code"] if person else None,
                "name": person["name"] if person else e,
                "name_dtx14": e,
                "name_disp": core,
                "date": d.isoformat(),
                "band": band,
                "frame_dept": dept,
                "frame_id": fidx.get((src, band)) or (pre + str(src)),
                "role": "first" if sheet_role == "first" else "second",
                "oncall": oncall_flag,
                "src_row": src,
                "note": "",
            })

    mk(run(rows1, ma1, rng1, "duty"), em_duty, False, "first", fid1)
    mk(run(rows2, ma2, rng2, "duty"), em_duty, False, "second", fid2)
    mk(run(rows1, ma1, rng1, "oncall"), em_on, True, "first", fid1)
    mk(run(rows2, ma2, rng2, "oncall"), em_on, True, "second", fid2)
    return records, unresolved
