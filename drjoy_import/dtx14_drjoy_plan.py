# -*- coding: utf-8 -*-
"""Dr.JOYの「勤務予定」エクスポート（at0029・種別=勤務予定）を読む。20260828

★当直の投入は **その日の勤務予定を丸ごと上書きする**ので、
  「いま登録されている勤務予定」を再現したうえで当直を足さないと、
  通常勤務・年休・救急部の遅出シフトなどを潰してしまう。
  給与連携項目一覧の「所定時間、休暇等」は文字列で休憩の位置が分からないが、
  こちらは **勤務種類・開始・終了・休憩** がそのまま入っている。

形（実測 R8.9月）:
  1シート＝1人。シート名は `00000000000 珊瑚 優子`（新採用者はコード無しで氏名だけ）
  1行目 : 勤務予定 / 出力期間 / 2026年09月度 / … / 職員コード / <code> / / 氏名 / <name>
  6-7行目: 日付 / 状況 / プリセット名 / 区分 / 所定時間(勤務種類・開始・終了) /
           休憩時間(開始・終了) / 休憩時間計 / 所定時間(休憩を除く)
  8行目〜 : `9/1 火` / … / `勤務日` / `通常勤務` / `08:30` / `17:15` / `12:00` / `13:00`
  ⚠週の区切りに `9/1〜9/6計` の小計行が挟まる（日付が空なので読み飛ばす）
  ★**1日に勤務種類が2つ以上あると、2つ目からは日付が空の「続きの行」に入る**。
    例）9/1 火 … 通常勤務 08:30-17:15 / （続きの行）当直(夜) 17:15-32:30
    20260828 まで日付の無い行を全部捨てていたため、
    **既に入っている当直が見えず**（＝入っていないと誤認して二重に書きに行く）、
    **予定を書き戻すときに2つ目の勤務種類・2つ目の休憩を落としていた**。
    → 日付の無い行でも、勤務種類＋時刻があれば直前の日に足す。
"""
from __future__ import annotations

import datetime as dt
import os
import re
from collections import defaultdict

import openpyxl

_DATE = re.compile(r"^(\d{1,2})/(\d{1,2})")
_HM = re.compile(r"^(\d{1,2}):(\d{2})$")
HDR_ROW = 7          # 見出しの2段目
DATA_ROW = 8


def _hm(v):
    """`08:30` / time / datetime → 分。読めなければ None。"""
    if v is None:
        return None
    if hasattr(v, "hour"):
        return v.hour * 60 + v.minute
    m = _HM.match(str(v).strip())
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def load(path, year=None, month=None, log=print):
    """→ {職員コード: {日付: {"区分":…, "work":[(種類,開始分,終了分)], "brk":[(開始分,終了分)]}}}

    ★職員コードはシート名の先頭から取る。無いシート（新採用者）は氏名で返す辞書にも入れる。
    """
    by_code, by_name, codes = defaultdict(dict), defaultdict(dict), {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    n_sheet = n_day = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) <= DATA_ROW:
            continue
        head = [str(x or "").strip() for x in (rows[0] or [])]
        code = ""
        if "職員コード" in head:
            i = head.index("職員コード")
            code = next((h for h in head[i + 1:i + 3] if h), "")
        name = ""
        if "氏名" in head:
            i = head.index("氏名")
            name = next((h for h in head[i + 1:i + 3] if h), "")
        if not code:
            m = re.match(r"^(\S+)\s", sn)
            if m and m.group(1).isdigit():
                code = m.group(1)
        if not (code or name):
            continue
        n_sheet += 1
        if code and name:
            codes[name.replace(" ", "").replace("　", "")] = code

        days = {}
        cur = None                             # ★続きの行を足す先
        for r in rows[DATA_ROW - 1:]:
            r = list(r) + [None] * 11
            m = _DATE.match(str(r[0] or "").strip())
            if m:
                mo, da = int(m.group(1)), int(m.group(2))
                y = year or dt.date.today().year
                if month and mo != month:
                    cur = None                 # 月をまたぐ表記（前後月の端）は対象外
                    continue
                try:
                    d = dt.date(y, mo, da)
                except ValueError:
                    cur = None
                    continue
                cur = days[d] = {"区分": str(r[3] or "").strip(), "work": [], "brk": []}
                n_day += 1
            elif cur is None:
                continue                       # 見出しの直後・月度計より後
            # ★日付のある行も無い行も、ここで同じように拾う
            k, s2, e2 = str(r[4] or "").strip(), _hm(r[5]), _hm(r[6])
            if k and s2 is not None and e2 is not None:
                cur["work"].append((k, s2, e2))
            bs, be = _hm(r[7]), _hm(r[8])
            if bs is not None and be is not None:
                cur["brk"].append((bs, be))
        if code:
            by_code[code] = days
        if name:
            by_name[name.replace(" ", "").replace("　", "")] = days
    wb.close()
    log(f"  Dr.JOYの勤務予定: {n_sheet}名 / {n_day}人日  （職員コードの分かる人 {len(codes)}名）")
    return by_code, by_name, codes
