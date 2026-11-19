# -*- coding: utf-8 -*-
"""dtx14_fill.py — アプリの当直表xlsx（または共有マスター）→ DTX14 に書き込む。20261104

⚠ 書けないときは SystemExit ではなく FillError（マスターに当直が無いときは NoDuties）を投げる（20260923）。
  SystemExit は呼び出し側（毎日15時の dtx14_daily_sync）の except Exception をすり抜けて、残りの月の PDF・
  交代記録の更新まで止めていた。CLI（main）だけが SystemExit に直す。
⚠ run_master は出力先（15時の反映では正本の DTX14 そのもの）を途中で書き換えない（20260923 第2陣）。
  当直表2シート・人別3シート・読み戻しの突合はすべてローカルの一時フォルダの写しで行い、全部が通って突合 ✅ のときだけ
  出力先を差し替える（publish: 同じフォルダの一時名に写して os.replace）。途中で落ちた／突合 ❌ なら出力先は元のまま FillError。
⚠ 雛形を読んだのに見出し（A列の科名・C列の昼/夜/全日）で引けない枠は、枠マスタの固定 row に落とさない（20260923 第2陣）。
  2行に割った雛形（11月以降）では固定 row が別の科の行を指すので、その科を書く月は FillError で止める（推測で書かない）。
⚠ 当番当直のグレー地に白字は、セルを塗るのではなく雛形の条件付き書式（当番当直シート × AK列の科名 × C列の昼/夜）が付ける
  （20260925 調べ: R8.10 の現物で 595959 地・白字。セルの塗りはどこにも無い）。昼/夜の行の色も条件付き書式
  （stopIfTrue）なので、セルを塗っても表に出ない。R8.11/R8.12 の雛形は規則の順番が崩れて当番当直と祝日の色が
  昼/夜の行の色の後ろに回り、一度も出ていなかった → fix_cf_order が毎回 R8.10 と同じ順に置き直し、
  check_toban が当番当直のマスの見え方（DisplayFormat）を data/toban.csv と照らして記録する。

用途（移行期・管理者 2026-09-14）: 10月はまだ従来どおり DTX14 に書いて事務担当へ出す運用なので、
変換器やアプリが作った当直表 xlsx の原子レコードを、事務担当が配った雛形
（例 ★R8.10月用 DTX14.9.5.12.xlsm）の「当直表 ファースト」「当直表 セカンド以下」の該当セルに写す。

書き方は現物 DTX14 と同じ:
  行 = 雛形の見出しから割り出す（dtx14_extract.locate_rows。枠マスタの固定 row は落とし先）
  列 = 4行目の日付シリアル（D列=1日）
  セル = 姓だけ（同姓は名を1文字ずつ足す＝duty_sheet.display_names）。当直を先、オンコールは (姓) で後ろ。複数は改行区切り

**全日タイプ／昼夜タイプは月ごとにデータから決める（20260922 管理者決定）**。
雛形が昼/夜の2行を持つ科（tools\dtx14_split_rows.py で割った11月以降の雛形）では:
  全日タイプ … その月の土日祝のすべての日で昼と夜の顔ぶれが同じ
                → 上の行にまとめて書き、上の行のC列を「全日」に戻し、下の行は空にして隠す
                （科名がA列の上の行にあるので、隠すのは必ず下の行）
  昼夜タイプ … 1日でも違う → C列は「昼」「夜」、両方の行を表示して行ごとに書く
毎回データから決め直す（15時の書き出しで交代が入るとタイプが変わる）ので、書き込みは
必ずラベルと表示/非表示をデータから上書きする＝前の状態には一切依存しない。
2行ある雛形では「昼:姓」「夜:姓」の書き分けは使わない（マスが分かれているので不要）。

全日の1行しか無い雛形（R8.9/R8.10 の現物）では今までどおり、昼と夜の顔ぶれが違う日だけ
「昼:姓」「夜:姓」（オンコールは「(昼:姓)」「(夜:姓)」）で書き分ける。両方のバンドにいる人は素の姓のまま。
並びは 当直 → オンコール、それぞれ 共通 → 昼だけ → 夜だけ。
読む側 dtx14_extract.post_chuya がこの記法を解くので、ここはその逆になっている。
雛形は Excel COM で開いて値だけ入れる（マクロ・入力規則・書式は触らない）。マクロは走らせない。

共有マスター1本から全科まとめて書くこともできる（--master。11月以降の移行期に、事務担当の
決裁・配布・手当計算が DTX14 のままでも回るようにするための橋渡し）。表示名（姓＋必要なら名）は
DTX14 のセルが科ごとなので**科ごとに**計算する（同姓の区別は科の中だけで足りる）。

usage: python -X utf8 tools\\dtx14_fill.py --in <当直表xlsx> --template <DTX14.xlsm> --out <出力.xlsm>
                                         [--year 2026 --month 10] [--staffs PATH] [--verify]
       python -X utf8 tools\\dtx14_fill.py --master <当直予定表_R8年10月.json> --template <DTX14.xlsm>
                                         --out <出力.xlsm> [--verify]
"""
import argparse
import calendar
import datetime as dt
import glob
import json
import os
import re
import shutil
import sys
import tempfile
import time
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402
import duty_sheet as D        # noqa: E402

XL_MACRO_ENABLED = 52         # xlOpenXMLWorkbookMacroEnabled
ROW_SERIAL = 4                # 日付シリアルの行
COL_DAY1 = 4                  # D列 = 1日


class FillError(Exception):
    """DTX14 に書けない（入力・雛形・マスターの問題）。呼び出し側が記録して次へ進めるように Exception の子。"""


class NoDuties(FillError):
    """マスターに当直の記録がまだ無い（どの科も出していない月）。毎日の反映では異常ではなく「書かない」。"""


class HeadingMissing(FillError):
    """雛形の見出し（A列の科名）で行が引けない枠がある（20260925: 事務へ「元の科名に戻して」と知らせるため）。
    depts = 見つからなかった科名（DTX14 の表記）。"""

    def __init__(self, msg, depts=()):
        super().__init__(msg)
        self.depts = list(depts)


class FileBusy(FillError):
    """正本の DTX14 が開かれていて差し替えられない（20260925: 事務へ「ファイルを閉じて」と知らせるため）。"""


def xl_text(v):
    """Excel に文字列として入れる: = + - @ で始まる文字列は先頭に ' を付ける（数式に化けない。' は表示されない）。
    氏名・注記など外から来た文字列をセルに入れるときは必ず通す（20260925）。"""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
        return "'" + v
    return v


def load_records(path, year, month, staffs):
    """当直表xlsx（新形式）→ 原子レコードとメンバー。旧テンプレ（当直1〜3/オンコール1〜6）も読める。"""
    frames = C.load_frames()
    holidays = C.load_holidays()
    people, exact, canon_map = C.load_staffs(staffs)
    try:
        bk = D.read_book(path)
    except Exception:
        bk = None
    if bk and bk.get("owner") and bk.get("rows"):          # 新形式（_メタ に owner がある）
        owner = bk["owner"]
        members = bk["members"]
        pidx = D.people_index(members)
        recs, rejects = C.import_rows(bk["rows"], owner, year, month, frames, holidays, *pidx,
                                      roster_label="この科のメンバー")
        C.apply_daytime_oncall(recs)
        return owner, recs, rejects, members, frames
    else:                                                    # 旧テンプレ（当直1〜3/オンコール1〜6）
        from aggregate import dept_from_filename
        rows = C.read_import_file(path, year, month)
        owner = dept_from_filename(path, frames)
        if not owner:
            raise FillError(f"ファイル名から科名が分かりません: {path}")
        recs, rejects = C.import_rows(rows, owner, year, month, frames, holidays, people, exact, canon_map)
        by_login = {p["login"]: p for p in people}
        members, seen = [], set()
        for r in recs:
            lg = r["staff_login"]
            if lg in seen:
                continue
            seen.add(lg)
            p = by_login.get(lg, {})
            sei, _, mei = C.namekey_display(p.get("name") or r["name"]).partition(" ")
            members.append({"login": lg, "code": p.get("code", ""), "sei": sei, "mei": mei,
                            "name": p.get("name") or r["name"], "kana": p.get("kana", ""),
                            "title": p.get("title", ""), "in_house": True, "unregistered": False})
        return owner, recs, rejects, members, frames


def with_sei(members):
    """姓名が分かれていないメンバー（マスターの roster は name だけ）に sei/mei を補う。"""
    out = []
    for m in members:
        if m.get("sei"):
            out.append(m)
            continue
        sei, _, mei = C.namekey_display(m.get("name") or "").partition(" ")
        out.append({**m, "sei": sei, "mei": mei})
    return out


def disp_map(members):
    """メンバー → {login: 表示名}。同姓の区別は科の中だけで足りる（DTX14のセルは科ごと）。"""
    ms = with_sei(members)
    return {m["login"]: d[0] for m, d in zip(ms, D.display_names(ms))}


# ------------------------------------------------------------------ 雛形の行の割り出し
def read_layout(template):
    """雛形の両シートを openpyxl（読み取り専用）で読み、{シート名: locate_rows の索引} を返す。"""
    import dtx14_extract as X
    loc = {}
    for sheet in (X.SHEET_FIRST, X.SHEET_SECOND):
        rows, _ = X.load_sheet_grid(template, sheet)
        loc[sheet] = X.locate_rows(rows, sheet)
    return loc


def frame_rows(f, loc):
    """枠 → (昼の行, 夜の行, 2行か)。前回の書き込みで全日に畳んだ状態（上の行が「全日」・すぐ下に「夜」の行）も 2行として拾う。
    見出しで引けないとき:
      雛形のそのシートを読んでいない（loc にシートが無い＝雛形なしの呼び方・テスト）→ 枠マスタの固定 row（後方互換）
      雛形のシートを読んだのに見出しが無い → None（20260923 第2陣: 2行に割った雛形では固定 row が別の科の行を指す。
      R8.9/R8.10/R8.11 の雛形はどれも全枠が見出しで引けることを確かめた＝現物では固定 row に落ちていなかった）"""
    import dtx14_extract as X
    sheet = f.get("sheet")
    lc = loc.get(sheet) or {}
    k = X.namekey(X.nfc(f.get("dept_dtx14") or f.get("dept") or ""))
    day, night, whole = lc.get((k, "昼")), lc.get((k, "夜")), lc.get((k, "全日"))
    if day and night:
        return day, night, True
    if whole and night == whole + 1:
        return whole, night, True
    r = whole or day or night
    if not r:
        if sheet in loc:
            return None
        r = int(f.get("row") or 0)
    return r, r, False


def unit_of(fid, fmap, rowmap):
    """枠 → 判定の単位 (sheet, 上の行, 下の行)。単位=科×役割（管理者 20260922）。
    2行ある全日枠（1枠で昼と夜）も、昼枠＋夜枠の対（ファーストの多く。同じ2行を指す）も同じ単位になる。
    2行が無い枠（全日1行の旧雛形）は None。"""
    f = fmap.get(fid)
    if not f or fid not in rowmap:
        return None
    d_row, n_row, two = rowmap[fid]
    if not two or f.get("shape") not in ("全日", "昼", "夜"):
        return None
    return (f["sheet"], d_row, n_row)


def decide_types(recs, fmap, rowmap, year, month, holidays, log=print, label=""):
    """2行ある単位（科×役割）ごとに 全日タイプ(True)／昼夜タイプ(False) をその月のデータから決める
    （20260922 管理者決定）。判定は土日祝だけ: 昼と夜の顔ぶれ（人＋当直/オンコール）が全部同じなら全日。
    平日に昼のレコードがある単位は（全日の展開規則に反するので）昼夜タイプに倒す。
    戻り: {frame_id: 全日タイプか}（単位に属する枠すべてに同じ値）。"""
    by = defaultdict(lambda: defaultdict(lambda: {"昼": set(), "夜": set()}))
    ufr = defaultdict(list)
    for fid, f in fmap.items():
        u = unit_of(fid, fmap, rowmap)
        if u:
            ufr[u].append(fid)
    for r in recs:
        u = unit_of(r.get("frame_id"), fmap, rowmap)
        if u:
            by[u][r["date"]][r.get("band") or ""].add((r.get("staff_login") or r.get("name"), bool(r.get("oncall"))))
    # 当番当直に当たる可能性がある科のファーストは全日に畳まない（管理者 20260922）:
    #   当番当直は必ず昼か夜のどちらかで、DTX14 の色付け（グレー地に白字）は行＝バンドごとに付く。
    #   全日1行だと昼と夜を分けて示せない。セカンド以下には当番当直が無いので従来どおりデータで決める。
    toban_depts = {C.namekey(d) for _dt, _b, d in C.load_toban()}
    out = {}
    for u, days in by.items():
        hol = [d for d in days if is_hol(d, holidays)]
        diff = [d for d in hol if days[d]["昼"] != days[d]["夜"]]
        wd_day = [d for d in days if not is_hol(d, holidays) and days[d]["昼"]]
        allday = not diff and not wd_day
        f = fmap[ufr[u][0]]
        toban_first = (f["sheet"] == "当直表 ファースト"
                       and C.namekey(f.get("dept_dtx14") or f["dept"]) in toban_depts)
        keep = bool(f.get("keep_day_night_rows"))     # 管理当直は昼・夜の2行のまま（元の雛形どおり。管理者 20260923）
        if allday and (toban_first or keep):
            allday = False
        for fid in ufr[u]:
            out[fid] = allday
        why = (f"土日祝 {len(hol)}日すべて昼=夜" if allday
               else "当番当直の対象科なので昼夜のまま" if toban_first and not diff and not wd_day
               else "管理当直は昼夜のまま" if keep and not diff and not wd_day
               else f"昼≠夜: {', '.join(str(int(d[8:10])) + '日' for d in (diff + wd_day)[:4])}"
               + ("…" if len(diff + wd_day) > 4 else ""))
        log(f"  {label}{f.get('dept_dtx14') or f['dept']}（{f['sheet'][4:]}）: "
            f"{'全日タイプ' if allday else '昼夜タイプ'}（{why}）")
    return out


def is_hol(iso, holidays):
    return C.is_holiday(dt.date.fromisoformat(iso), holidays)


def _fill_buckets(cell, recs, disp, fmap, rowmap=None, types=None):
    """原子レコードを {(sheet,row,day): {'duty':[(band,姓)],'oncall':[…]}} に積む。
    rowmap = {frame_id: (昼の行, 夜の行, 2行か)}（None なら枠マスタの固定 row。渡した rowmap に無い枠は FillError）。
    types = {frame_id: 全日タイプか}（2行ある全日枠だけ）。
    戻り: (置いた数, 飛ばした数)"""
    types = types or {}
    placed = skipped = 0
    for r in recs:
        f = fmap.get(r.get("frame_id"))
        if not f or not f.get("sheet") or not f.get("row"):
            skipped += 1                    # DTX14 に行が無い枠（管理者枠など）は書かない
            continue
        day = int(r["date"][8:10])
        band = str(r.get("band") or "")
        name = _disp_name(r, disp)
        if rowmap is None:
            d_row, n_row, two = int(f["row"]), int(f["row"]), False
        elif f["frame_id"] in rowmap:
            d_row, n_row, two = rowmap[f["frame_id"]]
        else:                               # plan_layout が先に止めるはず。万一でも固定 row に落とさない
            raise FillError(f"枠 {f['frame_id']}（{f.get('dept_dtx14') or f['dept']}）の行が雛形で見つかりません")
        if not two:
            row = d_row
        elif types.get(f["frame_id"]):
            row = d_row                     # 全日タイプ: 上の行にまとめる
        else:
            row = d_row if band == "昼" else n_row
        key = (f["sheet"], int(row), day)
        if f.get("shape") == "全日" and not two:
            cell[key]["allday"] = True      # 全日の1行しか無い雛形: 顔ぶれが違う日だけ 昼:/夜: で書き分ける
        bucket = cell[key]["oncall" if r.get("oncall") else "duty"]
        item = (band, name)
        if item not in bucket:
            bucket.append(item)
        placed += 1
    return placed, skipped


def _line(name, band, oncall):
    t = f"{band}:{name}" if band else name
    return f"({t})" if oncall else t


def cell_text(b):
    """1マスぶんの文字列。昼と夜の顔ぶれが同じなら今までどおり姓だけ。
    **全日の行**（昼と夜が同じ1マスに入る枠）で顔ぶれが違う日だけ、現物DTX14と同じ
    「昼:姓」「夜:姓」で書き分ける。昼・夜が別々の行にある枠（f-r39/f-r40 など）は
    マスが分かれているので書き分けない。平日の全日行（夜しか立たない）も書き分けない。"""
    def side(kind, band):
        return [n for bd, n in b.get(kind) or [] if bd == band]

    day = {k: side(k, "昼") for k in ("duty", "oncall")}
    night = {k: side(k, "夜") for k in ("duty", "oncall")}
    plain = (not b.get("allday")
             or not (day["duty"] or day["oncall"])
             or ({k: set(v) for k, v in day.items()} == {k: set(v) for k, v in night.items()}))
    out = []
    for kind in ("duty", "oncall"):
        oc = kind == "oncall"
        if plain:
            seen = set()
            for _bd, n in b.get(kind) or []:
                if n not in seen:
                    seen.add(n)
                    out.append(_line(n, "", oc))
            continue
        ds, ns = set(day[kind]), set(night[kind])
        out += [_line(n, "", oc) for n in day[kind] if n in ns]          # 昼も夜も = 素の姓
        out += [_line(n, "昼", oc) for n in day[kind] if n not in ns]
        out += [_line(n, "夜", oc) for n in night[kind] if n not in ds]
        # 昼/夜 以外のバンド（想定外）は素のまま末尾へ
        out += [_line(n, "", oc) for bd, n in b.get(kind) or [] if bd not in ("昼", "夜")]
    return "\n".join(out)


def _join(cell):
    return {k: cell_text(b) for k, b in cell.items()}


def plan_layout(frames, loc, owners, recs, holidays, log=print):
    """雛形の行割り + データからのタイプ判定 → (rowmap, types, units, clear_rows)。
    units = 2行ある単位（科×役割: 全日枠1つ、または昼枠＋夜枠の対）の [(sheet, 上の行, 下の行, 全日タイプか)]
    （ラベルと表示/非表示を毎回上書きする）。
    clear_rows = 書く前に日付セルを空にする行（データにある科の枠の行だけ。看護部・部長などは触らない）。"""
    fmap = {f["frame_id"]: f for f in frames}
    rowmap, missing = {}, []
    for f in frames:
        if f.get("sheet") and f.get("row"):
            rr = frame_rows(f, loc)
            if rr is None:
                missing.append(f)
            else:
                rowmap[f["frame_id"]] = rr
    # データに出てくる枠の owner も対象（マスターの科名が枠マスタの owner と違う時期のデータでも落とさない）
    owners = set(owners) | {fmap[r["frame_id"]].get("owner", fmap[r["frame_id"]]["dept"])
                            for r in recs if r.get("frame_id") in fmap}
    # 書く（＝空にして書き直す）科の枠が雛形の見出しで引けない → 推測で書かない（20260923 第2陣）
    miss = [f for f in missing if f.get("owner", f["dept"]) in owners]
    if miss:
        depts = []
        for f in miss:
            n = f.get("dept_dtx14") or f["dept"]
            if n not in depts:
                depts.append(n)
        raise HeadingMissing("雛形に見出し（A列の科名・C列の昼/夜/全日）が見つからない枠があるので書きません"
                             "（別の科の行に書かないため）: "
                             + "、".join(f"{f['frame_id']} {f.get('dept_dtx14') or f['dept']}（{f['sheet']}・{f.get('shape')}）"
                                         for f in miss), depts)
    types = decide_types(recs, fmap, rowmap, 0, 0, holidays, log)
    units, clear, seen_units = [], set(), set()
    for f in frames:
        if f["frame_id"] not in rowmap or f.get("owner", f["dept"]) not in owners:
            continue
        d_row, n_row, two = rowmap[f["frame_id"]]
        clear.add((f["sheet"], d_row))
        clear.add((f["sheet"], n_row))
        u = unit_of(f["frame_id"], fmap, rowmap)
        if u and u not in seen_units:
            seen_units.add(u)
            default = not (f["sheet"] == "当直表 ファースト"
                           and C.namekey(f.get("dept_dtx14") or f["dept"]) in {C.namekey(d) for _t, _b, d in C.load_toban()})
            units.append((f["sheet"], d_row, n_row, bool(types.get(f["frame_id"], default))))
    return rowmap, types, units, clear


def cells_from_records(recs, members, frames, loc=None, holidays=None, log=print):
    """原子レコード → plan {cells:{(sheet,row,day):'A\\nB\\n(C)'}, units, clear}。
    loc（read_layout）が無ければ枠マスタの固定 row に書く（旧来どおり）。"""
    fmap = {f["frame_id"]: f for f in frames}
    owners = {fmap[r["frame_id"]].get("owner", fmap[r["frame_id"]]["dept"]) for r in recs if r.get("frame_id") in fmap}
    rowmap, types, units, clear = plan_layout(frames, loc or {}, owners, recs, holidays or C.load_holidays(), log)
    cell = defaultdict(lambda: {"duty": [], "oncall": [], "allday": False})
    _fill_buckets(cell, recs, disp_map(members), fmap, rowmap, types)
    return {"cells": _join(cell), "units": units, "clear": clear}


def load_master(path, year=0, month=0):
    """共有マスターを読む。戻り: (master, year, month)。年月は指定が無ければマスター自身の値。
    path に dict を渡せばその写しを使う（20260924: 15時の反映が決裁済みの月に渡す承認済みの状態
    ＝kotai_record.approved_master。元の割当＋承認した交代）。"""
    if isinstance(path, dict):
        import copy
        m = copy.deepcopy(path)
        m.setdefault("depts", {})
    else:
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import master_store as MS
        m = MS.load(path, year or None, month or None)
    n_oc = sum(C.apply_daytime_oncall(ent.get("duties") or []) for ent in (m.get("depts") or {}).values())
    if n_oc:
        print(f"休診日の昼 {n_oc} 件を当直→オンコールに変えた（宿日直手当なしの日・holidays.csv daytime_oncall_only）")
    year = year or int(m.get("year") or 0)
    month = month or int(m.get("month") or 0)
    if not (year and month):
        raise FillError(f"マスターに年月がありません: {path if isinstance(path, str) else '（承認済みの状態）'}")
    return m, year, month


def cells_from_master(master, frames, loc=None, holidays=None, log=print):
    """マスターの全科 → (plan, 科ごとの表示名, 科ごとの件数)。表示名は科ごとに計算する。"""
    fmap = {f["frame_id"]: f for f in frames}
    depts = master.get("depts") or {}
    allrecs = [r for ent in depts.values() for r in (ent.get("duties") or [])]
    rowmap, types, units, clear = plan_layout(frames, loc or {}, set(depts), allrecs, holidays or C.load_holidays(), log)
    cell = defaultdict(lambda: {"duty": [], "oncall": [], "allday": False})
    disps, counts = {}, []
    for owner, ent in sorted(depts.items()):
        recs = ent.get("duties") or []
        members = ent.get("roster") or []
        disps[owner] = disp_map(members)
        placed, skipped = _fill_buckets(cell, recs, disps[owner], fmap, rowmap, types)
        counts.append((owner, placed, skipped, len(members)))
        log(f"  {owner}: {placed} 件"
            + (f"（DTX14に行が無い枠 {skipped} 件は飛ばした）" if skipped else "")
            + f" / メンバー {len(members)} 名")
    cells = _join(cell)
    return {"cells": cells, "units": units, "clear": clear,
            "changed": changed_cells(master, cells, fmap, rowmap, types, disps, log)}, disps, counts


def changed_cells(master, cells, fmap, rowmap, types, disps, log=print):
    """決裁済みの月だけ: 決裁のときの割当（各科の baseline）を、今日と同じ行の組み方（rowmap・types）で並べたときと
    文字が違うマス → 赤字にする（管理者 20260927。決裁後の DTX14 には承認された交代しか書かないので、
    違うマス＝承認された交代）。未確定の月・baseline の無い科は空。
    戻り: {(sheet, row, day): 赤字にする行の番号の集合（0 から）}。連名のマスは決裁のときに無かった行だけを赤にする
    （20261104 管理者: 片方だけ交代したのにマス全体が赤字だった）。足した行が無い（外しただけ）マスは全部の行。"""
    if not master.get("confirmed"):
        return set()
    import copy
    base = defaultdict(lambda: {"duty": [], "oncall": [], "allday": False})
    any_base = False
    for owner, ent in sorted((master.get("depts") or {}).items()):
        if ent.get("baseline") is None:
            continue
        any_base = True
        recs = copy.deepcopy(ent.get("baseline") or [])
        C.apply_daytime_oncall(recs)                      # 今日の割当と同じ扱い（休診日の昼はオンコール）
        _fill_buckets(base, recs, disps.get(owner) or disp_map(ent.get("roster") or []), fmap, rowmap, types)
    if not any_base:
        return set()
    btxt = _join(base)

    def same(a, b):          # 20261024: 氏名の並びが変わっただけのマスは交代ではない（顔ぶれと当/オで比べる）
        return sorted((a or "").split("\n")) == sorted((b or "").split("\n"))
    changed = {}
    for k, t in cells.items():
        if not t or same(t, btxt.get(k, "")):
            continue
        left = (btxt.get(k) or "").split("\n")
        new = set()
        for i, ln in enumerate(t.split("\n")):
            if ln in left:
                left.remove(ln)          # 同じ行が2つあっても1つずつ消し込む
            else:
                new.add(i)
        changed[k] = new or set(range(len(t.split("\n"))))
    if changed:
        log(f"  決裁のときから変わったマス（承認された交代）: {len(changed)} → 赤字")
    return changed


XL_EDGE_BOTTOM = 9
XL_AUTO = -4105                 # xlColorIndexAutomatic
RED_BGR = 0x0000FF              # 赤（COM の Color は BGR）
XL_NONE = -4142


def red_lines(c, text, lines):
    """セル c（text を入れたあと）の lines 番目の行だけ赤字にする。全部の行なら Font で一度に。
    Characters の位置は UTF-16 の数え方（改行も1字）。"""
    parts = text.split("\n")
    if set(range(len(parts))) <= set(lines):
        c.Font.Color = RED_BGR
        return
    pos = 1
    for i, p in enumerate(parts):
        n = len(p.encode("utf-16-le")) // 2
        if i in lines and n:
            c.GetCharacters(pos, n).Font.Color = RED_BGR
        pos += n + 1


XL_EDGE_TOP = 8


def _style(b):
    """罫線 → (LineStyle, Weight, ColorIndex, Color, TintAndShade)。線が無ければ None。
    雛形の線は ColorIndex 16 + TintAndShade -0.5（濃いグレー）。Color だけ写すと Tint が 0 に戻って薄くなる（20260922 実測）。"""
    ls = b.LineStyle
    if ls in (None, XL_NONE):
        return None
    try:
        tint = float(b.TintAndShade or 0.0)
    except Exception:
        tint = 0.0
    return (ls, b.Weight, b.ColorIndex, b.Color, round(tint, 4))


def _put_border(b, style):
    if style is None:
        b.LineStyle = XL_NONE
        return
    ls, wt, ci, col, tint = style
    b.LineStyle = ls
    b.Weight = wt
    if ci == -4105:                     # xlColorIndexAutomatic
        b.ColorIndex = -4105
    else:
        b.Color = col
        try:
            b.TintAndShade = tint
        except Exception:
            pass


def _ref_inner(ws, units):
    """昼と夜の間の線の手本: 展開されている単位の 上の行の下端（下端＝下の行の下端と違うもの）か 下の行の上端。無ければ細線。"""
    for (_sh, d_row, n_row, _a) in units:
        if ws.Rows(n_row).Hidden:
            continue
        ub, lb, lt = (_style(ws.Cells(d_row, COL_DAY1).Borders(XL_EDGE_BOTTOM)),
                      _style(ws.Cells(n_row, COL_DAY1).Borders(XL_EDGE_BOTTOM)),
                      _style(ws.Cells(n_row, COL_DAY1).Borders(XL_EDGE_TOP)))
        if ub and ub != lb:
            return ub
        if lt:
            return lt
    return (1, 1, -4105, 0, 0.0)               # xlContinuous / xlHairline / 自動色（雛形の昼夜の間の線）


def _ref_outer(ws, units):
    """ブロックの区切りの線の手本: どれかの単位の先頭行の上端。無ければ細線のグレー（ColorIndex 16・暗さ -0.5）。"""
    for (_sh, d_row, _n, _a) in units:
        st = _style(ws.Cells(d_row, COL_DAY1).Borders(XL_EDGE_TOP))
        if st:
            return st
    return (1, 2, 16, 8421504, -0.5)


def _shape_unit(ws, d_row, n_row, allday, ndays, ref_inner, ref_outer=None):
    """罫線と A 列の結合を、畳む／展開する に合わせて列ごとに直す（毎回上書き・前の状態に依存しない）。
    畳む: 下の行を隠すと、下の行が持っていた下端の線（ブロックの区切り）と A 列の結合セルの下端が消える。
          → A 列の結合を解き、上の行の下端に 下の行の下端の線 を写す。
    展開: 上の行の下端が「畳んだときに写した外側の線」になっていたら、昼と夜の間の線に戻す
          （下の行の上端に線があればそれ、無ければ手本）。元から昼夜の間に線を持つ行（管理当直の点線など）は触らない。"""
    last = COL_DAY1 + ndays - 1
    # A〜C 列の上下結合（A=科名は全ブロック、B=電話番号は一部のブロックで結合されている。結合セルの下端の線は
    # 隠した下の行に付くので線が消える＝20260922 実測: 脳神経外科 B列）。畳むときは解き、展開時は結合し直す
    for col in (1, 2, 3):
        rng = ws.Range(ws.Cells(d_row, col), ws.Cells(n_row, col))
        top, low = ws.Cells(d_row, col), ws.Cells(n_row, col)
        if allday:
            if rng.MergeCells:
                rng.UnMerge()                            # 値は上のセルに残る
        elif not top.MergeCells:
            if col == 1 or (col == 2 and top.Value not in (None, "") and low.Value in (None, "")):
                rng.Merge()
    nxt = n_row + 1                                      # 次のブロックの先頭行（隠し区切り行は飛ばす）
    while ws.Rows(nxt).Hidden and nxt < n_row + 4:
        nxt += 1
    prev = d_row - 1                                     # 前のブロックの末尾の見える行
    while prev > 4 and ws.Rows(prev).Hidden:
        prev -= 1
    for c in range(1, last + 1):
        # ブロックの先頭行の上端には必ず線を持たせる（管理者 20260922 の規則。前の行の下端、無ければ手本の外側の線）
        tb_b = ws.Cells(d_row, c).Borders(XL_EDGE_TOP)
        if not _style(tb_b):
            _put_border(tb_b, _style(ws.Cells(prev, c).Borders(XL_EDGE_BOTTOM)) or ref_outer)
        ub_b = ws.Cells(d_row, c).Borders(XL_EDGE_BOTTOM)
        ub = _style(ub_b)
        lb = _style(ws.Cells(n_row, c).Borders(XL_EDGE_BOTTOM))
        nt = _style(ws.Cells(nxt, c).Borders(XL_EDGE_TOP))
        if allday:
            # ブロックの下端の線を上の行の下端に持たせる（下の行の下端、無ければ次のブロックの上端。
            # Excel は隠した行をはさむと次の行の上端を描かないことがある＝20260922 実測）
            _put_border(ub_b, lb or nt)
        else:
            lt = _style(ws.Cells(n_row, c).Borders(XL_EDGE_TOP))
            if ub and (ub == lb or ub == nt) and ub != lt:  # 畳んだときに写した外側の線 → 内側の線へ戻す
                _put_border(ub_b, lt or ref_inner)
            elif c >= COL_DAY1 and not ub and not lt:      # 昼と夜の間に線が無い雛形（管理当直）→ 他と同じ内側の線を引く
                _put_border(ub_b, ref_inner)


PAPER = {"当直表 ファースト": (9, "A4"), "当直表 セカンド以下": (8, "A3")}   # xlPaperA4 / xlPaperA3（管理者 20260922）
PAGE_PT = {9: (842.0, 595.0), 8: (1191.0, 842.0)}                            # 横置きの (幅, 高さ) pt


def _page_setup(ws, sheet, ndays, log=print):
    """印刷設定: ファースト=A4横・セカンド以下=A3横、横1枚×縦1枚に収める（管理者 20260922）。
    印刷範囲は雛形のまま（行を挿入・非表示にしても Excel が追従する）。
    縦横どちらで縮尺が決まるかを計算して記録する（高さで決まっているあいだは行を高くすると縮尺が下がる）。"""
    paper, name = PAPER.get(sheet, (8, "A3"))
    ps = ws.PageSetup
    ps.PaperSize = paper
    ps.Orientation = 2                     # xlLandscape
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    try:
        rng = ws.Range(ps.PrintArea) if ps.PrintArea else ws.UsedRange
        w = rng.Width
        h = sum(ws.Rows(r).RowHeight for r in range(rng.Row, rng.Row + rng.Rows.Count) if not ws.Rows(r).Hidden)
        pw, ph = PAGE_PT[paper]
        aw = pw - ps.LeftMargin - ps.RightMargin
        ah = ph - ps.TopMargin - ps.BottomMargin
        sw, sh = aw / w, ah / h
        sc = min(sw, sh)
        lim = "高さ" if sh < sw else "幅"
        log(f"  印刷 {sheet}: {name}横 1枚  内容 {w:.0f}×{h:.0f}pt → 縮尺 {sc * 100:.0f}%（{lim}で決まる。"
            f"幅の使用 {w * sc / aw * 100:.0f}% / 高さの使用 {h * sc / ah * 100:.0f}%）")
    except Exception as e:                 # 記録だけなので失敗しても書き出しは続ける
        log(f"  印刷 {sheet}: 縮尺の計算に失敗 {e}")


def _text_width_pt(text, size):
    """文字列のおよその幅（pt）。全角＝size×1.05、半角＝size×0.72。
    画面の描画は印刷より広く、100% 表示では印刷で1行の電話番号が折り返した（20260922 実測）ので画面に合わせて多めに見る。"""
    return sum((size * 1.05 if ord(ch) > 0x7F else size * 0.72) for ch in text)


def _shrink_to_width(ws, row, col, min_size=10.0):
    """畳んだ行の B列（電話番号2人分など）が列幅で折り返して欠けないよう、折り返さない大きさまで文字を小さくする。
    文字ごとに大きさが違うセルは全体を同じ大きさにそろえる。戻り: 変えたか。"""
    cell = ws.Cells(row, col)
    v = cell.Value
    if not isinstance(v, str) or not v.strip() or not cell.WrapText:
        return False
    fs = cell.Font.Size
    if not isinstance(fs, (int, float)):
        fs = 0.0
        for i in range(1, min(len(v), 40) + 1):
            try:
                fs = max(fs, float(cell.GetCharacters(i, 1).Font.Size or 0))
            except Exception:
                break
    size = float(fs or 0) or 11.0
    avail = float(ws.Columns(col).Width) - 4.0
    longest = max(v.split("\n"), key=lambda t: _text_width_pt(t, 1.0))
    new = size
    while new > min_size and _text_width_pt(longest, new) > avail:
        new -= 1.0
    if new != size:
        cell.Font.Size = new
        return True
    return False


def _needed_height(ws, row, ndays, from_col=2):
    """行の中で最も行数の多いセル（B列の電話番号や積んだ氏名）が収まる高さ（pt）。
    Rows.AutoFit は文字ごとに大きさが違うセルで小さく出る（歯科口腔外科の B列）ので、改行の数と
    折り返し（WrapText のセルは列幅で折り返す）から見積もる。
    from_col=COL_DAY1 なら日付のセルだけを見る（展開した2行ブロック: B列の電話番号は2行にまたがって見せるので数えない）。"""
    need = 0.0
    for c in range(from_col, COL_DAY1 + ndays - 1 + 1):
        cell = ws.Cells(row, c)
        v = cell.Value
        if not isinstance(v, str) or not v.strip():
            continue
        fs = cell.Font.Size
        if not isinstance(fs, (int, float)):          # 文字ごとに大きさが違うセルは文字ごとの最大
            fs = 0.0
            for i in range(1, min(len(v), 40) + 1):
                try:
                    fs = max(fs, float(cell.GetCharacters(i, 1).Font.Size or 0))
                except Exception:
                    break
        size = float(fs or 0) or 11.0
        avail = max(float(ws.Columns(c).Width) - 4.0, size)
        lines = 0
        for line in v.split("\n"):
            if cell.WrapText:
                lines += max(1, -(-_text_width_pt(line, size) // avail))   # ceil
            else:
                lines += 1
        need = max(need, lines * size * 1.3 + 4)
    return need or 11.0 * 1.3 + 4


def fit_note(ws, log=print):
    """当直表の下の注記（『※１：内科(…)※2:…』）を印刷範囲の右端までに収める（管理者 20260922『注記が右にはみ出す』）。
    雛形は D:AH を結合し、先頭に全角空白を23個入れて右へ寄せていた → 30日の月は印刷範囲（AG まで）から末尾がはみ出し、
    31日の月も文字が結合幅より長い。空白を取って右寄せにし、結合を印刷範囲の右端までに縮め、入り切る大きさまで文字を小さくする。
    何度呼んでも同じ結果になる。戻り: 直した注記の数。"""
    pa = ws.PageSetup.PrintArea
    if not pa:
        return 0
    area = ws.Range(pa)
    last_col = area.Column + area.Columns.Count - 1
    n = 0
    for r in range(area.Row, area.Row + area.Rows.Count):
        cell = ws.Cells(r, COL_DAY1)
        v = cell.Value
        if not isinstance(v, str) or not v.lstrip(" 　").startswith("※"):
            continue
        text = v.lstrip(" 　")
        ma = cell.MergeArea
        end = ma.Column + ma.Columns.Count - 1
        if end > last_col:                                  # 印刷範囲の外へ出ている結合を縮める（右端の線は写す）
            rb = ws.Cells(r, end).Borders(10)               # xlEdgeRight
            style, weight, color = rb.LineStyle, rb.Weight, rb.Color
            ma.UnMerge()
            ws.Range(ws.Cells(r, COL_DAY1), ws.Cells(r, last_col)).Merge()
            if style not in (None, -4142):
                nb = ws.Cells(r, last_col).Borders(10)
                nb.LineStyle, nb.Weight, nb.Color = style, weight, color
        rng = cell.MergeArea
        if text != v:
            cell.Value = xl_text(text)
        rng.HorizontalAlignment = -4152                     # xlRight
        rng.WrapText = False
        rng.ShrinkToFit = False
        size = float(cell.Font.Size or 11)
        new = size
        while new > 8 and _text_width_pt(text, new) > float(rng.Width) - 8:
            new -= 0.5
        if new != size:
            rng.Font.Size = new
        log(f"  注記 {ws.Name} {rng.Address.replace('$', '')}: 右寄せ・{size:g}→{new:g}pt")
        n += 1
    return n


def fit_days(ws, ndays, log=print):
    """月にない日の列を隠し、ある日の列は見せる。印刷範囲の右端も月の最終日の列にそろえる（管理者 20261014:
    30日の11月で AH 列に翌月の『1 火』が出ていた。雛形は前の月から写すので、31日の月では見せ直す）。
    日の列は D（1日）〜AH（31日）。印刷範囲の右端が28〜31日の列にあるときだけ動かす（2月の正本を写した3月は
    右端が28日の列なので、広げ直す）。何度呼んでも同じ結果。戻り: 変えたか。"""
    last = COL_DAY1 + ndays - 1
    changed = False
    for c in range(COL_DAY1 + 28, COL_DAY1 + 31):            # 29〜31日の列
        hide = c > last
        if bool(ws.Columns(c).Hidden) != hide:
            ws.Columns(c).Hidden = hide
            changed = True
    pa = ws.PageSetup.PrintArea
    if pa:
        area = ws.Range(pa)
        right = area.Column + area.Columns.Count - 1
        if COL_DAY1 + 27 <= right <= COL_DAY1 + 30 and right != last:
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(area.Row, area.Column),
                                              ws.Cells(area.Row + area.Rows.Count - 1, last)).Address
            changed = True
    if changed:
        log(f"  {ws.Name}: {ndays}日の月 → 月にない日の列を隠し、印刷範囲を {ndays}日の列まで")
    if ndays == 31:
        n = _copy_row_lines(ws, COL_DAY1 + 29, COL_DAY1 + 30)
        if n:
            log(f"  {ws.Name}: 31日の列の横線を30日の列にそろえた（{n} か所）")
            changed = True
    return changed


def _copy_row_lines(ws, src_col, dst_col):
    """src_col の各行の上下の線（種類・太さ・色）を dst_col に写す。違うところだけ。戻り: 写した数。
    20261025 管理者「ファースト表の31日だけ罫線が違う」: 雛形の31日の列（AH）だけ、横線が薄い灰色の細線だった
    （30日の列は黒の極細線）。表の外枠（右端の縦線）は触らない。何度呼んでも同じ結果。"""
    pa = ws.PageSetup.PrintArea
    last = ws.Range(pa).Row + ws.Range(pa).Rows.Count - 1 if pa else int(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)
    n = 0
    for r in range(3, last + 1):
        for edge in (8, 9):                 # xlEdgeTop / xlEdgeBottom
            s, d = ws.Cells(r, src_col).Borders(edge), ws.Cells(r, dst_col).Borders(edge)
            want = (s.LineStyle, s.Weight, s.Color)
            if (d.LineStyle, d.Weight, d.Color) == want:
                continue
            if want[0] in (None, -4142):    # xlNone
                d.LineStyle = -4142
            else:
                d.LineStyle, d.Weight, d.Color = want
            n += 1
    return n


# セカンド以下の右上の凡例（管理者 20261022: 事務担当『毎回乱れる』→ ファーストと同じ作りにする）。
# 雛形はセルの文字（氏名：当直＋宿日直／氏名：当直／(氏名)：オンコール）に、右中かっこと「院内常駐」の文字枠の図形を重ねていた。
# 図形は画面の拡大率や PC で文字とずれ、日付を打ち直すたびに手で動かしていた。ファーストは全部セルの文字（・・・院内常駐）なので、
# 図形を外して同じ書き方にする。・・・の位置は BIZ UDPゴシック 太字の字幅で半角空白を詰めて揃えた（PIL で実測。大きさに依らない）。
# 大きさは雛形のセカンドのまま 20pt（A3 に縮めて刷るので、ファーストの 18pt にすると小さく見える）。
# 更新の印（管理者 20261025）: 凡例には日付を入れない（凡例の中に入れると当直・オンコールの説明に紛れて見にくい）。
#   右上の角（凡例の右隣 AF1:AG1）に「R8.10.25／16時更新」を1つだけ出す（stamp_updated）。凡例の結合は AE までに縮める
#   （セカンドの雛形は Z1:AG1 で、印の置き場が凡例の中に隠れていた）。ファーストの凡例も決まった文字で書き直す
#   （当番当直の灰色の見本の図形は3行目の頭にあるので残す）。
#   ⚠ 28・29日の月（2月）は fit_days が AF〜AH を隠すので、印も隠れる（置き場は R9.2月分の前に決める）。
LEGEND_FIRST_LINES = (
    "     氏名  ：  当直         ・・・院内常駐 ",
    "   (氏名) ：  オンコール ・・・呼ばれたら来る",
    "             :  当番当直",
)
LEGEND_SECOND_LINES = (
    "     氏名  ：  当直 + 宿日直 ・・・院内常駐",
    "     氏名  ：  当直               ・・・院内常駐",
    "   (氏名) ：  オンコール      ・・・呼ばれたら来る",
)
LEGEND_SIZES = {"当直表 ファースト": 18, "当直表 セカンド以下": 20}        # セカンドは A3 に縮めて刷るので大きめ
STAMP_SIZES = {"当直表 ファースト": 14, "当直表 セカンド以下": 18}
LEGEND_LAST_COL = 31                     # 凡例の結合の右端＝AE（AF・AG は更新の印）
STAMP_COL = 32                           # AF
MSO_SHAPE_RIGHT_BRACE = 32


def legend_text(sheet):
    """凡例の文字（日付は入れない）。"""
    return "\n".join(LEGEND_FIRST_LINES if sheet == SHEET_FIRST else LEGEND_SECOND_LINES)


def stamp_text(now=None):
    """更新の印『R8.10.25』『16時更新』の2行。時刻は実際に書いた時刻（手で流せば 18時にもなる。それでよい＝管理者 20260922）。"""
    now = now or dt.datetime.now()
    return f"R{now.year - 2018}.{now.month}.{now.day}\n{now.hour}時更新"


def _legend_cell(ws):
    for c in range(COL_DAY1, COL_DAY1 + 34):
        v = ws.Cells(1, c).Value
        if isinstance(v, str) and "オンコール" in v:
            return ws.Cells(1, c)
    return None


def fit_legend(ws, log=print, now=None):
    """1行目の凡例を決まった文字で書き直し、結合の右端を AE までにする。何度呼んでも同じ結果。戻り: 変えたか。
    セカンド以下は、雛形の右中かっこと「院内常駐」の図形も外す。15時の書き込みと、公開用 PDF の写し（書き込まない月も）で呼ぶ。"""
    if ws.Name not in (SHEET_FIRST, SHEET_SECOND):
        return False
    cell = _legend_cell(ws)
    if cell is None:
        return False
    old = cell.Value
    removed = 0
    for i in range(int(ws.Shapes.Count), 0, -1):
        sh = ws.Shapes.Item(i)
        try:
            if int(sh.TopLeftCell.Row) != 1:
                continue
            txt = ""
            try:
                txt = sh.TextFrame2.TextRange.Text or ""
            except Exception:
                pass
            if ws.Name == SHEET_SECOND and ("院内常駐" in txt or int(sh.AutoShapeType) == MSO_SHAPE_RIGHT_BRACE):
                sh.Delete()
                removed += 1
        except Exception:
            continue
    rng = cell.MergeArea
    narrowed = False
    if rng.Column + rng.Columns.Count - 1 > LEGEND_LAST_COL and cell.Column < LEGEND_LAST_COL:
        rng.UnMerge()
        ws.Range(ws.Cells(1, cell.Column), ws.Cells(1, LEGEND_LAST_COL)).Merge()
        rng = cell.MergeArea
        narrowed = True
    text = legend_text(ws.Name)
    rng.Value = xl_text(text)
    rng.HorizontalAlignment = -4131      # xlLeft
    rng.VerticalAlignment = -4108        # xlCenter
    rng.WrapText = True
    rng.ShrinkToFit = False
    font = rng.Font
    font.Name, font.Size, font.Bold, font.Italic, font.Color = "BIZ UDPゴシック", LEGEND_SIZES[ws.Name], True, False, 0
    changed = old != text or removed > 0 or narrowed
    if changed:
        log(f"  凡例 {ws.Name} {rng.Address.replace('$', '')}: 決まった文字で書き直し（図形 {removed} 個を外した"
            + ("・結合を AE までに縮めた" if narrowed else "") + "）")
    return changed


def stamp_updated(ws, now=None):
    """右上の角（AF1:AG1。ファーストの雛形は AF1:AH1 の結合のまま）に更新の印『R8.10.25／16時更新』を1つだけ書く
    （管理者 20260922。20261025: 凡例の中に入れるのをやめて、ここに戻した）。凡例の結合の中なら書かない（先に fit_legend）。"""
    if ws.Name not in (SHEET_FIRST, SHEET_SECOND):
        return False
    c = ws.Cells(1, STAMP_COL)
    ma = c.MergeArea
    if ma.Column != STAMP_COL:
        return False
    if ma.Columns.Count == 1:
        ws.Range(ws.Cells(1, STAMP_COL), ws.Cells(1, STAMP_COL + 1)).Merge()
        ma = c.MergeArea
    ma.Value = stamp_text(now)
    ma.WrapText = True
    ma.HorizontalAlignment = -4108      # xlCenter
    ma.VerticalAlignment = -4108
    ma.Font.Name = "BIZ UDPゴシック"
    ma.Font.Size = STAMP_SIZES[ws.Name]
    ma.Font.Bold = True
    ma.Font.Color = 0
    return True


def _serial(d):
    return (d - dt.date(1899, 12, 30)).days


# ------------------------------------------------------------------ 当番当直のグレー地に白字（20260925）
SHEET_FIRST = "当直表 ファースト"
SHEET_SECOND = "当直表 セカンド以下"
TOBAN_BG = 0x595959       # sheet_style.C_TOBAN_BG と同じ灰（R =G =B なので Excel の BGR でも同じ値）
TOBAN_INK = 0xFFFFFF
# 交代のあった当番当直のマスは赤地に白字（管理者 20260929）。雛形の当番当直の条件付き書式（灰地に白字）が文字の赤より
# 勝つので、そのマスだけに一番上の条件付き書式を足す。目印の式で見分け、毎回すべて消してから付け直す
KOTAI_MARK = "当番当直の交代"
KOTAI_MARK_FORMULA = f'=ISTEXT("{KOTAI_MARK}")'


def clear_kotai_marks(ws):
    """前の回に足した「交代のあった当番当直」の条件付き書式を消す。戻り: 消した数。"""
    fcs = ws.Cells.FormatConditions
    n = 0
    for i in range(fcs.Count, 0, -1):
        try:
            if KOTAI_MARK in str(fcs(i).Formula1):
                fcs(i).Delete()
                n += 1
        except Exception:
            continue
    return n


def mark_kotai_toban(ws, cells):
    """cells {(行, 日)} を赤地に白字の太字にする（一番上の条件付き書式・ここで止める）。"""
    for row, day in sorted(cells):
        fc = ws.Cells(row, COL_DAY1 + day - 1).FormatConditions.Add(2, 1, KOTAI_MARK_FORMULA)   # xlExpression（演算子は使われない）
        fc.SetFirstPriority()
        fc.Interior.Color = RED_BGR
        fc.Font.Color = TOBAN_INK
        fc.Font.Bold = True
        fc.StopIfTrue = True


def toban_cells(frames, loc, year, month, toban=None):
    """当番当直のマス {(行, 日)}（その月の data/toban.csv × ファースト表の見出しで引いた昼/夜の行）。
    DTX14 では色は雛形の条件付き書式が付けるので、これは check_toban の期待値（COM を使わない）。
    昼と夜の行が無い科（全日1行の旧雛形）は昼夜を分けて示せないので入れない。"""
    ym = f"{year:04d}-{month:02d}-"
    firsts = [f for f in frames if f.get("sheet") == SHEET_FIRST and f.get("role") == "first"]
    out = set()
    for d, b, dept in (C.load_toban() if toban is None else toban):
        if not str(d).startswith(ym) or b not in ("昼", "夜"):
            continue
        k = C.namekey(dept)
        for f in firsts:
            if C.namekey(f.get("dept_dtx14") or f.get("dept") or "") != k:
                continue
            rr = frame_rows(f, loc)
            if rr and rr[2]:
                out.add((rr[1] if b == "夜" else rr[0], int(str(d)[8:10])))
                break
    return out


def _cf_kind(formula):
    """条件付き書式の式 → 種類（fix_cf_order が並べる4つ。ほかは None）。"""
    f = str(formula or "").replace(" ", "")
    if "当番当直!" in f:
        return "toban"
    if f == "=WEEKDAY(D$4)=7":
        return "sat"
    if f == "=WEEKDAY(D$4)=1":
        return "sun"
    if "休日判定用!" in f:
        return "hol"
    return None          # 交代のあった当番当直の目印（KOTAI_MARK）もここ＝並べ替えの対象外


CF_ORDER = ("toban", "sat", "sun", "hol")      # R8.10 の現物の順（当番当直 → 土 → 日 → 祝日 → 昼/夜の行）


def _main_rules(ws):
    """表の本体（D5 を含む範囲）に掛かった条件付き書式 [(優先順位, 種類, FormatCondition)]。管理当直の行・AD9 は除く。"""
    out = []
    fcs = ws.Cells.FormatConditions
    for i in range(1, fcs.Count + 1):
        fc = fcs(i)
        try:
            area = fc.AppliesTo
            if not (area.Row <= 5 <= area.Row + area.Rows.Count - 1 and area.Column <= COL_DAY1):
                continue
            out.append((int(fc.Priority), _cf_kind(fc.Formula1), fc))
        except Exception:                 # 式を持たない種類（データバーなど）は並べない
            continue
    return sorted(out, key=lambda t: t[0])


def fix_cf_order(ws, log=print):
    """「当直表 ファースト」の条件付き書式を R8.10 の現物と同じ順（当番当直 → 土 → 日 → 祝日 → 昼/夜の行）に置き直す。
    すでにその順なら何もしない（何度呼んでも同じ結果）。戻り: 置き直したか。当番当直の規則が無ければ記録だけ。"""
    rules = _main_rules(ws)
    have = [k for k in CF_ORDER if any(r[1] == k for r in rules)]
    if "toban" not in have:
        log(f"  ⚠ {ws.Name}: 当番当直の条件付き書式が雛形にありません（グレー地に白字が出ません）")
    if [r[1] for r in rules[:len(have)]] == have:
        return False
    for k in reversed(have):             # 後ろの種類から順に先頭へ → 最後に当番当直が一番上になる
        fc = next(r[2] for r in _main_rules(ws) if r[1] == k)
        fc.SetFirstPriority()
    log(f"  {ws.Name}: 条件付き書式の順番を直した（{' → '.join(have)} → 昼/夜の行。当番当直と祝日の色が出るように）")
    return True



# 条件付き書式が見ている表の範囲を広げる（20260925）: 雛形は 休日判定用!$A$2:$A$50（祝日の色は $A$1:$A$50）＝2027-01-01 まで、
# 当番当直!$A$2:$A$449＝いまの表の最後の行まで しか見ていない。1月分以降の祝日・当番当直の色が出なくなるので、
# 休日判定用は 1000 行、当番当直は 5000 行まで見るように書き換える（何度呼んでも同じ結果）。
CF_RANGE_MIN = {"休日判定用": 1000, "当番当直": 5000}
# 当番当直の規則（灰地）の字は太字の白（管理者 20261031: 細い白字は灰地で弱く見える）。雛形は太字の指定が無く、
# 字の色はテーマの「背景1」。⚠ COM で x14 の規則（別シートを引く式）の Font.Color を変えると Excel ごと落ちた
# （20261031 実測）ので、範囲の書き換えと同じく XML の dxf を直す。何度流しても同じ結果。
TOBAN_BG_HEX = "595959"
TOBAN_FONT_XML = '<font><b/><color rgb="FFFFFFFF"/></font>'
TOBAN_RULE_RE = re.compile(r"<x14:cfRule\b.*?</x14:cfRule>", re.S)


def widen_cf_refs_file(src, log=print):
    """src（xlsm）の中の各シートの条件付き書式の式で、休日判定用・当番当直 の範囲の終わりを CF_RANGE_MIN まで広げた
    写しを一時フォルダに作る。戻り: (写しのパス or src, 書き換えた式の数)。
    ⚠ COM の FormatCondition.Modify は相対参照（D$3・$C3）を「選んでいるセル」基準で解釈し、見えない Excel では
    選択が効かず D$3→G$3 のようにずれた（20260925 実測）。ファイルの XML の中では式は適用範囲の左上基準で
    保存されているので、範囲の数字だけを書き換えれば相対参照には触れない。"""
    import re as _re
    import tempfile
    import zipfile
    pats = [(_re.compile(_re.escape(sh) + r"!(\$[A-Z]+\$\d+):(\$[A-Z]+)\$(\d+)"), sh, lo)
            for sh, lo in CF_RANGE_MIN.items()]
    fpat = _re.compile(r"(<(?:x14:|xm:)?f(?:ormula)?>)(.*?)(</(?:x14:|xm:)?f(?:ormula)?>)", _re.S)
    changed = 0
    bolded = 0

    def bold_toban(m):
        nonlocal bolded
        rule = m.group(0)
        if "当番当直!" not in rule or f'bgColor rgb="FF{TOBAN_BG_HEX}"' not in rule:
            return rule
        new = _re.sub(r"<font>.*?</font>|<font/>", TOBAN_FONT_XML, rule, count=1, flags=_re.S)
        if new != rule:
            bolded += 1
        return new

    def fix_formula(text):
        nonlocal changed
        new = text
        for pat, sh, lo in pats:
            new = pat.sub(lambda m: f"{sh}!{m.group(1)}:{m.group(2)}${max(int(m.group(3)), lo)}", new)
        if new != text:
            changed += 1
        return new

    with zipfile.ZipFile(src) as zin:
        items = [(zi, zin.read(zi.filename)) for zi in zin.infolist()]
    out_items = []
    for zi, data in items:
        if zi.filename.startswith("xl/worksheets/sheet") and zi.filename.endswith(".xml"):
            x = data.decode("utf-8")
            if "conditionalFormatting" in x:
                def one_block(m):
                    return fpat.sub(lambda f: f.group(1) + fix_formula(f.group(2)) + f.group(3), m.group(0))
                x2 = _re.sub(r"<(?:x14:)?conditionalFormatting\b.*?</(?:x14:)?conditionalFormatting>", one_block, x, flags=_re.S)
                x2 = TOBAN_RULE_RE.sub(bold_toban, x2)
                data = x2.encode("utf-8")
        out_items.append((zi, data))
    if not changed and not bolded:
        return src, 0
    dst = os.path.join(tempfile.mkdtemp(prefix="touchoku_cf_"), os.path.basename(src))
    with zipfile.ZipFile(dst, "w") as zout:
        for zi, data in out_items:
            zout.writestr(zi, data, compress_type=zi.compress_type)
    if changed:
        log(f"  条件付き書式が見る 休日判定用・当番当直 の範囲を広げた（{changed} 式。1月分以降の祝日・当番当直の色が出るように）")
    if bolded:
        log(f"  当番当直の条件付き書式の字を太字の白にした（{bolded} 規則）")
    return dst, changed + bolded


def _is_toban_look(cell):
    df = cell.DisplayFormat
    return int(df.Interior.Color) == TOBAN_BG and int(df.Font.Color) == TOBAN_INK


def check_toban(ws, expect, ndays, log=print, full=False):
    """当番当直のマスが グレー地に白字 で見えているか（DisplayFormat＝条件付き書式を当てたあとの見え方）を、
    data/toban.csv から作った expect {(行, 日)} と照らす。記録するだけで書き込みは止めない。
    full=False（15時の反映）は expect のマスだけ見る（全マスは1か月 8 秒かかる＝20260925 実測）。
    full=True（テスト）は AK列に科名がある見えている行（ファーストの診療科）の全マスを見て、余分な灰も拾う。
    戻り: (出ていないマス, 余分なマス)。"""
    if full:
        rows = [r for r in range(5, ws.UsedRange.Row + ws.UsedRange.Rows.Count)
                if ws.Cells(r, 37).Value not in (None, "") and not ws.Rows(r).Hidden]
        todo = [(r, day) for r in rows for day in range(1, ndays + 1)]
    else:
        todo = sorted(expect)
    got = {(r, day) for r, day in todo if _is_toban_look(ws.Cells(r, COL_DAY1 + day - 1))}
    miss, extra = sorted(set(expect) - got), sorted(got - set(expect))
    if not miss and not extra:
        log(f"  当番当直: {len(expect)} マスがグレー地に白字（data/toban.csv と一致）")
    else:
        log(f"  ⚠ 当番当直: data/toban.csv と見え方が違う（出ていない {len(miss)}・余分 {len(extra)}）"
            f" 例 {', '.join(f'行{r} {d}日' for r, d in (miss + extra)[:5])}"
            "（雛形の当番当直シート・C列の昼/夜・AK列の科名を確かめる）")
    return miss, extra


def write_with_com(template, out, plan, year, month, log=print):
    """雛形（xlsm）を Excel COM で開いて値だけ入れ、out に保存する。
    plan = {cells, units, clear}（cells_from_records / cells_from_master）。
    書く前に、データにある科の行の日付セルを空にする（前回の書き込みや雛形の残りを消す）。
    2行ある全日枠は、タイプに合わせて C列のラベルと下の行の表示/非表示を毎回上書きする。"""
    import calendar
    import win32com.client
    cells = plan["cells"] if isinstance(plan, dict) else plan
    units = plan.get("units", []) if isinstance(plan, dict) else []
    clear = plan.get("clear", set()) if isinstance(plan, dict) else set()
    toban = plan.get("toban") if isinstance(plan, dict) else None
    changed = plan.get("changed", set()) if isinstance(plan, dict) else set()
    ndays = calendar.monthrange(year, month)[1]
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3          # msoAutomationSecurityForceDisable（マクロを走らせない）
    xl.ScreenUpdating = False
    opened = template
    try:
        # 1月分以降の祝日・当番当直も色が出るように、条件付き書式の範囲を広げた写しを開く（XML で。20260925）
        opened, _n = widen_cf_refs_file(template, log)
        wb = xl.Workbooks.Open(os.path.abspath(opened), UpdateLinks=0, ReadOnly=True)
        xl.Calculation = -4135         # xlCalculationManual（ブックを開いてからでないと設定できない。保存前に自動へ戻す）
        written = 0
        names = [w.Name for w in wb.Worksheets]
        if SHEET_FIRST in names:       # 当番当直のグレー地に白字・祝日の色が出る順に（20260925）
            clear_kotai_marks(wb.Worksheets(SHEET_FIRST))     # 前の回の「交代のあった当番当直」を消す（20260929）
            fix_cf_order(wb.Worksheets(SHEET_FIRST), log)
        for sn in (SHEET_FIRST, SHEET_SECOND):                 # 月にない日の列を隠す（印刷範囲も。20261014）
            if sn in names:
                fit_days(wb.Worksheets(sn), ndays, log)
        tob_chg ={(r, d) for (s, r, d) in changed if s == SHEET_FIRST and (r, d) in (toban or ())}
        sheets = sorted({k[0] for k in cells} | {u[0] for u in units} | {c[0] for c in clear})
        for sheet in sheets:
            ws = wb.Worksheets(sheet)
            # 列の確認: 4行目のシリアルが対象月の1日か
            first = dt.date(year, month, 1)
            v = ws.Cells(ROW_SERIAL, COL_DAY1).Value
            v = _serial(v.date()) if hasattr(v, "date") else int(v or 0)
            if v != _serial(first):
                raise FillError(f"{sheet}: D4 のシリアル {v} が {first} ({_serial(first)}) ではありません。雛形の月が違います")
            for (sh, row) in sorted(clear):
                if sh == sheet:
                    rg = ws.Range(ws.Cells(row, COL_DAY1), ws.Cells(row, COL_DAY1 + ndays - 1))
                    rg.ClearContents()
                    rg.Font.ColorIndex = XL_AUTO       # 前の日の赤字（承認された交代）を戻してから付け直す（20260927）
            my_units = [u for u in units if u[0] == sheet]
            ref_inner = _ref_inner(ws, my_units)
            ref_outer = _ref_outer(ws, my_units)
            base_h = {}
            for (sh, d_row, n_row, allday) in my_units:
                ws.Cells(d_row, 3).Value = "全日" if allday else "昼"
                ws.Cells(n_row, 3).Value = "夜"
                ws.Rows(n_row).Hidden = False
                base_h[n_row] = ws.Rows(n_row).RowHeight     # 隠すと RowHeight が 0 を返すので先に控える
                ws.Rows(n_row).Hidden = bool(allday)
                _shape_unit(ws, d_row, n_row, allday, ndays, ref_inner, ref_outer)
            # ⚠ Excel の画面は「隠した行のすぐ下が雛形の隠し区切り行（高さ0）」だと、そのブロックの下端の線を描かない
            #   （20260922 実測: ファースト 放射線科→区切り行45→検査技師）。区切り行を 0.75pt の見える行にすると描く。
            hidden_n = {n for (_s, _d, n, a) in my_units if a}
            for n in hidden_n:
                r = n + 1
                if r not in hidden_n and ws.Rows(r).Hidden:
                    ws.Rows(r).Hidden = False
                    ws.Rows(r).RowHeight = 0.75
            for (sh, row, day), text in cells.items():
                if sh != sheet:
                    continue
                c = ws.Cells(row, COL_DAY1 + day - 1)
                if c.Value not in (None, ""):
                    log(f"  ⚠ {sheet} 行{row} {day}日 に既に『{c.Value}』があります → 上書き")
                c.Value = xl_text(text)     # 氏名が = + - @ で始まっても数式にしない（20260925）
                if (sh, row, day) in changed:
                    # 決裁のときから変わった行（承認された交代）だけ赤字（管理者 20260927・行ごとは 20261104）
                    red_lines(c, text, changed[(sh, row, day)])
                written += 1
            # 畳んだ行の高さ: B列の電話番号（2行ぶん）や積んだ氏名が収まるように内容に合わせて広げる。
            #   基準の高さは隠した下の行が持ち続ける（展開時は上の行を下の行と同じ高さに戻す）
            for (sh, d_row, n_row, allday) in my_units:
                if allday:
                    for col in (2, 3):
                        _shrink_to_width(ws, d_row, col)
                if sheet == "当直表 セカンド以下":
                    # 事務担当のマクロ「行の高さを7ポイント増やす_非表示行除外」と同じ: 見える行を AutoFit して +7pt
                    for r in ((d_row,) if allday else (d_row, n_row)):
                        ws.Rows(r).AutoFit()
                        ws.Rows(r).RowHeight = float(ws.Rows(r).RowHeight) + 7
                elif allday:
                    est = _needed_height(ws, d_row, ndays)
                    ws.Rows(d_row).AutoFit()                  # 折り返しの実測は Excel に任せ、見積もりと大きい方を採る
                    ws.Rows(d_row).RowHeight = max(base_h[n_row], est, float(ws.Rows(d_row).RowHeight))
                else:                                         # 展開: 昼と夜を同じ基準の高さに。氏名が積まれた行だけ広げる
                    ws.Rows(d_row).RowHeight = max(base_h[n_row], _needed_height(ws, d_row, ndays, COL_DAY1))
                    ws.Rows(n_row).RowHeight = max(base_h[n_row], _needed_height(ws, n_row, ndays, COL_DAY1))
            if sheet == SHEET_FIRST and tob_chg:
                mark_kotai_toban(ws, tob_chg)
                log(f"  交代のあった当番当直のマス: {len(tob_chg)} → 赤地に白字")
            _page_setup(ws, sheet, ndays, log)
            fit_note(ws, log)
            fit_legend(ws, log)                       # セカンド以下の凡例（20261022）
            stamp_updated(ws)
        xl.Calculation = -4105          # xlCalculationAutomatic
        if toban is not None and SHEET_FIRST in names:
            try:
                check_toban(wb.Worksheets(SHEET_FIRST), set(toban) - tob_chg, ndays, log)   # 赤地のマスは灰の確かめから外す
            except Exception as e:     # 確かめだけなので書き込みは続ける
                log(f"  当番当直の見え方を確かめられませんでした: {e}")
        wb.SaveAs(os.path.abspath(out), FileFormat=XL_MACRO_ENABLED)
        wb.Close(SaveChanges=False)
        return written
    finally:
        xl.Quit()
        # 広げた写しの一時フォルダ（touchoku_cf_*）を消す（20260926: 以前は毎回 %TEMP% に残っていた）
        if os.path.normcase(os.path.abspath(opened)) != os.path.normcase(os.path.abspath(template)):
            shutil.rmtree(os.path.dirname(opened), ignore_errors=True)


def _disp_name(r, disp):
    """セルに書く表示名。科の名簿にいない人（名簿から外れたが当直が残っている等）は姓だけ。"""
    return disp.get(r["staff_login"]) or C.namekey_display(r["name"]).split(" ")[0]


def _key_rec(r, disp):
    # 表示名は書く側（_fill_buckets）と同じ決め方にする（20260923 第2陣: 名簿にいない人をフルネームで期待して
    # 突合が ❌ になっていた。❌ は差し替えを止めるので、書いたものと同じ表記で比べる）
    return (r["date"], r["band"], r["frame_id"],
            C.canon(C.nfc(_disp_name(r, disp))), bool(r.get("oncall")))


def _key_back(r):
    return (r["date"], r["band"], r["frame_id"],
            C.canon(C.nfc(str(r.get("name_disp") or r.get("name") or "").replace(" ", ""))),
            bool(r.get("oncall")))


def verify_out(out, staffs, year, month, own_ids, expect, log=print):
    """書いた xlsm を dtx14_extract で読み戻して、表記（姓）で突合する。"""
    import dtx14_extract as X
    back, unresolved = X.extract(out, staffs, year, month)
    back = [r for r in back if r["frame_id"] in own_ids]
    # 雛形には所属医師・勤怠名簿の裏付けが無いのでログインIDは引けない。セルの表記（姓）で突合する
    a_, b_ = sorted(expect), sorted(map(_key_back, back))
    ok = (a_ == b_)
    log(f"読み戻し: {len(back)} 件（ログイン未解決 {len(unresolved)}＝雛形には名簿が無いので想定内）"
        f"  表記で一致: {'✅' if ok else '❌'}")
    if not ok:
        sa, sb = set(a_), set(b_)
        log(f"  書いたのに読めない: {sorted(sa - sb)[:5]}")
        log(f"  読めたのに書いていない: {sorted(sb - sa)[:5]}")
    return ok


# ------------------------------------------------------------------ 人別3シート（20260922 管理者）
#   事務担当のマクロ（所属医師の転記 → 人別生成 → 名前チェック）の代わりに、書いた当直表2シートを読み戻して
#   人別当直日／人別オンコール日／人別ファースト を DTX14 に書く。A〜E は従来どおり（下流が読む）、
#   F=教員（旧マクロ「教員チェック」と同じ意味・Dr.JOY の職名から判定）・G=ログインID（Dr.JOY の ID = M365 の UPN）・
#   H=職員コード を足す。下流（手当計算）は H（職員コード）で突き合わせて同姓同名の事故を防ぐ（給与連携にログインIDが無い）。
#   当直表から読み戻すので、事務担当が手で入れた行（管理当直など）も人別に載る。
#   人別を作る旧マクロ・top20 は dtx14_trim_vba.py で雛形から消した（20260922）。人別3シートはこの関数だけが書く。
JINBETSU_HEADER_EXTRA = {5: "フルネーム", 6: "教員", 7: "ログインID", 8: "職員コード"}
RANK_LABEL = {"教": "教員", "研": "大学院生", "専": "専攻医", "初": "初期研"}


def jinbetsu_row(r, hit, title_of, code_of=None):
    """読み戻した1件 r（dtx14_extract）→ 人別の1行 (A科, B表記, C日付, D昼夜, E氏名, F教員, GログインID, H職員コード)。
    hit = マスターの名簿で引けた (login, 氏名, 職員コード, 職名)（無ければ None）。
    名簿外の医師（ログインが unreg:氏名）は G・H を空にする（20260925: 下流は G を主キーにするので、
    Dr.JOY の ID でないものを入れない。空なら下流は E列の氏名で照合する）。"""
    disp = str(r.get("name_disp") or "").replace(" ", "")
    login, name, code, title = hit or ("", "", "", "")
    if not login:
        login, name, code = r.get("staff_login") or "", r.get("name") or "", r.get("staff_code") or ""
    if login.startswith("unreg:"):
        login, code = "", ""
    # 20261015: 予定表と config から職員コードを外したので、ログインIDで 管理ツール\名簿\ の職員コードを引く
    code = code or (code_of or {}).get(login, "")
    title = title or title_of.get(login, "")
    rank = C.rank_of(title)
    e = C.namekey(name) if name else ""
    b = f"({disp})" if r["oncall"] else disp
    return (r["frame_dept"], b, r["date"], r["band"], e, "教員" if rank == "教" else "", login, code)


def jinbetsu_rows(out, master, frames, staffs, year, month, log=print):
    """書き出した DTX14 の当直表2シート → 人別3シートの行 {sheet: [(A,B,C,D,E,F教員,GログインID,H職員コード)]}。
    氏名→ID はマスターの名簿（科ごとの表示名）で引き、引けなければフルネームの完全一致（異体字は寄せる）だけで引く。
    それでも引けなければ G/H は空にする（20260923 管理者: 下流は G列＝ログインIDを主キーにするので、
    姓の前方一致で「たぶんこの人」を入れると別人の当直・手当になる。空なら下流は E列の氏名で照合する）。"""
    import dtx14_extract as X
    fmap = {f["frame_id"]: f for f in frames}
    back, _unres = X.extract(out, staffs, year, month)
    idx = {}
    for owner, ent in (master.get("depts") or {}).items():
        disp = disp_map(ent.get("roster") or [])
        by_login = {m["login"]: m for m in ent.get("roster") or []}
        for r in ent.get("duties") or []:
            m = by_login.get(r.get("staff_login"), {})
            idx[_key_rec(r, disp)] = (r.get("staff_login") or "", m.get("name") or r.get("name") or "",
                                      m.get("code") or r.get("staff_code") or "", m.get("title") or "")
    people, _e, _c = C.load_staffs(staffs)
    title_of = {p["login"]: p.get("title", "") for p in people}
    code_of = {p["login"]: p["code"] for p in people if p.get("login") and p.get("code")} or C.staff_code_map()
    rows = {C.P_DUTY: [], C.P_ONCALL: [], C.P_FIRST: []}
    miss = []
    for r in back:
        row = jinbetsu_row(r, idx.get(_key_back(r)), title_of, code_of)
        if not row[6]:
            miss.append((r["frame_dept"], str(r.get("name_disp") or "").replace(" ", ""), r["date"], r["band"]))
        rows[C.P_ONCALL if r["oncall"] else C.P_DUTY].append(row)
        f = fmap.get(r.get("frame_id"))
        first_ok = (f["first_report"] if f else
                    (r.get("role") == "first" and "放射線科" not in r["frame_dept"] and "管理当直" not in r["frame_dept"]))
        if first_ok:
            rows[C.P_FIRST].append(row)
    for k in rows:
        rows[k].sort(key=lambda t: (t[2], t[0], t[3], t[1]))
    log(f"人別3シート: 当直日 {len(rows[C.P_DUTY])} / オンコール日 {len(rows[C.P_ONCALL])} / ファースト {len(rows[C.P_FIRST])}"
        + (f"  ⚠ IDが引けない {len(miss)} 件: {miss[:5]}" if miss else "  ID は全件解決"))
    return rows


def write_jinbetsu(out, rows, log=print):
    """人別3シートを DTX14（out）へ書く。1行目の見出しは残し、2行目以降を入れ替える。F/G の見出しが無ければ足す。"""
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(out), UpdateLinks=0, ReadOnly=False)
        xl.Calculation = -4135
        for sheet, data in rows.items():
            ws = wb.Worksheets(sheet)
            last = max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, 2)
            body = ws.Range(ws.Cells(2, 1), ws.Cells(max(last, len(data) + 1), 8))
            body.ClearContents()
            body.Font.ColorIndex = -4105           # 旧マクロの色分け（職位ごとの緑・青）は不要（管理者 20260922）。残すと別人の行に色がずれる
            for col, title in JINBETSU_HEADER_EXTRA.items():
                if ws.Cells(1, col).Value in (None, ""):
                    ws.Cells(1, col).Value = title
            if data:
                vals = [[xl_text(a), xl_text(b), float(_serial(dt.date.fromisoformat(c))), xl_text(d), xl_text(e),
                         xl_text(f), xl_text(g), xl_text(h)] for a, b, c, d, e, f, g, h in data]
                rng = ws.Range(ws.Cells(2, 1), ws.Cells(len(vals) + 1, 8))
                rng.Value = vals
                ws.Range(ws.Cells(2, 3), ws.Cells(len(vals) + 1, 3)).NumberFormat = "yyyy/mm/dd"
            log(f"  {sheet}: {len(data)} 行")
        xl.Calculation = -4105
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


FAILED_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "failed")
KEEP_FAILED = 5


def keep_failed(work, out):
    """突合 ❌ の書き上がりを調べるためにローカルへ残す（%LOCALAPPDATA%\\Touchoku\\failed。直近 5 本）。
    人別シートにログインIDと職員コードがあるので共有・同期フォルダには置かない。戻り: 残した場所（失敗なら ""）。"""
    try:
        os.makedirs(FAILED_DIR, exist_ok=True)
        stem, ext = os.path.splitext(os.path.basename(out))
        dst = os.path.join(FAILED_DIR, f"{stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{ext}")
        shutil.copy2(work, dst)
        olds = sorted(glob.glob(os.path.join(FAILED_DIR, f"{stem}_*{ext}")))
        for p in olds[:-KEEP_FAILED]:
            try:
                os.remove(p)
            except OSError:
                pass
        return dst
    except OSError:
        return ""


def _sha(path):
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def publish(src, dst, log=print, tries=5, wait=3.0):
    """書き上がった src（ローカル）で dst を差し替える（20260923 第2陣）。
    dst と同じフォルダに一時名（~dtx14_*.tmp）で写し、中身を確かめてから os.replace（同じボリューム内の置き換えなので
    途中の状態が dst に出ない）。dst を Excel・OneDrive が掴んでいて置き換えられなければ少し待って数回やり直し、
    それでもだめなら一時ファイルを消して FillError（dst は元のまま）。"""
    d = os.path.dirname(os.path.abspath(dst))
    tmp = os.path.join(d, f"~dtx14_{os.getpid()}_{int(time.time() * 1000)}.tmp")
    try:
        shutil.copyfile(src, tmp)
        if _sha(tmp) != _sha(src):
            raise FillError(f"差し替え用の写しが元と一致しません: {tmp}")
        last = None
        for i in range(tries):
            try:
                os.replace(tmp, dst)
                return dst
            except OSError as e:
                last = e
                log(f"  差し替えを待つ ({i + 1}/{tries}): {e}")
                time.sleep(wait)
        raise FileBusy(f"{dst} を差し替えられません（開かれている？）: {last}")
    except OSError as e:
        raise FillError(f"{dst} の差し替えに失敗: {type(e).__name__}: {e}") from e
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def run_master(master_path, template, out, year=0, month=0, staffs=None, verify=False, log=print, guard=None):
    """共有マスター1本 → DTX14 を全科まとめて埋める。戻り: {cells, records, owners, verified}
    20260923 第2陣: 書き込みはすべてローカルの一時フォルダの写しで行い（当直表2シート → 人別3シート → 読み戻しの突合）、
    全部が通ったときだけ out を差し替える（publish）。verify=True で突合が ❌ なら out に触らず FillError
    （書き上がりは keep_failed で %LOCALAPPDATA%\\Touchoku\\failed に残す）。途中の例外も FillError に包む（out は元のまま）。
    guard = 差し替えの直前に呼ぶ関数（呼び出し側の最終確認。FillError を投げれば差し替えない）。"""
    staffs = staffs or C.newest_staffs()
    frames = C.load_frames()
    master, year, month = load_master(master_path, year, month)
    depts = master.get("depts") or {}
    log(f"マスター: {master_path if isinstance(master_path, str) else '承認済みの状態（元の割当＋承認した交代）'}")
    log(f"R{year - 2018}年{month}月  {len(depts)} 科  rev {master.get('rev')}")
    if not any(ent.get("duties") for ent in depts.values()):
        raise NoDuties("マスターに当直の記録がありません")      # 雛形を開く前に止める（COM を使わない）
    loc = read_layout(template)
    plan, disps, counts = cells_from_master(master, frames, loc, C.load_holidays(), log)
    plan["toban"] = toban_cells(frames, loc, year, month)      # 当番当直の見え方の確かめ用（20260925）
    total = sum(c[1] for c in counts)
    log(f"合計 {total} 件 → 書き込むセル {len(plan['cells'])}"
        f"（2行ある全日枠 {len(plan['units'])}: 全日タイプ {sum(1 for u in plan['units'] if u[3])}"
        f" / 昼夜タイプ {sum(1 for u in plan['units'] if not u[3])}）")
    if not plan["cells"]:
        raise FillError("書き込むセルがありません（マスターの当直が雛形のどの行にも当たらない）")
    work_dir = tempfile.mkdtemp(prefix="dtx14_fill_")
    work = os.path.join(work_dir, "work" + (os.path.splitext(out)[1] or ".xlsm"))
    stage = "当直表2シート"
    try:
        n = write_with_com(template, work, plan, year, month, log)
        log(f"書き上がり（ローカルの写し）: {n} セル")
        stage = "人別3シート"
        write_jinbetsu(work, jinbetsu_rows(work, master, frames, staffs, year, month, log), log)
        ok = None
        if verify:
            stage = "読み戻しの突合"
            fmap = {f["frame_id"]: f for f in frames}
            # 突合の対象 = マスターに出てくる枠（科名が枠マスタの owner と違う時期のマスターでも落とさない）
            own_ids = {r.get("frame_id") for ent in depts.values() for r in (ent.get("duties") or [])
                       if fmap.get(r.get("frame_id"), {}).get("sheet") and fmap[r["frame_id"]].get("row")}
            expect = []
            for owner, ent in sorted(depts.items()):
                for r in ent.get("duties") or []:
                    f = fmap.get(r.get("frame_id"))
                    if f and f.get("sheet") and f.get("row"):
                        expect.append(_key_rec(r, disps[owner]))
            ok = verify_out(work, staffs, year, month, own_ids, expect, log)
            if not ok:
                kept = keep_failed(work, out)
                raise FillError(f"読み戻しの突合が一致しません（❌）→ {out} は差し替えていません"
                                + (f"（書き上がりの写し: {kept}）" if kept else ""))
        stage = "差し替え"
        if guard:
            guard()
        publish(work, out, log)
    except FillError:
        raise
    except Exception as e:              # COM の失敗など。out には触っていない
        raise FillError(f"{stage}で失敗: {type(e).__name__}: {e} → {out} は差し替えていません") from e
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    log(f"出力: {out}  ({n} セル)")
    return {"cells": n, "records": total, "owners": [c[0] for c in counts], "verified": ok}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_")
    ap.add_argument("--master", help="共有マスター 当直予定表_R8年N月.json（--in と排他）")
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--year", type=int, default=0)
    ap.add_argument("--month", type=int, default=0)
    ap.add_argument("--staffs", default=None)
    ap.add_argument("--verify", action="store_true", help="書いた後に dtx14_extract で読み戻して突合")
    a = ap.parse_args()
    if bool(a.in_) == bool(a.master):
        ap.error("--in か --master のどちらか一方を指定してください")
    try:
        _main(a)
    except FillError as e:
        raise SystemExit(str(e))


def _main(a):
    staffs = a.staffs or C.newest_staffs()
    if a.master:
        run_master(a.master, a.template, a.out, a.year, a.month, staffs, a.verify)
        return
    if not (a.year and a.month):
        t = dt.date.today()
        a.year, a.month = (t.year + (t.month == 12), t.month % 12 + 1)
    owner, recs, rejects, members, frames = load_records(a.in_, a.year, a.month, staffs)
    print(f"読み取り: {owner}  レコード {len(recs)} 件 / 却下 {len(rejects)} 件 / メンバー {len(members)} 名")
    for rj in rejects[:10]:
        print("  却下:", rj)
    loc = read_layout(a.template)
    plan = cells_from_records(recs, members, frames, loc)
    plan["toban"] = toban_cells(frames, loc, a.year, a.month)
    print(f"書き込むセル: {len(plan['cells'])}")
    n = write_with_com(a.template, a.out, plan, a.year, a.month)
    print(f"出力: {a.out}  ({n} セル)")
    if a.verify:
        own = {f["frame_id"] for f in frames if f.get("owner", f["dept"]) == owner}
        disp = disp_map(members)
        verify_out(a.out, staffs, a.year, a.month, own, [_key_rec(r, disp) for r in recs])


if __name__ == "__main__":
    main()
