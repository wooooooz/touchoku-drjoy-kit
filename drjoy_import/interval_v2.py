# -*- coding: utf-8 -*-
"""勤務間インターバル（代償休息）— Dr.JOYのアルゴリズム × 当院のデータ。20260827

★方針（ユーザー決定 2026-08-25）
  **判定ロジックはDr.JOYに追従し、データは当院のもの（DTX14の当直＋当直情報.xlsm）を使う。**
  Dr.JOYの勤務予定には当直が入っていない（給与連携の「当直」列は7月の全行が空）。
  庶務の予定入力も、医師の研鑽「通常当直業務」の選択も当てにならない。
  **信頼できる当直の事実はDTX14（当直勤務表）だけ**というのが当院の実情。

★2つの量を別々に出す（20260826 ★重要）
  - `short_h` … **9時間連続が取れなかった量** = 必要 − 窓内の**最長連続**休息。
                 「措置が守れたか」の充足判定に使う。
  - `comp_h`  … **付与すべき代償休息** = 必要 − 窓内の休息の**合計**。
                 厚労省 医師の働き方改革Q&A（R5.6.7事務連絡）Q_F-15 が
                 「9時間の連続休息を超える休息時間部分は代償休息として充当できる」と明言しており、
                 付与量は「インターバル中に従事した労働時間に相当する時間」。Q_F-11 は
                 9時間確保後の呼び出しに義務なし、Q_F-17 はオンコールも同じ整理。
                 → **返すべき量は comp_h**。short_h を返させるのは過大。

★判定の手順（Dr.JOYのアルゴリズム。実測で確認済み）
  1. **勤務予定の始業**を起点に窓を取る（予定が無い日に当直があれば**当直の開始**を起点にする）
  2. 窓の中の休息を求める（連続の最長／合計の2通り）
  3. 不足 = 必要 − 休息。0以下は不足なし。1つの窓につき1回だけ計上
  ⚠帯を窓の末尾に固定しているのではない。Dr.JOYは置き場所を探している（黄金弘二5日・水晶幸子7日で実測。
    前半に11時間15分の休息がある日は計上ゼロ）。履歴画面の「時間」欄が末尾に見えるのは、
    **休息の直前から不足分だけ遡った区間**を表示しているため。

★窓の長さと必要な連続休息（ユーザー決定＝制度どおり。Dr.JOY実装とは意図的に異なる）
  | 場面 | 窓 | 必要 |
  |---|---|---|
  | 通常日・宿日直許可ありの当直 | 24時間 | 9時間 |
  | 宿日直許可なしの当直（B水準） | 46時間 | 18時間 |
  | 同上で初期研修医（C-1水準 臨床研修医） | 48時間 | 24時間 |
  ⚠**Dr.JOYは小児科を通常扱いにし、初期研のC-1も適用していない**（実測）。当院は制度どおりにする。

★「宿日直許可なし」の判定＝**当直枠の科**（ユーザー決定）
  当直枠が救急科・小児科なら、**本人の所属に関係なく**許可なしとして扱う。
  ⚠**枠の科は人別シートの「当直科」列から読む**（20260826 に DTX_DrJOY Module1/Module3 で新設）。
    人別シートの当直は**昼夜が合成済み**なので、DTX14を直接読むより窓が正確。
    併任（小児科当直 兼 管理当直）は "小児科/管理当直" と連結されており、1つでも救急・小児なら許可なし。
    DTX14は「人別シートが無い人」の補完と、E列未記入の件数を出すために読む。
  宿日直許可は当直の種類の話であって所属の話ではない。
  ⚠所属基準にすると、麻酔科所属の真珠弘一が務めた救急科当直6件を取りこぼす（7月で18件が該当）。

★労働とみなすもの（ユーザー決定＝当院ルール維持）
  所定勤務（勤務予定 出勤〜退勤）∪ 時間外申請（当直情報.xlsm＝変換後）∪ **宿日直許可なしの当直窓**
  それ以外はすべて休息（研鑽の全種別・許可ありの当直・空き時間）
  ⚠**研鑽「通常当直業務」は休息**。宿日直許可ありの当直は労働時間ではない（Dr.JOY準拠・制度的にも自然）。
    これを労働に数えていたのがDr.JOYとの差の59.8%だった。
  ⚠**在席打刻は使わない**。ビーコン置きっぱなしで汚染される（打刻を労働源にすると差が4倍悪化＝実測）。
  ⚠当直中の「その他（手当なし）」は DTX_DrJOY Module4 が既に自己研鑽へ移してある（救急部・小児科は対象外）。
    当院は「診療（手当あり）だけを勤務と見なす」運用なので、変換後のデータをそのまま使う。

★既知の限界
  - **窓が翌月へ食い込む日は判定しない**（ユーザー決定 2026-08-25）。翌月のデータが無いと窓の後半が
    丸ごと「休息」に見え、不足が過小に出るため。結果として**月末の1〜2日が未判定**になる（件数はログと
    戻り値の skipped に出る）。恒久対応するなら翌月ファイルを連結する。
  - **DTX14のE列（フルネーム）が空の行は人物に紐づかない**（7月253行）。ログに件数を出す。

★読むファイル（20260826）
  **出退勤入力情報_..._当直情報.xlsm だけでほぼ足りる**（勤務予定・職名・所定時間・当直科を内蔵）。
  DTX14と給与連携は、人別シートが無い人（Dr.JOY未使用者）の補完にだけ使う。

使い方: python interval_v2.py "<月フォルダ>"
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from collections import defaultdict

sys.path.insert(0, r"C:\tools\touchoku\drjoy_import")
import openpyxl
import dtx18_stats as A

WIN_NORMAL, NEED_NORMAL = 24, 9        # 通常・宿日直許可あり当直
WIN_NOPERMIT, NEED_NOPERMIT = 46, 18   # 宿日直許可なし（B水準）
WIN_C1, NEED_C1 = 48, 24               # 宿日直許可なし（C-1水準 臨床研修医＝初期研修医）
RESIDENT = "初期研修医"
# ★当直の実体は「人別当直日」だけ。20260829
#   「人別ファースト」は**当直日誌を提出する義務のある人を書き出すための表**で、
#   当直勤務の実体とは別のもの（ユーザー説明 2026-08-29）。当直の材料に使わない。
#   実測＝R8.5〜9月の5か月とも、人別ファーストが足す当直は0人日・中身の違いも0件だった
#   （5月708／6月616／7月617／8月727／9月597人日がどちらでも同じ）。
#   ＝これまでの数字は変わらないが、将来ずれる余地があるので外す。
DUTY_SHEETS = ("人別当直日",)

# ★略称が括弧つき = オンコール（病院に泊まらない）。当直ではないので当直として数えない。
#   「人別ファースト」には当直とオンコールが混ざっている（R8.9月=758行中156行が括弧つき）。
#   「人別当直日」は括弧つき0行、「人別オンコール日」は全行が括弧つき。20260827
#   ⚠人別ファーストは _first_keys() で「その枠がファーストか」の手がかりにだけ使う
#     （氏名解決の補助）。当直の材料としては読まない。
_PAREN = re.compile(r"^[\(（].*[\)）]$")

# 直近の compute() の付帯情報（スライドの脚注に出す）。20260825
STATS = {"cut": 0, "empty": 0, "duty_days": 0, "duty_people": 0}


def _dtv(v, base):
    """セル値 → datetime。datetime / 'YYYY/MM/DD HH:MM(:SS)' / 時刻(timedelta) / time に対応。"""
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.time):
        return dt.datetime.combine(base, v)
    if isinstance(v, dt.timedelta):
        return dt.datetime.combine(base, dt.time()) + v
    s = str(v).strip()
    for f in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M", "%H:%M"):
        try:
            d = dt.datetime.strptime(s, f)
            return dt.datetime.combine(base, d.time()) if f == "%H:%M" else d
        except ValueError:
            pass
    return None


def _clean(s):
    from dtx14_names import clean_abbr
    return clean_abbr(s)


def _first_keys(dtx14_path):
    """人別ファースト（＝当直本体）に出てくる (枠, 略称)。初期研はここに入らない。"""
    out = set()
    wb = openpyxl.load_workbook(dtx14_path, read_only=True, data_only=True)
    if "人別ファースト" in wb.sheetnames:
        ws = wb["人別ファースト"]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i and r and len(r) > 1 and r[0] and r[1]:
                out.add((str(r[0]).strip(), _clean(r[1])))
    wb.close()
    return out


def load_duty(dtx14_path, log=print, resolve=None):
    """DTX14 → {氏名norm: {date: (許可なしか, 勤務形態)}}。**当直枠の科**で判定する。

    resolve = dtx14_names.build_resolver() の戻り。E列(フルネーム)が空の行を氏名まで解決する。
    ⚠渡さないと空の行はまるごと落ちる（R8.8月は435行＝当直の3割強がこれに当たった）。
    """
    out = defaultdict(dict)
    empty = 0
    solved = 0
    oncall = 0
    merged = 0
    first_keys = _first_keys(dtx14_path) if resolve else set()
    wb = openpyxl.load_workbook(dtx14_path, read_only=True, data_only=True)
    for sn in DUTY_SHEETS:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i == 0 or not r or not r[0]:
                continue
            full = str(r[4]).strip() if len(r) > 4 and r[4] is not None else ""
            d = A.cell_date(r[2] if len(r) > 2 else None)
            band = str(r[3] or "").strip() if len(r) > 3 else ""
            # ★略称が括弧つき = オンコール。当直ではないので落とす。20260827
            #   これを入れる前は「人別ファースト」の括弧つき156行(R8.9月)が当直として
            #   数えられ、Dr.JOYへの投入では自宅待機の医師に泊まり当直が登録されていた。
            if _PAREN.match(str(r[1] or "").strip() if len(r) > 1 else ""):
                oncall += 1
                continue
            # ★E列に**略称がそのまま**入っていることがある（事務担当が埋めかけの状態）。
            #   「フルネームがある」と誤認すると姓だけで下流に流れ、職員コードが引けない。
            if full and (full.replace(" ", "") == _clean(r[1] if len(r) > 1 else "")
                         or len(full.replace(" ", "").replace("　", "")) < 3):
                full = ""
            if not full and resolve is not None:
                frame = str(r[0] or "").strip()
                ab = str(r[1] or "").strip()
                got, _why, _c = resolve(frame, ab, d,
                                        first=(frame, _clean(ab)) in first_keys)
                if got:
                    full = got
                    solved += 1
            if not full:
                empty += 1
                continue
            if d is None or band not in A.DUTY_TYPES:
                continue
            nopermit = A.is_er_peds(str(r[0] or ""))   # ★枠の科で判定
            n = A.norm_name(full)
            prev = out[n].get(d)
            if prev is None:
                out[n][d] = (nopermit, band)
            else:
                pnop, pband = prev
                # ★同じ人・同じ日に「昼」と「夜」が別行で来たら「昼夜」に合成する。20260827
                #   休日の全日当直はDTX14に昼・夜の2行で入っている。合成しないと先に来る
                #   「昼」が勝ち、泊まり(17:15〜翌朝)が丸ごと落ちる（R8.9月=116人日）。
                #   duty_kind() には「昼夜」の分岐が用意されているのに、ここが生成しないため
                #   死にコードになっていた。
                if pband == band:
                    nb = pband
                elif "昼夜" in (pband, band) or {pband, band} == {"昼", "夜"}:
                    nb = "昼夜"
                    if pband != "昼夜":
                        merged += 1
                else:
                    nb = pband
                # 複数枠なら「許可なし」を優先（重いほうに寄せる）
                out[n][d] = (pnop or nopermit, nb)
    wb.close()
    STATS["empty"] = empty
    STATS["name_solved"] = solved
    STATS["oncall_skipped"] = oncall
    STATS["day_night_merged"] = merged
    STATS["duty_days"] = sum(len(v) for v in out.values())
    STATS["duty_people"] = len(out)
    log(f"  DTX14の当直: {STATS['duty_days']}人日 / {STATS['duty_people']}名"
        + (f"　E列が空→氏名を解決 {solved}行" if solved else "")
        + (f"　オンコール(括弧つき)を除外 {oncall}行" if oncall else "")
        + (f"　昼＋夜→昼夜に合成 {merged}人日" if merged else "")
        + (f"　⚠E列(フルネーム)が空で紐づかない行 {empty}件" if empty else ""))
    return out


def load_pay(pay_xlsx):
    """給与連携 → ({氏名norm: {date: (予定始業, 予定終業)}}, {氏名norm: 職名})。"""
    sched = defaultdict(dict)
    jobs = {}
    wb = openpyxl.load_workbook(pay_xlsx, read_only=True, data_only=True)
    ws = A._ws_full(wb.worksheets[0])
    it = ws.iter_rows(values_only=True)
    hdr = [str(h or "").strip() for h in next(it)]

    def col(n):
        return hdr.index(n) if n in hdr else -1

    cN, cD, cJ = col("氏名"), col("日付"), col("職名")
    cS, cE = col("勤務予定出勤時刻"), col("勤務予定退勤時刻")
    for r in it:
        n = A.norm_name(r[cN]) if 0 <= cN < len(r) else ""
        d = A.cell_date(r[cD]) if 0 <= cD < len(r) else None
        if not n or d is None:
            continue
        if 0 <= cJ < len(r) and str(r[cJ] or "").strip():
            jobs[n] = str(r[cJ]).strip()
        s = _dtv(r[cS], d) if 0 <= cS < len(r) else None
        e = _dtv(r[cE], d) if 0 <= cE < len(r) else None
        if s and e:
            if e < s:
                e += dt.timedelta(days=1)
            sched[n][d] = (s, e)
    wb.close()
    return sched, jobs


def load_overtime(src_xlsm):
    """当直情報.xlsm の人別シート → 必要なものを全部読み取る。20260826

    返り値 (ot, who, sched, jobs, shotei)
      ot     … {氏名norm: [(開始,終了)]}  時間外のみ（研鑽は休息なので読まない）
      who    … {氏名norm: (氏名, 科)}
      sched  … {氏名norm: {date: (予定始業, 予定終業)}}   ← 勤務予定出勤/退勤時刻
      jobs   … {氏名norm: 職名}                        ← 1行目3列目（「初期研N期」等）
      shotei … {氏名norm: {date: 所定時間文字列}}        ← 「所定時間、休暇等」
      duty   … {氏名norm: {date: (許可なしか, 勤務形態)}} ← 「当直」＋「当直科」
      punch  … {氏名norm: {date: (出勤, 退勤)}}          ← 打刻（代償休息の消化にだけ使う）
    ★人別シートの当直は**昼夜が合成済み**なので、DTX14を直接読むより正確
      （DTX14は昼と夜が別行で、片方しか採れず窓が短くなる）。
    ★2026-08-26にDr.JOYのエクスポート項目を増やしたので、給与連携を読まずに済むようになった。
      古いファイル（列が無い）でも落ちないよう、無ければ空で返す（呼び手が給与連携で補う）。
    """
    ot = defaultdict(list)
    who = {}
    sched = defaultdict(dict)
    jobs = {}
    shotei = defaultdict(dict)
    duty = defaultdict(dict)
    punch = defaultdict(dict)
    wb = openpyxl.load_workbook(src_xlsm, read_only=True, data_only=True)
    for sn in wb.sheetnames:
        ws = wb[sn]
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 5 or not rows[1] or not rows[1][0]:
            continue
        dept = str(rows[0][0] or "") if rows[0] else ""
        name = str(rows[1][0]).strip()
        hdr = [str(c or "").strip() for c in rows[3]]
        if "日付" not in hdr:
            continue
        cDate = hdr.index("日付")
        cS = [hdr.index(f"時間外開始{k}") if f"時間外開始{k}" in hdr else -1 for k in range(1, A.MAX_SLOTS + 1)]
        cE = [hdr.index(f"時間外終了{k}") if f"時間外終了{k}" in hdr else -1 for k in range(1, A.MAX_SLOTS + 1)]
        n = A.norm_name(name)
        who[n] = (name, dept)
        pos_raw = str(rows[0][2] or "").strip() if len(rows[0]) > 2 else ""
        if pos_raw:
            jobs[n] = pos_raw
        cPS = hdr.index("勤務予定出勤時刻") if "勤務予定出勤時刻" in hdr else -1
        cPE = hdr.index("勤務予定退勤時刻") if "勤務予定退勤時刻" in hdr else -1
        cSH = hdr.index("所定時間、休暇等") if "所定時間、休暇等" in hdr else -1
        cDU = hdr.index("当直") if "当直" in hdr else -1
        cDK = hdr.index("当直科") if "当直科" in hdr else -1
        cIN = hdr.index("出勤") if "出勤" in hdr else -1
        cOT = hdr.index("退勤") if "退勤" in hdr else -1
        for r in rows[4:]:
            if not r or cDate >= len(r) or not r[cDate]:
                continue
            d = A.cell_date(r[cDate])
            if d is None:
                continue
            for k in range(A.MAX_SLOTS):
                if cS[k] < 0 or cS[k] >= len(r) or cE[k] < 0 or cE[k] >= len(r):
                    continue
                a, b = _dtv(r[cS[k]], d), _dtv(r[cE[k]], d)
                if a and b:
                    if b < a:
                        b += dt.timedelta(days=1)
                    ot[n].append((a, b))
            if 0 <= cSH < len(r):
                shotei[n][d] = str(r[cSH] or "").strip()
            if 0 <= cIN < len(r) and 0 <= cOT < len(r):
                punch[n][d] = (r[cIN], r[cOT])
            if 0 <= cDU < len(r) and 0 <= cDK < len(r):
                band = str(r[cDU] or "").strip()
                if band in A.DUTY_TYPES:
                    frame = str(r[cDK] or "").strip()
                    # 併任は "小児科/管理当直" のように連結されている。1つでも救急・小児なら許可なし。
                    nopermit = any(A.is_er_peds(x) for x in frame.split("/") if x)
                    duty[n][d] = (nopermit, band)
            if 0 <= cPS < len(r) and 0 <= cPE < len(r):
                ps, pe = _dtv(r[cPS], d), _dtv(r[cPE], d)
                if ps and pe:
                    if pe < ps:
                        pe += dt.timedelta(days=1)
                    sched[n][d] = (ps, pe)
    wb.close()
    return ot, who, sched, jobs, shotei, duty, punch


# ============================================================
# 代償休息の消化（ユーザー決定 2026-08-26）
#   平日（土日祝を除く）の 8:30-17:15 から 12:00-13:00 を除いた 1日7時間45分 を基準に、
#   働いていない時間を消化とみなす。
#   ・年休の日もそのまま数える（休暇の種別が判別できないため。院内の目安として了承済み）
#   ・打刻が無い日は**まるごと**計上（ビーコンを持ち歩かない人は救済対象にしない）
#   ・兼業日（＝所定時間が変則）は対象外
# ============================================================
CORE_BANDS = [(8 * 60 + 30, 12 * 60), (13 * 60, 17 * 60 + 15)]
CORE_MIN = 7 * 60 + 45
STD_START, STD_END = 8 * 60 + 30, 17 * 60 + 15
_RE_SEG = None


def is_kengyo_day(text):
    """「所定時間、休暇等」が変則なら割振（兼業）の日とみなす。空欄は変則ではない。
    ①開始が8:30より遅い ②終了が17:15より早い ③中抜けがある（区間が2つ以上）"""
    global _RE_SEG
    if _RE_SEG is None:
        import re
        _RE_SEG = re.compile(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})")
    segs = _RE_SEG.findall(str(text or ""))
    if not segs:
        return False
    ss = [(int(a) * 60 + int(b), int(c) * 60 + int(d)) for a, b, c, d in segs]
    return len(ss) >= 2 or ss[0][0] > STD_START or ss[0][1] < STD_END


def _as_min(v):
    """セル値 → 0時からの分。時刻・datetime・シリアル値に対応。"""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, dt.time):
        return v.hour * 60 + v.minute
    if isinstance(v, dt.timedelta):
        return int(v.total_seconds() // 60) % 1440
    if isinstance(v, (int, float)):
        return int(round((float(v) % 1) * 1440))
    return None


def consumed_minutes(punch_of_person, shotei_of_person, month_days):
    """その人がその月に消化できた代償休息（分）と内訳。"""
    total = 0
    n_kengyo = n_nopunch = n_punch = 0
    for d in month_days:
        if d.weekday() >= 5 or A.is_holiday_date(d):
            continue
        if is_kengyo_day(shotei_of_person.get(d, "")):
            n_kengyo += 1
            continue
        p = punch_of_person.get(d)
        a = _as_min(p[0]) if p else None
        b = _as_min(p[1]) if p else None
        if a is None or b is None:
            total += CORE_MIN          # 打刻が無い日はまるごと消化とみなす
            n_nopunch += 1
            continue
        if b <= a:
            b = 24 * 60                # 日をまたぐ退勤は当日いっぱい在席とみなす
        worked = sum(max(0, min(b, e) - max(a, st)) for st, e in CORE_BANDS)
        total += max(0, CORE_MIN - worked)
        n_punch += 1
    return total, n_kengyo, n_nopunch, n_punch


def _punch_dt(v, base):
    """打刻セル → datetime。20260904

    ⚠24時を超える打刻は Excel のシリアル値が1を超えるため、openpyxl が
      `datetime(1900,1,1,2,0)` のような値で返す。これは壊れているのではなく
      **24:00＋2:00＝翌2:00** の意味（当月708件ある）。base の翌日に読み替える。
    """
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        if v.year == 1900:                      # 24時超（1900-01-01 が基準日）
            return dt.datetime.combine(base, dt.time(0, 0)) \
                + dt.timedelta(days=1, hours=v.hour, minutes=v.minute)
        return v
    if isinstance(v, dt.time):
        return dt.datetime.combine(base, v)
    return None


def subtract_spans(spans, cuts):
    """spans（労働区間）から cuts（休息とみなす区間）を差し引く。20260904

    宿日直許可のある当直帯を労働から外すために使う。区間は (開始, 終了) の並び。
    """
    out = []
    for s, e in spans:
        pieces = [(s, e)]
        for c1, c2 in cuts:
            nxt = []
            for a, b in pieces:
                if c2 <= a or c1 >= b:          # 重ならない
                    nxt.append((a, b))
                    continue
                if a < c1:                       # 前が残る
                    nxt.append((a, c1))
                if c2 < b:                       # 後ろが残る
                    nxt.append((c2, b))
            pieces = nxt
            if not pieces:
                break
        out.extend(p for p in pieces if p[1] > p[0])
    return out


def total_rest(win_start, win_end, work):
    """窓の中で休息だった時間の**合計**（分）。20260826

    ★代償休息の量はこちらで測る（厚労省 医師の働き方改革FAQ 事務連絡 R5.6.7）。
      Q_F-15: 予定インターバル中に発生した労働については**その労働時間に相当する時間**の
      代償休息を付与すればよく、**9時間を超える休息時間部分は充当できる**。
      Q_F-11: 9時間の連続休息を確保した後の呼び出しには代償休息の義務は生じない。
      Q_F-17: オンコールで呼ばれた場合も「呼び出し後に従事した労働時間に相当する代償休息」。
    したがって 付与すべき代償休息 = max(0, 必要時間 − 窓内の休息の合計)。
    連続で見る longest_rest は「9時間連続が取れたか」という**要件の充足判定**には使うが、
    代償休息の量としては過大になる。
    """
    segs = []
    for a, b in work:
        a2, b2 = max(a, win_start), min(b, win_end)
        if b2 > a2:
            segs.append((a2, b2))
    segs.sort()
    worked = 0.0
    cur = None
    for a, b in segs:
        if cur and a <= cur[1]:
            if b > cur[1]:
                worked += (b - cur[1]).total_seconds() / 60.0
                cur[1] = b
        else:
            worked += (b - a).total_seconds() / 60.0
            cur = [a, b]
    return max(0.0, (win_end - win_start).total_seconds() / 60.0 - worked)


def longest_rest(win_start, win_end, work):
    """窓の中で実際に取れた最長の連続休息(分)。窓の先頭・末尾の空きも休息に数える。"""
    segs = []
    for a, b in work:
        a2, b2 = max(a, win_start), min(b, win_end)
        if b2 > a2:
            segs.append((a2, b2))
    segs.sort()
    merged = []
    for a, b in segs:
        if merged and a <= merged[-1][1]:
            if b > merged[-1][1]:
                merged[-1][1] = b
        else:
            merged.append([a, b])
    best, cur = 0.0, win_start
    for a, b in merged:
        gap = (a - cur).total_seconds() / 60.0
        if gap > best:
            best = gap
        if b > cur:
            cur = b
    return max(best, (win_end - cur).total_seconds() / 60.0, 0.0)


def compute(folder, log=print):
    src = A.find_latest(folder, "*_当直情報*.xlsm")
    az14 = A.find_latest(folder, "*DTX14*.xlsm")
    for p, nm_ in ((src, "当直情報"), (az14, "DTX14")):
        if not p:
            raise FileNotFoundError(f"{nm_} が見つかりません")
    ot, who, sched, jobs, shotei, duty, punch = load_overtime(src)
    # 「当直科」列があれば人別シートの当直を正とする（昼夜が合成済みで正確）。
    # DTX14は「人別シートが無い人」の補完と、E列未記入の件数を出すために読む。
    if duty:
        az_duty = load_duty(az14, log=log)
        add = 0
        for n, v in az_duty.items():
            if n not in duty:
                duty[n] = v
                add += 1
        log(f"  当直は人別シートの「当直科」列から取得 "
            f"({sum(len(v) for v in duty.values())}人日/{len(duty)}名"
            + (f"・人別シートが無い{add}名はDTX14で補完" if add else "") + ")")
    else:
        log("  ※「当直科」列が無いのでDTX14を直接読みます")
        duty = load_duty(az14, log=log)
    # 2026-08-26以降のエクスポートは当直情報だけでほぼ足りる。ただし
    # **人別シートが無い人**（出退勤の記録が1件も無い＝Dr.JOYを使っていない）は落ちるので、
    # 給与連携があれば「人別シートに居ない人だけ」補う。7月は50名が該当し、
    # うち3名は救急・小児の当直（宿日直許可なし）を持っていた。
    pay = A.find_latest(folder, "*給与連携項目一覧*.xlsx")
    if not sched or not jobs:
        if not pay:
            raise FileNotFoundError("当直情報に勤務予定・職名が無く、給与連携項目一覧も見つかりません")
        log("  ※古い当直情報のため勤務予定・職名は給与連携から読みます")
        sched, jobs = load_pay(pay)
    elif pay:
        ps, pj = load_pay(pay)
        add = 0
        for n, v in ps.items():
            if n not in sched:
                sched[n] = v
                jobs.setdefault(n, pj.get(n, ""))
                add += 1
        if add:
            log(f"  人別シートが無い{add}名を給与連携で補完")
    log(f"  勤務予定 {len(sched)}名 / 時間外のある人別シート {len(ot)}名"
        + (f" / 所定時間の記載 {sum(len(v) for v in shotei.values())}人日" if shotei else ""))

    names = set(sched) | set(ot) | set(duty)
    month_end = max((d for v in sched.values() for d in v), default=None)
    out, cut = [], 0
    for n in names:
        sc = sched.get(n, {})
        dd = duty.get(n, {})
        pos = A.normalize_position(jobs.get(n, ""))
        is_res = pos == RESIDENT

        # ★20260904 許可なし当直の「翌日」の勤務予定は、実際の退勤打刻で置き換える。
        #   Dr.JOYの勤務予定は当直明けも 08:30-17:15 のままだが、実際は帰っている
        #   （打刻のある59件の退勤は中央値12:00・43件が12時以前・17時以降は7件だけ）。
        #   予定のまま数えると1日7〜8時間を余計に積み、救急部の発生が2〜3倍に膨らむ。
        #   **打刻が無い日は休んだとみなす**（ユーザー決定 2026-09-04）。
        #   ⚠打刻で予定を「延長」はしない。ビーコン置きっぱなしが多く（当月72名）、
        #     そのまま信じると逆方向に大きく狂う。頭打ちにだけ使う。
        ake = {d + dt.timedelta(days=1) for d, (nop, _b) in dd.items() if nop}
        pu = punch.get(n, {})
        work = list(ot.get(n, []))
        for d, (ps, pe) in sc.items():
            if d in ake:
                out_dt = _punch_dt(pu.get(d, (None, None))[1], d)
                if out_dt is None:
                    continue                             # 打刻なし＝帰宅扱い
                pe = min(pe, out_dt)
                if pe <= ps:
                    continue
            work.append((ps, pe))
        permitted = []                                   # 許可ありの当直窓＝労働ではない
        for d, (nopermit, band) in dd.items():
            w1, w2 = A.duty_window(d, band, A.is_holiday_date(d + dt.timedelta(days=1)))
            if w1 is None:
                continue
            if nopermit:
                work.append((w1, w2))                    # 許可なしの当直窓だけ労働
            else:
                permitted.append((w1, w2))
        # ★20260904 宿日直許可のある当直帯は、そこに時間外の申告があっても休息として扱う
        #   （ユーザー指示＝Dr.JOYのルールに合わせる。当直中は労働時間ではないので、
        #     本人が残業を申告していてもインターバル不足には計上しない）。
        #   当直の有無と枠は当直勤務表(DTX14)を正とする。
        if permitted:
            work = subtract_spans(work, permitted)

        viol, short_min, comp_min, detail = 0, 0.0, 0.0, []
        for d in sorted(set(sc) | set(dd)):
            nopermit = dd.get(d, (False, ""))[0]
            if d in sc:
                start = sc[d][0]
            elif nopermit:                               # 予定が無い当直日は当直の開始を起点に
                w1, _w2 = A.duty_window(d, dd[d][1], A.is_holiday_date(d + dt.timedelta(days=1)))
                if w1 is None:
                    continue
                start = w1
            else:
                continue                                 # 予定も許可なし当直も無い日は判定しない
            if nopermit:
                win, need = (WIN_C1, NEED_C1) if is_res else (WIN_NOPERMIT, NEED_NOPERMIT)
            else:
                win, need = WIN_NORMAL, NEED_NORMAL
            end = start + dt.timedelta(hours=win)
            if month_end and end.date() > month_end:
                # 窓が翌月へ食い込む日は**判定しない**（ユーザー決定 2026-08-25）。
                # 翌月のデータが無いと窓の後半が丸ごと「休息」に見え、不足が過小に出るため。
                # 月末の1〜2日が未判定になるぶんは、脚注に件数を書いて明示する。
                cut += 1
                continue
            lack = need * 60 - longest_rest(start, end, work)      # 要件(9h連続)の未達の程度
            comp = need * 60 - total_rest(start, end, work)        # 付与すべき代償休息（厚労省FAQ）
            if comp > 0.5:
                comp_min += comp
            if lack > 0.5:
                viol += 1
                short_min += lack
                detail.append((d, need, round(lack), round(max(0.0, comp))))
        if viol or comp_min > 0.5:
            nm_disp, dept = who.get(n, (n, ""))
            out.append({"name": nm_disp, "nm": n, "dept": A._dept_leaf(dept) or dept or "(不明)",
                        "pos": pos or jobs.get(n, ""), "viol": viol,
                        "short_h": round(short_min / 60.0, 2),
                        "comp_h": round(comp_min / 60.0, 2), "detail": detail})
    STATS["cut"] = cut
    if cut:
        log(f"  窓が翌月へ食い込むため未判定にした日 {cut}件（月末の1〜2日。脚注に明示すること）")
    STATS["shotei"] = shotei          # 兼業日の判定に使う（呼び手が参照）

    # ---- 代償休息の消化を各人に付ける（20260826）----
    all_days = sorted({d for v in shotei.values() for d in v} | {d for v in sched.values() for d in v})
    if month_end:
        all_days = [d for d in all_days if d.year == month_end.year and d.month == month_end.month]
    kg = np_ = pc = 0
    for o in out:
        # その人が実際に行を持っている日だけを見る（人別シートが無い人を満額計上しない）
        own = sorted((set(shotei.get(o["nm"], {})) | set(punch.get(o["nm"], {}))) & set(all_days))
        used, a1, a2, a3 = consumed_minutes(punch.get(o["nm"], {}), shotei.get(o["nm"], {}), own)
        o["used_h"] = round(used / 60.0, 2)
        o["left_h"] = round(max(0.0, o["comp_h"] - used / 60.0), 2)
        kg += a1
        np_ += a2
        pc += a3
    STATS["kengyo_days"] = kg
    STATS["nopunch_days"] = np_
    STATS["punch_days"] = pc
    STATS["comp_total"] = round(sum(o["comp_h"] for o in out), 1)
    STATS["comp_people"] = len([o for o in out if o["comp_h"] > 0.01])
    STATS["used_total"] = round(sum(o["used_h"] for o in out), 1)
    STATS["left_total"] = round(sum(o["left_h"] for o in out), 1)
    STATS["left_people"] = len([o for o in out if o["left_h"] > 0.01])
    log(f"  代償休息 発生 {STATS['comp_total']}h/{STATS['comp_people']}名"
        f" ／ 消化 {STATS['used_total']}h ／ **未消化 {STATS['left_total']}h/{STATS['left_people']}名**")
    log(f"  平日の内訳: 兼業日として除外 {kg}日 / 打刻なしで満額計上 {np_}日 / 打刻あり {pc}日")

    out.sort(key=lambda x: -x["left_h"])
    log(f"  不足あり {len(out)}名 / 合計 {round(sum(o['short_h'] for o in out), 1)}h")
    return out


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\user003\Desktop\R8.7月分ローカル"
    res = compute(folder)
    print(f"\n{'不足計':>8} {'回数':>4}  {'職位':<14} {'科':<20} 氏名")
    for o in res[:25]:
        print(f"{o['short_h']:>8.2f} {o['viol']:>4}  {o['pos']:<14} {o['dept']:<20} {o['name']}")
