# -*- coding: utf-8 -*-
"""DTX14の当直 → Dr.JOY「勤務予定」インポート用xlsx を作る。20261028

★2026-10-28 管理者決定: 救急科・小児科の当直の前後に外す日中勤務は、当直勤務表アプリの
  day_off_plan（hospital.day-off-plan/1）を正にする。build(day_off=...) に渡すと、
  下の「当直明けの日を上限まで削る」（B=12:00で切る・C-1=明けでオフ）は使わず、その一覧で書き換える。
  2026-10-28 の試験（管理者の12月の予定）: 時刻を書くとその日はまるごと置き換わる（2件にならない）。
  区分だけで時刻の無い行は失敗。全日休みは 区分=明け・時刻なし で通る（C-1 のオフと同じ）。

★2026-08-26 に本番で実測して確定した仕様

  ・ひな形 = Dr.JOY `at0069`「勤務予定」→テンプレート（work-schedule_v3.xlsx）
  ・★**実データは16行目から**。11〜15行は記入例。11行目に書くと「読み込み件数0」で黙って終わる
  ・列 2=職員コード / 3=日付(yyyy/mm/dd) / 8=区分 /
    9,10,11=勤務種類1・開始・終了 / 12..=2つ目 / 5枠まで / 24,25=休憩1の開始・終了
  ・★**その日の勤務予定を上書きする**。当直だけ書くと通常勤務が消えるので所定も一緒に書く
  ・★**時間帯が重なる勤務種類は「重複しています」で弾かれる**
  ・★区分は必須。平日=`勤務日` / 土日祝=`所定休日`（`休日` は「存在しません」で弾かれる）
  ・★**代償休息が発生するシフトは登録を拒否される**
    （「YYYY/M/Dインターバル(18時間)が不足しています」）。過去実績の再現では一部が入らない
  ・当直の勤務種類（`at0046` で定義済み・宿日直許可の設定も済み）
      当直(夜) 17:15-32:30 / 当直(昼) 09:00-17:15         … 宿日直許可あり
      診療夜間等勤務(夜) 17:15-32:30 / 同(昼) 09:00-17:15  … 許可なし
  ・★Dr.JOYの定義は翌8:30固定。**翌日が休日なら 33:00 を明示的に書いて当院ルールに揃える**

使い方:
  python dtx14_to_drjoy_schedule.py "<月フォルダ>" [出力.xlsx]
"""
from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import sys
from collections import Counter, defaultdict

sys.path.insert(0, r"C:\tools\touchoku\drjoy_import")
import openpyxl
import dtx18_stats as A
import interval_v2 as V

TEMPLATE = r"C:\Users\user003\Downloads\work-schedule_v3.xlsx"
DATA_ROW = 16                      # ★実データの開始行
LUNCH = ("12:00", "13:00")
# ★当直明けでオフにする日の区分。
#   ★2026-09-05 訂正: インポートでも「明け」が使える。20260905
#     2026-08-28に「勤務日/所定休日だけ」と実測したのは誤りで、真因は
#     設定→基本(at0027)の「利用する区分の設定」で**明けがOFF**だったこと。
#     ONにしてから 1行→成功1、25行→成功25。エクスポートで区分＝明けを確認済み。
#   ⚠「勤務日」で勤務種類を空にする = 「必須項目なし（勤務種類）。」は今も同じ。
#   ⚠明けも所定休日も月の所定を7:45減らす（8/30実測）。違うのは意味と休日数のカウント。
#   ⚠法定休日は当院の事情で使わない方針（設定でOFF）。
AKE_KUBUN = "明け"

# ★B水準（専攻医・医師）の当直明けの終業時刻。20260907
#   連続勤務28時間の上限は当直日の始業＋28時間（始業8:30なら翌12:30）だが、
#   **12:00〜13:00は昼休みなので、予定を12:30で切ると昼休みの真ん中で終わる**。
#   現場も12:00で回している（R8.8月の半日勤務予定43日はすべて 08:30-12:00。
#   12:30は2026-09-05の投入で私が初めて入れた形で20日ぶんだけ）。
#   → 上限と12:00の**早いほう**で切る（ユーザー決定 2026-09-07）。
#   ⚠始業が早い科（歯科口腔外科は7:30始業）は上限のほうが早くなるので min が要る。
AKE_B_END = 12 * 60                       # 12:00
# ★当直の「昼」と「夜」の境目。昼夜通しの当直をここで2本に切る。20260829
DUTY_SPLIT = 17 * 60 + 15
_RE = re.compile(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})")
_CODE = re.compile(r"0?5\d{10}")   # Dr.JOYの職員コードの形


def hm(mins):
    return f"{mins // 60:02d}:{mins % 60:02d}"


def duty_kind(nopermit, band):
    """DTX14の枠 → (Dr.JOYの勤務種類, 開始, 終了分)。終了は翌日を跨ぐと24時超で表す。"""
    base = "診療夜間等勤務" if nopermit else "当直"
    if band == "昼":
        return f"{base}(昼)", 9 * 60, 17 * 60 + 15
    if band in ("夜", "休"):
        return f"{base}(夜)", 17 * 60 + 15, None      # 終了は呼び手が翌日休日で決める
    if band == "昼夜":
        return f"{base}(夜)", 9 * 60, None
    return None, None, None


def load_pay(pay_xlsx):
    """給与連携項目一覧 → ({氏名norm: 職員コード}, {氏名norm: {日付: 所定時間、休暇等}})

    ★当直情報.xlsm は月が終わってからでないと作れないので、**未来の月では使えない**。
      給与連携項目一覧は **at0029 で任意の期間を指定でき、未来の勤務予定も出せる**ので、
      投入（＝これから起きる当直を先に入れる）にはこちらを使う。
      必要な列は 職員コード / 氏名 / 日付 / 所定時間、休暇等 の4つだけ。
    """
    codes, shotei = {}, defaultdict(dict)
    wb = openpyxl.load_workbook(pay_xlsx, read_only=True, data_only=True)
    ws = A._ws_full(wb.worksheets[0])
    it = ws.iter_rows(values_only=True)
    h = [str(x or "").strip() for x in next(it)]
    for need in ("職員コード", "氏名", "日付", "所定時間、休暇等"):
        if need not in h:
            raise SystemExit(f"給与連携に「{need}」列がありません: {os.path.basename(pay_xlsx)}")
    iC, iN, iD, iS = (h.index(x) for x in ("職員コード", "氏名", "日付", "所定時間、休暇等"))
    for r in it:
        if len(r) <= max(iC, iN, iD, iS):
            continue
        nm = A.norm_name(r[iN])
        d = A.cell_date(r[iD])
        if not nm or d is None:
            continue
        v = r[iC]
        if v not in (None, ""):
            codes.setdefault(nm, str(int(v)) if isinstance(v, float) else str(v).strip())
        shotei[nm][d] = str(r[iS] or "").strip()
    wb.close()
    return codes, shotei


def load_staff_codes(staffs_csv):
    """staffs.csv（Dr.JOYのスタッフ一覧・CP932・28列） → {氏名norm: 職員コード}。20260827

    ★勤務予定エクスポートに載るのは「Dr.JOYに勤務予定が登録されている人」だけ。
      新採用者などは1シートも無いので職員コードが引けず、当直が丸ごと落ちていた
      （R8.9月＝18名67人日）。名簿から補う。
    """
    import csv
    import io
    out = {}
    with io.open(staffs_csv, encoding="cp932", errors="replace") as f:
        rd = csv.reader(f)
        h = next(rd)
        iC, iS, iM = h.index("職員コード"), h.index("姓"), h.index("名")
        for r in rd:
            if not any(r) or len(r) <= max(iC, iS, iM):
                continue
            c = r[iC].strip()
            if _CODE.fullmatch(c):
                out.setdefault(A.norm_name(r[iS].strip() + r[iM].strip()), c)
    return out


def load_staff_levels(staffs_csv):
    """staffs.csv → {職員コード: 特例水準指定}。20260828

    ★連続勤務の上限は特例水準で違う。
      C-1（初期研修医）＝24時間 ／ B（専攻医・医師ほか）＝28時間。
      宿日直許可の有無ではない（許可の有無は当直の種類＝当直/診療夜間等勤務を決める）。
    """
    import csv
    import io
    out = {}
    with io.open(staffs_csv, encoding="cp932", errors="replace") as f:
        rd = csv.reader(f)
        h = next(rd)
        iC, iL = h.index("職員コード"), h.index("特例水準指定")
        for r in rd:
            if not any(r) or len(r) <= max(iC, iL):
                continue
            c = r[iC].strip()
            if _CODE.fullmatch(c):
                out[c] = r[iL].strip()
    return out


# ★DTX14のE列(フルネーム)の表記ゆれ。Dr.JOY側の氏名に寄せる。20260827
#   例: 当直表の誤字の氏名を、Dr.JOY の正しい氏名に寄せる（ユーザー確認済み 2026-08-27）。
NAME_ALIAS = {
    # 見本: DTX14 の当直表だけ字が違う医師を、Dr.JOY 側の氏名に寄せる。
    # 例 "見本修二朗": "見本修治朗"。自院の表記ゆれをここに足す。
    # ⚠ どちらが正しいかは人事給与システム・統合ファイルなど正本で確かめてから足すこと
    #   （Dr.JOY 側の手打ち間違いだったことがある＝変換が逆に壊す）。
}


def _finish(code, d, kubun, work, brk, nm, kind, ks, ke, ends, skip, band=None):
    """所定の後ろに当直をつなぐ。★時間帯が重なると Dr.JOY に弾かれるので、
    所定の終わりが当直の開始より後ならその分だけ当直をずらす。"""
    if ends:
        last_end = max(ends)
        if last_end > ks:
            ks = last_end
    if ke <= ks:
        skip["当直の時間帯が作れない"] += 1
        return None
    work = list(work)
    # ★休日の昼夜通し当直は「昼」と「夜」の2本に分けて書く（ユーザー指示 2026-08-29）。
    #   Dr.JOYの勤務種類に「昼夜」が無いため、従来は1本にまとめて
    #     当直(夜) 09:00〜翌朝
    #   と書いていた。時間は正しいが「夜なのに朝から」で画面が読めない。
    #   → 17:15 で切って 当直(昼) 09:00-17:15 ＋ 当直(夜) 17:15〜翌朝 にする。
    #   ⚠境目に接するだけ（重ならない）ので弾かれない。所定＋当直と同じつなぎ方。
    #   ⚠所定が 17:15 を越える日は昼の枠が作れないので、従来どおり1本で書く。
    if band == "昼夜" and ks < DUTY_SPLIT < ke:
        work.append((kind.replace("(夜)", "(昼)"), hm(ks), hm(DUTY_SPLIT)))
        work.append((kind, hm(DUTY_SPLIT), hm(ke)))
        skip["（当直）昼夜を昼と夜の2本に分けた"] += 1
    else:
        work.append((kind, hm(ks), hm(ke)))
    if len(work) > 5:
        skip["勤務種類が5枠を超える"] += 1
        return None
    return (code, d, kubun, work, brk, nm)


# 20261028 the list is built from the swap ledger's approved assignments (what DTX14 and Dr.JOY hold), not from the
# app's master (which already contains swaps still waiting for approval). The rules are the app's day_off_plan.plan().
DAY_OFF_TOOL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tools", "day_off_from_approved.py")
DAY_OFF_OWNERS = ("救急科", "小児科")


def _mins(v):
    """'8:30' / '08:30' -> minutes; None stays None."""
    if v in (None, ""):
        return None
    h, m = str(v).split(":")
    return int(h) * 60 + int(m)


def load_day_off(json_path, staffs_csv):
    """当直勤務表アプリの day_off_plan（--json）→ [{code, date, kind, start, end, name, why}]。20261028

    職員コードは入っていないので、staffs.csv のログインIDで引く（無ければ氏名で引く）。
    引けない人は code=None のまま返し、build() が「要確認」に数える。"""
    import csv
    import io
    import json
    plan = json.load(open(json_path, encoding="utf-8"))
    if plan.get("schema") != "hospital.day-off-plan/1":
        raise SystemExit(f"day_off_plan の形が違います: {plan.get('schema')}")
    by_login, by_name = {}, {}
    with io.open(staffs_csv, encoding="cp932", errors="replace") as f:
        rd = csv.reader(f)
        h = next(rd)
        iC, iL, iS, iM = h.index("職員コード"), h.index("ログインID"), h.index("姓"), h.index("名")
        for r in rd:
            if len(r) <= max(iC, iL, iS, iM):
                continue
            c = r[iC].strip()
            if not _CODE.fullmatch(c):
                continue
            by_login.setdefault(r[iL].strip().lower(), c)
            by_name.setdefault(A.norm_name(r[iS].strip() + r[iM].strip()), c)
    out = []
    for o in plan.get("offs", []):
        if o.get("owner") not in DAY_OFF_OWNERS:
            continue
        code = by_login.get(str(o.get("login") or "").strip().lower()) or by_name.get(A.norm_name(o.get("name") or ""))
        out.append({"code": code, "date": dt.date.fromisoformat(o["date"]), "kind": o["kind"],
                     "start": _mins(o.get("start")), "end": _mins(o.get("end")),
                     "name": o.get("name") or "", "why": o.get("why") or ""})
    return out, plan.get("notes") or []


def make_day_off(month, out_json):
    """tools/day_off_from_approved.py を走らせて JSON を作る（決め方は当直勤務表アプリの day_off_plan.plan、
    割当は交代記録台帳の approved。Dr.JOY には触らない）。"""
    import subprocess
    r = subprocess.run([sys.executable, "-X", "utf8", DAY_OFF_TOOL, "--month", month, "--json", out_json],
                       cwd=os.path.dirname(DAY_OFF_TOOL), capture_output=True)
    if r.returncode != 0 or not os.path.exists(out_json):
        raise SystemExit("day_off_plan.py が失敗しました:\n" + r.stdout.decode("utf-8", "replace")[-800:]
                         + r.stderr.decode("utf-8", "replace")[-800:])
    return out_json


def apply_day_off(rows, offs, plan_code, skip, log, plan_next=None, run_start=None, duty_end=None):
    """救急科・小児科の外す日中勤務を rows に当てる。20261028

    早帰り・遅出: その日の通常勤務を [start, end] と重なる部分だけに縮める（勤務の無い日に勤務を作らない）。
    全日休み    : その日の通常勤務を消す。何も残らなければ 区分=明け・時刻なし（試験で通った形）。
    当直など通常勤務でない行（その夜の当直）はそのまま残す。休憩は残る勤務の中に入るものだけ残す。"""
    idx = {(r0[0], r0[1]): i for i, r0 in enumerate(rows)}
    for o in offs:
        code, d = o["code"], o["date"]
        if not code:
            skip["★（休み）職員コードが引けない（要確認）"] += 1
            log(f"  ★要確認 職員コードが引けない: {o['name']} {d} {o['kind']}")
            continue
        if (code, d) in idx:
            _c, _d, kubun, work_s, brk_s, nm = rows[idx[(code, d)]]
            work = [(k, _mins(s), _mins(e)) for k, s, e in work_s]
            brk = [(_mins(b0), _mins(b1)) for b0, b1 in brk_s]
        else:
            day = (plan_code or {}).get(code, {}).get(d) or (plan_next or {}).get(code, {}).get(d)
            if day is None:
                if o["kind"] == "全日休み" and d.weekday() < 5 and not A.is_holiday_date(d):
                    # 翌月ぶんの明け（月末の当直）で予定が見えないとき。全日休みは中身を見なくても書ける
                    # （以前の C-1 の「翌月ぶんをオフにした」と同じ）。
                    rows.append((code, d, AKE_KUBUN, [], [], o["name"]))
                    idx[(code, d)] = len(rows) - 1
                    skip["（休み）予定が見えない翌月の全日休みを書いた"] += 1
                    continue
                skip["★（休み）Dr.JOYにその日の勤務予定が無い（要確認）"] += 1
                log(f"  ★要確認 予定が見えない: {o['name']} {d} {o['kind']}")
                continue
            kubun, nm = day["区分"] or "勤務日", o["name"]
            work, brk = list(day["work"]), list(day["brk"])
        normal = [w for w in work if w[0] == "通常勤務"]
        other = [w for w in work if w[0] != "通常勤務"]
        if not normal:
            skip["（休み）その日は元から通常勤務が無い"] += 1
            continue
        if o["kind"] == "全日休み":
            new_normal = []
        elif o["start"] is None or o["end"] is None:
            skip["★（休み）時刻の無い早帰り・遅出（要確認）"] += 1
            log(f"  ★要確認 時刻が無い: {o['name']} {d} {o['kind']}")
            continue
        else:
            w0, w1 = o["start"], o["end"]
            prev = (code, d - dt.timedelta(days=1))
            if "早帰り" in o["kind"] and run_start and prev in run_start:
                # ★一覧の「12:30まで」は当直日の始業 8:30 が前提（8:30＋28時間）。始業が早い人
                #   （6:30・7:20 の早番）は当直が早く終わり、12:30 まで書くと28時間を超えて弾かれうる。
                #   → 当直の終わりから始め、当直日の始業＋28時間で切る（8:30始業なら一覧どおり）。20261028
                cap = run_start[prev] + 28 * 60 - 1440
                lo = (duty_end or {}).get(prev, 32 * 60 + 30) - 1440
                w0, w1 = max(0, lo), min(w1, cap)
                if (w0, w1) != (o["start"], o["end"]):
                    skip["★（休み）当直日の始業が早いので28時間で切った（一覧と違う・要確認）"] += 1
                    log(f"  ★要確認 28時間で切った: {o['name']} {d} 一覧 {hm(o['start'])}-{hm(o['end'])} → {hm(w0)}-{hm(w1)}")
            new_normal = [(k, max(s, w0), min(e, w1)) for k, s, e in normal
                          if min(e, w1) > max(s, w0)]
        new_work = sorted(other + new_normal, key=lambda w: w[1])
        if new_work == sorted(work, key=lambda w: w[1]):
            skip["（休み）元から一覧どおり"] += 1
            continue
        new_brk = [(b0, b1) for b0, b1 in brk
                   if any(s <= b0 and b1 <= e for _k, s, e in new_work)]
        new_kubun = AKE_KUBUN if not new_work else kubun
        row = (code, d, new_kubun, [(k, hm(s), hm(e)) for k, s, e in new_work],
               [(hm(b0), hm(b1)) for b0, b1 in new_brk], nm)
        if (code, d) in idx:
            rows[idx[(code, d)]] = row
        else:
            idx[(code, d)] = len(rows)
            rows.append(row)
        skip[f"（休み）{o['kind']}を書いた"] += 1
    rows.sort(key=lambda r0: (r0[1], r0[0]))
    return rows


def build(folder, out_path, dtx14=None, resolve=None, plan_xlsx=None,
          tougou_codes=None, staffs_csv=None, ake=True, force=None,
          duty=None, plan=None, extra_codes=None, busy=None, day_off=None, plan_next_xlsx=None):
    """dtx14 を渡すとそのファイルを使う（SharePointの同期が遅れるので実体を渡せるように）。

    20260923 月の途中の当直交代（tools/drjoy_kotai_sync.py）のための入口:
      duty        = {氏名キー: {日付: (許可なしか, "昼"/"夜"/"昼夜")}} を直接渡す（DTX14 を読まない）
      plan        = dtx14_drjoy_plan.load() の戻り (by_code, by_name, codes) を読み込み済みで渡す
                    （2か月ぶんを合わせたものを渡せる。月末の当直の明けが翌月になるため）
      extra_codes = {氏名キー: 職員コード}。交代記録は職員コードを持っているのでそれを優先する
      busy        = {氏名キー: {日付,...}}。duty に入れていない当直日。翌日が本人の当直日なら
                    明けとして書き換えないための守り（duty に無い日を明けで潰さない）
      out_path=None なら書き出さずに行だけ返す。
    20261028 day_off = load_day_off() の1つ目（救急科・小児科の外す日中勤務）。渡すと
      許可なし当直の明けを自分で削らず、この一覧で書き換える。一覧に無い明けは「要確認」に数える。
      翌月の一覧も混ぜて渡してよい。翌月の日付は、この月の当直の明けに当たる行だけ使う（月末の当直の明け）。
      plan_next_xlsx = 翌月の勤務予定エクスポート。翌月の明けの日の予定を読むのに使う。
    resolve = dtx14_names.build_resolver() … E列(フルネーム)が空の当直行を氏名まで解決する。
    plan_xlsx = Dr.JOYの勤務予定エクスポート（at0029・種別=勤務予定）。

    ★投入は**その日の勤務予定を丸ごと上書きする**ので、いま登録されている予定を
      そのまま書き戻したうえで当直を足す。plan_xlsx があればそれを使う
      （勤務種類・開始・終了・休憩がそのまま入っており、年休や遅出シフトも潰さない）。
      無ければ給与連携の「所定時間、休暇等」から組み立てる（休憩の位置は推定になる）。
    """
    if duty is None:
        duty = V.load_duty(dtx14 or A.find_latest(folder, "*DTX14*.xlsm"), log=print, resolve=resolve)
    plan_code = plan_name = shotei = None
    if plan is not None:
        plan_code, plan_name, codes = plan[0], plan[1], dict(plan[2])
        if staffs_csv:
            for nm2, c2 in load_staff_codes(staffs_csv).items():
                codes.setdefault(nm2, c2)
    elif plan_xlsx:
        import dtx14_drjoy_plan as P
        plan_code, plan_name, codes = P.load(plan_xlsx)
        # ★統合ファイルの職員コードは **線形合同スクランブル＋Base64 の Token**。
        #   `tougou_token.decode_token()` で平文に戻せる（check_zaiseki と同じ方式）。
        #   実データ596名で往復検証し594名一致を確認済み。
        #   ⚠復号しないまま使うと `MjcwOTUyOTQ=` のようなでたらめなコードで投入する。
        n0 = len(codes)
        for nm, c in (tougou_codes or {}).items():
            if _CODE.fullmatch(str(c).strip()):
                codes.setdefault(nm, str(c).strip())
        if len(codes) > n0:
            print(f"  職員コードを補完: {len(codes) - n0}名")
        if staffs_csv:
            n1 = len(codes)
            for nm2, c2 in load_staff_codes(staffs_csv).items():
                codes.setdefault(nm2, c2)
            if len(codes) > n1:
                print(f"  staffs.csvで職員コードを補完: {len(codes) - n1}名")
    else:
        codes, shotei = load_pay(A.find_latest(folder, "*給与連携項目一覧*.xlsx"))
        print(f"  給与連携: 職員コード {len(codes)}名 / 所定の分かる人日 "
              f"{sum(len(v) for v in shotei.values())}")

    if extra_codes:
        codes.update(extra_codes)
    rows, skip = [], Counter()
    # ★連続勤務の上限は特例水準で違う（C-1=24時間／B=28時間）。20260828
    levels = load_staff_levels(staffs_csv) if staffs_csv else {}
    run_start = {}      # (職員コード, 日付) → 実際に書いた当直日の勤務開始（分）
    duty_end = {}       # (職員コード, 日付) → 実際に書いた当直の終了（分・24時超え）
    for nm in sorted(duty):
        for d in sorted(duty[nm]):
            nopermit, band = duty[nm][d]
            key = NAME_ALIAS.get(nm, nm)          # ★E列の表記ゆれをDr.JOY側に寄せる
            code = codes.get(key)
            if not code:
                skip["職員コードが分からない"] += 1
                continue
            kind, ks, _ = duty_kind(nopermit, band)
            if kind is None:
                skip["勤務形態が不明"] += 1
                continue
            ke = 33 * 60 if A.is_holiday_date(d + dt.timedelta(days=1)) else 32 * 60 + 30
            if band == "昼":
                ke = 17 * 60 + 15

            if plan_code is not None and band != "昼":
                # ★翌朝の勤務が 8:30 より早い人は、当直の終わりをそこまでにする。20260828
                #   歯科口腔外科は 6:45〜7:39 始業で、32:30(=翌8:30)まで書くと
                #   翌日の通常勤務と重なり **その人の行が全部弾かれる**
                #   （Dr.JOYのメッセージは「重複しています（9/29と9/30）」と1組だけ出るが、
                #     実際には同じ人の他の日も道連れになる）。R8.9月は4名21件が該当。
                sheet = plan_code.get(code) or plan_name.get(key) or {}
                nxt = sheet.get(d + dt.timedelta(days=1))
                if nxt is not None:
                    starts = [s3 for _k3, s3, _e3 in nxt["work"]]
                    if starts:
                        ke = min(ke, min(starts) + 24 * 60)
                elif not A.is_holiday_date(d + dt.timedelta(days=1)):
                    # 月末＝翌日がエクスポートに無い。その人の当月の最早始業で代用する
                    early = [s3 for day2 in sheet.values()
                             for k3, s3, _e3 in day2["work"] if k3 == "通常勤務"]
                    if early:
                        ke = min(ke, min(early) + 24 * 60)

            if plan_code is not None:
                # ★いま登録されている勤務予定をそのまま書き戻す
                day = (plan_code.get(code) or plan_name.get(key) or {}).get(d)
                if day is None:
                    # ★Dr.JOYにその日の勤務予定が無い＝**潰すものが無い**ので当直だけ書く。
                    #   勤務予定エクスポートに1シートも無い人（新採用者など）が該当する。
                    #   従来はここで見送っており、R8.9月は18名67人日が丸ごと落ちていた。
                    #   ⚠区分は必須なので土日祝=所定休日・平日=勤務日を自分で決める。20260827
                    kubun = "所定休日" if (d.weekday() >= 5 or A.is_holiday_date(d)) else "勤務日"
                    row = _finish(code, d, kubun, [], [], nm, kind, ks, ke, [], skip, band)
                    if row:
                        rows.append(row)
                        skip["（当直のみ）Dr.JOYに勤務予定が無いので当直だけ書いた"] += 1
                    continue
                kubun = day["区分"] or "勤務日"
                # ★その日に**既に当直が登録されている**なら足さない。
                #   足すと当直が2つ並び、`当直(夜) 32:30-33:00` のような断片ができる。
                if any(k.startswith(("当直", "診療夜間等勤務")) for k, _s, _e in day["work"]):
                    # ★force に入れた日は書き直す。既にある当直は落としてから足す
                    #   （落とさないと当直が2本になる）。20260828
                    if not force or (code, d) not in force:
                        skip["既にDr.JOYに当直が入っている"] += 1
                        continue
                    day = dict(day)
                    day["work"] = [(k, s5, e5) for k, s5, e5 in day["work"]
                                   if not k.startswith(("当直", "診療夜間等勤務"))]
                    skip["★既にある当直を書き直した"] += 1
                w0 = list(day["work"])
                # ★当直の終わりを、連続勤務の上限まで縮める。
                #   （ユーザー指示 2026-08-28「金曜朝の開始を9:00にするのではなく、
                #     土曜朝の終了を8:30にしてほしい」）
                #   金曜に当直だと明けが土曜9:00（=33:00）になるので
                #     08:30-17:15 通常勤務 ＋ 17:15-33:00 当直 ＝ 24時間30分
                #   となり、C-1(24時間)を30分だけ超える。翌日は元から休みで削れない。
                #   → 当直の終わりを 08:30＋24時間＝32:30（翌8:30）にする。
                #   ⚠許可ありの当直は労働時間に算入しないので触らない。
                if nopermit and levels and band != "昼":
                    limit = (24 if str(levels.get(code, "")).upper().startswith("C-1") else 28) * 60
                    s_now = min([s2 for _k2, s2, _e2 in w0] or [ks])
                    # ★許可なしの当直は終わりを**常に翌8:30**にする（休日でも9:00にしない）。
                    #   ユーザー指示 2026-08-28「土曜朝の終了を8:30にしてほしい」。
                    #   ①連続勤務: 08:30始業＋当直で 24時間ちょうどに収まる
                    #   ②インターバル: 翌8:30に終われば翌々日8:30まで24時間空く
                    #     （9:00終わりだと23時間30分で C-1 の24時間に30分足りない）
                    new_ke = min(ke, 32 * 60 + 30, s_now + limit)
                    if new_ke != ke:
                        ke = new_ke
                        skip["（当直）終わりを翌8:30までに縮めた"] += 1
                work = [(k, hm(s2), hm(e2)) for k, s2, e2 in w0]
                # ★休憩は1枠だけでなく全部書き戻す（ひな形は4枠）。20260828
                brk = [(hm(b0), hm(b1)) for b0, b1 in day["brk"]
                       if not w0 or (b0 >= min(s2 for _k2, s2, _e2 in w0))]
                ends = [e2 for _k, _s, e2 in w0]
                row = _finish(code, d, kubun, work, brk, nm, kind, ks, ke, ends, skip, band)
                if row:
                    run_start[(code, d)] = min([s2 for _k2, s2, _e2 in w0] or [ks])
                    duty_end[(code, d)] = ke
                if row:
                    rows.append(row)
                continue

            if nm not in shotei:
                skip["給与連携に居ない(Dr.JOY未登録)"] += 1
                continue
            s = str(shotei[nm].get(d, "") or "").strip()
            segs = [(int(a) * 60 + int(b), int(c) * 60 + int(e)) for a, b, c, e in _RE.findall(s)]
            work, brk, kubun = [], [], "勤務日"

            if not segs:
                # ⚠所定が空＝土日祝とは限らない。**年休・振替休日も空になる**。
                #   平日なのに空の日を「所定休日」で上書きすると、年休の記録を潰す。
                #   → 土日祝だけ所定休日にし、平日で空の日は見送る。
                if not (d.weekday() >= 5 or A.is_holiday_date(d)):
                    skip["平日なのに所定が空（年休・振替の疑い）"] += 1
                    continue
                kubun = "所定休日"
            elif len(segs) == 1 and segs[0] == (8 * 60 + 30, 17 * 60 + 15):
                work = [("通常勤務", "08:30", "17:15")]
                brk = [LUNCH]
            elif len(segs) >= 2:                            # 中抜け＝休憩は区間の間
                work = [("通常勤務", hm(a), hm(b)) for a, b in segs]
            elif segs[0][1] - segs[0][0] <= 360:            # 6時間以下は休憩不要
                work = [("通常勤務", hm(segs[0][0]), hm(segs[0][1]))]
            else:
                skip["変則で休憩の位置が分からない"] += 1     # 6時間超の1区間
                continue

            row = _finish(code, d, kubun, work, brk, nm, kind, ks, ke,
                          [b for _a, b in segs], skip, band)
            if row:
                rows.append(row)

    # ================================================================
    # ★当直明けの日を上限まで削る（ユーザー決定 2026-08-28）
    # ================================================================
    #   宿日直許可なしの当直（当直枠が救急科・小児科）は労働時間なので、
    #   当直日の始業から連続勤務が切れずに続く。上限は特例水準で違う。
    #     B水準   28時間 → 当直日の始業＋28時間まで。ただし 12:00 で頭打ち
    #                    （＝翌 08:30-12:00 の午前だけ勤務。AKE_B_END を見ること）
    #     C-1水準 24時間 → 当直日の始業＋24時間まで（＝翌 8:30。オフ）
    #   これを超えると Dr.JOY が「連続勤務(24h/28h超過)を超過しています」で弾く。
    #   ⚠許可ありの当直（外科・内科などの枠）は触らない。労働時間に算入しないため。
    #   ⚠翌日が本人の当直日でもあるときは触らない（当直の行が優先）。
    #   ⚠翌日がエクスポートに無いとき（月末＝翌月）は触れない。
    # 20261028 day_off があれば、ここでは明けを書かずに「一覧に入っているか」だけ確かめる。
    off_keys = {(o["code"], o["date"]) for o in (day_off or []) if o["code"]}
    ake_keys = set()            # (code, 明けの日) for every night duty without permission
    if ake and plan_code is not None and staffs_csv:
        have = {(r0[0], r0[1]) for r0 in rows}
        add = []
        for nm in sorted(duty):
            for d in sorted(duty[nm]):
                nopermit, band = duty[nm][d]
                if not nopermit or band == "昼":
                    continue                       # 許可あり／昼だけ＝明けが無い
                key = NAME_ALIAS.get(nm, nm)
                code = codes.get(key)
                if not code:
                    continue
                nxt = d + dt.timedelta(days=1)
                if (code, nxt) in have or nxt in duty[nm] or (busy and nxt in busy.get(nm, ())):
                    skip["（明け）翌日が本人の当直日なので触らない"] += 1
                    continue
                sheet = plan_code.get(code) or plan_name.get(key) or {}
                day2 = sheet.get(nxt)
                if day_off is not None:
                    # ★明けは当直勤務表アプリの一覧が正（管理者 2026-10-28）。自分では書かない。
                    #   平日の明けで通常勤務があるのに一覧に無ければ、上限超えで弾かれうるので知らせる。
                    ake_keys.add((code, nxt))
                    if (code, nxt) in off_keys:
                        skip["（明け）当直勤務表アプリの一覧に任せた"] += 1
                    elif (day2 is not None and day2["work"]) or (
                            day2 is None and nxt.weekday() < 5 and not A.is_holiday_date(nxt)):
                        skip["★（明け）一覧に無い明け（要確認）"] += 1
                        print(f"  ★要確認 一覧に無い明け: {nm} {d}→{nxt}")
                    continue
                isc1_ = str(levels.get(code, "")).upper().startswith("C-1")
                if day2 is None:
                    # ★月末の当直は明けが翌月になり、その月のエクスポートには出てこない。
                    #   C-1 は「明けはオフ」と決まっているので、中身を見なくても書ける。
                    #   （R8.9月＝冬野五郎 9/30 の明け 10/1 がこれ。翌10/1は木曜で通常勤務が
                    #     入っており、書かないと 9/30 08:30〜10/1 17:15＝32時間45分で弾かれる）
                    #   B は翌日の勤務を 12:00 で切るので、見えないと書けない。
                    if not (isc1_ and nxt.weekday() < 5 and not A.is_holiday_date(nxt)):
                        continue
                    add.append((code, nxt, AKE_KUBUN, [], [], nm))
                    skip["（明け）翌月ぶんをC-1=オフにした"] += 1
                    continue
                if not day2["work"]:
                    continue                       # 翌日が休み＝することが無い
                # ★起点は「実際に書いた当直日の勤務開始」。始業を遅らせた日はそれに合わせる。
                day0 = sheet.get(d)
                run_s = run_start.get((code, d))
                if run_s is None:
                    run_s = min([s4 for _k4, s4, _e4 in day0["work"]] or [8 * 60 + 30]) if day0 else 8 * 60 + 30
                if band in ("昼夜", "夜") and duty_kind(nopermit, band)[1] < run_s:
                    run_s = duty_kind(nopermit, band)[1]
                isc1 = str(levels.get(code, "")).upper().startswith("C-1")
                cur = [(k4, s4, e4) for k4, s4, e4 in day2["work"]]
                if isc1:
                    # ★C-1はオフ。上限で切ると休日の昼夜当直の明けが
                    #   「08:30-09:00＝30分だけ出勤」になり実務にならない。
                    work2, brk2 = [], []
                else:
                    # ★上限（28時間）と 12:00 の早いほうで切る。20260907
                    cap = min(run_s + 28 * 60 - 1440, AKE_B_END)   # 翌日のローカル時刻
                    # ★当直が終わる時刻（翌日側）。ここは main ループの外なので
                    #   ke をそのまま使うと別人の値が残っている。実際に書いた行から取る。
                    lo = max(0, duty_end.get((code, d), 32 * 60 + 30) - 1440)
                    if max(e4 for _k4, _s4, e4 in cur) <= cap and min(
                            s4 for _k4, s4, _e4 in cur) >= lo:
                        continue                       # 元から上限内
                    work2 = [(k4, max(s4, lo), min(e4, cap)) for k4, s4, e4 in cur
                             if min(e4, cap) > max(s4, lo)]
                    brk2 = [(b0, b1) for b0, b1 in day2["brk"] if b0 >= lo and b1 <= cap]
                kubun2 = day2["区分"] or "勤務日"
                if not work2:
                    kubun2 = AKE_KUBUN                # オフ（当直明け）
                add.append((code, nxt, kubun2,
                            [(k4, hm(s4), hm(e4)) for k4, s4, e4 in work2],
                            [(hm(b0), hm(b1)) for b0, b1 in brk2], nm))
                skip[f"（明け）{'C-1=オフ' if isc1 else 'B=午後休'}にした"] += 1
        rows.extend(add)
        rows.sort(key=lambda r0: (r0[1], r0[0]))

    if day_off:
        months = {(d.year, d.month) for nm in duty for d in duty[nm]}
        use = [o for o in day_off
               if (o["date"].year, o["date"].month) in months or (o["code"], o["date"]) in ake_keys]
        if len(use) < len(day_off):
            skip["（休み）翌月の一覧のうち、この月の当直の明けでないもの（翌月の投入で書く）"] += len(day_off) - len(use)
        plan_next = None
        if plan_next_xlsx:
            import dtx14_drjoy_plan as P
            plan_next = P.load(plan_next_xlsx, log=lambda *_a: None)[0]
        rows = apply_day_off(rows, use, plan_code, skip, print, plan_next=plan_next,
                             run_start=run_start, duty_end=duty_end)

    if out_path:
        write(out_path, rows)
    print(f"\n出力 {len(rows)}行 → {os.path.basename(out_path) if out_path else '（書き出さない）'}")
    for k, v in skip.most_common():
        print(f"  見送り {k}: {v}件")
    return rows


def write(out_path, rows):
    if os.path.exists(out_path):
        os.remove(out_path)
    shutil.copy2(TEMPLATE, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb["勤務予定"]
    # ★ひな形の11〜15行には記入例(例1〜例5)が入っている。消さずに投入すると
    #   2022年の架空データを一緒に投げることになる。
    for r in range(11, DATA_ROW):
        for c in range(1, 26):
            ws.cell(r, c).value = None
    for i, (code, d, kubun, work, brk, _nm) in enumerate(rows):
        r = DATA_ROW + i
        ws.cell(r, 2).value = code
        ws.cell(r, 3).value = d.strftime("%Y/%m/%d")
        ws.cell(r, 8).value = kubun
        for j, (k, s, e) in enumerate(work):
            ws.cell(r, 9 + j * 3).value = k
            ws.cell(r, 10 + j * 3).value = s
            ws.cell(r, 11 + j * 3).value = e
        for j, (bs, be) in enumerate(brk or []):
            if j >= 4:                       # ひな形の休憩は4枠まで
                break
            ws.cell(r, 24 + j * 2).value = bs
            ws.cell(r, 25 + j * 2).value = be
    wb.save(out_path)
    wb.close()


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\user003\Desktop\R8.7月分"
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "DrJOY勤務予定_当直.xlsx")
    build(folder, out)
