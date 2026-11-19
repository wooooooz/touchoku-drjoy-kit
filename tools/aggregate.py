# -*- coding: utf-8 -*-
"""aggregate.py  20260927
集約器: 提出フォルダ（*.json / *.xlsx / *.csv）＋枠マスタ＋名簿 → 正規JSON → DTX14互換xlsx。

  python -X utf8 tools/aggregate.py --year 2026 --month 9 --in <提出dir> --out <出力dir>
      [--staffs <staffs.csv>] [--frames data/frames.json]

  - 提出JSON       … アプリ（Win exe / Mac HTML）の提出。dept単位・原子レコード入り
  - 提出xlsx/csv   … 入力規則付きExcelテンプレ（主経路）／CSV（従経路）。
                     ファイル名に科名が入っている前提: 当直入力_R8年9月_救急科.xlsx など
                     （科名が読めなければ中の枠名で判定を試みる）
  - どの経路も同じ検証（名簿突合・枠実在・capacity・重複）を必ず通す。弾いた行は
    集約レポートに理由つきで全部出す（黙って捨てない）
  - 提出が無い科は必ず一覧に出す（現行Module5の「選び忘れた科が黙って残る」の対策）

出力:
  duty-YYYY-MM.json                    … 正規JSON（§1.5 月次パッケージ）
  互換3シート_R{reiwa}.{m}月_(アプリ生成).xlsx … 人別当直日/人別オンコール日/人別ファースト
  集約レポート.txt                       … 提出状況・当直交代・弾いた行・警告（UTF-8 BOM）
  ※ファイル名に DTX14 は入れない（本物と同じフォルダで FindLatestFile に拾われる事故の防止。
    切替の月に本物を抜いてからリネームする＝設計案§6-C-10）

report_only=True（管理ツール②「提出状況を確かめる」。管理者 20260925）:
  集約レポート.txt だけを作る（duty-YYYY-MM.json・互換3シートは作らない）。中身は 提出が無い科・「当直のみ」の申告・
  名簿外の医師・重複/弾いた行・提出があった科。当直交代の節は出さない（交代は 当直交代申請.xlsx が扱う）。
  レポートの全文を標準出力にも出す（管理ツールのログ欄に出る）。

20260927（管理者: 勤務の区分は フルタイム／パートタイム／院外研修専従 の3つ）:
  名簿の各人の区分は contract.work_of（work があればそれ、無い旧い名簿は in_house=False → パートタイム）。
  duty_only（JSON の名前は従来どおり）= フルタイム以外＝Dr.JOY の週パターンを「勤務なし」にする人。各項目に "work" を足した。
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402


def dept_from_filename(path, frames):
    base = os.path.splitext(os.path.basename(path))[0]
    owners = sorted(set(C.owners_of(frames)), key=len, reverse=True)
    for o in owners:
        if C.namekey(o) and C.namekey(o) in C.namekey(base):
            return o
    return None


def aggregate(year, month, in_dir, out_dir, staffs_path, frames_path=None, holidays_path=None,
              master_path=None, report_only=False):
    """戻り: (pkg, 互換3シートのパス)。report_only=True なら互換3シート・duty JSON は作らず (pkg, None)。
    pkg["report_path"] = 集約レポート.txt。"""
    frames = C.load_frames(frames_path)
    holidays = C.load_holidays(holidays_path)
    people, exact, canon_map = C.load_staffs(staffs_path)
    os.makedirs(out_dir, exist_ok=True)

    all_records = []
    submissions = []
    rejects_all = []          # (source, line, row, why)
    problems_all = []
    duty_only = []            # (dept, name, title, login, work) フルタイム以外の申告（2026-08-30「当直のみ」→ 20260927 区分3つ）
    unregistered = []         # (dept, name, title, resolved_login, how) Dr.JOY未登録の手入力医師（20260902）

    # 入力は共有マスター1本（20260912ユーザー決定。提出フォルダは廃止）。
    # in_dir に旧来の提出ファイルが残っていれば、移行のあいだは併せて読む
    import master_store as MS
    mpath = master_path or MS.master_path(os.path.dirname(os.path.abspath(in_dir)), year, month)
    master = MS.load(mpath, year, month) if os.path.exists(mpath) else MS.empty(year, month)
    changes_all = []          # (dept, 交代の記録)
    for dept, ent in sorted((master.get("depts") or {}).items()):
        sub = {"schema": "hospital.duty-submission/1", "year": year, "month": month, "dept": dept,
               "roster": ent.get("roster") or [], "duties": ent.get("duties") or []}
        recs, probs = C.validate_submission(sub, frames, holidays, people, exact, canon_map)
        n_oc = C.apply_daytime_oncall(recs)          # 開院記念日などの昼は全員オンコール（20260922）
        if n_oc:
            print(f"  {dept}: 休診日の昼 {n_oc} 件を当直→オンコールに変えた（宿日直手当なしの日）")
        for pr in probs:
            rejects_all.append((dept, None, pr["row"], pr["why"]))
        for p in sub["roster"]:
            if C.work_of(p) != "full":
                duty_only.append((dept, p.get("name") or "", p.get("title") or "", p.get("login") or "", C.work_of(p)))
            if p.get("unregistered") or str(p.get("login") or "").startswith("unreg:"):
                who, how = C.resolve_name(p.get("name") or "", people, exact, canon_map)
                unregistered.append((dept, p.get("name") or "", p.get("title") or "",
                                     who["login"] if who else "", how))
        for ch in ent.get("changes") or []:
            changes_all.append((dept, ch))
        submissions.append({"dept": dept, "file": os.path.basename(mpath),
                            "submitted_at": ent.get("updated", ""), "by": ent.get("by", ""),
                            "count": len(recs), "rejected": len(probs)})
        all_records.extend(recs)

    files = sorted(
        p for p in glob.glob(os.path.join(in_dir, "*"))
        if os.path.isfile(p) and os.path.splitext(p)[1].lower() in (".json", ".xlsx", ".csv")
        and not os.path.basename(p).startswith("~$"))
    for path in files:
        name = os.path.basename(path)
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".json":
                sub = C.load_submission_json(path)
                if (sub.get("year"), sub.get("month")) != (year, month):
                    problems_all.append((name, f"対象月が違います({sub.get('year')}/{sub.get('month')})。取り込みません"))
                    continue
                recs, probs = C.validate_submission(sub, frames, holidays, people, exact, canon_map)
                C.apply_daytime_oncall(recs)
                for p in probs:
                    rejects_all.append((name, None, p["row"], p["why"]))
                dept = sub.get("dept") or "?"
                for p in sub.get("roster") or []:
                    if C.work_of(p) != "full":
                        duty_only.append((dept, p.get("name") or "", p.get("title") or "",
                                          p.get("login") or "", C.work_of(p)))
                    if p.get("unregistered") or str(p.get("login") or "").startswith("unreg:"):
                        # Dr.JOY未登録の手入力医師 → 登録済みになっていれば氏名で本人に結ぶ
                        who, how = C.resolve_name(p.get("name") or "", people, exact, canon_map)
                        unregistered.append((dept, p.get("name") or "", p.get("title") or "",
                                             who["login"] if who else "", how))
                submissions.append({"dept": dept, "file": name,
                                    "submitted_at": sub.get("submitted_at", ""),
                                    "by": sub.get("by", ""), "count": len(recs),
                                    "rejected": len(probs)})
                all_records.extend(recs)
            else:
                owner = dept_from_filename(path, frames)
                if owner is None:
                    problems_all.append((name, "ファイル名から科名が読めません（当直入力_R8年9月_科名.xlsx の形にしてください）"))
                    continue
                rows = C.read_import_file(path, year, month)
                recs, rejs = C.import_rows(rows, owner, year, month, frames, holidays,
                                           people, exact, canon_map)
                for r in rejs:
                    rejects_all.append((name, r.get("line"), r.get("row"), r["why"]))
                submissions.append({"dept": owner, "file": name, "submitted_at": "",
                                    "by": "excel/csv", "count": len(recs), "rejected": len(rejs)})
                all_records.extend(recs)
        except Exception as e:
            problems_all.append((name, f"読み込みに失敗: {e}"))

    # 同じ科の提出が複数 → 後勝ちにはしない。両方読んで重複キーだけ警告
    keyc = Counter((r["staff_login"], r["date"], r["band"], r["frame_id"]) for r in all_records)
    dups = {k: c for k, c in keyc.items() if c > 1}
    if dups:
        seen = set()
        uniq = []
        for r in all_records:
            k = (r["staff_login"], r["date"], r["band"], r["frame_id"])
            if k in dups and k in seen:
                continue
            seen.add(k)
            uniq.append(r)
        all_records = uniq

    submitted = {s["dept"] for s in submissions}
    owners = sorted(C.owners_of(frames))
    missing = [o for o in owners if o not in submitted]

    reiwa = year - 2018
    pkg = {"schema": "hospital.duty-month/1", "year": year, "month": month,
           "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
           "frames": frames,
           "duties": all_records,
           "submissions": submissions,
           "missing_depts": missing,
           "duty_only": [{"dept": d, "name": n, "title": t, "login": lg, "work": w}
                         for d, n, t, lg, w in sorted(duty_only)],
           "unregistered": [{"dept": d, "name": n, "title": t, "resolved_login": lg, "how": how}
                            for d, n, t, lg, how in sorted(unregistered)],
           "changes": [dict(ch, dept=d) for d, ch in changes_all]}
    if report_only:
        duty_path = xlsx_path = None
    else:
        duty_path = os.path.join(out_dir, f"duty-{year}-{month:02d}.json")
        with open(duty_path, "w", encoding="utf-8") as f:
            json.dump(pkg, f, ensure_ascii=False, indent=1)
        xlsx_path = os.path.join(out_dir, f"互換3シート_R{reiwa}.{month}月_(アプリ生成).xlsx")
        counts = C.write_compat_xlsx(all_records, frames, xlsx_path, year, month)

    # レポート（人が読む・Excelで開かれうる → UTF-8 BOM + CRLF）
    if report_only:
        lines = report_lines(year, month, submissions, all_records, missing, duty_only, unregistered,
                             dups, rejects_all, problems_all)
    else:
        lines = [f"当直勤務表 集約レポート  R{reiwa}年{month}月  {dt.datetime.now():%Y-%m-%d %H:%M}", ""]
        lines.append(f"提出 {len(submissions)} 件 / 原子レコード {len(all_records)} 件"
                     f"（当直 {sum(1 for r in all_records if not r['oncall'])}"
                     f" / オンコール {sum(1 for r in all_records if r['oncall'])}）")
        lines.append(f"互換3シート: 人別当直日 {counts[C.P_DUTY]} / 人別オンコール日 {counts[C.P_ONCALL]}"
                     f" / 人別ファースト {counts[C.P_FIRST]}")
        lines.append("")
        lines.append("■ 提出があった科")
        for s in submissions:
            lines.append(f"  {s['dept']}: {s['count']}件"
                         + (f"（弾いた行 {s['rejected']}）" if s['rejected'] else "")
                         + f"  ← {s['file']}")
        if changes_all:
            lines.append("")
            lines.append(f"■ 当直交代（確定後に変わったもの・{len(changes_all)}件）　★下流へ流す前に確認")
            for dept, ch in changes_all:
                lines.append(f"  {dept}: {ch.get('date','')} 〔{ch.get('band','')}〕{ch.get('kind','')} "
                             f"{ch.get('from') or '(なし)'} → {ch.get('to') or '(なし)'}")
        lines.append("")
        lines.append(f"■ 提出が無い科（{len(missing)}）　★確認してから下流に渡すこと")
        for o in missing:
            lines.append(f"  {o}")
        if duty_only:
            lines.append("")
            lines.append(f"■ {DUTY_ONLY_HEAD}（{len(duty_only)}名）"
                         "　→ Dr.JOY勤務予定側の修正対象リスト")
            lines += duty_only_lines(duty_only)
        if unregistered:
            lines.append("")
            lines.append(f"■ 名簿外として手入力された医師（{len(unregistered)}名）"
                         "　→ Dr.JOYに登録する対象（登録済みなら氏名で本人に結び付けた）")
            for dept, nm, title, lg, how in sorted(unregistered):
                lines.append(f"  {dept}: {nm}" + (f"（{title}）" if title else "")
                             + (f"  → 名簿で解決 {lg}（{how}）" if lg else "  名簿外のまま（氏名で出力）"))
        if dups:
            lines.append("")
            lines.append(f"■ 提出間の重複（同一人・同日・同枠。1件だけ採用）: {len(dups)} 組")
            for k in list(dups)[:20]:
                lines.append(f"  {k}")
        if rejects_all:
            lines.append("")
            lines.append(f"■ 弾いた行（理由つき・全{len(rejects_all)}件）")
            for src, line, row, why in rejects_all:
                loc = f"{src}" + (f" {line}行目" if line else "")
                lines.append(f"  [{loc}] {why}")
                lines.append(f"      {json.dumps(row, ensure_ascii=False)}")
        if problems_all:
            lines.append("")
            lines.append("■ 読み込めなかったファイル")
            for name, why in problems_all:
                lines.append(f"  {name}: {why}")
    rep_path = os.path.join(out_dir, "集約レポート.txt")
    with open(rep_path, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(lines) + "\n")
    pkg["report_path"] = rep_path
    if report_only:
        print("\n".join(lines))
        print(f"\n（保存: {rep_path}）")
        return pkg, None

    print(f"集約: 提出{len(submissions)} レコード{len(all_records)} 弾き{len(rejects_all)} 未提出科{len(missing)}")
    print(f"  {duty_path}\n  {xlsx_path}\n  {rep_path}")
    return pkg, xlsx_path


# 見出し（旧「当直のみ」の語を残す＝以前のレポートと見比べる人・selftest が探す語）
DUTY_ONLY_HEAD = "パートタイム・院外研修専従（旧「当直のみ」。日中は当院勤務なし）の申告"


def duty_only_lines(duty_only):
    """フルタイム以外の申告の行（科: 氏名（職名）【区分】 ログインID）。"""
    return [f"  {dept}: {nm}" + (f"（{title}）" if title else "") + f"【{C.WORK_LABEL.get(w, w)}】"
            + (f"  {login}" if login else "")
            for dept, nm, title, login, w in sorted(duty_only)]


def report_lines(year, month, submissions, all_records, missing, duty_only, unregistered,
                 dups, rejects_all, problems_all):
    """管理ツール②「提出状況を確かめる」のレポート（report_only。管理者 20260925）。
    読む順: まだの科 → 当直のみ → 名簿外 → 重複・弾いた行 → 提出があった科。当直交代は載せない（当直交代申請.xlsx が扱う）。"""
    reiwa = year - 2018
    zero = [s["dept"] for s in submissions if not s["count"]]
    out = [f"当直勤務表 提出状況  R{reiwa}年{month}月分  {dt.datetime.now():%Y-%m-%d %H:%M} 現在", ""]
    out.append(f"提出があった科 {len(submissions)} ／ まだの科 {len(missing)}"
               f"　（当直 {sum(1 for r in all_records if not r['oncall'])} 件・"
               f"オンコール {sum(1 for r in all_records if r['oncall'])} 件）")
    out += ["", f"■ まだ提出が無い科（{len(missing)}）"]
    out += [f"  {o}" for o in missing] or ["  （なし。全科そろっています）"]
    if zero:
        out += ["", f"■ 提出はあるが当直が1件も入っていない科（{len(zero)}）"]
        out += [f"  {d}" for d in zero]
    if duty_only:
        out += ["", f"■ {DUTY_ONLY_HEAD}（{len(duty_only)}名）"
                    "　→ Dr.JOY の週パターンを「勤務なし」にする対象"]
        out += duty_only_lines(duty_only)
    else:
        out += ["", f"■ {DUTY_ONLY_HEAD}: なし"]
    if unregistered:
        out += ["", f"■ 名簿外の医師（{len(unregistered)}名）　→ Dr.JOY に登録してもらう対象"
                    "（登録済みなら氏名で本人に結び付けた）"]
        for dept, nm, title, lg, how in sorted(unregistered):
            out.append(f"  {dept}: {nm}" + (f"（{title}）" if title else "")
                       + (f"  → 名簿で解決 {lg}（{how}）" if lg else "  名簿外のまま（Dr.JOY に未登録）"))
    else:
        out += ["", "■ 名簿外の医師: なし"]
    if not (dups or rejects_all or problems_all):
        out += ["", "■ 重複・弾いた行: なし"]
    if dups:
        out += ["", f"■ 重複（同じ人・同じ日・同じ枠が2回。1件だけ数えた）: {len(dups)} 組"]
        for k in list(dups)[:20]:
            out.append(f"  {k[1]} {k[2]} {k[3]} {k[0]}")
    if rejects_all:
        out += ["", f"■ 弾いた行（理由つき・全{len(rejects_all)}件）"]
        for src, line, row, why in rejects_all:
            loc = f"{src}" + (f" {line}行目" if line else "")
            out.append(f"  [{loc}] {why}")
            out.append(f"      {json.dumps(row, ensure_ascii=False)}")
    if problems_all:
        out += ["", "■ 読み込めなかったファイル"]
        for name, why in problems_all:
            out.append(f"  {name}: {why}")
    out += ["", "■ 提出があった科"]
    for s in submissions:
        out.append(f"  {s['dept']}: {s['count']}件"
                   + (f"（弾いた行 {s['rejected']}）" if s["rejected"] else "")
                   + (f"  最終 {str(s['submitted_at'])[:16].replace('T', ' ')}" if s.get("submitted_at") else ""))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=0)
    ap.add_argument("--month", type=int, default=0)
    ap.add_argument("--in", dest="in_dir", default="")
    ap.add_argument("--out", dest="out_dir", default="")
    ap.add_argument("--staffs", default=None, help="名簿CSV（省略=Downloadsの最新 staffs*.csv・500行以上）")
    ap.add_argument("--frames", default=None)
    ap.add_argument("--master", default=None, help="共有マスター（省略=提出フォルダの親の master\）")
    a = ap.parse_args()
    a.staffs = a.staffs or C.newest_staffs()
    if not (a.year and a.month):
        import datetime as _dt
        _n = (_dt.date.today().replace(day=1) + _dt.timedelta(days=32)).replace(day=1)   # 既定=翌月（各科は翌月分を提出する）
        print("=== 当直勤務表 集約 ===")
        a.year = a.year or int(input(f"対象の年（西暦・空欄={_n.year}）: ").strip() or _n.year)
        a.month = a.month or int(input(f"対象の月（1-12・空欄={_n.month}）: ").strip() or _n.month)
    md = C.prod_month_dir(a.year, a.month)
    d_in = os.path.join(md, "提出") if os.path.isdir(md) else "提出"
    d_out = os.path.join(md, "集約") if os.path.isdir(md) else "出力"
    if not a.in_dir:
        a.in_dir = input(f"提出フォルダ（既定 {d_in}）: ").strip() or d_in
    if not a.out_dir:
        a.out_dir = input(f"出力フォルダ（既定 {d_out}）: ").strip() or d_out
    os.makedirs(a.out_dir, exist_ok=True)
    aggregate(a.year, a.month, a.in_dir, a.out_dir, a.staffs, a.frames, master_path=a.master)


if __name__ == "__main__":
    main()
