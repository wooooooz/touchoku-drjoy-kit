# -*- coding: utf-8 -*-
r"""new_month.py — 再来月の当直勤務表の置き場を作る（アプリの月フォルダ＋SharePoint の DTX14 雛形）。20261028

管理者 2026-09-23 決定:
  - SharePoint の `当直勤務表_正本\R8.N月 当直勤務表\` と `★R8.N月用 DTX14<版>.xlsm` を、こちらで自動的に作る（今までは事務担当）。
  - アプリの月フォルダ（本番共有 `R8年N月\`＝管理ツールの①）も自動で作る。締切は「前の月の20日 12:00」、
    土日祝に当たれば前の平日に寄せる。
  - 次の次の月の分を作る（15時の反映 dtx14_daily_sync の最後に呼ぶ）。

雛形の作り方: 前の月の正本 DTX14（SharePoint）をローカルに写し、dtx14_split_rows.split の月送り（A1 の月初シリアルを差し替え、
当直の行の日付セルを全部空にする。看護部などの手書きの行も空）→ 人別3シートを空にする → 同じ版の名前で置く。
既にあるものは作り直さない・上書きしない（フォルダがあっても正本の名前のファイルが無ければ置く）。
当番当直シートは年ごとの表なので、その月の分が data/toban.csv に無ければ知らせる（管理者が作って渡す。1月分は11/24ごろの依頼より前）。

20260925: month_setup.create は SystemExit で止まるので、ensure／15時の反映は (Exception, SystemExit) で受け止める
（1〜5日に作れないと、15時の反映の最後の警報とログが消えていた）。アプリの月フォルダが「ある」とみなすのは
config\current.json と共有マスターの両方があるときだけ（途中で落ちたフォルダを作り直す）。
DTX14 の雛形は、正本の名前のファイルが1本でもあれば作らない（2本あるときに3本目を作らない）。

20260926（レビュー）:
  - 1〜5日だけではなく毎日の15時に、次の次の月（と、欠けていれば来月）の置き場を揃える（ensure_targets・missing）。
    無いものだけ作り、あるものは上書きしない。作ったときだけ【お知らせ】メール（dtx14_daily_sync.maybe_new_month）。
    6日を過ぎてもまだ欠けていれば、毎日の警報（管理者へ）に載せる。
  - 手入力の部署（看護部・薬剤部など）に配る【診療部名】の写し（`【診療部名】R8.N月用 DTX14<版>.xlsm`＝正本の ★ を
    【診療部名】に替えた名前。中身は空の雛形のファースト表だけ。20261028 に【診療部科名】・全シートから変えた）も同じフォルダに置く（make_dtx14）。④の取り込みは表の月を確かめるので、
    写しも対象の月の雛形から作る。正本が先にあって写しだけ無いときは、正本を同じ月へ月送りして当直のマスと人別を空にして作る。
  - 当番当直が data/toban.csv に無い月の知らせは toban_note（15時の反映が毎日の警報に載せる）。

usage: python -X utf8 tools\new_month.py --month 2026-12 [--dry-run]
       python -X utf8 tools\new_month.py --auto          （来月〔欠けていれば〕と次の次の月）
"""
import argparse
import csv
import datetime as dt
import os
import re
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402

JINBETSU = ("人別当直日", "人別オンコール日", "人別ファースト")
WD = "月火水木金土日"
AUTO_DAYS = range(1, 6)        # 次の次の月が新しく対象になる日（この間は作れなくても「まだ無い」の警告にしない）
# 手入力の部署に配る写しの名前の頭（正本の ★ の代わり）。20261028 に【診療部科名】から変えた（管理者）:
#   診療科はアプリで入力から反映まで済むので、事務担当が R8.12 で自分で【診療部名】に直し、中身もファースト表だけにしていた
KAMEI = "【診療部名】"
SLIM_DROP = ("当直表 セカンド以下",) + JINBETSU          # 写しから消すシート（手入力の部署の行はファースト表にしか無い）
SLIM_HIDE = ("管理用⇒", "休日判定用", "当番当直")        # 写しで隠すシート（ファースト表の書式が参照するので消さない）


def add_months(y, m, k):
    n = y * 12 + (m - 1) + k
    return n // 12, n % 12 + 1


def deadline_for(y, m, holidays=None):
    """締切 = 前の月の20日 12:00。土日祝なら前の平日へ。戻り: '10月20日(火) 12:00' の形（config の deadline と同じ）。"""
    holidays = C.load_holidays() if holidays is None else holidays
    py, pm = add_months(y, m, -1)
    d = dt.date(py, pm, 20)
    while C.is_holiday(d, holidays):
        d -= dt.timedelta(days=1)
    return f"{d.month}月{d.day}日({WD[d.weekday()]}) 12:00"


def toban_has(y, m, path=None):
    path = path or os.path.join(C.DATA, "toban.csv")
    pre = f"{y:04d}-{m:02d}-"
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            return any((r.get("date") or "").startswith(pre) for r in csv.DictReader(f))
    except OSError:
        return False


def toban_note(y, m, path=None):
    """その月の当番当直が data/toban.csv に無ければ知らせる文。あれば None（20260926: 15時の反映が毎日の警報に載せる）。"""
    if toban_has(y, m, path):
        return None
    return (f"R{y - 2018}.{m}月の当番当直が data/toban.csv にありません（管理者が作って渡す。"
            f"DTX14 の「当番当直」シートにも入れる）")


# ---------------------------------------------------------------- アプリの月フォルダ（本番共有）

def make_app_month(y, m, dry=False, log=print, warnings=None):
    """warnings（リスト）= 作れたが知らせること（ACL の失敗。20260926）。"""
    import month_setup as MS
    md = MS.month_dir(y, m)
    cfg = os.path.join(md, "config", "current.json")
    if os.path.isfile(cfg) and os.path.isfile(MS.master_json(y, m)):
        log(f"  アプリの月フォルダはある: {md}")
        return None
    if os.path.isdir(md):
        log(f"  アプリの月フォルダはあるが config\\current.json か共有マスターが無い → 作り直す: {md}")
    dl = deadline_for(y, m)
    if dry:
        log(f"  （dry-run）アプリの月フォルダを作る: {md}（締切 {dl}）")
        return None
    MS.create(y, m, deadline=dl, log=lambda s: log("    " + s), warnings=warnings)
    return f"アプリの月フォルダ {os.path.basename(md)}（締切 {dl}）"


# ---------------------------------------------------------------- DTX14 の雛形（SharePoint）

def _clear_jinbetsu(path, log=print):
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3          # マクロは走らせない
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0)
        names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        for sh in JINBETSU:
            if sh not in names:
                continue
            ws = wb.Worksheets(sh)
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            if last >= 2:
                ws.Range(ws.Cells(2, 1), ws.Cells(last, 11)).ClearContents()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    log("    人別3シートを空にした")


def _slim(src, out, log=print):
    """空の雛形 src → 手入力の部署に配る写し out（ファースト表だけ。事務担当が R8.12 で手で作った形と同じ）。"""
    import win32com.client
    shutil.copy2(src, out)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3          # マクロは走らせない
    try:
        wb = xl.Workbooks.Open(os.path.abspath(out), UpdateLinks=0)
        names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        for sh in SLIM_DROP:
            if sh in names:
                wb.Worksheets(sh).Delete()
        for sh in SLIM_HIDE:
            if sh in names:
                wb.Worksheets(sh).Visible = 0          # xlSheetHidden
        wb.Worksheets("当直表 ファースト").Activate()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    log("    写しはファースト表だけにした")


def is_kamei_name(name, y, m):
    """作る写しの名前か（空白は無視。正本の名前の規則と同じ形で、★ の代わりに KAMEI）。"""
    core = name.replace(" ", "").replace("　", "")
    return re.fullmatch(rf"{re.escape(KAMEI)}R{y - 2018}\.{m}月用DTX14[0-9.]*\.xlsm", core) is not None


def is_copy_name(name, y, m):
    """手入力の部署向けの写しとみなす名前か（【…】で始まる。事務が名前を変えても写しはあるとみなす）。
    20261028: 事務担当が R8.12 の【診療部科名】を【診療部・科名】に変え【診療部名】を足したところ、15時の回が【診療部科名】を作り直した。"""
    core = name.replace(" ", "").replace("　", "")
    return re.fullmatch(rf"【[^】]+】R{y - 2018}\.{m}月用DTX14[0-9.]*\.xlsm", core) is not None


def kamei_candidates(d, y, m):
    """月フォルダにある手入力の部署向けの写しの一覧（【…】で始まるもの。フォルダが無ければ空）。"""
    try:
        names = os.listdir(d)
    except OSError:
        return []
    return [os.path.join(d, n) for n in sorted(names) if not n.startswith("~$") and is_copy_name(n, y, m)]


def kamei_name_of(master_name):
    """正本の名前 → 写しの名前（★ を KAMEI に替えるだけ。版と空白はそのまま）。"""
    return KAMEI + (master_name[1:] if master_name.startswith("★") else master_name)


def _empty_book(src, o_local, y, m, tmp, log=print):
    """src（前の月の正本、または同じ月の正本）→ 対象の月の空の雛形（当直のマスと人別3シートを空に）を o_local に作る。"""
    import dtx14_daily_sync as S
    import dtx14_split_rows as SR
    s_local = os.path.join(tmp, "src.xlsm")
    S.hydrate(src)                                   # OneDrive の「オンデマンド」なら中身を落としてから写す
    shutil.copy2(src, s_local)
    SR.split(s_local, o_local, roll=(y, m), keep_data=False, log=lambda s: log("    " + s))
    _clear_jinbetsu(o_local, log=log)


def _place(o_local, out, what, log=print):
    """作った雛形を置く。作っている間に誰かが置いていたら触らない（上書きしない）。戻り: 置いたか。"""
    os.makedirs(os.path.dirname(out), exist_ok=True)
    if os.path.exists(out):
        log(f"  {what}は作っている間に置かれていた（上書きしない）: {out}")
        return False
    shutil.copy2(o_local, out)
    return True


def make_dtx14(y, m, dry=False, log=print):
    """SharePoint の月フォルダに、正本の DTX14 の雛形と手入力の部署に配る写しを置く（無いものだけ。上書きしない）。
    正本が無ければ前の月の正本から作り、写しも同じ中身で置く。正本があって写しだけ無ければ、正本から空の写しを作る。
    戻り: 作ったものの説明（無ければ None）。作れなければ RuntimeError。"""
    import dtx14_daily_sync as S
    d, hits = S.dtx_candidates(y, m)
    kamei = kamei_candidates(d, y, m)
    if hits and kamei:
        log(f"  DTX14 の雛形はある: {', '.join(os.path.basename(h) for h in hits + kamei)}")
        return None
    name = None
    if hits:
        log(f"  DTX14 の雛形はある: {', '.join(os.path.basename(h) for h in hits)}（{KAMEI}の写しが無い）")
        if len(hits) > 1:
            raise RuntimeError(f"正本の名前の DTX14 が {len(hits)} 本あるので、{KAMEI}の写しの元を決められません: {d}")
        src = hits[0]
        kname = kamei_name_of(os.path.basename(src))
    else:
        py, pm = add_months(y, m, -1)
        _pd, src = S.dtx_file(py, pm)
        if not src:
            raise RuntimeError(f"前の月（R{S.reiwa(py)}.{pm}月）の正本 DTX14 が見つからないので、雛形を作れません: {_pd}")
        ver = re.sub(r"^.*DTX14", "", os.path.splitext(os.path.basename(src))[0]).replace(" ", "").replace("　", "")
        name = f"★R{S.reiwa(y)}.{m}月用 DTX14{ver}.xlsm"
        if not S.is_master_name(name, y, m):
            raise RuntimeError(f"作る名前が正本の規則に合いません: {name}")
        kname = None if kamei else kamei_name_of(name)
    if kname and not is_kamei_name(kname, y, m):
        raise RuntimeError(f"作る{KAMEI}の写しの名前が規則に合いません: {kname}")
    if dry:
        if name:
            log(f"  （dry-run）DTX14 の雛形を作る: {os.path.join(d, name)}（元 = {os.path.basename(src)}）")
        if kname:
            log(f"  （dry-run）{KAMEI}の写しを作る: {os.path.join(d, kname)}（元 = {os.path.basename(src)}）")
        return None
    made = []
    tmp = tempfile.mkdtemp(prefix="touchoku_newmonth_")
    try:
        o_local = os.path.join(tmp, "out.xlsm")
        _empty_book(src, o_local, y, m, tmp, log=log)
        if name and _place(o_local, os.path.join(d, name), "DTX14 の雛形", log):
            log(f"  DTX14 の雛形を置いた: {os.path.join(d, name)}（元 = {os.path.basename(src)}）")
            made.append(f"DTX14 の雛形 {name}")
        if kname:
            k_local = os.path.join(tmp, "copy.xlsm")
            _slim(o_local, k_local, log=log)
        if kname and _place(k_local, os.path.join(d, kname), f"{KAMEI}の写し", log):
            log(f"  {KAMEI}の写しを置いた: {os.path.join(d, kname)}（元 = {os.path.basename(src)}）")
            made.append(f"手入力の部署に配る写し {kname}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return f"{'・'.join(made)}（フォルダ {os.path.basename(d)}）" if made else None


def missing(y, m):
    """その月の置き場で欠けているもの（説明のリスト。空なら揃っている）。
    アプリの月フォルダ（config＋共有マスター）・正本の DTX14・手入力の部署に配る写し。
    本番の共有や SharePoint の同期フォルダに届かないときも「無い」と数える（届かないこと自体は15時の反映が別に知らせる）。"""
    import dtx14_daily_sync as S
    import month_setup as MS
    out = []
    md = MS.month_dir(y, m)
    if not (os.path.isfile(os.path.join(md, "config", "current.json")) and os.path.isfile(MS.master_json(y, m))):
        out.append(f"アプリの月フォルダ（{os.path.basename(md)}）")
    d, hits = S.dtx_candidates(y, m)
    if not hits:
        out.append(f"正本の DTX14 の雛形（{os.path.basename(d)}）")
    if not kamei_candidates(d, y, m):
        out.append(f"{KAMEI}の写し（{os.path.basename(d)}）")
    return out


def ensure(y, m, dry=False, log=print):
    """再来月の置き場を揃える（無いものだけ作る）。戻り: (作ったもの[], 知らせること[])。失敗は例外（呼び出し側が警報に載せる）。"""
    made, notes = [], []
    errs = []
    for fn in (make_app_month, make_dtx14):
        try:
            if fn is make_app_month:
                r = fn(y, m, dry=dry, log=log, warnings=notes)     # ACL の失敗は知らせることへ（20260926）
            else:
                r = fn(y, m, dry=dry, log=log)
            if r:
                made.append(r)
        except (Exception, SystemExit) as e:         # 片方の失敗でもう片方を止めない（month_setup は SystemExit）
            log(f"  ✗ {fn.__name__}: {type(e).__name__}: {e}")
            errs.append(f"{fn.__name__}: {e}")
    tn = toban_note(y, m)
    if tn:
        notes.append(tn)
    if errs:
        raise RuntimeError(" / ".join(errs))
    return made, notes


def auto_target(today=None):
    """次の次の月が新しく対象になった日（1〜5日）ならその月。ほかの日は None。"""
    t = today or dt.date.today()
    if t.day not in AUTO_DAYS:
        return None
    return add_months(t.year, t.month, 2)


def ensure_targets(today=None):
    """毎日揃える月（20260926）: [来月, 次の次の月]。来月は欠けているものがあるときだけ作る（呼び出し側が missing で見る）。"""
    t = today or dt.date.today()
    return [add_months(t.year, t.month, 1), add_months(t.year, t.month, 2)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="YYYY-MM")
    ap.add_argument("--auto", action="store_true", help="来月（欠けていれば）と次の次の月")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if a.month:
        targets = [tuple(int(x) for x in a.month.split("-"))]
    else:
        nxt, nn = ensure_targets()
        targets = ([nxt] if missing(*nxt) else []) + [nn]
    rc = 0
    for (y, m) in targets:
        print(f"■ R{y - 2018}.{m}月")
        try:
            made, notes = ensure(y, m, dry=a.dry_run)
        except RuntimeError as e:
            print("✗", e)
            rc = 1
            continue
        for s in made:
            print("作った:", s)
        for s in notes:
            print("⚠", s)
    return rc


if __name__ == "__main__":
    sys.exit(main())
