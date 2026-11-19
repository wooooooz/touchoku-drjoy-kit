# -*- coding: utf-8 -*-
"""dtx14_daily_sync.py — 毎日15時: 共有マスター → DTX14 現物へ反映し、PDF を公開用フォルダに置く。20261101

管理者 2026-09-22 決定:
  - 各科は月途中の交代もアプリで出し直す。管理側（管理者の PC）が毎日 15:00 に共有マスターを読み、
    Dr.JOY 投入の処理が読む現物 `当直勤務表_正本\\R8.N月 当直勤務表\\★R8.N月用 DTX14*.xlsm`（SharePoint 同期フォルダ）
    の診療科の行を上書きする（dtx14_fill.run_master。人別3シートも書く）。
  - 反映後、当直表2シート（ファースト=A4横／セカンド以下=A3横）を PDF にして
    `当直表チーム\\公開用\\当直勤務表_R8年N月.pdf` に置く。置くのは 前月・今月・来月 の2〜3本だけ。
    （SharePoint の共有リンクを各科に配る＝閲覧のみ）

対象の月 = 前月・今月・来月。DTX14 への書き込みは SYNC_FROM（2026-11）以降の月だけ（10月は予行で事務担当の手作業の現物）。
PDF は対象の月のファイルがあれば書き込みの有無にかかわらず作る。
ただし SYNC_FROM 以降の月は、共有マスターが確定（管理ツール③「この月を確定」＝事務担当）してからだけ PDF を公開する
（管理者 20260923）。未確定なら作らず、前に置いた PDF があれば消す。DTX14 への書き込みは確定前も毎日行う。
10月以前（事務担当の手作業の月）は従来どおり毎日作る。
Excel で誰かが開いていて（~$ ロック）書けないときは 5 分おきに最大 12 回待つ。
書く前に同じフォルダの `_上書き履歴\\` へ写す（直近 14 本を残す）。

20260923: 1か月の失敗で残りを止めない。月ごとに例外を受け止めてログに残し、次の月へ進む。
マスターに当直がまだ無い月（10/1 からの来月＝11月など）は異常ではなく「提出がまだ無い → 書かない」。
書きかけの PDF の掃除と 当直交代申請.xlsx の更新は、途中の月が失敗しても必ず走らせ、最後の行に失敗の有無を書く（古い月の PDF は 20261001 から消さない）。
PDF の作業用の写し（人別シートにログインIDと職員コードがある）は公開用フォルダではなくローカルの一時フォルダに作る。

20260923 第2陣（安全策）:
  - 正本を途中で書き換えない。dtx14_fill.run_master がローカルの写しに全部書いて読み戻しの突合 ✅ のときだけ差し替える。
    書いている間に正本が Excel で開かれた／更新された（mtime・大きさが変わった）ら差し替えない（guard）。
  - 失敗（月の書き込み・PDF・交代記録・本番の共有やログに届かない）と OneDrive の競合コピーは、最後に1通だけ
    user003@example.ac.jp へ警報メール（Outlook COM）。
    （当初は「件名に ※自動送信 を付けない＝振り分けで埋もれないように」としていたが、管理者が 20261012・20261020 に
    撤回した＝受信トレイのほうが埋もれる。いまは管理者だけ宛ての自動メールにはすべて付ける。下の 20261020 を参照）
    送信は ALERT_SENDER を差し替えればテストで止められる。Outlook が使えなくてもログに残して落ちない。
  - ログは本番の 管理ツール\\log\\ に加えてローカル（%LOCALAPPDATA%\\Touchoku\\log\\）にも書く（共有に届かない日も残る）。
  - 停止スイッチ: 本番の 管理ツール\\ か 月の SharePoint フォルダに STOP_DTX14_SYNC.txt があれば、その月は DTX14 に書かない
    （PDF は作る）。悪い日のあとに _上書き履歴\\ から戻している間に上書きされないように。
  - 書いたあと、月フォルダに正本の競合コピー（〜-PC名.xlsm・〜 - コピー・〜 (1)。1日以内）があれば警告してメールする。

20260924（管理者: 決裁後の変更は交代申請。承認したものだけを反映）:
  - 交代記録（kotai_record.update_all＝当直交代申請.xlsx の承認欄を読む・却下を元に戻す・承認済みの状態を作る）を
    DTX14 に書く **前に** 回す（今日の承認を今日のうちに反映する）。以前は最後に回していた。
  - 決裁済み（③確定）の月は、共有マスターそのものではなく kotai_record.approved_master（元の割当＋承認した交代）で書く。
    その月の交代記録の更新が失敗した／走らなかったときは、その月の DTX14 には書かない（失敗として警報。生のマスターでは書かない）。
  - 当直交代申請.xlsx が読めない月（承認欄を読めず作り直してもいない）は警告としてメールに載せる。

20260925（本番前の見直し）:
  - 確定しているかは 集約\\確定.json（管理者だけが書ける確定の正本）で決める（master_store.effective_confirmed）。
    マスターの confirmed と食い違えば警報を出して正本に従う（マスターは全科が書けるので、書き換えて生のマスターを
    DTX14 に書かせることができた）。確定.json が壊れていれば、その月は書かない（失敗）。
  - 決裁済みの月は、DTX14 に書けたあとで kotai_record.publish_approved を呼ぶ（台帳の approved と reflected_at を付ける
    ＝手当計算 の Dr.JOY 投入が読む）。書けなかった月・停止スイッチ・SYNC_FROM 前の月は付けない。
  - SYNC_FROM 以降で書く月に正本の DTX14 が無い／2本以上ある → 失敗として警報（以前は黙って飛ばしていた）。
  - 事務（事務担当・事務担当2）にも、自分で直せることだけをやさしい言葉で1通（CLERK_TO。管理者への警報とは別）:
    DTX14 が開かれていて書けなかった／正本が2つある／競合コピー（DTX14・当直交代申請.xlsx）／雛形の科名が変わって
    行が見つからない／同じ日の先の申請が決まっていないので保留。件名に ※自動送信 は付けない。
  - check_last_run: ローカルのログの最後の「=== 終了」が古すぎないか（kotai_record のすぐ載せる回が見張る）。
  - 次の次の月の置き場づくりの SystemExit も受け止める（1〜5日に最後の警報・ログが消えていた）。

20260926（レビュー）:
  - 置き場づくり（new_month）は1〜5日だけでなく毎日: 次の次の月と、欠けていれば来月（SYNC_FROM 以降だけ）。無いものだけ作る。
    作ったら【お知らせ】。作れなければ失敗として警報。6日を過ぎても欠けていれば（来月はいつでも）警告として毎日載せる。
    ACL の失敗・当番当直が無いなど new_month の「知らせること」も捨てずに警告へ（以前は _notes を捨てていた）。
  - 当番当直が data/toban.csv に無い月（前月〜来月・次の次の月）は毎日の警告に載せる（toban_warnings）。
  - 前月〜来月の外の月でも、承認したのに DTX14 に反映していない交代（台帳の approval=承認 で reflected_at が無い、
    または approved_pending が approved と違う）があれば、その日の書き込みに加える（late_months。SYNC_FROM 以降・正本が1本の月だけ。
    PDF は作らない＝公開は前月〜来月だけ）。反映できない月は警告。
  - SYNC_FROM 以降の書く月に共有マスターが無い → 失敗（以前は「書かない」で成功扱い）。
  - 競合コピーの新しさを 24 → 72 時間に。書く前にも見る（書き込みが失敗した日も知らせる）。
  - PDF はローカルの作業フォルダに書き、公開用フォルダへは一時名で写してから差し替える。残った一時ファイル（~*_pdf_*.tmp・以前の *.tmp.pdf）は cleanup_public が消す。
  - OneDrive が動いていなければ警告（SharePoint の正本・公開 PDF が更新されない）。check_roots が _STATE["root_warn"] に積む。
  - 毎日の見回り（置き場・当番当直・後からの承認）は --month を付けない回（タスクの本番の回）だけ。

20260927（管理者: アプリの名簿 = Dr.JOY の所属医師の申告）:
  - 所属の変更メール（shozoku_notice。事務担当へ・cc 管理者・甲野・己野・事務担当2）は、15時の回では送らない。
    10分おきの回（kotai_record --quick）が月末の日と毎週水曜の朝9時に、来月の分だけ送る（morning_shozoku_notice。20260927）。
  - outlook_send に cc（省略可）を足した。

20261020（管理者: 管理者だけ宛ての自動メールの書式）:
  - outlook_send が宛先（To と CC）を見て、全員が管理者（user003 / admin）のときだけ
    件名の末尾に「 ※自動送信」（付いていなければ）・本文の最後の行に「問い合わせ先: 当直勤務表の担当」を付け、DeleteAfterSubmit で送信済みアイテムに残さない。
    対象 = 【警報】15時の反映・【警報】所属の変更メールを送れませんでした・【お知らせ】置き場を作りました・
    kotai_record の見張りの【警報】。事務へのお知らせ（CLERK_TO）と所属の変更メール（事務担当宛て・cc つき）は変えない。
    1日1通の判定は送信済みアイテムを見ず、状態ファイル（shozoku_notice_day.json・watchdog.json）で行っている。

usage: python -X utf8 tools\\dtx14_daily_sync.py [--dry-run] [--no-fill] [--month 2026-11 ...]
起動: tools\\dtx14_daily_sync.bat（タスク スケジューラ「当直勤務表_DTX14反映」毎日 15:00・管理者の PC）
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
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402

SYNC_ROOT = r"C:\tools\touchoku\sharepoint"
DTX_ROOT = os.path.join(SYNC_ROOT, "当直勤務表_正本")
PUBLIC_DIR = os.path.join(SYNC_ROOT, "公開用")
PROD_ROOT = r"\\fileserver\share$\00_DrJOY\当直勤務表"
SYNC_FROM = (2026, 11)                     # この月から DTX14 へ書く（10月は予行）
HISTORY = "_上書き履歴"
KEEP_HISTORY = 14
LOCK_WAIT = (12, 300)                      # 回数, 秒
SHEETS = ("当直表 ファースト", "当直表 セカンド以下")
# 公開用 PDF の用紙。20261025 管理者: 掲示板でページの大きさが違って見えるので、ファーストも A3 にそろえる（以前は A4 / A3）。
#   正本そのものの印刷設定（dtx14_fill.PAPER）は変えない
PAPER = {"当直表 ファースト": 8, "当直表 セカンド以下": 8}   # 両方 A3
ADMIN_DIR = C.PROD_ADMIN                   # 本番の 管理ツール\（停止スイッチを置く所）
STOP_NAME = "STOP_DTX14_SYNC.txt"          # これがあれば DTX14 に書かない（20260923 第2陣）
SHARE_LOG = os.path.join(ADMIN_DIR, "log", "dtx14_daily_sync.log")
LOCAL_LOG = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "log",
                         "dtx14_daily_sync.log")
ALERT_TO = "user003@example.ac.jp"
CLERK_TO = ["user004@example.ac.jp", "user005@example.ac.jp"]   # 事務担当・事務担当2（自分で直せることだけ。管理者 20260925）
CONFLICT_WITHIN = 72 * 3600                # 競合コピーとみなす新しさ（秒。20260926: 24 → 72 時間＝週末をまたいでも気づく）
# 正本の名前（空白を除く）のあとに付く競合コピーの印: -PC名（OneDrive・Office。-PC名-2 も）／-コピー（エクスプローラー）／(1)
CONFLICT_RE = re.compile(r"(?:-(?:コピー|copy|Copy)|-[A-Za-z0-9][A-Za-z0-9_-]*|\(\d+\)|（\d+）)(?:\(\d+\)|（\d+）)?")

_STATE = {"share_err": None, "run": [],  # 共有のログに書けなかった理由／この回のログ（メールの本文に使う）
          "clerk": [],                     # 事務へのお知らせ [((年, 月), やさしい文)]（20260925）
          "dtx_hits": {},                  # dtx_file が見つけた正本の名前の候補 {(年, 月): [パス]}
          "root_warn": []}                 # check_roots が見つけた警告（OneDrive が動いていない。20260926）


def clerk_note(y, m, text):
    """事務（事務担当・事務担当2）へのお知らせを1つ積む（同じ文は1つに）。"""
    item = ((y, m), text)
    if item not in _STATE["clerk"]:
        _STATE["clerk"].append(item)


def _append(path, line):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def log(msg):
    line = f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    _STATE["run"].append(line)
    try:                                   # ローカルには必ず残す（共有に届かない日も失敗が記録される）
        _append(LOCAL_LOG, line)
    except OSError:
        pass
    if _STATE["share_err"] is None:        # 共有は一度失敗したらこの回はもう試さない（UNC の待ちを繰り返さない）
        try:
            _append(SHARE_LOG, line)
        except OSError as e:
            _STATE["share_err"] = f"{type(e).__name__}: {e}"


def months_window(today=None):
    t = today or dt.date.today()
    out = []
    for k in (-1, 0, 1):
        y, m = t.year, t.month + k
        if m < 1:
            y, m = y - 1, 12
        elif m > 12:
            y, m = y + 1, 1
        out.append((y, m))
    return out


def reiwa(y):
    return y - 2018


def is_master_name(name, y, m):
    """正本の名前の規則（管理者 20260922）: 『★R8.N月用 DTX14』で始まり、DTX14 のあとは半角数字と . だけで .xlsm。
    ★の前に何か付いたもの（〜版・（事務担当2）など）や、後ろに文字が付いたものは正本とみなさない。
    半角・全角の空白は入っていても正本（同日の追加指示。10月の現物は『DTX14.9.5.12 .xlsm』だった）。"""
    core = name.replace(" ", "").replace("　", "")
    return re.fullmatch(rf"★R{reiwa(y)}\.{m}月用DTX14[0-9.]*\.xlsm", core) is not None


def dtx_candidates(y, m):
    """月フォルダと、正本の名前の規則に合う DTX14 の一覧（0本・2本以上もそのまま）。"""
    d = os.path.join(DTX_ROOT, f"R{reiwa(y)}.{m}月 当直勤務表")
    if not os.path.isdir(d):
        return d, []
    names = [n for n in os.listdir(d) if n.lower().endswith(".xlsm") and not n.startswith("~$")]
    hits = [os.path.join(d, n) for n in sorted(names) if is_master_name(n, y, m)]
    others = [n for n in names if "DTX14" in n and n not in {os.path.basename(h) for h in hits}
              and not n.startswith("【")]          # 手入力の部署に配る写し（【診療部名】など。new_month が置く）は毎日並べない（20261028）
    if others:
        log(f"  正本の名前の規則に合わない DTX14（無視）: {others}")
    return d, hits


def dtx_file(y, m):
    """正本の DTX14（規則に合うものがちょうど1本のときだけ）。0本・2本以上は None（sync_month が失敗として知らせる）。"""
    d, hits = dtx_candidates(y, m)
    _STATE["dtx_hits"][(y, m)] = hits
    if len(hits) > 1:
        log(f"  ✗ 正本の規則に合うファイルが複数ある → 今日は触らない: {[os.path.basename(h) for h in hits]}")
        return d, None
    return (d, hits[0]) if hits else (d, None)


def master_file(y, m):
    return os.path.join(PROD_ROOT, f"R{reiwa(y)}年{m}月", "master", f"当直予定表_R{reiwa(y)}年{m}月.json")


def hydrate(path):
    """OneDrive の「ファイル オンデマンド」で中身がクラウドにしか無いファイルは Excel が開けない
    （『アクセスできません』）。読み切ればローカルに落ちてくる。"""
    for i in range(3):
        try:
            with open(path, "rb") as f:
                while f.read(1 << 20):
                    pass
            return True
        except OSError as e:
            log(f"  ダウンロード待ち ({i + 1}/3): {e}")
            time.sleep(10)
    return False


def locked(path):
    return os.path.exists(os.path.join(os.path.dirname(path), "~$" + os.path.basename(path)))


def backup(path):
    hd = os.path.join(os.path.dirname(path), HISTORY)
    os.makedirs(hd, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    bak = os.path.join(hd, f"{stem}_{dt.datetime.now():%Y%m%d_%H%M}{ext}")
    shutil.copy2(path, bak)
    olds = sorted(glob.glob(os.path.join(hd, f"{stem}_*{ext}")))
    for p in olds[:-KEEP_HISTORY]:
        try:
            os.remove(p)
        except OSError:
            pass
    return bak


def stop_file(month_dir):
    """停止スイッチ（STOP_DTX14_SYNC.txt）の場所。本番の 管理ツール\\ か 月の SharePoint フォルダ。無ければ None。"""
    for d in (ADMIN_DIR, month_dir):
        p = os.path.join(d, STOP_NAME)
        try:
            if os.path.isfile(p):
                return p
        except OSError:
            pass
    return None


def _nospace(t):
    return t.replace(" ", "").replace("　", "")


def conflict_copies(month_dir, master_path, now=None, within=CONFLICT_WITHIN):
    """正本の競合コピー（OneDrive・Office が作る『名前-PC名.xlsm』、エクスプローラーの『名前 - コピー』、『名前 (1)』）で、
    within 秒以内に作られた／更新されたものの名前。〜_0902〆版・〜（事務担当2） のような人が付けた名前は拾わない。"""
    now = now or time.time()
    stem = _nospace(os.path.splitext(os.path.basename(master_path))[0])
    out = []
    try:
        names = os.listdir(month_dir)
    except OSError:
        return out
    for n in sorted(names):
        p = os.path.join(month_dir, n)
        base, ext = os.path.splitext(n)
        if n.startswith("~$") or ext.lower() not in (".xlsm", ".xlsx") or not os.path.isfile(p):
            continue
        if os.path.normcase(os.path.abspath(p)) == os.path.normcase(os.path.abspath(master_path)):
            continue
        s = _nospace(base)
        if not s.startswith(stem) or not CONFLICT_RE.fullmatch(s[len(stem):]):
            continue
        try:
            if now - os.path.getmtime(p) > within:
                continue
        except OSError:
            continue
        out.append(n)
    return out


def note_conflicts(y, m, d, path, warn):
    """正本の競合コピーがあれば警告と事務へのお知らせに足す（同じ文は1つ）。書く前と書いたあとに呼ぶ（20260926）。"""
    cc = conflict_copies(d, path)
    if not cc:
        return cc
    item = f"R{reiwa(y)}年{m}月 競合コピー: {'、'.join(cc)}（{d}）"
    if item not in warn:
        log(f"  ⚠ 正本の競合コピーがある（OneDrive）: {cc} → 中身を正本と見比べてください")
        warn.append(item)
    clerk_note(y, m, f"当直勤務表（DTX14）の写し（{'、'.join(cc)}）ができています。"
                     "2か所で同時に書いたときにできるものです。写しの方には書き込まず、管理者に知らせてください。")
    return cc


class KotaiBlocked(Exception):
    """決裁済みの月だが、承認済みの状態（交代記録）が使えない → その月の DTX14 は書かない。"""


def is_confirmed(y, m, master_path, mst, warn=None):
    """確定しているか（確定の正本 集約\\確定.json で決める。20260925）。戻り: confirmed の dict か None。
    正本とマスターが食い違えば warn に足し（警報）、正本に従う。正本が壊れていれば KotaiBlocked。"""
    import master_store as MST
    try:
        conf, probs = MST.effective_confirmed(master_path, mst)
    except MST.TruthError as e:
        raise KotaiBlocked(f"確定の正本が読めない（確定しているか分からない）: {e}") from e
    for level, text in probs:
        log(("  ⚠ " if level == "warn" else "  ✗ ") + text)
        if level == "alert" and warn is not None:
            item = f"R{reiwa(y)}年{m}月 {text}"
            if item not in warn:
                warn.append(item)
    return conf


def fill_source(y, m, master, kotai, dry, warn=None):
    """DTX14 に書く元。未確定の月 = 共有マスターのパス。決裁済みの月 = 元の割当＋承認した交代（kotai_record.approved_master）。
    決裁済みの月は、この回の交代記録の更新が通っていなければ KotaiBlocked（生のマスターでは書かない。管理者 20260924）。
    kotai = run_kotai の戻り {(年, 月): 状態}（走らなかったら None）。
    確定しているかは確定の正本で決める（is_confirmed。20260925）。"""
    import kotai_record as K
    import master_store as MST
    mst = MST.load(master, y, m)
    conf = is_confirmed(y, m, master, mst, warn)
    if not conf:
        return master
    r = (kotai or {}).get((y, m))
    if not dry and (r is None or r in K.FAILED):
        raise KotaiBlocked(f"決裁済みの月なのに交代記録の更新が通っていない（{r or '走っていない'}）")
    mst["confirmed"] = conf                  # 読んだ写しの確定の印を正本に揃える（approved_master はこれを見る）
    try:
        am = K.approved_master(os.path.dirname(os.path.dirname(master)), y, m, mst)
    except K.LedgerError as e:
        raise KotaiBlocked(f"承認済みの状態が読めない: {e}") from e
    log("  決裁済み → 元の割当＋承認した交代で書く（承認待ち・却下の申請は書かない）")
    drift = K.order_drift(mst, am)           # マスの中の氏名の並びが各科の出した順から変わっていないか（20261025）
    if drift:
        msg = (f"R{reiwa(y)}年{m}月 決裁済みの割当で、マスの中の氏名の並びが各科の出した順と違う {len(drift)} マス"
               f"（例: {'、'.join(f'{o} {d} {b}' for o, d, b, _f in drift[:3])}）→ そのまま書くが、並びの作り直しを確かめること")
        log("  ⚠ " + msg)
        if warn is not None and msg not in warn:
            warn.append(msg)
    return am


CLERK_OPEN = ("当直勤務表（DTX14）が開かれていたため、今日の分を書き込めませんでした。ファイルを閉じてください。"
              "明日の15時に今日の分も含めて書き込みます。")


def fill_month(y, m, path, master, dry):
    """master = 共有マスターのパス、または承認済みの状態（dict。fill_source の戻り）。"""
    import dtx14_fill as F
    for i in range(LOCK_WAIT[0]):
        if not locked(path):
            break
        log(f"  Excel で開かれています（~$）。{LOCK_WAIT[1]}秒待つ ({i + 1}/{LOCK_WAIT[0]})")
        time.sleep(LOCK_WAIT[1])
    if locked(path):
        log(f"  ✗ ロックが外れないので今日は書かない: {path}")
        clerk_note(y, m, CLERK_OPEN)
        return False
    if dry:
        log(f"  (dry-run) fill: {master if isinstance(master, str) else '承認済みの状態'} → {path}")
        return True
    hydrate(path)
    bak = backup(path)
    log(f"  退避: {bak}")
    st0 = os.stat(path)
    # 雛形の写しはローカルに（20260923 第2陣: 以前は SharePoint の _上書き履歴\ に置いていた）
    work_dir = tempfile.mkdtemp(prefix="dtx14_tpl_")
    tmp = os.path.join(work_dir, "_template.xlsm")
    shutil.copy2(path, tmp)

    def guard():                            # 差し替えの直前: 書いている間に開かれた／更新されたら差し替えない
        if locked(path):
            raise F.FileBusy("書いている間に Excel で開かれた（~$）→ 差し替えない（次の15時にやり直す）")
        st = os.stat(path)
        if (st.st_mtime, st.st_size) != (st0.st_mtime, st0.st_size):
            raise F.FillError("書いている間に正本が更新された → 差し替えない（次の15時にやり直す）")

    try:
        # 突合 ❌・途中の失敗は FillError（正本は元のまま）。ここへ戻れば差し替え済み・突合 ✅
        r = F.run_master(master, tmp, path, y, m, staffs=C.newest_staffs(), verify=True,
                         log=lambda s: log("    " + s), guard=guard)
        log(f"  書き込み: {r['cells']} セル / {r['records']} 件 / 突合 ✅ → 正本を差し替えた")
        return True
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


def pdf_path(y, m):
    return os.path.join(PUBLIC_DIR, f"R{reiwa(y)}年{m}月 病院当直勤務表.pdf")   # 名前は管理者指定（日付は付けない＝毎日上書き）


def confirmed_state(y, m):
    """その月の共有マスターが確定済み（管理ツール③）か。True/False。
    本番の共有に届かないときは None（分からない＝今ある PDF に触らない）。マスターが壊れて読めなければ例外。"""
    if not os.path.isdir(PROD_ROOT):
        return None
    import master_store as MST
    mp = master_file(y, m)
    return bool(is_confirmed(y, m, mp, MST.load(mp, y, m)))   # 確定の正本で決める（20260925）


def remove_pdf(y, m, dry):
    """未確定の月の公開用 PDF を消す（前に置いたものが残っていれば）。"""
    p = pdf_path(y, m)
    if not dry:
        drop_pending(p)      # とっておいた確定中の PDF を、10分おきの回が置き直さないように（20261009 レビュー）
    if os.path.exists(p):
        log(f"  未確定の月の PDF を消す: {os.path.basename(p)}" + ("（dry-run）" if dry else ""))
        if not dry:
            try:
                os.remove(p)
            except PermissionError:
                if not release_pdf(p):     # 開かれていれば、そのタブだけ閉じてから消す（20261009）
                    raise
                os.remove(p)


def export_pdf(y, m, path, dry):
    os.makedirs(PUBLIC_DIR, exist_ok=True)
    out = pdf_path(y, m)
    if dry:
        log(f"  (dry-run) pdf: {path} → {out}")
        return out
    hydrate(path)
    # 現物は誰かが開いていたり OneDrive が掴んでいたりして Excel が開けないことがある（9月で実測）→ 写しから作る
    # ⚠ 写しは公開用フォルダに置かない（全科に閲覧で共有され OneDrive が上げる。人別シートにログインIDと職員コードがある）
    work_dir = tempfile.mkdtemp(prefix="dtx14_pdf_")
    work = os.path.join(work_dir, "_work.xlsm")
    xl = None
    try:
        shutil.copy2(path, work)
        import win32com.client
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.AutomationSecurity = 3
        wb = xl.Workbooks.Open(work, UpdateLinks=0, ReadOnly=True)
        import dtx14_fill as F
        for sn in SHEETS:                          # 印刷設定はメモリ上だけ（写しなので保存しない）
            ws = wb.Worksheets(sn)
            ps = ws.PageSetup
            ps.PaperSize = PAPER[sn]
            ps.Orientation = 2
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            F.fit_days(ws, calendar.monthrange(y, m)[1], log=lambda s: None)   # 月にない日の列を隠す（20261014）
            F.fit_note(ws, log=lambda s: None)     # 下の注記を印刷範囲に収める（事務担当の手作業の月も PDF では収まる）
            F.fit_legend(ws, log=lambda s: None)   # 凡例を決まった文字にし、右下に『(R8.10.25 15時更新)』（20261025。書き込まない月の写しも）
            F.stamp_updated(ws)                    # 右上の古い小さな更新の印を消す（20261025。以前は AF1:AG1 に書いていた）
        wb.Worksheets(SHEETS).Select()
        local_pdf = os.path.join(work_dir, "_out.pdf")   # PDF もまずローカルに（20260926: 以前は公開用フォルダに一時ファイルを作った）
        xl.ActiveSheet.ExportAsFixedFormat(0, local_pdf)   # 選択した2シートを1本の PDF に
        geos = None
        try:                                               # 氏名に Teams のチャットへのリンクを重ねるためのマスの位置（20261007）
            import pdf_links as PL
            geos = PL.collect_geometry(wb, SHEETS, y, m)
        except Exception as e:
            log(f"  ⚠ リンク: マスの位置を読めない（{type(e).__name__}: {e}）→ リンクなしで置く")
        wb.Close(SaveChanges=False)
        xl.Quit()
        xl = None
        if geos:
            local_pdf = link_names(local_pdf, work, geos, y, m)
        try:
            try:
                publish_file(local_pdf, out)
            except PermissionError:
                # 20261009 管理者: 開いたままの公開 PDF は、15時の回がそのタブだけを閉じてよい（PDF を編集することはない）
                if not release_pdf(out):
                    raise
                publish_file(local_pdf, out)
            drop_pending(out)
        except PermissionError:
            # 誰かが（この PC で）開いていて差し替えられない（20261007 実測: PDF-XChange で開いていた）。
            # 手元にとっておき、10分おきの回（kotai_record --quick → flush_pending_pdfs）が閉じられたあとに置く。
            # 前の日の分がまだ置けていなければ失敗にする（警報）。管理者 20261007
            since = pending_since(out)
            stash_pdf(local_pdf, out)
            log(f"  PDF: 開かれていて差し替えられない → 手元にとっておき、閉じられたら置く（10分おきに確かめる）: {out}")
            if since and (dt.datetime.now() - since).total_seconds() > PENDING_ALERT_HOURS * 3600:
                raise RuntimeError(f"{since:%m/%d %H:%M} から開かれたままで、新しい PDF を置けていない")
            return out
    finally:
        if xl is not None:
            xl.Quit()
        shutil.rmtree(work_dir, ignore_errors=True)
    log(f"  PDF: {out}")
    return out


CLOSE_TAB_PS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "close_pdf_tab.ps1")


def release_pdf(out):
    """out を開いているビューアーの、そのタブだけを閉じる（tools/close_pdf_tab.ps1。UI Automation で「タブを閉じる」を押す）。
    20261009 管理者「15時前に強制的に閉じて」。PDF-XChange Editor はほかに60ほどのタブを開いているので、アプリごとは閉じない。
    戻り: 手放されたか。失敗しても落ちない（とっておいて閉じられたら置く、の道に戻る）。"""
    import json as _json
    import subprocess
    pwsh = shutil.which("pwsh") or r"C:\Program Files\PowerShell\7\pwsh.exe"
    try:
        r = subprocess.run([pwsh, "-NoProfile", "-NonInteractive", "-File", CLOSE_TAB_PS, "-Path", out],
                           capture_output=True, timeout=90, stdin=subprocess.DEVNULL,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        line = r.stdout.decode("utf-8", "replace").strip().splitlines()[-1]
        d = _json.loads(line)
    except Exception as e:
        log(f"  PDF: 開いているタブを閉じられない（{type(e).__name__}: {e}）")
        return False
    who = "、".join(d.get("held_by") or []) or "不明"
    if d.get("released"):
        log(f"  PDF: 開かれていたタブを閉じた（{who}・{d.get('closed')} 個）: {os.path.basename(out)}")
        return True
    log(f"  PDF: 開かれていたタブを閉じられない（{who}・{d.get('note')}）: {os.path.basename(out)}")
    return False


PENDING_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "pending_pdf")
PENDING_ALERT_HOURS = 20


def _pending_paths(out):
    """とっておく PDF と、その置き先の控え（拡張子は1つずつ。二重拡張子は EDR が嫌う）。"""
    name = os.path.basename(out)
    stem = os.path.splitext(name)[0]
    return os.path.join(PENDING_DIR, name), os.path.join(PENDING_DIR, stem + "_pending.json")


def stash_pdf(src, out):
    os.makedirs(PENDING_DIR, exist_ok=True)
    pdf, meta = _pending_paths(out)
    since = pending_since(out) or dt.datetime.now()
    shutil.copyfile(src, pdf)
    with open(meta, "w", encoding="utf-8") as f:
        json.dump({"out": out, "since": since.isoformat(timespec="seconds"),
                   "made": dt.datetime.now().isoformat()}, f, ensure_ascii=False)   # full precision: compared with the public PDF's mtime


def pending_since(out):
    """out 向けにとっておいた PDF が、いつから置けていないか（無ければ None）。"""
    _pdf, meta = _pending_paths(out)
    try:
        with open(meta, encoding="utf-8") as f:
            return dt.datetime.fromisoformat(json.load(f)["since"])
    except (OSError, ValueError, KeyError):
        return None


def drop_pending(out):
    for p in _pending_paths(out):
        try:
            os.remove(p)
        except OSError:
            pass


def flush_pending_pdfs(log=None):
    """とっておいた PDF を置き直す（10分おきの回から呼ぶ）。まだ開かれていれば黙って次の回へ。
    公開用フォルダの方が新しければ（15時の回が置き直した）捨てる。戻り: 置いた数。"""
    say = log or globals()["log"]
    placed = 0
    for meta in glob.glob(os.path.join(PENDING_DIR, "*_pending.json")):
        try:
            with open(meta, encoding="utf-8") as f:
                d = json.load(f)
            out = d["out"]
            made = dt.datetime.fromisoformat(d["made"])
        except (OSError, ValueError, KeyError):
            continue
        pdf, _meta = _pending_paths(out)
        if not os.path.isfile(pdf):
            drop_pending(out)
            continue
        if os.path.isfile(out) and dt.datetime.fromtimestamp(os.path.getmtime(out)) > made:
            drop_pending(out)
            continue
        try:
            publish_file(pdf, out)
        except PermissionError:
            continue
        except OSError as e:
            say(f"✗ とっておいた PDF を置けない: {os.path.basename(out)}: {type(e).__name__}: {e}")
            continue
        drop_pending(out)
        placed += 1
        say(f"とっておいた PDF を置いた（閉じられたので）: {os.path.basename(out)}（{made:%m/%d %H:%M} に作ったもの）")
    return placed


def link_names(local_pdf, work, geos, y, m):
    """PDF の氏名に Teams のチャットへのリンクを重ねた写しを返す（管理者 20261007。書式は変えない＝見えない注釈だけ）。
    うまくいかなければ元の PDF を返す（リンクが無いだけで公開は止めない）。"""
    try:
        import pdf_links as PL
        logins = PL.collect_logins(work, C.newest_staffs(), y, m)
        linked = os.path.join(os.path.dirname(local_pdf), "_linked.pdf")
        n = PL.link_pdf(local_pdf, linked, SHEETS, geos, logins, log=log)
        return linked if n else local_pdf
    except Exception as e:
        log(f"  ⚠ リンク: 付けられない（{type(e).__name__}: {e}）→ リンクなしで置く")
        return local_pdf


TMP_GLOBS = ("~*_pdf_*.tmp", "*.tmp.pdf")  # 公開用フォルダへ写す途中の名前（今の名前・20261002 より前の名前）。残っていれば cleanup_public が消す


def publish_file(src, out):
    """ローカルの src を公開用フォルダの out へ: 一時名（~<名前>_pdf_<pid>.tmp）で写してから差し替える（書きかけを見せない）。
    20261002: 以前の一時名 out + .tmp.pdf は二重拡張子で、EDR がランサムウェア疑いで通報する形だったのでやめた。"""
    import master_store as MST
    tmp = MST.tmp_name(out)
    try:
        shutil.copyfile(src, tmp)
        os.replace(tmp, out)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def cleanup_public(dry):
    """途中で落ちた回の一時ファイル（TMP_GLOBS）だけを消す。
    20261001 管理者: 古い月の PDF は自動では消さない（人が消したいときに消す）。以前は前月〜来月の外を消していた。"""
    if not os.path.isdir(PUBLIC_DIR):
        return
    for pat in TMP_GLOBS:                                              # 途中で落ちた回の一時ファイル（20260926）
        for p in glob.glob(os.path.join(PUBLIC_DIR, pat)):
            log(f"  書きかけの PDF を消す: {os.path.basename(p)}" + ("（dry-run）" if dry else ""))
            if not dry:
                os.remove(p)


def sync_month(y, m, a, warn=None, kotai=None, pdf=True):
    """1か月分（DTX14 へ書く → PDF）。戻り: 失敗したことの一覧（空なら成功・飛ばしたのも成功）。
    warn = 失敗ではないがメールで知らせること（競合コピー）を足すリスト。
    kotai = この回の交代記録の更新の結果（run_kotai）。決裁済みの月はこれが通っていないと書かない（20260924）。
    pdf=False = PDF に触らない（前月〜来月の外で、後からの承認を反映するだけの月。20260926）。"""
    import dtx14_fill as F
    warn = [] if warn is None else warn
    bad = []
    d, path = dtx_file(y, m)
    fill = not a.no_fill and (y, m) >= SYNC_FROM
    if not path:
        hits = _STATE["dtx_hits"].get((y, m)) or []
        if fill and not stop_file(d):        # 書く月に正本が無い／2本以上 → 失敗（20260925: 以前は黙って飛ばしていた）
            if len(hits) > 1:
                names = [os.path.basename(h) for h in hits]
                log(f"R{reiwa(y)}年{m}月: ✗ 正本の名前の DTX14 が {len(names)} 本ある（{d}）→ 書かない: {names}")
                bad.append(f"書き込み（正本の名前の DTX14 が {len(names)} 本: {'、'.join(names)}）")
                clerk_note(y, m, f"R{reiwa(y)}年{m}月の当直勤務表（DTX14）が {len(names)} つあります"
                                 f"（{'、'.join('「' + n + '」' for n in names)}）。正しい方を1つだけ残してください。"
                                 "ほかは名前を変えるか、別の場所へ移してください。今日の15時の書き込みはしていません。")
            else:
                log(f"R{reiwa(y)}年{m}月: ✗ 正本の DTX14 が無い（{d}）→ 書けない")
                bad.append("書き込み（正本の DTX14 が無い）")
            return bad
        log(f"R{reiwa(y)}年{m}月: DTX14 が無い（{d}）→ 飛ばす")
        return bad
    log(f"R{reiwa(y)}年{m}月: {os.path.basename(path)}")
    master = master_file(y, m)
    if fill and not a.dry_run:
        note_conflicts(y, m, d, path, warn)      # 書く前にも見る（書き込みが失敗した日も知らせる。20260926）
    stop = stop_file(d) if fill else None
    if stop:
        log(f"  ⏸ 停止スイッチがある（{stop}）→ この月は DTX14 に書かない（PDF は作る）")
    elif fill:
        if os.path.isfile(master):
            wrote = False
            src = None
            try:
                src = fill_source(y, m, master, kotai, a.dry_run, warn)
                wrote = fill_month(y, m, path, src, a.dry_run)
                if not wrote:
                    bad.append("書き込み")
            except KotaiBlocked as e:
                log(f"  ✗ {e} → この月の DTX14 は書かない（決裁後は承認済みの状態でだけ書く）")
                bad.append("書き込み（交代記録が使えず止めた）")
            except F.NoDuties:
                log("  提出がまだ無い → 書かない")
            except F.HeadingMissing as e:
                log(f"  ✗ 書き込みに失敗（正本は元のまま）: {e}")
                bad.append("書き込み（雛形の科名で行が見つからない）")
                clerk_note(y, m, f"雛形の科名（A列）が変わっていて、{'、'.join(e.depts)}の行が見つかりませんでした。"
                                 "元の科名に戻してください。")
            except F.FileBusy as e:
                log(f"  ✗ 書き込みに失敗（正本は元のまま）: {e}")
                bad.append("書き込み（DTX14 が開かれていた）")
                clerk_note(y, m, CLERK_OPEN)
            except (Exception, SystemExit) as e:
                log(f"  ✗ 書き込みに失敗（正本は元のまま）: {type(e).__name__}: {e}")
                bad.append("書き込み")
            snap = getattr(src, "kotai_snapshot", None)
            if wrote and snap and not a.dry_run:
                # 書けた月だけ: 台帳の approved（Dr.JOY 投入が読む）と反映済みの印（20260925）
                import kotai_record as K
                al = []
                try:
                    # out_dir は DTX14 と同じ月フォルダ（d）。本番では out_dir_of(y, m) と同じ所
                    K.publish_approved(os.path.dirname(os.path.dirname(master)), y, m, snap, out_dir=d,
                                       log=lambda s: log("  " + s), alerts=al)
                except Exception as e:
                    log(f"  ✗ 反映済みの印（交代記録の approved）を付けられない: {type(e).__name__}: {e}")
                    bad.append("反映済みの印（交代記録）")
                route_alerts(al, bad, warn)
            if wrote and not a.dry_run:
                note_conflicts(y, m, d, path, warn)
        else:
            # 20260926: 書く月（SYNC_FROM 以降）に共有マスターが無いのは失敗（以前は「書かない」で成功扱い）
            log(f"  ✗ アプリの共有マスターが無い（{master}）→ 書けない")
            bad.append("書き込み（アプリの共有マスターが無い）")
    else:
        log(f"  書き込み対象外（{'--no-fill' if a.no_fill else 'SYNC_FROM 前'}）")
    if not pdf:
        log("  PDF は作らない（前月〜来月の外の月）")
        return bad
    # 公開用の PDF は、SYNC_FROM 以降の月は確定（管理ツール③）してから（管理者 20260923）。DTX14 への書き込みは確定前も毎日
    conf = True
    if (y, m) >= SYNC_FROM:
        try:
            conf = confirmed_state(y, m)
        except Exception as e:
            log(f"  ✗ 確定しているか読めない（マスター {master_file(y, m)}）: {type(e).__name__}: {e} → PDF は触らない")
            bad.append("PDF（確定の確認）")
            return bad
    if conf is None:
        log(f"  本番の共有に届かず確定しているか分からない → PDF は触らない（{pdf_path(y, m)}）")
    elif not conf:
        log("  未確定（管理ツール③「この月を確定」の前）→ PDF は公開しない")
        try:
            remove_pdf(y, m, a.dry_run)
        except OSError as e:
            log(f"  ✗ 未確定の月の PDF を消せない: {type(e).__name__}: {e}")
            bad.append("PDF の削除")
    else:
        try:
            export_pdf(y, m, path, a.dry_run)
        except (Exception, SystemExit) as e:
            log(f"  ✗ PDF に失敗: {type(e).__name__}: {e}")
            bad.append("PDF")
    return bad


def route_alerts(alerts, failed, warn):
    """kotai_record の警報（{"ym", "text", "clerk", "fail"}）を 失敗／警告／事務へのお知らせ に振り分ける。"""
    for x in alerts or []:
        ym = x.get("ym")
        head = f"R{reiwa(ym[0])}年{ym[1]}月 " if ym else ""
        item = head + (x.get("text") or "")
        dst = failed if x.get("fail") else warn
        if item not in dst:
            dst.append(item)
        if x.get("clerk") and ym:
            clerk_note(ym[0], ym[1], x["clerk"])


BUCHO_STATE = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "bucho_notice_state.json")


def bucho_check(months, state_path=None):
    """管理当直の名簿で Dr.JOY の医師と結び付かない部長を、事務担当・事務担当2へのお知らせに積む（20261028 管理者:
    事務担当が自分で気づけるように）。その月の「管理当直」の名簿に同じ姓の医師がいれば、もう足してあるとみなして知らせない。
    同じ月・同じ部長は1回だけ（BUCHO_STATE に覚える）。読めなくても落ちない（ログに残すだけ）。"""
    import bucho_roster as B
    state_path = state_path or BUCHO_STATE
    try:
        people = C.load_staffs(C.newest_staffs())[0]
        miss = B.unresolved(people)
    except Exception as e:
        log(f"  ⚠ 管理当直の名簿を確かめられない: {type(e).__name__}: {e}")
        return []
    if not miss:
        return []
    try:
        with open(state_path, encoding="utf-8") as f:
            told = set(json.load(f).get("told") or [])
    except (OSError, ValueError):
        told = set()
    import master_store as MST
    out = []
    for (y, m) in months:
        if (y, m) < tuple(SYNC_FROM):
            continue
        try:
            mst = MST.load(MST.master_path(C.prod_month_dir(y, m), y, m), y, m)
        except Exception:
            continue
        have = {B._surname(x.get("name")) for x in (((mst.get("depts") or {}).get(B.OWNER) or {}).get("roster") or [])}
        for name, dept, why, key in miss:
            k = f"{y:04d}-{m:02d}|{name}"
            if key in have or k in told:
                continue
            clerk_note(y, m, f"管理当直の名簿の「{name}（{dept}）」が、Dr.JOY の医師と結び付きません（{why}）。"
                             "アプリで「管理当直」を選び、医師リストの下の検索欄から足してください。")
            told.add(k)
            out.append(k)
    if out:
        try:
            os.makedirs(os.path.dirname(state_path), exist_ok=True)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump({"told": sorted(told)}, f, ensure_ascii=False, indent=1)
        except OSError as e:
            log(f"  ⚠ 管理当直のお知らせを覚えておけない（{e}）→ 次の回も同じお知らせが出る")
        log(f"  管理当直の名簿で結び付かない部長を事務へ知らせる: {', '.join(out)}")
    return out


def run_rebind(months, warn):
    """名簿外（手入力）で入れた医師を、Dr.JOY に登録された本人に付け替える（unreg_rebind。20261028 管理者）。
    交代記録の前に呼ぶ。付け替えられなかった医師はログに残すだけ（同姓同名・所属が違う・Dr.JOY にいない）。"""
    import unreg_rebind as UR
    try:
        people = C.load_staffs(C.newest_staffs())[0]
        import shozoku as SZ
        hist, dmap = SZ.load_history(), SZ.load_dept_map()
    except Exception as e:
        log(f"  ⚠ 名簿外の付け替え: 名簿が読めない（{type(e).__name__}: {e}）→ 今日は付け替えない")
        return
    for (y, m) in months:
        try:
            rows, skipped = UR.run_month(y, m, log=log, people=people, hist=hist, dept_map=dmap)
        except Exception as e:
            msg = f"R{reiwa(y)}年{m}月 名簿外の付け替えに失敗: {type(e).__name__}: {e}"
            log(f"  ⚠ {msg}")
            if msg not in warn:
                warn.append(msg)
            continue
        for o, n, lg in rows:
            log(f"  名簿外の付け替え: R{reiwa(y)}年{m}月 {o} {n} → {lg.split('@')[0]}")
        for o, n, why in skipped:
            log(f"  名簿外のまま: R{reiwa(y)}年{m}月 {o} {n}（{why}）")


XL_PROBE_N = 50              # 空のブックのセルに書く回数
XL_SLOW_SEC = 2.0            # これより遅ければ Windows 入力エクスペリエンスを再起動する（正常 約0.5秒・暴走中 約4秒）


def excel_probe(n=XL_PROBE_N):
    """Excel の COM 呼び出しの速さ = 空のブックのセルに n 回書く秒数（Excel の起動の時間は含めない）。"""
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        t = time.time()
        for i in range(n):
            ws.Cells(1 + i, 1).Value = i
        sec = time.time() - t
        wb.Close(SaveChanges=False)
        return sec
    finally:
        xl.Quit()


def restart_text_input():
    """Windows 入力エクスペリエンス（TextInputHost.exe）を止める。Windows が数秒で起動し直す（20261028 実測）。"""
    import subprocess
    r = subprocess.run(["taskkill", "/F", "/IM", "TextInputHost.exe"], capture_output=True, timeout=30,
                       stdin=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return r.returncode == 0


def check_excel_speed(warn, probe=excel_probe, restart=restart_text_input, wait=5):
    """Excel が遅ければ Windows 入力エクスペリエンスを再起動する（管理者 20261028）。
    10/25 から TextInputHost が暴走して COM の呼び出しが約25倍遅くなり、15時の反映が3分→42分に延びた
    。
    再起動しても遅ければ警報に載せる（PC の再起動を勧める）。測れなければログだけ。"""
    try:
        sec = probe()
    except Exception as e:
        log(f"  ⚠ Excel の速さを測れない: {type(e).__name__}: {e}")
        return None
    if sec < XL_SLOW_SEC:
        log(f"  Excel の速さ: セル{XL_PROBE_N}回 {sec:.2f}秒（正常）")
        return sec
    log(f"  ⚠ Excel が遅い: セル{XL_PROBE_N}回 {sec:.2f}秒 → Windows 入力エクスペリエンスを再起動する")
    try:
        ok = restart()
    except Exception as e:
        ok = False
        log(f"  ⚠ 再起動に失敗: {type(e).__name__}: {e}")
    time.sleep(wait)
    try:
        sec2 = probe()
    except Exception as e:
        log(f"  ⚠ 再起動のあと Excel の速さを測れない: {type(e).__name__}: {e}")
        return None
    if sec2 < XL_SLOW_SEC:
        log(f"  再起動{'した' if ok else 'を試みた'} → セル{XL_PROBE_N}回 {sec2:.2f}秒（戻った）")
    else:
        msg = (f"Excel が遅いまま（セル{XL_PROBE_N}回 {sec:.2f}秒 → 入力エクスペリエンスの再起動後 {sec2:.2f}秒）。"
               "15時の反映が長引くので、PC の再起動を")
        log(f"  ⚠ {msg}")
        if msg not in warn:
            warn.append(msg)
    return sec2


def run_kotai(failed, warn):
    """当直交代申請.xlsx の承認を読み、交代記録（台帳）を更新する（全部の月フォルダ）。DTX14 に書く前に呼ぶ（20260924）。
    戻り: {(年, 月): 状態}。丸ごと失敗したら None（決裁済みの月は書かない）。"""
    import kotai_record as K
    al = []
    try:
        res = K.update_all(log=lambda s: log("  " + s), alerts=al) or {}
    except (Exception, SystemExit) as e:
        log(f"  ✗ 交代記録（当直交代申請.xlsx）の更新に失敗: {type(e).__name__}: {e}")
        failed.append("交代記録")
        return None
    failed += [f"交代記録 R{reiwa(y)}年{m}月（{r}）" for (y, m), r in sorted(res.items()) if r in K.FAILED]
    warn += [f"R{reiwa(y)}年{m}月 {K.FILE_NAME} が読めない（承認欄を読めず、作り直してもいない。前の承認のまま反映）"
             for (y, m), r in sorted(res.items()) if r in K.WARNED]
    route_alerts(al, failed, warn)
    return res


ONEDRIVE_WARN = "OneDrive が動いていません。SharePoint の正本・公開 PDF が更新されません"


def onedrive_running():
    """OneDrive.exe が動いているか。True/False。確かめられなければ None（警告しない）。"""
    import subprocess
    try:
        r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OneDrive.exe", "/FO", "CSV", "/NH"], capture_output=True,
                           text=True, encoding="cp932", errors="replace", timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.SubprocessError):
        return None
    if r.returncode != 0:
        return None
    return "onedrive.exe" in (r.stdout or "").lower()


def check_roots():
    """本番の共有と SharePoint の同期フォルダに届くか。届かないものの説明のリスト（失敗）。
    20260926: OneDrive が動いていなければ _STATE["root_warn"] に警告を積む（失敗ではない。main が警告に回す）。"""
    bad = []
    if not os.path.isdir(PROD_ROOT):
        bad.append(f"本番の共有に届かない（{PROD_ROOT}）→ マスターが読めず DTX14 に書けない")
    if not os.path.isdir(DTX_ROOT):
        bad.append(f"SharePoint の同期フォルダが無い（{DTX_ROOT}）")
    if onedrive_running() is False:
        _STATE["root_warn"].append(ONEDRIVE_WARN)
    return bad


def _norm_assign(a):
    """approved / approved_pending（{科: [記録...]}）を比べられる形に（並び・空の科を気にしない）。"""
    import json
    if not isinstance(a, dict):
        return None
    return {o: sorted(json.dumps(x, ensure_ascii=False, sort_keys=True) for x in (v or []))
            for o, v in a.items() if v}


def unreflected(lg):
    """台帳に、承認したのに DTX14 に反映していない交代があるか（approval=承認 で reflected_at が無い、
    または approved_pending が approved と違う）。"""
    import kotai_record as K
    if any(r.get("approval") == K.APPROVE and not r.get("reflected_at") and not r.get("reverted")
           for r in lg.get("records") or []):
        return True
    ap = lg.get("approved_pending")
    return isinstance(ap, dict) and _norm_assign(ap) != _norm_assign(lg.get("approved"))


def late_months(window, root=None):
    r"""前月〜来月（window）の外で、承認したのに DTX14 に反映していない交代がある月（20260926）。
    台帳（集約\交代記録.json）を読むだけ（kotai_record の load_ledger・month_dirs）。
    戻り: (書く月 [(年, 月)], 警告 [文])。SYNC_FROM より前の月は見ない。正本の DTX14 がちょうど1本の月だけ書く。"""
    import kotai_record as K
    root = root or PROD_ROOT
    out, probs = [], []
    try:
        dirs = K.month_dirs(root)
    except OSError:
        return out, probs                   # 本番の共有に届かない（check_roots が失敗として知らせる）
    for md, y, m in dirs:
        if (y, m) in window or (y, m) < SYNC_FROM:
            continue
        lp = os.path.join(md, "集約", K.LEDGER_NAME)
        if not os.path.isfile(lp):
            continue
        try:
            lg = K.load_ledger(lp, y, m)
        except K.LedgerError as e:
            probs.append(f"R{reiwa(y)}年{m}月 交代記録の台帳が読めない（後からの承認を反映できるか分からない）: {e}")
            continue
        if not unreflected(lg):
            continue
        _d, hits = dtx_candidates(y, m)
        if len(hits) == 1:
            log(f"  R{reiwa(y)}年{m}月: 前月〜来月の外だが、承認したのに DTX14 に反映していない交代がある → 今日書く")
            out.append((y, m))
        else:
            probs.append(f"R{reiwa(y)}年{m}月 承認したのに DTX14 に反映していない交代があるが、正本の DTX14 が "
                         f"{len(hits)} 本なので反映できない（{_d}）")
    return sorted(out), probs


def toban_warnings(months):
    """当番当直が data/toban.csv に無い月の警告（毎日。20260926）。"""
    import new_month as N
    out = []
    for (y, m) in months:
        t = N.toban_note(y, m)
        if t and t not in out:
            out.append(t)
    return out


ADMIN_ADDRS = {"user003@example.ac.jp", "admin@example.ac.jp"}   # 管理者の普段使いのアカウント（同じ1つのアカウントの別名）
AUTO_TAG = " ※自動送信"
CONTACT_LINE = "問い合わせ先: 当直勤務表の担当"


def _addrs(s):
    """Split a recipient string ("a; b" / "a, b") into lower-cased addresses."""
    return [x.strip().lower() for x in re.split(r"[;,]", s or "") if x.strip()]


def admin_only(to, cc=None):
    """True when every recipient (To and CC) is 管理者 (user003 / admin)."""
    rs = _addrs(to) + _addrs(cc)
    return bool(rs) and all(r in ADMIN_ADDRS for r in rs)


def admin_mail(subject, body):
    """管理者だけに送る自動メールの書式（管理者 20261020）: 件名の末尾に ※自動送信、本文の最後の行に問い合わせ先。"""
    subject = subject if subject.rstrip().endswith(AUTO_TAG.strip()) else subject.rstrip() + AUTO_TAG
    body = body.rstrip("\r\n")
    if body.splitlines()[-1:] != [CONTACT_LINE]:
        body = body + "\n" + CONTACT_LINE
    return subject, body


def outlook_send(to, subject, body, cc=None):
    """Outlook COM で1通送る（管理者の PC・対話ログオンのタスクで動く）。cc = 「; 」区切りの宛先（20260927。省略可）。
    20261020 管理者: 宛先（To と CC）が管理者だけのときは、件名の末尾に ※自動送信・本文の最後の行に問い合わせ先を付け、
    送信済みアイテムに残さない（DeleteAfterSubmit）。ほかの人が入るメール（事務へのお知らせ・所属の変更メール）は従来どおり。"""
    import win32com.client
    me_only = admin_only(to, cc)
    if me_only:
        subject, body = admin_mail(subject, body)
    ol = win32com.client.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)                  # olMailItem
    mail.To = to
    if cc:
        mail.CC = cc
    mail.Subject = subject
    mail.Body = body
    if me_only:
        mail.DeleteAfterSubmit = True        # do not keep a copy in Sent Items (mail to 管理者 only)
    mail.Send()


ALERT_SENDER = outlook_send                 # テストでは差し替える（本物のメールを送らない）


def send_cc(to, subject, body, cc=None):
    """cc つきで1通（所属の変更メール。20260927）。cc は宛先のリストか「; 」区切り。
    ALERT_SENDER がテストで差し替えられていれば、本物は送らずにそちらへ渡す（宛先に cc を書き足す）。"""
    cc = "; ".join(cc) if isinstance(cc, (list, tuple)) else (cc or "")
    if ALERT_SENDER is outlook_send:
        return outlook_send(to, subject, body, cc=cc)
    return ALERT_SENDER(f"{to}; cc: {cc}" if cc else to, subject, body)


def run_shozoku_notice(dry, today=None):
    """所属の変更メール（20260927 管理者）。月末の日と毎週水曜の朝9時（morning_shozoku_notice から） → 来月だけ（今月は比べない）。SYNC_FROM 以降で共有マスターがある月だけ。
    戻り: 管理者への警告に載せる文のリスト（送れた・送らなかったは log に）。"""
    import shozoku_notice as SN
    t = today or dt.date.today()
    targets = SN.targets_for(t, SYNC_FROM, exists=lambda y, m: os.path.isfile(master_file(y, m)))
    if not targets:
        return []
    log(f"  所属の変更メール: 対象 {', '.join(f'R{reiwa(y)}年{m}月' for y, m in targets)}")
    res = SN.run(targets, dry=dry, master_file=master_file, log=log, today=t,
                 sender=lambda to, subject, body, cc: send_cc(to, subject, body, cc))
    return [f"所属の変更メール: {x}" for x in res["problems"]]


SHOZOKU_HOUR = 9                                # 所属の変更メールを送る時刻（管理者 20260927: 月末と毎週水曜の9時）
SHOZOKU_STATE = os.path.join(os.path.dirname(LOCAL_LOG), "shozoku_notice_day.json")


def morning_shozoku_notice(now=None, state_path=None):
    """10分おきの回（kotai_record --quick）から: 月末と毎週水曜、9時を過ぎて最初の回で所属の変更メールを1日1回だけ送る
    （20260927 管理者。以前は15時の反映の中で送っていた）。送れなければ次の回（10分後）にやり直し、12時を過ぎても
    送れなければ管理者へ1通知らせて、その日はやめる。戻り: "sent"／"skip"／"retry"／"gave-up"。"""
    now = now or dt.datetime.now()
    today = now.date()
    if now.hour < SHOZOKU_HOUR:
        return "skip"
    import shozoku_notice as SN
    if not SN.targets_for(today, SYNC_FROM):
        return "skip"
    state_path = state_path or SHOZOKU_STATE
    try:
        with open(state_path, encoding="utf-8") as f:
            st = json.load(f)
    except (OSError, ValueError):
        st = {}
    if st.get("day") == today.isoformat() and st.get("status") in ("sent", "gave-up"):
        return "skip"
    status, err = "sent", ""
    try:
        for x in run_shozoku_notice(False, today=today):
            log(f"  ⚠ {x}")
    except Exception as e:
        err = f"{type(e).__name__}: {e}"
        log(f"  ✗ 所属の変更メールを作れない・送れない: {err}")
        status = "retry"
        if now.hour >= 12:
            status = "gave-up"
            try:
                ALERT_SENDER(ALERT_TO, "【警報】当直勤務表 所属の変更メールを送れませんでした",
                             f"月末・水曜の朝の所属の変更メールを、12時までに送れませんでした。\n\n  {err}\n\n"
                             f"ログ: {LOCAL_LOG}\n手で送るとき: python -X utf8 tools\\shozoku_notice.py --month YYYY-MM")
            except Exception:
                pass
    try:
        os.makedirs(os.path.dirname(state_path), exist_ok=True)
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump({"day": today.isoformat(), "status": status, "error": err}, f, ensure_ascii=False)
    except OSError:
        pass
    return status


def alert_subject(failed, warn):
    if failed:
        m = re.search(r"R\d+年\d+月", failed[0])
        head = m.group(0) if m else failed[0][:30]
        return f"【警報】当直勤務表 15時の反映に失敗（{head}{' ほか' if len(failed) > 1 else ''}）"
    m = re.search(r"R\d+年\d+月", warn[0])
    what = "競合コピーを検出" if all("競合コピー" in w for w in warn) else "確かめることがあります"
    return f"【警報】当直勤務表 15時の反映で{what}（{m.group(0) if m else warn[0][:30]}）"


def alert_body(failed, warn, notes):
    out = ["当直勤務表の毎日15時の反映（tools\\dtx14_daily_sync.py）で、次のことがありました。", ""]
    if failed:
        out += ["■ 失敗"] + [f"  - {x}" for x in failed] + [""]
    if warn:
        out += ["■ 警告"] + [f"  - {x}" for x in warn] + [""]
    if notes:
        out += ["■ そのほか"] + [f"  - {x}" for x in notes] + [""]
    errs = [ln for ln in _STATE["run"] if "✗" in ln or "⚠" in ln]
    if errs:
        out += ["■ エラーと警告の行（この回のログから）"] + [f"  {ln}" for ln in errs[-40:]] + [""]
    out += ["■ ログ",
            f"  本番: {SHARE_LOG}" + (f"（この回は書けなかった: {_STATE['share_err']}）" if _STATE["share_err"] else ""),
            f"  ローカル: {LOCAL_LOG}", "",
            "■ 戻し方・止め方",
            "  書き込みに失敗した月の正本は差し替えていません（元のまま）。",
            "  書き込み前の写しは各月フォルダの _上書き履歴\\ にあります。",
            f"  書き込みを止めるには 本番の 管理ツール\\ か 月のフォルダに {STOP_NAME} を置いてください。",
            "", "（このメールは管理者の PC のタスク「Touchoku_DTX14_sync」が自動で送っています）"]
    return "\n".join(out)


def notify(failed, warn, notes, dry):
    """失敗か競合コピーがあれば1通だけ警報メール。送れなくても落ちない。戻り: 状態の一言（送らなければ ""）。"""
    if not failed and not warn:
        return ""
    subject = alert_subject(failed, warn)
    if dry:
        log(f"  (dry-run) 警報メールは送らない: {subject}")
        return "メールなし（dry-run）"
    try:
        ALERT_SENDER(ALERT_TO, subject, alert_body(failed, warn, notes))
        log(f"  警報メールを送った: {ALERT_TO}『{subject}』")
        return "警報メール送信済み"
    except Exception as e:
        log(f"  ✗ 警報メールを送れない（Outlook）: {type(e).__name__}: {e}")
        return "警報メール送信失敗"


def clerk_subject(items):
    months = sorted({ym for ym, _t in items})
    y, m = months[0]
    return f"【当直勤務表】今日の15時の書き込みについて（R{reiwa(y)}年{m}月{'ほか' if len(months) > 1 else ''}）"


def clerk_body(items):
    """事務（事務担当・事務担当2）へのお知らせの本文。パス・エラーの細部は書かない。短い文で。"""
    out = ["事務担当、事務担当2", "",
           "当直勤務表の、毎日15時の自動の書き込みで、見ていただきたいことがあります。", ""]
    for ym in sorted({ym for ym, _t in items}):
        out.append(f"■ R{reiwa(ym[0])}年{ym[1]}月分")
        out += [f"・{t}" for k, t in items if k == ym]
        out.append("")
    out += ["分からないときは、管理者に知らせてください。",
            "（このメールは自動で送っています。管理者にも同じ内容が届いています）"]
    return "\n".join(out)


def clerk_notify(items, dry):
    """事務へのお知らせを1通だけ送る（CLERK_TO。自分で直せることだけ。管理者 20260925）。送れなくても落ちない。"""
    if not items or not CLERK_TO:
        return ""
    subject = clerk_subject(items)
    if dry:
        log(f"  (dry-run) 事務へのお知らせは送らない: {subject}")
        return "事務へのメールなし（dry-run）"
    try:
        ALERT_SENDER("; ".join(CLERK_TO), subject, clerk_body(items))
        log(f"  事務へのお知らせを送った: {'; '.join(CLERK_TO)}『{subject}』")
        return "事務へのお知らせ送信済み"
    except Exception as e:
        log(f"  ✗ 事務へのお知らせを送れない（Outlook）: {type(e).__name__}: {e}")
        return "事務へのお知らせ送信失敗"


def _last_end(p):
    """ログの最後の「=== 終了」の時刻（無ければ None・読めなければ OSError）。"""
    with open(p, "rb") as f:
        f.seek(0, os.SEEK_END)
        size = f.tell()
        f.seek(max(0, size - 512 * 1024))
        lines = f.read().decode("utf-8", errors="replace").splitlines()
    for ln in reversed(lines):
        if "=== 終了" in ln:
            mt = re.match(r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]", ln)
            if mt:
                return dt.datetime.strptime(mt.group(1), "%Y-%m-%d %H:%M:%S")
    return None


def check_last_run(max_age_hours=26, log_path=None, now=None):
    """15時の反映の見張り（20260925）: ログの最後の「=== 終了」がいつか。
    max_age_hours より古い（または見つからない）なら問題の一言を返す。問題が無ければ ""。
    20261028: ローカルと共有の両方のログを見て新しい方で決める（ローカルのログが古い版に戻っていて、動いているのに
    警報が出た＝10/1・10/28）。log_path を渡せばそのログだけを見る。"""
    now = now or dt.datetime.now()
    paths = [log_path] if log_path else [LOCAL_LOG, SHARE_LOG]
    ends, missing = [], []
    for p in paths:
        try:
            t = _last_end(p)
        except OSError:
            missing.append(p)
            continue
        if t:
            ends.append(t)
    if not ends:
        if missing and len(missing) == len(paths):
            return f"15時の反映のログがありません（{'・'.join(missing)}）"
        return f"15時の反映のログに終了の行がありません（{'・'.join(paths)}）"
    t = max(ends)
    hours = (now - t).total_seconds() / 3600
    if hours > max_age_hours:
        return f"15時の反映が最後に終わったのは {t:%m/%d %H:%M}（{hours:.0f} 時間前）"
    return ""


def new_month_mail(y, m, made, notes):
    """置き場を作ったときのお知らせ（事務担当・事務担当2宛て・cc 管理者。管理者 20261101: 事務にも送る・自動処理と分かる
    文体で、簡潔に要件だけ）。made = new_month.ensure の作ったもの（文の頭で種類を見分ける）、
    notes = 知らせること（管理者が対応すること）。戻り: (件名, 本文)。"""
    import new_month as N
    r = reiwa(y)
    app = any(x.startswith("アプリの月フォルダ") for x in made)
    tname = next((mm.group(1) for mm in (re.search(r"DTX14 の雛形 (.+?\.xlsm)", x) for x in made) if mm), "")
    kname = next((mm.group(1) for mm in (re.search(r"手入力の部署に配る写し (.+?\.xlsm)", x) for x in made) if mm), "")
    out = ["事務担当、事務担当2", "",
           f"当直勤務表の自動処理です。R{r}年{m}月分の置き場を作りました。", ""]
    if app:
        out += [f"■ アプリの月フォルダ（提出の締切 {N.deadline_for(y, m)}）",
                os.path.join(PROD_ROOT, f"R{r}年{m}月"), ""]
    if tname or kname:
        # 管理者の PC の OneDrive のパスはお二人には意味がないので、SharePoint の中の場所で書く
        out.append(f"■ SharePoint「{os.path.basename(os.path.dirname(DTX_ROOT))}」"
                   f" {os.path.basename(DTX_ROOT)} ＞ R{r}.{m}月 当直勤務表")
        if tname:
            out.append(f"・{tname}（正本。診療科の行は毎日15時に自動で書き込みます）")
        if kname:
            out.append(f"・{kname}（手入力の部署に配る写し）")
        out.append("")
    if notes:
        out.append("■ 管理者が対応すること（お二人の作業はありません）")
        out += ["・" + ("当番当直の表がまだありません。" if "当番当直" in x and "toban.csv" in x else x) for x in notes]
        out.append("")
    out.append(CONTACT_LINE)
    # 所属の変更メールと同じく、管理者以外の宛先でも自動で送ったと分かるように件名の末尾に付ける（管理者 20261031・20261101）
    return f"【お知らせ】当直勤務表 R{r}.{m}月分の置き場を作りました{AUTO_TAG}","\n".join(out)


def maybe_new_month(dry, today=None):
    """毎日（20260926。以前は毎月1〜5日だけ）: 次の次の月と、欠けていれば来月の置き場（アプリの月フォルダ＋SharePoint の
    DTX14 雛形＋【診療部科名】の写し）を揃える（管理者 20260923）。無いものだけ作る。SYNC_FROM より前の月は触らない。
    戻り: (作ったもの[], 知らせること[], 失敗[])。作ったら【お知らせ】メール。作れなければ失敗（毎日）。
    知らせること = new_month の知らせ（ACL の失敗・当番当直が無い）と、6日を過ぎても（来月はいつでも）まだ欠けているもの。"""
    import new_month as N
    t = today or dt.date.today()
    nxt, nn = N.ensure_targets(t)
    made_all, notes, errs = [], [], []
    for (y, m) in (nxt, nn):
        if (y, m) < SYNC_FROM:
            continue
        tag = f"R{reiwa(y)}年{m}月"
        try:
            miss = N.missing(y, m)
        except (Exception, SystemExit) as e:
            miss = [f"欠けているかを確かめられない（{type(e).__name__}: {e}）"]
        if not miss:
            continue
        if dry:
            log(f"  (dry-run) {tag} の置き場で欠けているもの: {'、'.join(miss)}（作らない）")
            continue
        log(f"  {tag} の置き場で欠けているもの: {'、'.join(miss)} → 作る")
        failed_here = False
        try:
            made, ns = N.ensure(y, m, log=lambda s: log("  " + s))
        except (Exception, SystemExit) as e:        # month_setup は SystemExit を投げる（20260925）
            log(f"  ✗ {tag} の置き場を作れない: {type(e).__name__}: {e}")
            errs.append(f"{tag} の置き場づくり（{e}）")
            made, ns, failed_here = [], [], True
        for x in ns:
            if x not in notes:
                notes.append(x)
        if made:
            made_all += made
            subject, body = new_month_mail(y, m, made, ns)
            try:
                ALERT_SENDER("; ".join(CLERK_TO), subject, body, cc=ALERT_TO)
                log(f"  置き場のお知らせを送った: {'; '.join(CLERK_TO)}（cc {ALERT_TO}）『{subject}』")
            except Exception as e:
                log(f"  ✗ お知らせのメールを送れない: {e}")
        if failed_here:
            continue                                  # 失敗として載せた（作れない日は毎日）
        try:
            still = N.missing(y, m)
        except (Exception, SystemExit):
            still = []
        if still and ((y, m) == nxt or t.day not in N.AUTO_DAYS):
            item = f"{tag} の置き場がまだ揃っていない: {'、'.join(still)}"
            log(f"  ⚠ {item}")
            notes.append(item)
    return made_all, notes, errs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-fill", action="store_true", help="DTX14 へ書かず PDF だけ")
    ap.add_argument("--month", action="append", help="対象の月 YYYY-MM（省略=前月・今月・来月）")
    a = ap.parse_args()
    months = [tuple(int(x) for x in s.split("-")) for s in a.month] if a.month else months_window()
    daily = not a.month                          # タスクの本番の回（--month なし）だけ毎日の見回りをする（20260926）
    _STATE["share_err"], _STATE["run"], _STATE["clerk"], _STATE["dtx_hits"] = None, [], [], {}
    _STATE["root_warn"] = []
    log(f"=== 開始 対象 {['%d-%02d' % ym for ym in months]}{'（dry-run）' if a.dry_run else ''}")
    failed = []                                  # 失敗したこと（最後の行に書く・警報メール）
    warn, notes = [], []                         # 競合コピー（メールする）／停止スイッチ（メールの本文に添える）
    try:
        for msg in check_roots():
            log(f"  ✗ {msg}")
            failed.append(msg)
        for msg in _STATE["root_warn"]:          # OneDrive が動いていない（20260926）
            log(f"  ⚠ {msg}")
            if msg not in warn:
                warn.append(msg)
        # 当直交代申請.xlsx の承認を読んで交代記録を更新する。DTX14 に書く前に（今日の承認を今日反映。管理者 20260924）
        if daily and not a.dry_run:
            check_excel_speed(warn)           # 遅ければ Windows 入力エクスペリエンスを再起動（20261028）
            run_rebind(months, warn)          # 名簿外の医師を本人に付け替える（交代記録の前に。20261028）
            bucho_check(months)               # 管理当直の名簿で結び付かない部長を事務へ（20261028）
        kotai = None if a.dry_run else run_kotai(failed, warn)
        late = []
        if daily:
            # 前月〜来月の外で、承認したのに反映していない交代がある月も書く（PDF は作らない。20260926）
            try:
                late, probs = late_months(months)
            except Exception as e:
                log(f"  ✗ 前月〜来月の外の承認を確かめられない: {type(e).__name__}: {e}")
                late, probs = [], [f"前月〜来月の外の承認を確かめられない（{e}）"]
            for x in probs:
                log(f"  ⚠ {x}")
                if x not in warn:
                    warn.append(x)
        for (y, m) in months + late:
            try:
                res = (sync_month(y, m, a, warn, kotai) if (y, m) in months
                       else sync_month(y, m, a, warn, kotai, pdf=False))
                failed += [f"R{reiwa(y)}年{m}月 {x}" for x in res]
            except (Exception, SystemExit) as e:  # 1か月の失敗で残りの月を止めない
                log(f"  ✗ R{reiwa(y)}年{m}月 の処理に失敗: {type(e).__name__}: {e}")
                failed.append(f"R{reiwa(y)}年{m}月")
    except BaseException as e:                   # 想定外（Ctrl+C など）もメールに載せてから止まる
        log(f"  ✗ 反映が途中で止まった: {type(e).__name__}: {e} {traceback.format_exc(limit=3)}")
        failed.append("反映が途中で止まった")
        raise
    finally:                                     # 途中で何があっても掃除は走らせる（交代記録は月の前に回した。20260924）
        try:
            cleanup_public(a.dry_run)
        except Exception as e:
            log(f"  ✗ 書きかけの PDF の掃除に失敗: {type(e).__name__}: {e}")
            failed.append("書きかけの PDF の掃除")
        if daily and not a.dry_run and PROD_ROOT == C.PROD_ROOT and ADMIN_DIR == C.PROD_ADMIN:
            # 全職員が読む config・予定表・_履歴 に職員コードが入っていれば消す（20261015 管理者。ふだんは0件で何もしない。
            #   開いたままの古いアプリが書き戻した分を拾う）。試験（置き場を差し替えた selftest）では本番を触らない
            try:
                import strip_codes as SC
                _n, bad = SC.strip_all(root=PROD_ROOT, log=log)
                for b in bad:
                    msg = f"予定表などから職員コードを消せない: {b}"
                    log(f"  ⚠ {msg}")
                    if msg not in warn:
                        warn.append(msg)
            except Exception as e:
                log(f"  ⚠ 職員コードの見回りに失敗: {type(e).__name__}: {e}")
            # 院内掲示板（WordPress）の当直勤務表のページに公開用の PDF を載せる（20261025 管理者。変わった月だけ上げ直す。
            #   資格情報が無ければ何もしない）
            try:
                import wp_publish as WP
                if "selftest" not in os.path.basename(sys.argv[0] if sys.argv else ""):   # 試験では本物の掲示板に書かない
                    WP.publish(PUBLIC_DIR, log=log)
            except Exception as e:
                msg = f"院内掲示板に当直勤務表を載せられない: {type(e).__name__}: {e}"
                log(f"  ⚠ {msg}")
                if msg not in warn:
                    warn.append(msg)
        if daily:
            # 次の次の月（と、欠けていれば来月）の置き場（管理者 20260923。20260926 に1〜5日から毎日へ）
            try:
                _made, nm_notes, _errs = maybe_new_month(a.dry_run)
                failed += _errs
                for x in nm_notes:                # 20260926: 捨てずに警告へ（ACL の失敗・まだ揃っていない・当番当直）
                    if x not in warn:
                        warn.append(x)
            except (Exception, SystemExit) as e:
                log(f"  ✗ 次の次の月の置き場づくりに失敗: {type(e).__name__}: {e}")
                failed.append("次の次の月の置き場づくり")
            # 当番当直が data/toban.csv に無い月（前月〜来月・次の次の月）は毎日の警告に（20260926）
            try:
                import new_month as N
                for x in toban_warnings(months + [N.add_months(*months_window()[1], 2)]):
                    log(f"  ⚠ {x}")
                    if x not in warn:
                        warn.append(x)
            except Exception as e:
                log(f"  ✗ 当番当直の有無を確かめられない: {type(e).__name__}: {e}")
            # 所属の変更メールは 20260927 から朝9時（10分おきの回の morning_shozoku_notice）。15時の回では送らない
        stops = sorted({p for p in (stop_file(os.path.join(DTX_ROOT, f"R{reiwa(y)}.{m}月 当直勤務表"))
                                    for (y, m) in months) if p})
        notes += [f"停止スイッチで DTX14 への書き込みを止めている: {p}" for p in stops]
        if _STATE["share_err"]:
            failed.append(f"本番の共有のログに書けない（{_STATE['share_err']}）")
        sent = notify(failed, warn, notes, a.dry_run)
        sent_c = clerk_notify(_STATE["clerk"], a.dry_run)       # 事務へ（自分で直せることだけ。20260925）
        if sent_c:
            sent = f"{sent}・{sent_c}" if sent else sent_c
        log("=== 終了" + (f"（✗ 失敗あり: {' / '.join(failed)}）" if failed else "（失敗なし）")
            + (f"（⚠ {' / '.join(warn)}）" if warn else "") + ("（⏸ 停止スイッチあり）" if stops else "")
            + (f"【{sent}】" if sent else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    # 20261007: the task runs under pythonw (no console window). On 10/7 15:00 the visible console got a Ctrl+C
    # (exit 0xC000013A) and the run stopped mid-way. pythonw has no stderr, so send it to the same .err file the .bat used.
    if sys.stderr is None:
        try:
            _err_dir = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "log")
            os.makedirs(_err_dir, exist_ok=True)
            sys.stderr = open(os.path.join(_err_dir, "dtx14_daily_sync.err"), "a", encoding="utf-8")
        except OSError:
            pass
    sys.exit(main())
