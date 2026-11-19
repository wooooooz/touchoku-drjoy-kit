# -*- coding: utf-8 -*-
r"""kotai_record.py — 当直交代の公式記録 `当直交代申請.xlsx` を作る。20261028

管理者 2026-09-23: 従来の Forms「当直交替申請書」の代わりに、公式な交代記録として残す。
置き場所は SharePoint（同期フォルダ）の `当直勤務表_正本\R8.N月 当直勤務表\`（DTX14 の現物と同じ月フォルダ。同日の指示で本番共有の月フォルダからは外した）。
SharePoint の月フォルダがまだ無い月は作らずに待つ（フォルダは new_month.py が毎月1〜5日に作る。20260923）。
シートは「交代記録」1枚（記録した順）。説明の文言は載せない（同日の指示）。
更新は毎日15時の反映（dtx14_daily_sync.py の最後）のときだけ（同日の指示。最初は10分おきのタスクにしていた）。

記録の決め方
  * 起点 = 管理ツール③「この月を確定」（病院長決裁）。確定のときに焼いた各科の baseline が「元の割当」。
    確定前の書き直しは予定づくりなので記録しない。
  * 各科がアプリで直して提出するたびに、その版と直前の版の差を 1 件ずつ記録する（1日×昼/夜×枠ごと）。
    版は共有マスターの `master\_履歴\`（上書き前の複製）と現在のマスターから、科ごとの rev で拾う。
    前の回から何度出し直しても、途中の版は _履歴\ に残っているので取りこぼさない（記録日時は各版の提出時刻）。
  * 記録は本番共有の `R8年N月\集約\交代記録.json`（管理者だけが書ける）に積み増すだけで、消さない・番号を振り直さない。
    xlsx はそこから毎回作り直す（xlsx を手で直されても、次に書き直すときに台帳の内容へ戻る）。
  * 確定を解除して付け直したとき: 承認した記録があれば最初の確定からの続きとして積む（元の割当は焼き直さない）。
    承認した記録がまだ無ければ新しい確定から台帳を作り直す（20260925。下の「本番前の見直し」）。
  * 「確定後の版か」と「処理の順番」は各科の rev で決める（20260923。updated は各 PC の時計なのでずれる）。
    確定のときの各科の rev（マスターの confirmed.revs）を台帳の起点にし、それより大きい rev だけを記録する。
    confirmed.revs の無い古い確定だけは従来どおり updated と確定時刻の比較に落ちる。
    同じ科の版は必ず rev の順に処理する。科をまたいだ並び（seq）は updated の順だが、科の中の rev の順は崩さない。
  * マスターの科の rev が台帳の rev より小さい＝マスターが _履歴\ から戻された。その科は記録を出さずに
    今のマスターの版へ合わせ直し（ログに ⚠ を出し、台帳の resyncs[] に残す）、次の提出からの差を記録する。
    合わせ直した科は、合わせ直した時刻より前に書かれたファイル（戻す前の版）を以後は拾わない（同じ rev が2系統あるため）。
  * 台帳が在るのに読めないときは、その月には触らない（空の台帳で上書きすると記録と番号が失われる）。

区分: 交代（A→B）／追加／取り止め／当直→オンコール／オンコール→当直。
records[] の1件は 氏名に加えて from_login/from_code・to_login/to_code（職員コード11桁）を持つ。
手当計算 の Dr.JOY 投入（交代の自動登録。管理者 20260923）がこれを正として読む＝項目名を変えないこと。
同じ提出で同じ日の昼と夜が同じ変わり方をしたものは「昼夜」の1行にまとめる（記録は昼・夜の2件のまま）。

承認（管理者 2026-09-24）: 決裁後の変更は「交代申請」。事務担当・事務担当2が xlsx の「承認」欄（ドロップダウン 承認／却下。
空欄＝承認待ち）で決め、**承認したものだけ** を DTX14 と Dr.JOY に反映する。
  * 承認欄は xlsx を作り直す前に毎回読む（表示の行と記録は隠し列「記録番号」の seq で結ぶ＝並べ替え・絞り込みに強い）。
    読めない（壊れている・開けない）ときは台帳の値のまま・xlsx も作り直さない（書いてあった承認を消さない）。
    承認／却下／空欄 以外は読まない（ログに残す）。元に戻した却下は、あとで承認欄を変えても受け付けない。
    承認→却下は、まだ反映していない（DTX14 に書く前の）ものだけ受け付ける。反映済みは変えられない（20260926）。
  * 台帳に baseline（確定のときの元の割当・compact）を持ち、approved = baseline に承認した記録を seq の順に当てたもの。
    DTX14（dtx14_daily_sync が approved_master() で作る）と 手当計算 の Dr.JOY 投入はこの approved を読む
    （state は各科が出した今の予定＝承認待ち・却下を含む。Dr.JOY の登録に使わないこと）。
  * 却下: 15時の回に、その科の共有マスターの今の予定から打ち消して put_dept で書き戻す（by＝「却下で戻した（当直交代申請）」）。
    台帳の state もその版に合わせるので、戻したこと自体は記録にならない。競合（その間に科が出し直した）なら次の回にやり直す。
  * 承認待ちの通知メールは出さない。
  * 15時の回の順: 承認欄を読む → 版を集めて記録を積む → 却下を戻す → approved を作る → 台帳 → xlsx。
  * すぐ載せる回（--quick。10分おきのタスク。管理者 20260924: 新しい申請はすぐ事務の目に入るように）:
    マスターが変わっていなければ何もしない。記録が増えたときだけ、承認欄を読んでから xlsx を作り直す。
    却下の戻し・approved・「反映済み」は15時の回だけの仕事（それまで承認は「承認（15時に反映）」と出る）。
    15時の回と重ならないように、台帳の横に 交代記録.json.lock を置く（すぐ載せる回は、あれば黙って飛ばす）。

本番前の見直し（20260925）:
  * 確定しているかは 集約\確定.json（管理者だけが書ける確定の正本。master_store.confirm／unconfirm が書く）で決める
    （master_store.effective_confirmed）。マスターの confirmed と食い違えば警報を出して 確定.json に従う。
    確定.json が無い旧い月だけマスターの confirmed に落ちる（ログに ⚠）。未確定の間は何も記録しない。
  * 確定の付け直し（押し間違いの戻し）: 確定の時刻が台帳の confirmed_at と違い、承認した記録がまだ無ければ、
    台帳を新しい確定から作り直す（前の記録は discarded[] に残す。記録番号は前の続きから＝xlsx の古い行と混ざらない）。
    承認した記録があれば焼き直さずに警報。
  * 同じマスの申請が続いたとき（A→B のあと B→C）: 承認した記録でも、当てはまらない操作が1つでもあれば当てない
    （hold。状態「前の申請（No.n）が未決のため保留」・警報・事務へのお知らせ）。先の申請が承認待ちの後の申請は
    状態「先の申請（No.n）の承認待ち」。却下の打ち消しが当てはまらないときもマスターに書かない（警報）。
  * approved（手当計算 の Dr.JOY 投入が読む）と reflected_at は、その月の DTX14 への書き込みが通ったあとで
    だけ付ける（publish_approved。dtx14_daily_sync が呼ぶ）。15時の交代記録の回が作るのは approved_pending
    （これから DTX14 に書く割当）と hold まで。書き込みが失敗・停止スイッチ・SYNC_FROM 前の月は付けない。
  * 見落とし防止: マスターの更新時刻は読む前に控え、飛ばすかどうかは各科の rev（台帳の state）でも確かめる。
  * xlsx の書き直し: 自分が書いたあとに他で変わり、しかも RECENT_XLSX 秒以内なら書き直さない（事務が入力中／
    同期前）。承認→空欄は受け付けず警報（古い写しで上書きされた可能性）。反映済みの承認→却下は受け付けて警報（→ 20260926 に取りやめ。下を参照）。
    当直交代申請.xlsx の競合コピー（-PC名・(1)・- コピー）は警報と事務へのお知らせ。
    隠し列「記録番号」は、表の1行になり得る組（同じ提出・同じ日・同じ変わり方の昼と夜）でなければ読まない。
  * 版を保存したアカウント（NTFS の持ち主）を records[].saved_by_account に残し、by と違えば入力者欄に出して警報。
  * 警報は run_month／update_all の alerts（リスト）に積む。すぐ載せる回の警報は台帳の alerts_pending に貯め、
    15時の回がまとめて渡す（10分ごとにメールしない）。
  * すぐ載せる回は、15時の反映が 26 時間以上終わっていなければ1日1通だけ管理者へ警報（dtx14_daily_sync.check_last_run）。

20260926（レビューの続き）:
  * 反映済み（reflected_at あり＝DTX14 に書け、Dr.JOY の登録が読んだ）の承認は変えられない（管理者 20260926 案1）。
    承認欄が却下・空欄などに変えられても承認のまま・次に xlsx を書き直すときに「承認」へ戻す・事務と管理者へ知らせる。
    取り消すときは、元に戻す当直交代申請をアプリから出してもらう。反映済みの行の承認欄は灰色（シートの保護はしない）。
    保留（compute_approved）も反映済みの記録からは reflected_at を外さない。後の反映済みの記録とぶつかる前の申請の方を保留にする。
  * 台帳を確定から作るとき、マスターの確定の印が正本と違う・確定のときの rev があるのに元の割当が焼かれていない／
    予定があったのに元の割当が空、なら作らない（'baseline-wait'＝失敗・管理者へ警報・次の回にやり直す）。
    空の元の割当を焼くと、15時にその科の行が DTX14 から消えるため。
  * 却下の戻し（put_dept）のあと台帳を保存する前に落ちた回: 次の回に「却下で戻した」版を見たら、その版で効果が
    消えた却下の記録を「元に戻した」にする（以前は戻せないと警報を出し続けた）。
  * 台帳が見えない（FileNotFoundError）ときに新しく作るのは、集約\ の一覧が読めて、前に台帳があった跡
    （一覧に台帳・控え 交代記録_bak.json/_bak2.json（20261014 より前は 交代記録.json.bak/.bak2）・確定から LEDGER_GRACE 以上・xlsx に記録の行）が無いときだけ。
    跡があれば LedgerError（控えから戻すか、作り直してよければ 集約\交代記録.json.recreate を置く）。保存のたびに控えを取る。
  * 承認欄の読めない値は事務にも知らせる。xlsx が書けないのは PermissionError に限らず OSError なら「開かれている」扱い
    （次の回に書く。その月の DTX14 は止めない）。

usage: python -X utf8 tools\kotai_record.py [--month 2026-11 ...] [--force] [--quick] [--root 本番以外の当直勤務表フォルダ]
  10分おきのタスク: pythonw.exe -X utf8 tools\kotai_record.py --quick（作業フォルダ C:\tools\touchoku）
"""
import argparse
import copy
import datetime as dt
import glob
import hashlib
import json
import os
import re
import shutil
import socket
import sys
import time
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402
import master_store as MST    # noqa: E402
from dtx14_daily_sync import DTX_ROOT, conflict_copies, reiwa   # noqa: E402  SharePoint 同期フォルダの 当直勤務表_正本

FILE_NAME = "当直交代申請.xlsx"
# 試験用フォルダ（R8年N月(動作確認用テストフォルダ)）の交代記録の書き出し先（20261001 管理者: 交代申請を試す。本物と分ける）
TEST_FILE_NAME = "当直交代申請（動作確認用）.xlsx"
LEDGER_NAME = "交代記録.json"
SCHEMA = "hospital.duty-kotai/1"
SCAN_MARGIN = 600.0            # _履歴\ を読み直す余裕（秒）。複製は元の更新時刻を引き継ぐことがある
WD = "月火水木金土日"

KIND_SWAP, KIND_ADD, KIND_DROP = "交代", "追加", "取り止め"
KIND_TO_OC, KIND_TO_DUTY = "当直→オンコール", "オンコール→当直"
APPROVE, REJECT = "承認", "却下"                 # xlsx の「承認」欄（空欄＝承認待ち。管理者 20260924）
REVERT_BY = "却下で戻した（当直交代申請）"        # 却下を打ち消してマスターへ書き戻した版の by（記録にしない）
SHEET = "交代記録"
SEQ_HEADER = "記録番号"                          # 隠し列: 表示の1行 ↔ 台帳の記録（seq。「12,13」）
FAILED = ("ledger-error", "error", "busy", "truth-error", "baseline-wait")   # run_month／update_all の戻りのうち失敗を表すもの（15時の回）
WARNED = ("xlsx-unreadable",)                   # 失敗ではないが知らせるもの（承認欄が読めず xlsx を作り直していない）
LOCK_STALE = 30 * 60                            # 台帳のロックがこれより古ければ落ちた置き土産とみなす（秒）
LOCK_WAIT_FULL = 180.0                          # 15時の回がすぐ載せる回の終わりを待つ上限（秒）
LEDGER_GRACE = 5 * 86400                        # 確定からこれより経っても台帳が無い＝前にあったはず（秒。20260926）
BAK1, BAK2 = "_bak", "_bak2"                    # 台帳の控え（保存のたびに _bak、半日に1回 _bak2。20260926）
OLD_BAKS = (".bak", ".bak2")                    # 20261014 より前の控えの名前（交代記録.json.bak）。跡としてだけ見る
BAK2_AGE = 12 * 3600
RECREATE = ".recreate"                          # 管理者が置けば、跡があっても台帳を新しく作ってよい（作ったら消す）
RECENT_XLSX = 15 * 60                           # xlsx が他で変わってからこれより新しければ書き直さない（秒。20260925）
_LOCALAPP = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku")
LOCAL_LOG = os.path.join(_LOCALAPP, "log", "kotai_record.log")
WATCHDOG_STATE = os.path.join(_LOCALAPP, "watchdog.json")   # 15時の反映の見張りの警報を送った日（1日1通）
GROUP_ACCOUNTS = {"administrators", "system", "管理者", "domain admins"}   # 持ち主がグループのときは突き合わせない


class LedgerError(Exception):
    """台帳（交代記録.json）が在るのに読めない。空の台帳で上書きしてはいけない。"""


class LedgerNotReady(Exception):
    """確定から台帳を作る材料（マスターの元の割当）がまだ揃っていない。今回は作らずに次の回にやり直す（20260926）。"""


class LedgerBusy(Exception):
    """台帳を別の回（15時の回／すぐ載せる回）が使っている。"""


def log_line(msg, root):
    line = f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)          # pythonw（コンソールなし）では sys.stdout が None で、print は何もしない
    for d in (os.path.join(root, "管理ツール", "log"), os.path.dirname(LOCAL_LOG)):   # 本番の共有＋ローカル
        try:
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, "kotai_record.log"), "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass


def _alert(al, text, clerk=None, fail=False):
    """警報を積む（al が None なら何もしない）。text＝管理者へ（パス・細部を書いてよい）。
    clerk＝事務担当・事務担当2へのやさしい言葉（自分で直せるものだけ。無ければ None）。fail＝失敗として扱う。"""
    if al is not None:
        al.append({"text": text, "clerk": clerk, "fail": bool(fail)})


_SID_NAMES = {}


def file_owner(path):
    """ファイルの持ち主（NTFS）のアカウント名（ドメインなし）。分からなければ ""（20260925: 版を保存した人の目印）。"""
    try:
        import win32security
        sd = win32security.GetFileSecurity(path, win32security.OWNER_SECURITY_INFORMATION)
        sid = sd.GetSecurityDescriptorOwner()
        k = str(sid)
        if k not in _SID_NAMES:
            name, _dom, _t = win32security.LookupAccountSid(None, sid)
            _SID_NAMES[k] = name or ""
        return _SID_NAMES[k]
    except Exception:
        return ""


FILE_OWNER = file_owner          # テストでは差し替える


def account_mismatch(by, acct):
    """入力者（by＝「ログイン名 / PC名」）と保存したアカウントが違うか。どちらかが分からなければ False。"""
    if not acct or acct.lower() in GROUP_ACCOUNTS or not by or by in (REVERT_BY, "app"):
        return False
    m = re.match(r"[A-Za-z0-9._-]+", by.strip())
    return bool(m) and m.group(0).lower() != acct.lower()


class ledger_lock:
    """台帳（集約\\交代記録.json）の横に置く排他作成のロック（20260924: すぐ載せる回と15時の回を重ねない）。
    wait 秒まで待ち、取れなければ LedgerBusy。LOCK_STALE より古いロックは落ちた置き土産として外す。
    外すときは目印が自分のものの時だけ。"""

    def __init__(self, ledger_path, wait=0.0):
        self.p, self.wait, self.token = ledger_path + ".lock", wait, None

    def _info(self):
        try:
            st = os.stat(self.p)
            with open(self.p, encoding="utf-8", errors="replace") as f:
                return f.read(), st.st_mtime
        except OSError:
            return None

    def __enter__(self):
        os.makedirs(os.path.dirname(self.p), exist_ok=True)
        t0 = time.monotonic()
        token = f"{socket.gethostname()} {os.getpid()} {dt.datetime.now().isoformat(timespec='seconds')} {os.urandom(4).hex()}"
        while True:
            try:
                fd = os.open(self.p, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                try:
                    os.write(fd, token.encode("utf-8"))
                finally:
                    os.close(fd)
                self.token = token
                return self
            except FileExistsError:
                seen = self._info()
                if seen and time.time() - seen[1] > LOCK_STALE and self._info() == seen:
                    try:
                        os.remove(self.p)
                    except OSError:
                        pass
                    continue
                if time.monotonic() - t0 >= self.wait:
                    raise LedgerBusy(f"台帳を別の回が使っている（{seen[0] if seen else '?'}）")
                time.sleep(min(2.0, max(0.05, self.wait / 20)))

    def __exit__(self, *a):
        cur = self._info()
        if cur and cur[0] == self.token:
            try:
                os.remove(self.p)
            except OSError:
                pass
        self.token = None
        return False


# ---------------------------------------------------------------- 版を集めて差を取る

_UNREG = {}


def norm_login(login):
    """名簿外（unreg:氏名）を、氏名が同じ Dr.JOY の医師がちょうど1人ならその人のログインIDに寄せる（20261028）。
    名簿外の医師を本人に付け替えた（unreg_rebind）前後の版を同じ人として比べ、交代申請にしないため。"""
    lg = str(login or "")
    if not lg.startswith("unreg:"):
        return lg
    if "map" not in _UNREG:
        mp = {}
        try:
            import unreg_rebind as UR
            for k, ps in UR.doctors_by_name(C.load_staffs(C.newest_staffs())[0]).items():
                if len(ps) == 1:
                    mp["unreg:" + k] = ps[0]["login"]
        except Exception:                         # 名簿が読めなければ寄せない（比べ方は今までどおり）
            mp = {}
        _UNREG["map"] = mp
    return _UNREG["map"].get(lg, lg)


def _norm_compact(compact):
    """台帳に前から持っている compact（名簿外のログインのまま）を norm_login で寄せる。"""
    return [[r[0], r[1], r[2], norm_login(r[3]), *r[4:]] for r in compact or []]


def _compact(duties):
    """duties → [[date, band, frame_id, login, name, oncall, 職員コード]]（台帳に持つ最小限）。開院記念日の昼は全員オンコールに寄せる。
    20261028: 名簿外のログインは norm_login で本人に寄せる。"""
    recs = [dict(r) for r in duties or [] if r.get("date") and r.get("frame_id")]
    C.apply_daytime_oncall(recs)
    # 20261015: 予定表から職員コードを外したので、空ならログインIDで 管理ツール\名簿\ から引く（from_code/to_code は
    #   Dr.JOY への交代の自動登録が読む）。名簿外（unreg:）や引けない人は空のまま
    return sorted([r["date"], r.get("band") or "", r["frame_id"], norm_login(r.get("staff_login")),
                   r.get("name") or r.get("name_disp") or "", bool(r.get("oncall")),
                   r.get("staff_code") or C.code_of(r.get("staff_login"))]
                  for r in recs)


def _groups(compact):
    g = defaultdict(dict)
    for date, band, fid, login, name, oc, *rest in compact:
        g[(date, band, fid)][login or "name:" + name] = (name, oc, login, rest[0] if rest else "")
    return g


def diff(old, new):
    """2つの版（compact）の差 → 記録の元（1日×昼/夜×枠ごと）。"""
    go, gn = _groups(old), _groups(new)
    out = []
    for key in sorted(set(go) | set(gn)):
        o, n = go.get(key, {}), gn.get(key, {})
        if o == n:
            continue
        rem = [k for k in o if k not in n]
        add = [k for k in n if k not in o]
        for k in o:
            if k in n and o[k][1] != n[k][1]:
                out.append((key, KIND_TO_OC if n[k][1] else KIND_TO_DUTY, o[k], n[k]))
        # 同じ区分（当直どうし・オンコールどうし）を先に組み、余りを区分違いで組む。残りは追加／取り止め
        for same in (True, False):
            for r in list(rem):
                a = next((x for x in add if (o[r][1] == n[x][1]) == same), None)
                if a is not None:
                    out.append((key, KIND_SWAP, o[r], n[a]))
                    rem.remove(r)
                    add.remove(a)
        out += [(key, KIND_DROP, o[r], None) for r in rem]
        out += [(key, KIND_ADD, None, n[a]) for a in add]
    return [{"date": k[0], "band": k[1], "frame_id": k[2], "kind": kind,
             "from": f[0] if f else "", "from_oc": f[1] if f else None,
             "from_login": f[2] if f else "", "from_code": f[3] if f else "",
             "to": t[0] if t else "", "to_oc": t[1] if t else None,
             "to_login": t[2] if t else "", "to_code": t[3] if t else ""} for k, kind, f, t in out]


def _versions(md, mp, since_ts):
    r"""_履歴\ の複製（since_ts より新しいもの）と現在のマスターから、科ごとの版を集める。
    戻り: {owner: {rev: (updated, by, duties, ファイルの更新時刻, 保存したアカウント)}}
    保存したアカウント（20260925）: マスター全体の rev G の版を書いた人 =
      ・今のマスターが G なら、その持ち主（一時ファイルから差し替えるので書いた人のもの）
      ・G-1 の中身を持つ _履歴\ の複製の持ち主（上書きの直前に、G を書く人が複製する）
    科の版は、その科の updated がマスター全体の updated と同じファイル（＝その科を書いた版）で決める。分からなければ ""。"""
    out = defaultdict(dict)
    files = []
    hist = os.path.join(os.path.dirname(mp), MST.HISTORY)
    for p in glob.glob(os.path.join(hist, "*.json")):
        try:
            if os.path.getmtime(p) >= since_ts - SCAN_MARGIN:
                files.append(p)
        except OSError:
            pass
    author = {}                 # マスター全体の rev → 書いた人
    loaded = []
    for p in sorted(files) + [mp]:
        try:
            mt = os.path.getmtime(p)
            with open(p, encoding="utf-8-sig") as f:
                m = json.load(f)
        except (OSError, ValueError):
            continue
        g = int(m.get("rev") or 0)
        acct = FILE_OWNER(p)
        if p == mp:
            author.setdefault(g, acct)
        else:
            author.setdefault(g + 1, acct)
        loaded.append((p, mt, m, g))
    for p, mt, m, g in loaded:
        for owner, ent in (m.get("depts") or {}).items():
            rev = int(ent.get("rev") or 0)
            if rev:     # 同じ rev は後の（新しい）ファイルを採る＝マスターを戻したあとの版が勝つ。普段は同じ中身
                prev = out[owner].get(rev)
                acct = author.get(g, "") if (ent.get("updated") or "") == (m.get("updated") or "") else ""
                acct = (prev[4] if prev and prev[4] else acct)    # 最初に出てきたファイル（その版を書いたとき）を採る
                out[owner][rev] = (ent.get("updated") or "", ent.get("by") or "", ent.get("duties") or [], mt, acct)
    return out


def _epoch(iso):
    try:
        return dt.datetime.fromisoformat(iso).timestamp()
    except (TypeError, ValueError):
        return 0.0


def _order(todo):
    """(updated, owner, rev, …) の並び: 科をまたいでは updated（提出した PC の時計）の順、
    ただし同じ科の中は必ず rev の順（時計がずれていても rev6 を rev5 より先にしない）。
    updated で並べたうえで、各科が占める位置にその科の版を rev の順に入れ直す。"""
    todo = sorted(todo, key=lambda x: x[:3])
    slots = defaultdict(list)
    for i, x in enumerate(todo):
        slots[x[1]].append(i)
    out = list(todo)
    for owner, idx in slots.items():
        for i, x in zip(idx, sorted((todo[i] for i in idx), key=lambda x: x[2])):
            out[i] = x
    return out


_UNSET = object()
# 確定の付け直しで台帳を作り直すときに消す（discarded[] に残す）項目
_RESET_KEYS = ("confirmed_at", "confirmed_by", "rev_based", "state", "baseline", "scanned_ts", "approved", "approved_sig",
               "approved_pending", "approved_pending_sig", "approved_pending_applied", "resyncs", "reconfirm_seen")


def _next_seq(lg):
    """次の記録番号。作り直した台帳でも前の番号の続き（xlsx に残っている古い行の番号と混ざらない）。"""
    return max([int(r["seq"]) for r in lg.get("records") or []] + [int(lg.get("seq_base") or 0)]) + 1


def update_ledger(ledger, master, md, mp, log=None, conf=_UNSET, al=None):
    """台帳に新しい記録を積む。戻り: 積んだ件数。
    conf = 確定の正本（master_store.effective_confirmed の値）。渡さなければマスターの confirmed（旧い呼び方）。
    未確定なら何も記録しない。確定の時刻が台帳と違えば、承認した記録が無いときだけ台帳を作り直す（20260925）。"""
    log = log or (lambda s: None)
    if conf is _UNSET:
        conf = master.get("confirmed") or None
    if not conf:
        return 0               # 未確定（解除中を含む）: 予定づくりなので記録しない
    cat = conf.get("at") or ""
    n_ok = sum(1 for r in ledger.get("records") or [] if r.get("approval") == APPROVE)
    if not ledger.get("confirmed_at") or (ledger["confirmed_at"] != cat and not n_ok):
        probs = baseline_problems(master, conf, mp)   # 台帳を確定から作る前に（何も変えないうちに）確かめる。20260926
        if probs:
            raise LedgerNotReady("；".join(probs))
    if ledger.get("confirmed_at") and ledger["confirmed_at"] != cat:
        if not n_ok:
            old = {k: ledger.get(k) for k in ("confirmed_at", "confirmed_by", "resyncs")}
            recs = ledger.get("records") or []
            ledger["seq_base"] = _next_seq(ledger) - 1
            ledger.setdefault("discarded", []).append(dict(old, discarded_at=_now_iso(), new_confirmed_at=cat,
                                                           records=recs))
            for k in _RESET_KEYS:
                ledger.pop(k, None)
            ledger["records"] = []
            ledger["xlsx_dirty"] = True
            msg = (f"確定が付け直された（{old['confirmed_at']} → {cat}）。承認した申請はまだ無いので、元の割当を新しい確定で"
                   f"焼き直した（前の記録 {len(recs)} 件は台帳の discarded に残した）")
            log("⚠ " + msg)
            _alert(al, msg)
        elif ledger.get("reconfirm_seen") != cat:
            ledger["reconfirm_seen"] = cat
            msg = (f"確定が付け直された（{ledger['confirmed_at']} → {cat}）が、承認した申請が {n_ok} 件あるので元の割当は"
                   "焼き直さない（前の確定からの続きとして記録する）。付け直しが要るなら管理者が台帳を確かめる")
            log("⚠ " + msg)
            _alert(al, msg, fail=True)
    if not ledger.get("confirmed_at"):
        ledger["confirmed_at"] = cat
        ledger["confirmed_by"] = conf.get("by") or ""
        revs = conf.get("revs")
        # 起点の rev = 確定のときの各科の rev（confirmed.revs）。古い確定（revs 無し）は 0 から＋時刻で判定
        ledger["rev_based"] = isinstance(revs, dict)
        ledger["state"] = {o: {"rev": int((revs or {}).get(o) or 0), "duties": _compact(e.get("baseline") or [])}
                           for o, e in (master.get("depts") or {}).items()}
        # 元の割当（承認した交代を当てる起点。20260924）。あとから出てきた科の元の割当は空
        ledger["baseline"] = {o: copy.deepcopy(st["duties"]) for o, st in ledger["state"].items()}
        # 確定したときの割当は決裁済み＝反映してよい（手当計算 が読む approved の起点。20260925）
        ledger["approved"] = copy.deepcopy(ledger["baseline"])
        ledger["approved_sig"] = ""
        ledger["scanned_ts"] = _epoch(ledger["confirmed_at"])
    t0 = time.time()
    since = min(ledger.get("scanned_ts") or 0.0, t0)
    base_at = ledger["confirmed_at"]
    rev_based = bool(ledger.get("rev_based"))
    # マスターが _履歴\ から戻された（今の rev が台帳より小さい）科: 記録を出さずに今の版へ合わせ直す
    resynced = set()
    for owner, ent in (master.get("depts") or {}).items():
        st = ledger["state"].get(owner)
        cur = int(ent.get("rev") or 0)
        if st and cur < st["rev"]:
            log(f"⚠ {owner}: マスターの rev {cur} が台帳の rev {st['rev']} より小さい（マスターを戻した？）"
                f" → 記録を出さずに今のマスターへ合わせ直す")
            ledger.setdefault("resyncs", []).append({"at": dt.datetime.now().isoformat(timespec="seconds"),
                                                     "owner": owner, "from_rev": st["rev"], "to_rev": cur})
            st["rev"], st["duties"] = cur, _compact(ent.get("duties") or [])
            st["resync_ts"] = time.time()
            resynced.add(owner)
    todo = []                  # 科をまたいで提出の時刻順に積む（科ごとには rev の順）
    for owner, vers in _versions(md, mp, since).items():
        if owner in resynced:
            continue           # 戻す前の版（_履歴\ に残る大きい rev）を拾わない
        st = ledger["state"].setdefault(owner, {"rev": 0, "duties": []})
        floor = st.get("resync_ts") or 0.0     # 合わせ直した科: それより前のファイル＝戻す前の系統の版は拾わない
        # 確定後か: rev で決める。古い確定の台帳でまだ確定後の版を1つも処理していない科だけ時刻で判定
        todo += [(vers[rev][0], owner, rev, vers[rev][1], vers[rev][2], vers[rev][4]) for rev in vers
                 if rev > st["rev"] and vers[rev][3] >= floor
                 and (rev_based or st["rev"] > 0 or vers[rev][0] > base_at)]
    added = 0
    for updated, owner, rev, by, duties, acct in _order(todo):
        st = ledger["state"][owner]
        new = _compact(duties)
        if st["rev"] and rev > st["rev"] + 1:
            # 20260925: 途中の版が _履歴\ に無い（上書き前の複製に失敗した＝共有の不調・容量など）。
            #   差分は前の版からまとめて記録するので交代は落ちないが、途中の申請が1件にまとまる。気づけるように知らせる
            msg = (f"{owner}: rev {st['rev'] + 1}〜{rev - 1} の版が master\\_履歴\\ に見つからない"
                   f"（アプリの履歴の複製が失敗している？）→ rev {rev} までの変更を1つにまとめて記録")
            log("⚠ " + msg)
            _alert(al, msg)
        if by == REVERT_BY:    # 却下を打ち消した版（書き戻したあと台帳を保存する前に落ちた回の残り）は記録にしない
            got = recover_reverts(ledger, owner, st["duties"], new, updated)
            log(f"{owner}: rev {rev} は却下で戻した版 → 記録にせず合わせる"
                + (f"（戻したあと台帳を保存できていなかった却下 {len(got)} 件を「元に戻した」にした: "
                   f"{', '.join(map(str, got))}）" if got else ""))
        else:
            bad = account_mismatch(by, acct)
            n0 = added
            for r in diff(_norm_compact(st["duties"]), new):
                ledger["records"].append(dict(r, seq=_next_seq(ledger), at=updated, by=by,
                                              owner=owner, rev=rev, approval="", saved_by_account=acct,
                                              account_mismatch=bad))
                added += 1
            if bad and added > n0:
                msg = (f"{owner} rev {rev}（{updated}）: 入力者「{by}」と、版を保存したアカウント「{acct}」が違う"
                       f"（記録 {ledger['records'][n0 - added]['seq']}〜）→ 本人の申請か確かめる")
                log("⚠ " + msg)
                _alert(al, msg)
        st["rev"], st["duties"] = rev, new
    ledger["scanned_ts"] = t0
    return added


def baseline_problems(master, conf, mp):
    """確定から台帳を作ってよいか（20260926）。戻り: 問題の一覧（空なら作ってよい）。
      * 読んだマスターの確定の印が確定の正本（conf）と違う＝確定より前の版を読んだ／書き換えられた
      * 確定のときの rev（conf.revs）が 1 以上なのに、その科がマスターに無い・元の割当（baseline）が焼かれていない（None）
      * 元の割当が空なのに、確定のときの版（その rev）には予定があった
    どれかがあれば、空の元の割当を焼いて 15時に DTX14 のその科の行を消すことになるので作らない。"""
    out = []
    mc = master.get("confirmed") or None
    if not mc or (mc.get("at") or "") != (conf.get("at") or ""):
        out.append(f"共有マスターの確定の印（{(mc or {}).get('at') or '未確定'}）が確定の正本（{conf.get('at') or '?'}）と違う"
                   "（確定より前の版を読んだか、書き換えられた）")
    revs = conf.get("revs")
    if not isinstance(revs, dict):
        return out
    depts = master.get("depts") or {}
    vers = None
    for o, r0 in sorted(revs.items()):
        r0 = int(r0 or 0)
        if r0 <= 0:
            continue
        e = depts.get(o)
        if e is None:
            out.append(f"{o}: 確定のとき rev {r0} だったのにマスターに無い")
            continue
        b = e.get("baseline")
        if b is None:
            out.append(f"{o}: 確定のとき rev {r0} だったのに元の割当（baseline）が焼かれていない")
            continue
        if b:
            continue
        if int(e.get("rev") or 0) == r0:
            had = bool(e.get("duties"))
        else:
            if vers is None:
                vers = _versions(os.path.dirname(os.path.dirname(mp)), mp, 0.0)
            v = (vers.get(o) or {}).get(r0)
            had = bool(v and v[2])
        if had:
            out.append(f"{o}: 確定のとき（rev {r0}）には予定があったのに、元の割当が空")
    return out


def _effect(rows, r):
    """記録 r の効果が compact の行 rows にあるか: "after"（申請のあと）/"before"（申請の前）/"neither"。"""
    cell = _cell_of(r)
    have = {_pkey(x[3], x[4]): bool(x[5]) for x in rows if (x[0], x[1], x[2]) == cell}
    f = _pkey(r.get("from_login") or "", r.get("from") or "")
    t = _pkey(r.get("to_login") or "", r.get("to") or "")
    k = r.get("kind")
    if k == KIND_SWAP:
        if t in have and f not in have:
            return "after"
        if f in have and t not in have:
            return "before"
    elif k == KIND_ADD:
        return "after" if t in have else "before"
    elif k == KIND_DROP:
        return "before" if f in have else "after"
    elif k in (KIND_TO_OC, KIND_TO_DUTY) and f in have:
        return "after" if have[f] == bool(r.get("to_oc")) else "before"
    return "neither"


def recover_reverts(ledger, owner, old, new, at):
    """「却下で戻した」版（old→new）を台帳がまだ知らないとき（戻したあと台帳を保存する前に落ちた。20260926）:
    その科のまだ戻していない却下のうち、old では効果があり new では消えているものを「元に戻した」にする。戻り: seq の一覧。"""
    got = []
    for r in ledger.get("records") or []:
        if (r.get("owner") == owner and r.get("approval") == REJECT and not r.get("reverted")
                and _effect(old, r) != "before" and _effect(new, r) == "before"):
            r["reverted"], r["reverted_at"], r["reverted_recovered"] = True, at or _now_iso(), True
            got.append(r["seq"])
    return got


def _fresh_evidence(path, xlsx=None, now=None):
    """台帳（path）が見えないとき、前にあった跡（20260926）。戻り: 跡の説明の一覧（空なら新しく作ってよい）。"""
    d, base = os.path.dirname(path), os.path.basename(path)
    try:
        names = os.listdir(d)
    except FileNotFoundError:
        try:                   # 集約\ が無い＝まだ何も置いていない月（月フォルダは見えている）
            os.listdir(os.path.dirname(d))
            return []
        except OSError as e:
            return [f"月フォルダの一覧が読めない（{e}）"]
    except OSError as e:
        return [f"{d} の一覧が読めない（{e}）"]
    if base + RECREATE in names:
        return []              # 管理者が作り直しを許した
    ev = []
    if base in names:
        ev.append("一覧には台帳があるのに開けない")
    baks = [n for n in [os.path.basename(b) for b in bak_paths(base)] + [base + s for s in OLD_BAKS] if n in names]
    if baks:
        ev.append(f"控え（{'、'.join(baks)}）がある")
    if MST.TRUTH_NAME in names:
        try:
            with open(os.path.join(d, MST.TRUTH_NAME), encoding="utf-8-sig") as f:
                t = json.load(f)
            ats = [h.get("at") for h in t.get("history") or [] if h.get("action") == "confirm"]
            ats.append((t.get("confirmed") or {}).get("at"))
            old = sorted(a for a in ats if a and _epoch(a) and (now or time.time()) - _epoch(a) > LEDGER_GRACE)
            if old:
                ev.append(f"{old[0]} に確定してから {LEDGER_GRACE // 86400} 日以上たっている")
        except (OSError, ValueError, AttributeError, TypeError):
            pass               # 確定.json が読めないのは確定を決める側が止める（truth-error）
    if xlsx and os.path.isfile(xlsx):
        kind, rows = read_approvals(xlsx, log=lambda s: None)
        if kind == "ok" and rows:
            ev.append(f"{FILE_NAME} に記録が {len(rows)} 行ある")
        elif kind in ("bad", "old"):
            ev.append(f"{FILE_NAME} がある（読めない・旧い形）")
    return ev


def load_ledger(path, year, month, xlsx=None):
    """台帳を読む。新しく作るのはファイルが無い（FileNotFoundError）うえに、前にあった跡が無いときだけ（20260926。
    _fresh_evidence。xlsx＝当直交代申請.xlsx のパス＝記録の行があれば跡）。
    在るのに読めない・形が違う・跡があるのに見えないときは LedgerError（空の台帳で上書きすると記録と番号が失われる）。"""
    try:
        with open(path, encoding="utf-8") as f:
            lg = json.load(f)
    except FileNotFoundError:
        ev = _fresh_evidence(path, xlsx)
        if ev:
            raise LedgerError(f"{path} が見えませんが、前にあった跡があります（{'・'.join(ev)}）。"
                              f"控え（{os.path.basename(bak_paths(path)[0])}）から戻すか、作り直してよければ "
                              f"{os.path.basename(path)}{RECREATE} を同じフォルダに置いてください")
        return {"schema": SCHEMA, "year": year, "month": month, "confirmed_at": "", "confirmed_by": "",
                "state": {}, "records": [], "scanned_ts": 0.0, "master_mtime": 0.0, "xlsx_sig": ""}
    except (OSError, ValueError) as e:
        raise LedgerError(f"{path} が読めません（{type(e).__name__}: {e}）") from e
    if not isinstance(lg, dict) or lg.get("schema") != SCHEMA:
        raise LedgerError(f"{path} の形が違います（schema {lg.get('schema') if isinstance(lg, dict) else '?'}）")
    return lg


def bak_paths(path):
    """台帳の控えの名前（交代記録_bak.json・交代記録_bak2.json）。20261014: 以前の 交代記録.json.bak は二重拡張子で、
    Python が作ると EDR がランサムウェアの疑いとして通報する形だった。"""
    stem, ext = os.path.splitext(path)
    return stem + BAK1 + ext, stem + BAK2 + ext


def _backup_ledger(path):
    """保存の前に今の台帳を _bak へ（半日に1回は _bak を _bak2 へ）。控えは best-effort（20260926）。"""
    if not os.path.isfile(path):
        return
    b1, b2 = bak_paths(path)
    try:
        if os.path.isfile(b1) and (not os.path.isfile(b2) or time.time() - os.path.getmtime(b2) > BAK2_AGE):
            shutil.copy2(b1, b2)
            os.utime(b2)
        shutil.copyfile(path, b1)
    except OSError:
        pass


def _save_json(path, obj):
    """台帳を書く（一時ファイル → 差し替え）。書く前に控えを取り、作り直しの許し（.recreate）は書けたら消す。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = MST.tmp_name(path)                      # 20261002: 二重拡張子にしない（EDR がランサムウェア疑いで通報する）
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)
    _backup_ledger(path)
    os.replace(tmp, path)
    if os.path.exists(path + RECREATE):
        try:
            os.remove(path + RECREATE)
        except OSError:
            pass


# ---------------------------------------------------------------- 承認（管理者 20260924）

def _now_iso():
    return dt.datetime.now().isoformat(timespec="seconds")


def _pkey(login, name):
    """1マスの中で人を見分けるキー（_groups と同じ: ログインID、無ければ氏名）。"""
    return login or "name:" + (name or "")


def _ops(rec, inverse):
    """記録1件を当てる操作（inverse=True は打ち消し）。("rm"|"add", 側) / ("oc", 側, 値)。側は "from"/"to"。"""
    k = rec.get("kind")
    if k == KIND_SWAP:
        return [("rm", "to"), ("add", "from")] if inverse else [("rm", "from"), ("add", "to")]
    if k == KIND_ADD:
        return [("rm", "to")] if inverse else [("add", "to")]
    if k == KIND_DROP:
        return [("add", "from")] if inverse else [("rm", "from")]
    if k in (KIND_TO_OC, KIND_TO_DUTY):
        return [("oc", "from", bool(rec.get("from_oc") if inverse else rec.get("to_oc")))]
    return []


def _apply(items, rec, inverse, key_of, cell_of, set_oc, make):
    """items（compact の行か duty の dict のリスト。その場で変える）に記録1件を当てる。
    当てはまらない操作（外す人がいない・足す人がもういる）は飛ばす。戻り: 飛ばした数。"""
    cell = (rec["date"], rec.get("band") or "", rec["frame_id"])
    miss = 0
    slot = None              # 20261025: 交代（外して足す）は、外した人の位置に足す＝マスの中の氏名の並びを変えない（却下の戻しも）
    for op in _ops(rec, inverse):
        side = op[1]
        pk = _pkey(rec.get(side + "_login") or "", rec.get(side) or "")
        hit = [i for i, x in enumerate(items) if cell_of(x) == cell and key_of(x) == pk]
        if op[0] == "rm":
            if hit:
                slot = hit[0]
                del items[hit[0]]
            else:
                miss += 1
        elif op[0] == "add":
            if hit:
                miss += 1
            elif slot is not None:
                items.insert(slot, make(rec, side))
            else:
                items.append(make(rec, side))
        elif hit:
            set_oc(items[hit[0]], op[2])
        else:
            miss += 1
    return miss


def _set5(x, v):
    x[5] = v


COMPACT_OPS = dict(      # compact = [date, band, frame_id, login, name, oncall, 職員コード]
    key_of=lambda x: _pkey(x[3], x[4]), cell_of=lambda x: (x[0], x[1], x[2]), set_oc=_set5,
    make=lambda rec, s: [rec["date"], rec.get("band") or "", rec["frame_id"], rec.get(s + "_login") or "",
                         rec.get(s) or "", bool(rec.get(s + "_oc")), rec.get(s + "_code") or ""])


def _duty_key(r):
    return _pkey(r.get("staff_login") or "", r.get("name") or r.get("name_disp") or "")


def make_duty(date, band, fid, login, name, oncall, code, frame=None, like=None):
    """duty 1件（schema hospital.duty/1。アプリの buildDuties と同じ形）。
    frame_dept・role は同じ枠の既存の duty（like）に合わせ、無ければ枠マスタ（frames.json）から。"""
    f, like = frame or {}, like or {}
    # 20261015: 職員コードは予定表に書かない（却下の戻しで全職員が読む予定表に入るため）。人別シートは書くときに引く
    d = {"schema": "hospital.duty/1", "staff_login": login or "", "staff_code": "", "name": name or "",
         "name_disp": C.namekey(name or ""), "date": date, "band": band or "",
         "frame_dept": like.get("frame_dept") or f.get("dept") or "", "frame_id": fid,
         "role": like.get("role") or f.get("role") or "", "oncall": bool(oncall), "note": ""}
    if (login or "").startswith("unreg:"):
        d["unregistered"] = True
    return d


def _duty_ops(fmap, like):
    def set_oc(r, v):
        r["oncall"] = bool(v)

    def make(rec, s):
        return make_duty(rec["date"], rec.get("band"), rec["frame_id"], rec.get(s + "_login"), rec.get(s),
                         rec.get(s + "_oc"), rec.get(s + "_code"), fmap.get(rec["frame_id"]), like.get(rec["frame_id"]))
    return dict(key_of=_duty_key, cell_of=lambda r: (r.get("date"), r.get("band") or "", r.get("frame_id")),
                set_oc=set_oc, make=make)


def migrate(lg, log=print):
    """20260924 より前に確定した台帳: baseline が無ければ今の state から記録を逆に当てて作る。
    それまでの記録は当時の決まり（記録＝反映）で既に反映していたので「承認」とみなす（legacy_approval）。"""
    if not lg.get("confirmed_at") or isinstance(lg.get("baseline"), dict):
        return False
    base = {o: copy.deepcopy(st.get("duties") or []) for o, st in (lg.get("state") or {}).items()}
    miss = 0
    for r in sorted(lg.get("records") or [], key=lambda x: -x["seq"]):
        miss += _apply(base.setdefault(r["owner"], []), r, True, **COMPACT_OPS)
    lg["baseline"] = {o: sorted(v) for o, v in base.items()}
    now = _now_iso()
    for r in lg.get("records") or []:
        if "approval" not in r:
            r.update(approval=APPROVE, approval_seen_at=now, reflected_at=now, legacy_approval=True)
    log(f"台帳を承認つきの形に直した（元の割当を逆算・それまでの記録 {len(lg.get('records') or [])} 件は承認とみなす"
        + (f"・当てはまらない {miss} か所" if miss else "")
        + ("・⚠ マスターの戻しの合わせ直しがあるので元の割当は要確認" if lg.get("resyncs") else "") + "）")
    return True


def read_approvals(xp, log=print):
    """当直交代申請.xlsx の「承認」「備考」欄を読む。
    戻り: ("none", [])＝ファイルが無い ／ ("old", [])＝承認欄の無い旧い形 ／ ("bad", [])＝読めない ／
          ("ok", [(seq のリスト, 承認欄の値, 備考, Excel の行番号)])"""
    if not os.path.isfile(xp):
        return "none", []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xp, read_only=True, data_only=True)
        try:
            ws = wb[SHEET] if SHEET in wb.sheetnames else wb.worksheets[0]
            rows = list(ws.iter_rows(values_only=True))
        finally:
            wb.close()
    except Exception as e:
        log(f"⚠ {FILE_NAME} が読めない（{type(e).__name__}: {e}）→ 承認欄は台帳の値のまま・作り直さない")
        return "bad", []
    for hi, row in enumerate(rows[:10]):
        head = [str(v).strip() if v is not None else "" for v in row]
        if SEQ_HEADER in head and APPROVE in head:
            cs, ca = head.index(SEQ_HEADER), head.index(APPROVE)
            cn = head.index("備考") if "備考" in head else None
            out = []
            for i, r in enumerate(rows[hi + 1:], hi + 2):
                r = list(r) + [None] * (len(head) - len(r))
                seqs = [int(x) for x in re.findall(r"\d+", str(r[cs] if r[cs] is not None else ""))]
                if seqs:
                    out.append((seqs, "" if r[ca] is None else str(r[ca]),
                                "" if cn is None or r[cn] is None else str(r[cn]).strip(), i))
            return "ok", out
    return "old", []


def _struct_key(r, fi):
    """表の1行にまとめられる記録の組の目印（display_rows と同じ: 同じ提出・同じ日・同じ当直名と表・同じ変わり方）。
    昼と夜の違いだけが許される。承認・状態・備考は入れない（事務の入力で変わるため）。"""
    name, table, _o = fi.get(r["frame_id"], (r["frame_id"], "", 999))
    return (r["at"], r["owner"], r["rev"], r["date"], name, table, r["kind"], r["from"], r["to"],
            r["from_oc"], r["to_oc"])


def valid_group(recs, fi):
    """隠し列「記録番号」の組が、表の1行になり得るか（1件、または同じ目印で昼夜が重ならない組）。20260925"""
    if len(recs) == 1:
        return True
    keys = {_struct_key(r, fi) for r in recs}
    bands = [r.get("band") for r in recs]
    return len(keys) == 1 and len(set(bands)) == len(bands)


def apply_approvals(lg, rows, log=print, frames=None, al=None):
    """read_approvals の行を台帳の記録へ移す。戻り: (台帳が変わった数, xlsx が台帳と食い違う＝作り直したい数)。
    20260925: 記録番号の組が表の1行になり得ないものは読まない（警報）。承認→空欄は受け付けない（警報。古い写しで
    上書きされた可能性。取り消すなら「却下」）。
    20260926: 反映済み（reflected_at）の承認は変えられない（却下・空欄・ほかの文字でも承認のまま。xlsx は次に書き直すとき
    「承認」に戻る。事務と管理者へ知らせる）。読めない文字も事務へ知らせる。"""
    bys = {r["seq"]: r for r in lg.get("records") or []}
    frames = frames if frames is not None else C.load_frames()
    fi = _frame_info(frames)
    no_of = {}
    if any(r.get("reflected_at") for r in bys.values()):     # 事務が見ている No.（書き換える前の台帳＝最後に書いた xlsx）
        no_of = {q: x["no"] for x in display_rows(lg.get("records") or [], frames, lambda b: b) for q in x["seqs"]}
    changed = stale = 0
    now = _now_iso()
    for seqs, raw, note, line in rows:
        recs = [bys[q] for q in seqs if q in bys]
        if not recs:
            continue
        if len(recs) != len(seqs) or len(set(seqs)) != len(seqs) or not valid_group(recs, fi):
            msg = (f"{FILE_NAME} {line}行目: 隠し列の記録番号「{','.join(map(str, seqs))}」が表の1行の組と合わない"
                   "（書き換えられた？）→ この行の承認欄・備考は読まない")
            log("⚠ " + msg)
            _alert(al, msg)
            stale += 1
            continue
        v = raw.replace(" ", "").replace("　", "")
        if v not in ("", APPROVE, REJECT):
            msg = f"{FILE_NAME} {line}行目: 承認欄の「{raw}」は読めない（承認／却下のどちらか）→ 前の値のまま"
            log("⚠ " + msg)
            _alert(al, msg, clerk=f"{FILE_NAME} の {line} 行目の承認欄は『承認』か『却下』を選んでください"
                                  f"（入っていた文字: {raw}）")
            stale += 1
            v = None
        locked = []
        for r in recs:
            if note != (r.get("approval_note") or ""):
                r["approval_note"] = note
                changed += 1
            cur = r.get("approval") or ""
            if v is None or v == cur:
                continue
            if r.get("reverted"):
                log(f"⚠ {line}行目（記録 {r['seq']}）: 却下で元に戻したあとなので、承認欄の変更（{v or '空欄'}）は受け付けない")
                stale += 1
                continue
            if r.get("reflected_at"):
                locked.append(r)          # 反映済みの承認は変えられない（管理者 20260926 案1）
                continue
            if cur == APPROVE and v == "":
                msg = (f"{FILE_NAME} {line}行目（記録 {r['seq']}）: 承認欄が「承認」から空欄に変わっていた → 承認のままにした"
                       "（古い写しで上書きされた可能性。取り消すなら「却下」を選ぶ）")
                log("⚠ " + msg)
                _alert(al, msg)
                stale += 1
                continue
            r["approval"], r["approval_seen_at"] = v, now
            changed += 1
        if locked:
            n = no_of.get(locked[0]["seq"], "?")
            msg = (f"{FILE_NAME} {line}行目（No.{n}・記録 {','.join(str(r['seq']) for r in locked)}）: 反映済みの承認が"
                   f"「{v or '空欄'}」に変えられていた → 受け付けず承認のまま（反映済みの承認は取り消せない。"
                   "承認欄は次に書き直すときに「承認」へ戻す）")
            log("⚠ " + msg)
            _alert(al, msg, clerk=f"No.{n} はすでに当直表と Dr.JOY に反映済みなので、承認を取り消せません。"
                                  "取り消すときは、元に戻す当直交代申請をアプリから出してください")
            stale += 1
            lg["xlsx_dirty"] = True       # 承認欄を「承認」に戻す（次に書き直すとき）
    return changed, stale


def revert_rejected(lg, mp, year, month, frames, log=print, al=None):
    """却下でまだ戻していない記録を、その科の共有マスターの今の予定から打ち消して書き戻す（科ごとに1回の put_dept）。
    台帳の state をその版に合わせる（戻したこと自体は記録にならない）。競合・書けないときは次の回に。戻り: 戻した記録の数。
    20260925: 打ち消しが当てはまらない記録（今の予定が申請のときと違う＝同じマスに後の申請が出ている等）は書かずに警報。"""
    todo = defaultdict(list)
    for r in lg.get("records") or []:
        if r.get("approval") == REJECT and not r.get("reverted"):
            todo[r["owner"]].append(r)
    if not todo:
        return 0
    try:
        master = MST.load(mp, year, month)
    except Exception as e:
        log(f"⚠ 却下を戻せない（マスターが読めない: {e}）→ 次の回に")
        return 0
    fmap = {f["frame_id"]: f for f in frames}
    done = 0
    for owner, recs in sorted(todo.items()):
        ent = MST.get_dept(master, owner) or {}
        st = lg["state"].get(owner) or {}
        cur = int(ent.get("rev") or 0)
        if cur != int(st.get("rev") or 0):
            log(f"⚠ {owner}: マスターの rev {cur} が台帳の rev {st.get('rev')} と違う（提出が届いたばかり）→ 却下の戻しは次の回に")
            continue
        duties = copy.deepcopy(ent.get("duties") or [])
        like = {}
        for d in list(ent.get("baseline") or []) + duties:
            like.setdefault(d.get("frame_id"), d)
        ops = _duty_ops(fmap, like)
        ok = []
        for r in sorted(recs, key=lambda x: -x["seq"]):     # 新しい順に打ち消す。当てはまらない記録は当てない
            trial = copy.deepcopy(duties)
            if _apply(trial, r, True, **ops) == 0:
                duties = trial
                ok.append(r)
            else:
                msg = (f"{owner} 記録 {r['seq']}（{r['date']} {r.get('band')} {r['kind']} {r['from']}→{r['to']}）の却下を"
                       "元に戻せない（今の予定が申請のときと違う。同じマスに後の申請が出ている等）→ 予定は書き換えていない。"
                       "管理者が確かめる")
                log("⚠ " + msg)
                _alert(al, msg)
        miss = len(recs) - len(ok)
        recs = ok
        if not recs:
            continue
        try:
            _m, e2 = MST.put_dept(mp, owner, duties=duties, by=REVERT_BY, expect_rev=cur, year=year, month=month,
                                  retries=1)
        except (MST.Conflict, MST.Busy, OSError) as ex:
            log(f"⚠ {owner}: 却下の戻しを書けない（{type(ex).__name__}: {ex}）→ 次の回に")
            continue
        now = _now_iso()
        for r in recs:
            r["reverted"], r["reverted_at"] = True, now
        st["rev"], st["duties"] = int(e2["rev"]), _compact(e2.get("duties") or [])
        done += len(recs)
        log(f"{owner}: 却下 {len(recs)} 件を元に戻した（マスター rev {e2['rev']}）"
            + (f"・当てはまらない {miss} 件は戻さなかった" if miss else ""))
    return done


def _cell_of(r):
    return (r["date"], r.get("band") or "", r["frame_id"])


def _persons(r):
    return {_pkey(r.get(s + "_login") or "", r.get(s) or "") for s in ("from", "to") if r.get(s) or r.get(s + "_login")}


def _earlier_on_cell(records, r, pred):
    """r より前（seq が小さい）で、同じマス・同じ人がかかわる記録のうち pred に合う最も新しいもの（無ければ None）。"""
    mine = _persons(r)
    hit = None
    for e in records or []:
        if e["seq"] < r["seq"] and _cell_of(e) == _cell_of(r) and (_persons(e) & mine) and pred(e):
            if hit is None or e["seq"] > hit["seq"]:
                hit = e
    return hit


def _fits(rows, recs):
    """recs（承認した記録。seq の順）が rows に順に全部当てはまるか（rows は変えない）。"""
    trial = [list(x) for x in rows]
    return all(_apply(trial, x, False, **COMPACT_OPS) == 0 for x in recs)


def compute_approved(lg, al=None, log=None):
    """approved_pending = 元の割当（baseline）＋承認した記録（seq の順）。15時にこれを DTX14 に書き、書けたら
    publish_approved が approved（手当計算 が読む）と reflected_at を付ける（20260925: 以前はここで付けていた）。
    当てはまらない操作が1つでもある承認の記録は丸ごと当てず、hold（保留）にする（同じマスの先の申請が未決など。
    当てると A と C が両方立つ二重割当になる）。戻り: 保留にした記録のリスト。
    20260926: 反映済みの記録は保留にしない（reflected_at を外さない）。まだ反映していない承認を当てると、後の反映済みの
    記録が当てはまらなくなるときは、まだ反映していない方を保留にする。反映済みの記録がそれでも当てはまらなければ、
    当てずに反映済みのまま管理者へ警報（失敗）。"""
    log = log or (lambda s: None)
    out = {o: [list(x) for x in rows] for o, rows in (lg.get("baseline") or {}).items()}
    applied, held = [], []
    now = _now_iso()
    recs = sorted(lg.get("records") or [], key=lambda x: x["seq"])
    for r in recs:
        if r.get("approval") != APPROVE:
            r.pop("hold", None)
            continue
        rows = out.setdefault(r["owner"], [])
        trial = [list(x) for x in rows]
        fits = _apply(trial, r, False, **COMPACT_OPS) == 0
        after = None
        if fits and not r.get("reflected_at"):
            later = [x for x in recs if x["seq"] > r["seq"] and x["owner"] == r["owner"]
                     and x.get("approval") == APPROVE and x.get("reflected_at")]
            if later and not _fits(trial, later) and _fits(rows, later):
                fits = False           # 当てると後の反映済みの記録が当てはまらなくなる → こちらを保留
                after = next((x for i, x in enumerate(later) if not _fits(trial, later[:i + 1])), later[0])
        if fits:
            out[r["owner"]] = trial
            applied.append(r["seq"])
            r.pop("hold", None)
            continue
        if r.get("reflected_at"):
            r.pop("hold", None)
            msg = (f"記録 {r['seq']}（{r['owner']} {r['date']} {r.get('band')} {r['kind']} {r['from']}→{r['to']}）は反映済みだが、"
                   "承認済みの割当に当てはまらない → 当てずにおく（反映済みの印は残す）。台帳と DTX14 を管理者が確かめる")
            log("⚠ " + msg)
            _alert(al, msg, fail=True)
            continue
        if after is not None:
            r["hold"] = {"at": (r.get("hold") or {}).get("at") or now, "blocked_by": after["seq"],
                         "blocker_undecided": False, "after_reflected": True}
        else:
            ok = set(applied)
            b = _earlier_on_cell(lg.get("records"), r, lambda e: e["seq"] not in ok and not e.get("reverted"))
            r["hold"] = {"at": (r.get("hold") or {}).get("at") or now, "blocked_by": b["seq"] if b else None,
                         "blocker_undecided": bool(b) and not b.get("approval")}
        held.append(r)
    for o in lg.get("state") or {}:
        out.setdefault(o, [])
    lg["approved_pending"] = {o: sorted(v) for o, v in out.items()}
    lg["approved_pending_sig"] = approval_sig(lg)
    lg["approved_pending_applied"] = applied
    return held


def approval_sig(lg):
    """approved_pending を作ったときの承認の並び（すぐ載せる回で承認→空欄に変わったら15時の回が作り直す目印）。"""
    return ",".join(str(r["seq"]) for r in lg.get("records") or [] if r.get("approval") == APPROVE)


class ApprovedMaster(dict):
    """approved_master の戻り（マスターの形の dict）。kotai_snapshot に、書いた割当（publish_approved に渡す）を持つ。"""
    kotai_snapshot = None


def pending_snapshot(md, year, month):
    """これから DTX14 に書く承認済みの割当（approved_pending）の写し。publish_approved に渡す。
    台帳が無い・読めない・まだ作っていないときは LedgerError。"""
    lp = os.path.join(md, "集約", LEDGER_NAME)
    if not os.path.isfile(lp):
        raise LedgerError(f"{lp} がありません")
    lg = load_ledger(lp, year, month)
    ap = lg.get("approved_pending")
    applied = lg.get("approved_pending_applied")
    if not isinstance(ap, dict):          # 20260925 より前の台帳（15時の回がまだ作り直していない）
        ap = lg.get("approved")
        applied = [r["seq"] for r in lg.get("records") or [] if r.get("approval") == APPROVE and r.get("reflected_at")]
    if not lg.get("confirmed_at") or not isinstance(ap, dict):
        raise LedgerError(f"{lp} に承認済みの状態（approved）がまだありません")
    return {"confirmed_at": lg["confirmed_at"], "approved": copy.deepcopy(ap), "applied": list(applied or []),
            "sig": lg.get("approved_pending_sig") or approval_sig(lg)}


def approved_master(md, year, month, master, frames=None, conf=_UNSET):
    """決裁済みの月: master の写しの各科の duties を、承認済みの割当（approved_pending＝元の割当＋承認した交代で、
    これから DTX14 に書くもの）で作り直して返す（ApprovedMaster。kotai_snapshot を publish_approved に渡す）。
    未確定の月は master をそのまま返す。conf＝確定の正本（effective_confirmed）。渡さなければマスターの confirmed。
    台帳が無い・読めない・approved がまだ無いときは LedgerError（生のマスターに落とさない。管理者 20260924）。
    ⚠ 手当計算 の Dr.JOY 投入は、台帳の approved（DTX14 に書けたもの）を読むこと。state（各科が出した今の予定。
      承認待ち・却下を含む）や approved_pending（まだ書けていないかもしれない）は読まない。"""
    if conf is _UNSET:
        conf = master.get("confirmed")
    if not conf:
        return master
    snap = pending_snapshot(md, year, month)
    fmap = {f["frame_id"]: f for f in (frames if frames is not None else C.load_frames())}
    out = ApprovedMaster(copy.deepcopy(master))
    out.kotai_snapshot = snap
    approved = snap["approved"]
    depts = out.setdefault("depts", {})
    for owner in sorted(set(depts) | set(approved)):
        ent = depts.setdefault(owner, {"duties": [], "roster": []})
        pool, like, pos = {}, {}, {}
        src = list(ent.get("duties") or []) + list(ent.get("baseline") or [])
        for i, d in enumerate(src):
            if d.get("frame_id"):
                k = (d.get("date"), d.get("band") or "", d["frame_id"], _duty_key(d))
                pool.setdefault(k, d)
                pos.setdefault(k, i)
                like.setdefault(d["frame_id"], d)
        # 20261024: 並びは各科が出した順（今の duties → baseline）に戻す。台帳の approved は並べ替えた形（compact）なので、
        # そのまま並べるとマスの中の氏名の順が変わり、changed_cells が交代の無い 227 マスを赤字にした（R8.11 確定後の初回）。
        # 交代で入った人は、同じマスで抜けた人の位置を継ぐ（無ければそのマスの後ろ）
        rows = []
        for date, band, fid, login, name, oc, *rest in approved.get(owner, []):
            k = (date, band, fid, _pkey(login, name))
            d = pool.get(k)
            d = copy.deepcopy(d) if d else make_duty(date, band, fid, login, name, oc, rest[0] if rest else "",
                                                     fmap.get(fid), like.get(fid))
            d["oncall"] = bool(oc)
            rows.append((k, d))
        kept = {k for k, _d in rows}
        freed = defaultdict(list)                        # マス → 抜けた人の位置
        for k, i in sorted(pos.items(), key=lambda kv: kv[1]):
            if k not in kept and i < len(ent.get("duties") or []):
                freed[k[:3]].append(i)
        order = []
        for n, (k, d) in enumerate(rows):
            if k in pos:
                order.append((pos[k], 0, n, d))
            elif freed[k[:3]]:
                order.append((freed[k[:3]].pop(0), 1, n, d))
            else:
                order.append((len(src), 1, n, d))
        ent["duties"] = [d for *_x, d in sorted(order, key=lambda t: t[:3])]
    return out


def order_drift(master, am):
    """見張り（20261025 管理者『セカンド以下の氏名の並びは重要。壊れないように』）: 承認済みの割当（am）の各マスの氏名の並びが、
    各科が出した並び（master の duties）と違うマス。顔ぶれが違うマス（承認待ち・却下・交代）は比べない。
    戻り: [(科, 日付, 昼夜, 枠)]。"""
    def cells(duties):
        out = defaultdict(list)
        for d in duties or []:
            out[(d.get("date"), d.get("band") or "", d.get("frame_id"))].append(_duty_key(d))
        return out
    bad = []
    for owner, ent in sorted((am.get("depts") or {}).items()):
        mine = cells(((master.get("depts") or {}).get(owner) or {}).get("duties"))
        for k, seq in cells(ent.get("duties")).items():
            src = mine.get(k)
            if src and sorted(src) == sorted(seq) and src != seq:
                bad.append((owner,) + k)
    return bad


def xlsx_sig(lg):
    """xlsx を作り直すかの目印: 確定・記録の数・承認と状態。"""
    st = [(r["seq"], r.get("approval") or "", bool(r.get("reflected_at")), bool(r.get("reverted")),
           r.get("approval_note") or "", (r.get("hold") or {}).get("blocked_by", 0) if r.get("hold") else -1)
          for r in lg.get("records") or []]
    h = hashlib.sha1(json.dumps(st, ensure_ascii=False).encode("utf-8")).hexdigest()[:16]
    return f"{lg.get('confirmed_at')}|{len(lg.get('records') or [])}|{h}"


# ---------------------------------------------------------------- 表示用の行

# 入力者の氏名と所属を引く先（20260929）。AD はこの PC（管理側）のドメインで引く。テストでは AD_LOOKUP=False にする
AD_LOOKUP = True
PEOPLE_CACHE = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "people_cache.json")
PEOPLE_CACHE_DAYS = 7
AD_TIMEOUT = 120               # Outlook が起動していないと、画面なしで立ち上げるのに時間がかかる（20261001）
PEOPLE_MISS_HOURS = 6          # 引けなかった人は6時間だけ覚える（AD が一時的に引けなかったときに7日待たない）
# AD と Outlook は別々の PowerShell で呼ぶ（20261001: 同じ処理の中でドメインへの問い合わせのあとに Outlook の宛先を引くと、
# Python の子から呼んだときに止まった。分ければ1秒で返る）
_AD_PS = r"""
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$out = @{}
foreach ($u in @(__LOGINS__)) {
  try {
    $s = [adsisearcher]"(sAMAccountName=$u)"
    [void]$s.PropertiesToLoad.AddRange(@("displayName","department","mail","userPrincipalName"))
    $r = $s.FindOne()
    if ($r) { $out[$u] = @{ name = [string]$r.Properties["displayname"][0]; dept = [string]$r.Properties["department"][0];
                            mail = [string]$r.Properties["mail"][0]; upn = [string]$r.Properties["userprincipalname"][0] } }
  } catch {}
}
$out | ConvertTo-Json -Compress
"""
# Entra の表示名 = Outlook のアドレス帳（Exchange Online）の表示名。まずログイン名で引き、別名（Alias）がログイン名と同じ人だけ採る
# （20261001: 管理者のアカウントは主アドレスが admin@ で、user003@ では Exchange の利用者として引けないが、別名は user003）。
# だめならメールアドレス（AD の mail・UPN・ログイン名@example.ac.jp）で引く
_OL_PS = r"""
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$out = @{}
$ns = (New-Object -ComObject Outlook.Application).GetNamespace("MAPI")
foreach ($item in @(__ITEMS__)) {
  $parts = $item.Split("|"); $lg = $parts[0]
  try {
    $rc = $ns.CreateRecipient($lg)
    if ($rc.Resolve()) { $x = $rc.AddressEntry.GetExchangeUser()
      if ($x -and $x.Name -and ([string]$x.Alias).ToLower() -eq $lg) { $out[$lg] = @{ name = [string]$x.Name; dept = [string]$x.Department; via = "alias" }; continue } }
  } catch {}
  foreach ($a in $parts[1..($parts.Count - 1)]) {
    if (-not $a) { continue }
    try {
      $rc = $ns.CreateRecipient($a)
      if ($rc.Resolve()) { $x = $rc.AddressEntry.GetExchangeUser(); if ($x -and $x.Name) { $out[$lg] = @{ name = [string]$x.Name; dept = [string]$x.Department; via = $a }; break } }
    } catch {}
  }
}
$out | ConvertTo-Json -Compress
"""


def _run_ps(script):
    import subprocess
    r = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
                       capture_output=True, timeout=AD_TIMEOUT, stdin=subprocess.DEVNULL, creationflags=0x08000000)
    s = r.stdout.decode("utf-8", "replace").strip()
    return json.loads(s) if s.startswith("{") else {}


def ad_people(logins):
    """Windows のログイン名 → {login: {"name", "dept", "entra", "entra_dept"}}。name・dept = Active Directory、
    entra = Entra の表示名（Outlook のアドレス帳。AD の mail・UPN で引く。20261001）。引けなければ空。"""
    logins = [x for x in logins if re.fullmatch(r"[A-Za-z0-9._-]+", x or "")]
    if not logins:
        return {}
    try:
        ad = {k.lower(): v for k, v in _run_ps(_AD_PS.replace("__LOGINS__", ",".join(f"'{x}'" for x in logins))).items()}
    except Exception:
        ad = {}
    addrs = {}
    for lg in logins:
        v = ad.get(lg) or {}
        for a in (v.get("mail"), v.get("upn"), f"{lg}@example.ac.jp"):
            if a and re.fullmatch(r"[A-Za-z0-9._@-]+", a):
                addrs.setdefault(lg, []).append(a.lower())
    try:
        items = ",".join("'" + "|".join([lg] + addrs.get(lg, [])) + "'" for lg in logins)
        ol = {k.lower(): v for k, v in _run_ps(_OL_PS.replace("__ITEMS__", items)).items()}
    except Exception:
        ol = {}
    out = {}
    for lg in logins:
        v = dict(ad.get(lg) or {})
        hit = ol.get(lg)
        if hit:
            v["entra"], v["entra_dept"] = hit.get("name") or "", hit.get("dept") or ""
        v = {k: x for k, x in v.items() if k in ("name", "dept", "entra", "entra_dept") and x}
        if v:
            out[lg] = v
    return out


def _dept_disp(dept):
    """'/循環器内科/集中治療部/' → '循環器内科・集中治療部'。"""
    return "・".join(x for x in str(dept or "").split("/") if x.strip())


class Names:
    r"""入力者（by＝「ログイン名 / PC名」）の表示 = 「氏名（所属部署）」（20260929 管理者: ログイン名・PC名は出さない。所属部署を出す）。
    氏名は Entra の表示名（Outlook のアドレス帳。20261001 管理者）、引けなければ下の順。
    所属の引く順（20260929 管理者「Dr.JOY の所属で大丈夫」）: Dr.JOY の全職員の一覧（Downloads の staffs*.csv）→ 医師の名簿
    （名簿\staff_*.csv）→ Dr.JOY に載っていない人（事務の一部など）だけ Active Directory の所属。
    引けなければ by のまま。AD の結果はこの PC に7日ためる（PEOPLE_CACHE）。
    resolve(by) → (表示, 引けたか)。記録したときの表示は台帳の by_disp に残す（あとで異動しても当時の所属が分かる）。"""

    def __init__(self, ad=None):
        self.ad = ad
        self.map = None

    def _load(self):
        self.map = {}
        for src in (self._doctors, self._drjoy_all):          # 後に読んだものが勝つ（全職員の一覧 > 医師の名簿）
            try:
                self.map.update(src())
            except Exception:
                pass

    @staticmethod
    def _doctors():
        out = {}
        p = C.newest_staffs()
        if p:
            for x in C.load_staffs(p)[0]:
                if x.get("login"):
                    out[x["login"].split("@")[0].lower()] = {"name": x["name"], "dept": _dept_disp(x.get("dept"))}
        return out

    @staticmethod
    def _drjoy_all():
        import make_staff_csv as MSC
        out = {}
        p = MSC.newest_raw()
        if p:
            for x in C.load_staffs(p)[0]:
                if x.get("login"):
                    out[x["login"].split("@")[0].lower()] = {"name": x["name"], "dept": _dept_disp(x.get("dept"))}
        return out

    def prefetch(self, bys):
        """これから引く入力者（by の並び）の AD・Entra を1回の問い合わせでまとめて引き、この PC にためる（20261001:
        1人ずつだと Outlook を立ち上げる待ちが人数分重なった）。"""
        logins = sorted({m.group(0).lower() for m in (re.match(r"[A-Za-z0-9._-]+", (b or "").strip()) for b in bys) if m})
        try:
            with open(PEOPLE_CACHE, encoding="utf-8") as f:
                cache = json.load(f)
        except (OSError, ValueError):
            cache = {}
        now = time.time()
        todo = [x for x in logins if x not in ("app",) and not (
            cache.get(x) and now - cache[x].get("at", 0) < (PEOPLE_CACHE_DAYS * 86400 if (cache[x].get("name") or cache[x].get("entra")) else PEOPLE_MISS_HOURS * 3600))]
        if not todo or (self.ad is None and not AD_LOOKUP):
            return
        got = self.ad(todo) if self.ad is not None else ad_people(todo)
        for x in todo:
            cache[x] = dict(got.get(x) or {}, at=now)
        try:
            os.makedirs(os.path.dirname(PEOPLE_CACHE), exist_ok=True)
            with open(PEOPLE_CACHE, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False)
        except OSError:
            pass

    def _ad(self, login):
        cache = {}
        try:
            with open(PEOPLE_CACHE, encoding="utf-8") as f:
                cache = json.load(f)
        except (OSError, ValueError):
            pass
        hit = cache.get(login)
        now = time.time()
        if hit and now - hit.get("at", 0) < (PEOPLE_CACHE_DAYS * 86400 if (hit.get("name") or hit.get("entra")) else PEOPLE_MISS_HOURS * 3600):
            return hit if (hit.get("name") or hit.get("entra")) else None
        if self.ad is not None:
            got = self.ad([login])
        elif AD_LOOKUP:
            got = ad_people([login])
        else:
            return None
        v = got.get(login)
        cache[login] = dict(v or {}, at=now)
        try:
            os.makedirs(os.path.dirname(PEOPLE_CACHE), exist_ok=True)
            with open(PEOPLE_CACHE, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False)
        except OSError:
            pass
        return v if v and (v.get("name") or v.get("entra")) else None

    def resolve(self, by):
        by = (by or "").strip()
        if by in ("app", ""):
            return {"app": "（アプリ・旧版）", "": "（不明）"}[by], False
        m = re.match(r"[A-Za-z0-9._-]+", by)
        if not m:
            return by, False
        login = m.group(0).lower()
        if self.map is None:
            self._load()
        dj = self.map.get(login) or {}
        ad = self._ad(login) or {}
        # 氏名 = Entra の表示名（20261001 管理者「申請者名は Entra の表示名」）→ Dr.JOY → AD。所属 = Dr.JOY → Entra → AD
        name = ad.get("entra") or dj.get("name") or ad.get("name") or ""
        dept = dj.get("dept") or ad.get("entra_dept") or ad.get("dept") or ""
        name = re.sub(r"[\s\u3000]+", " ", str(name)).strip()
        dept = str(dept).strip()
        if not name:
            return by, False
        return (f"{name}（{dept}）" if dept else name), True

    def __call__(self, by):
        return self.resolve(by)[0]


def _frame_info(frames):
    info = {}
    for i, f in enumerate(frames):
        name = "管理当直" if f.get("dept") == "管理当直" else f.get("dept") or f.get("owner") or ""
        table = "ファースト" if f.get("role") == "first" else "セカンド以下"
        info[f["frame_id"]] = (name, table, f.get("order") or i)
    return info


def _oc_label(oc):
    return "" if oc is None else ("オンコール" if oc else "当直")


def _md(iso):
    try:
        d = dt.datetime.fromisoformat(iso)
        return f"{d.month}/{d.day}"
    except (TypeError, ValueError):
        return "?"


def status_of(r):
    """「状態」欄: 承認待ち／承認（15時に反映）／反映済み（M/D）／却下・元に戻した（M/D）／却下（戻せていない）。"""
    a = r.get("approval") or ""
    if a == APPROVE:
        return f"反映済み（{_md(r['reflected_at'])}）" if r.get("reflected_at") else "承認（15時に反映）"
    if a == REJECT:
        return f"却下・元に戻した（{_md(r.get('reverted_at'))}）" if r.get("reverted") else "却下（戻せていない）"
    return "承認待ち"


def _status_class(r, records):
    """状態の種類（行をまとめる目印にも使う）: (status_of, 待っている先の申請の seq, 保留か)。
    20260925: 同じマスの先の申請が承認待ちの申請は「先の申請（No.n）の承認待ち」、当てられずに保留した承認は
    「前の申請（No.n）が未決のため保留」（No. は表示のあとで決まるので、ここでは seq を持つ）。"""
    a = r.get("approval") or ""
    if a == APPROVE and r.get("hold"):
        h = r["hold"] or {}
        return ("保留", h.get("blocked_by"), "after" if h.get("after_reflected") else True)
    if not a:
        b = _earlier_on_cell(records, r, lambda e: not e.get("approval"))
        if b:
            return ("承認待ち", b["seq"], False)
    return (status_of(r), None, False)


def _status_text(cls, no_of):
    st, dep, hold = cls
    if hold == "after":
        return f"反映済みの申請（No.{no_of[dep]}）と重なるため保留" if dep in no_of else "反映済みの申請と重なるため保留"
    if hold:
        return f"前の申請（No.{no_of[dep]}）が未決のため保留" if dep in no_of else "前の申請が未決のため保留"
    if dep is not None:
        return f"先の申請（No.{no_of[dep]}）の承認待ち" if dep in no_of else "先の申請の承認待ち"
    return st


def display_rows(records, frames, names):
    """台帳の記録 → 表の行（記録順）。同じ提出で昼と夜が同じ変わり方をしたものは「昼夜」にまとめる
    （承認・状態・備考も同じときだけ。まとめた行の承認は両方の記録に効く）。"""
    fi = _frame_info(frames)
    rows = []
    by_ver = defaultdict(list)
    for r in records:
        by_ver[(r["at"], r["owner"], r["rev"])].append(r)
    for ver in sorted(by_ver, key=lambda v: min(x["seq"] for x in by_ver[v])):
        merged = {}
        for r in by_ver[ver]:
            name, table, order = fi.get(r["frame_id"], (r["frame_id"], "", 999))
            cls = _status_class(r, records)
            k = (r["date"], name, table, r["kind"], r["from"], r["to"], r["from_oc"], r["to_oc"],
                 r.get("approval") or "", cls, r.get("approval_note") or "")
            if k in merged and r["band"] not in merged[k]["bands"]:
                merged[k]["bands"].append(r["band"])
                merged[k]["seqs"].append(r["seq"])
                continue
            by = r.get("by_disp") or names(r["by"])      # 記録したときの「氏名（所属）」（20260929）
            if r.get("account_mismatch"):         # 入力者と版を保存したアカウントが違う（20260925）
                by = f"{by}［⚠保存: {r.get('saved_by_account')}］"
            merged[k] = {"date": r["date"], "bands": [r["band"]], "name": name, "table": table, "order": order,
                         "kind": r["kind"], "from": r["from"], "to": r["to"],
                         "from_oc": r["from_oc"], "to_oc": r["to_oc"], "at": r["at"], "by": by,
                         "seq": r["seq"], "seqs": [r["seq"]], "ver": ver, "approval": r.get("approval") or "",
                         "cls": cls, "note": r.get("approval_note") or "",
                         "locked": bool(r.get("reverted") or r.get("reflected_at"))}
        for x in merged.values():
            x["band"] = "昼夜" if set(x["bands"]) == {"昼", "夜"} else "".join(x["bands"])
        rows += sorted(merged.values(), key=lambda x: (x["date"], 0 if x["band"].startswith("昼") else 1, x["order"]))
    no_of = {}
    for i, x in enumerate(rows, 1):
        for q in x["seqs"]:
            no_of[q] = i
    for i, x in enumerate(rows, 1):
        x["no"] = i
        x["status"] = _status_text(x.pop("cls"), no_of)
        if x["kind"] in (KIND_TO_OC, KIND_TO_DUTY):
            x["oc"] = x["kind"]
        elif x["kind"] == KIND_SWAP and x["from_oc"] != x["to_oc"]:
            x["oc"] = f"{_oc_label(x['from_oc'])}→{_oc_label(x['to_oc'])}"
        else:
            x["oc"] = _oc_label(x["to_oc"] if x["to_oc"] is not None else x["from_oc"])
    return rows


# ---------------------------------------------------------------- xlsx

FONT = "BIZ UDPゴシック"
COLS = [("No", 5), ("記録日時", 19.5), ("当直日", 11), ("時間帯", 7), ("表", 12.5), ("当直名", 18),
        ("区分", 17.5), ("元の勤務者", 15), ("", 3), ("新しい勤務者", 15), ("当直／オンコール", 16), ("申請者", 26),
        (APPROVE, 9), ("状態", 25), ("備考", 24), (SEQ_HEADER, 9)]     # 承認・備考＝事務の入力欄／記録番号＝隠し列
C_APPROVE, C_STATUS, C_NOTE, C_SEQ = 13, 14, 15, 16
APPROVE_FILL = "FFF2B3"          # 承認欄（入力する所）の淡い黄色
KIND_FILL = {KIND_SWAP: "DDEBF7", KIND_ADD: "E2EFDA", KIND_DROP: "FCE4E4", KIND_TO_OC: "FFF2CC", KIND_TO_DUTY: "FFF2CC"}


def _fmt_at(iso):
    try:
        return dt.datetime.fromisoformat(iso).strftime("%Y/%m/%d %H:%M")
    except (TypeError, ValueError):
        return iso or ""


def write_xlsx(path, year, month, ledger, rows, holidays):
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    from openpyxl.styles import Protection
    from openpyxl.worksheet.datavalidation import DataValidation

    ncol = len(COLS)
    last = get_column_letter(ncol)
    thin = Side(style="thin", color="BFBFBF")
    wb = openpyxl.Workbook()

    def sheet(ws, title, data):
        ws.title = title
        ws.sheet_view.showGridLines = False
        for i, (_, w) in enumerate(COLS, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws["A1"] = f"当直交代記録　令和{year - 2018}年{month}月分（{year}年{month}月）"
        ws["A1"].font = Font(name=FONT, size=16, bold=True)
        ws.row_dimensions[1].height = 26
        hr = 3
        for i, (h, _) in enumerate(COLS, 1):
            c = ws.cell(hr, i, h)
            c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="7F6000" if i in (C_APPROVE, C_NOTE) else "404040")
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(C_SEQ)].hidden = True
        ws.row_dimensions[hr].height = 22
        if not data:
            r = hr + 1
            ws.cell(r, 1, "記録はまだありません。")
            ws.merge_cells(f"A{r}:{last}{r}")
            c = ws.cell(r, 1)
            c.font = Font(name=FONT, size=10, italic=True, color="7F7F7F")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r].height = 30
        prev_ver, shade = None, False
        for j, x in enumerate(data):
            r = hr + 1 + j
            if x["ver"] != prev_ver:
                shade, prev_ver = not shade, x["ver"]
            d = dt.date.fromisoformat(x["date"])
            same = x["kind"] in (KIND_TO_OC, KIND_TO_DUTY)        # 同じ人の当直⇔オンコール
            vals = [x["no"], _fmt_at(x["at"]), f"{d.month}/{d.day}（{WD[d.weekday()]}）", x["band"], x["table"],
                    x["name"], x["kind"], x["from"], "→" if x["from"] and x["to"] else "",
                    "（同じ人）" if same else x["to"], x["oc"], x["by"],
                    x.get("approval") or None, x.get("status") or "承認待ち", x.get("note") or None,
                    ",".join(str(q) for q in x.get("seqs") or [x["seq"]])]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v)
                c.font = Font(name=FONT, size=10.5)
                c.alignment = Alignment(vertical="center", horizontal="center" if i in (1, 3, 4, 9) else "left")
                c.border = Border(bottom=thin)
                if shade:
                    c.fill = PatternFill("solid", fgColor="F7F7F7")
            ws.cell(r, 7).fill = PatternFill("solid", fgColor=KIND_FILL.get(x["kind"], "FFFFFF"))
            ws.cell(r, 7).font = Font(name=FONT, size=10.5, bold=True)
            ws.cell(r, 7).alignment = Alignment(horizontal="center", vertical="center")
            if C.is_holiday(d, holidays):
                ws.cell(r, 3).font = Font(name=FONT, size=10.5, bold=True,
                                          color="1F4E79" if d.weekday() == 5 and d not in holidays else "C00000")
            ws.cell(r, 8).font = Font(name=FONT, size=10.5, color="595959")
            ws.cell(r, 10).font = Font(name=FONT, size=10.5, bold=not same, color="7F7F7F" if same else "000000")
            ws.cell(r, 2).font = Font(name=FONT, size=9.5, color="595959")
            ws.cell(r, 12).font = Font(name=FONT, size=9.5, color="595959")
            # 承認欄（事務の入力。元に戻した却下・反映済みの承認はもう変えられないので灰色で鍵をかける＝見た目だけ）・状態・備考・記録番号
            ca = ws.cell(r, C_APPROVE)
            ca.font = Font(name=FONT, size=10.5, bold=True, color="C00000" if x.get("approval") == REJECT else "000000")
            ca.alignment = Alignment(horizontal="center", vertical="center")
            ca.fill = PatternFill("solid", fgColor="D9D9D9" if x.get("locked") else APPROVE_FILL)
            ca.protection = Protection(locked=bool(x.get("locked")))
            st = x.get("status") or ""
            ws.cell(r, C_STATUS).font = Font(name=FONT, size=10, bold=st.startswith("承認待ち"),
                                             color="375623" if st.startswith(("反映済み", "承認（")) else
                                             "C00000" if st.startswith("却下") else "7F6000")
            ws.cell(r, C_NOTE).protection = Protection(locked=False)
            ws.cell(r, C_NOTE).alignment = Alignment(vertical="center", wrap_text=True)
            ws.cell(r, C_SEQ).font = Font(name=FONT, size=9, color="A6A6A6")
            ws.row_dimensions[r].height = 19
        end = hr + max(1, len(data))
        if data:
            ws.auto_filter.ref = f"A{hr}:{last}{end}"
            al = get_column_letter(C_APPROVE)
            dv = DataValidation(type="list", formula1=f'"{APPROVE},{REJECT}"', allow_blank=True,
                                showErrorMessage=True, errorTitle="承認欄",
                                error="「承認」か「却下」を選んでください（空欄＝承認待ち）",
                                showInputMessage=True, promptTitle="承認欄",
                                prompt="承認＝15時に当直表へ反映／却下＝15時に科の予定を元に戻す")
            dv.add(f"{al}{hr + 1}:{al}{end}")
            ws.add_data_validation(dv)
        # シートの保護はしない（管理者 20260924: 事務担当・事務担当2はうっかり書き換えない。並べ替えを使えるほうがよい）。
        #   申請の中身を書き換えられても、次の書き直しで台帳の内容に戻る
        ws.freeze_panes = f"A{hr + 1}"
        ws.print_title_rows = f"{hr}:{hr}"
        ws.print_area = f"A1:{last}{end}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top, ws.page_margins.bottom = 0.5, 0.6
        ws.oddFooter.center.text = "&P / &N ページ"
        ws.oddFooter.right.text = f"当直交代記録 R{year - 2018}.{month}月"

    sheet(wb.active, SHEET, rows)
    wb.properties.title = f"当直交代記録 R{year - 2018}.{month}月"
    wb.properties.creator = "当直勤務表アプリ（kotai_record.py）"
    tmp = os.path.join(os.path.dirname(path), "~kotai_" + os.path.basename(path))
    wb.save(tmp)
    try:
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


# ---------------------------------------------------------------- 月ごと

def out_dir_of(year, month):
    """xlsx の置き場所 = SharePoint 同期フォルダの 当直勤務表_正本\\R8.N月 当直勤務表\\（DTX14 と同じ）。"""
    return os.path.join(DTX_ROOT, f"R{reiwa(year)}.{month}月 当直勤務表")


def _xlsx_stat(xp):
    try:
        st = os.stat(xp)
        return {"mtime": st.st_mtime, "size": st.st_size}
    except OSError:
        return None


def xlsx_busy(lg, xp, now=None):
    """xlsx を今は書き直さないほうがよいか（20260925）: 自分が最後に書いたあとに他で変わり（更新時刻か大きさが違う）、
    しかもその更新が RECENT_XLSX 秒以内 → True（事務が入力中／別の PC・Excel Online の入力がまだ同期しきっていない恐れ。
    ここで作り直すと入力が消えるか OneDrive の競合コピーになる）。ファイルが無ければ False。"""
    cur = _xlsx_stat(xp)
    if cur is None:
        return False
    mine = lg.get("xlsx_stat") or {}
    if mine and abs(float(mine.get("mtime") or 0) - cur["mtime"]) < 1e-3 and mine.get("size") == cur["size"]:
        return False
    return (now or time.time()) - cur["mtime"] < RECENT_XLSX


def _sync_approvals(lg, xp, frames, mlog, al):
    """承認欄を台帳へ（作り直す前に必ず読む＝書いてあった承認を消さない）。戻り: (kind, 変わった数, 食い違い)。"""
    if not lg.get("records"):
        return "none", 0, 0
    kind, rows = read_approvals(xp, mlog)
    ch, st = apply_approvals(lg, rows, mlog, frames=frames, al=al)
    return kind, ch, st + (1 if kind == "old" else 0)


def _write_xlsx_guarded(lg, lp, xp, out_dir, year, month, frames, holidays, names, mlog, quick):
    """xlsx を作り直し、台帳の xlsx_sig・xlsx_stat を更新して保存する。
    戻り: 'written' / 'no-folder' / 'locked'（開かれていて書けない）/ 'deferred'（ほかで更新されたばかり）。
    書けなかったときは xlsx_dirty を立てる（すぐ載せる回がやり直す）。"""
    if not os.path.isdir(out_dir):
        return "no-folder"          # SharePoint の月フォルダがまだ無い（作らない）
    if xlsx_busy(lg, xp):
        lg["xlsx_dirty"] = True
        _save_json(lp, lg)
        if not quick:
            mlog(f"{FILE_NAME} は {RECENT_XLSX // 60} 分以内にほかで更新されている（入力中・同期待ち）→ 今は書き直さない（次の回に）")
        return "deferred"
    holidays = holidays if holidays is not None else C.load_holidays()
    names = names or Names()
    resolve = getattr(names, "resolve", None)
    if hasattr(names, "prefetch"):
        try:
            names.prefetch([r.get("by") for r in lg["records"] if not r.get("by_disp")])
        except Exception:
            pass
    for r in lg["records"] if resolve else ():   # 記録したときの入力者の「氏名（所属）」を残す（引けたときだけ。20260929）
        if not r.get("by_disp"):
            disp, ok = resolve(r.get("by"))
            if ok:
                r["by_disp"] = disp
    rows = display_rows(lg["records"], frames, names)
    try:
        write_xlsx(xp, year, month, lg, rows, holidays)
    except OSError as ex:            # 20260926: PermissionError に限らない（同期フォルダの一時的な不調など）。月の失敗にしない
        lg["xlsx_dirty"] = True
        _save_json(lp, lg)
        if not quick:                 # すぐ載せる回は黙って次の回に（10分ごとにログを積まない）
            mlog(f"{FILE_NAME} が開かれていて書けない（{type(ex).__name__}: {ex}。次の回に書く）")
        return "locked"
    lg["xlsx_sig"], lg["xlsx_nrec"], lg["xlsx_stat"] = xlsx_sig(lg), len(lg["records"]), _xlsx_stat(xp)
    lg["xlsx_dirty"] = False
    _save_json(lp, lg)
    mlog(f"{FILE_NAME} を書いた（{len(rows)} 行{'・すぐ載せる回' if quick else ''}）")
    return "written"


def _merge_alerts(dst, src):
    """警報を足す（同じ文面は1つに）。"""
    seen = {a.get("text") for a in dst}
    for a in src or []:
        if a.get("text") not in seen:
            dst.append(a)
            seen.add(a.get("text"))


def _hold_alerts(lg, held, frames, al):
    """保留にした承認（compute_approved）を警報に。先の申請が承認待ちなら事務にもやさしい言葉で知らせる。"""
    rows = display_rows(lg["records"], frames, lambda b: b)
    no_of = {q: x["no"] for x in rows for q in x["seqs"]}
    for r in held:
        h = r.get("hold") or {}
        b = h.get("blocked_by")
        n, m = no_of.get(r["seq"]), no_of.get(b)
        text = (f"記録 {r['seq']}（No.{n}・{r['owner']} {r['date']} {r.get('band')} {r['kind']} {r['from']}→{r['to']}）は"
                "承認済みだが今の割当に当てはまらないので保留（当てると二重割当になる）"
                + (f"。同じマスの先の申請 No.{m}（記録 {b}）が{'承認待ち' if h.get('blocker_undecided') else '未反映'}"
                   if b else "。先の申請が見当たらない → 管理者が確かめる"))
        if h.get("after_reflected"):
            text = (f"記録 {r['seq']}（No.{n}・{r['owner']} {r['date']} {r.get('band')} {r['kind']} {r['from']}→{r['to']}）は"
                    f"承認済みだが、当てると反映済みの No.{m}（記録 {b}）が当てはまらなくなるので保留 → 管理者が確かめる")
        clerk = (f"No.{n} の申請は、同じ日の先の申請（No.{m}）が決まっていないので保留にしました。"
                 f"先に No.{m} を承認か却下してください。" if b and h.get("blocker_undecided") else None)
        _alert(al, text, clerk=clerk)


def run_month(md, year, month, *, out_dir=None, force=False, frames=None, holidays=None, names=None, log=print,
              quick=False, lock_wait=None, alerts=None, file_name=FILE_NAME):
    """1か月分: 台帳（md の 集約\\）を更新し、要れば xlsx（out_dir）を作り直す。
    quick=False（15時の回）: 承認欄を読む → 版を集めて記録を積む → 却下を戻す → approved_pending を作る → 台帳 → xlsx。
      approved（手当計算 が読む）と reflected_at は、DTX14 に書けたあとで publish_approved が付ける（20260925）。
    quick=True（すぐ載せる回。10分おき）: マスターが変わっていなければ何もしない。記録を積み、増えたときだけ
      承認欄を読んでから xlsx を作り直す。却下の戻し・approved は作らない。台帳が使用中なら黙って 'busy'。
    alerts = 警報を足すリスト（{"ym", "text", "clerk", "fail"}）。すぐ載せる回の警報は台帳に貯めて15時の回が渡す。
    戻り: 'skip' / 'same' / 'written' / 'locked' / 'deferred' / 'no-master' / 'no-folder' / 'ledger-error' /
          'truth-error' / 'busy' / 'xlsx-unreadable'。"""
    mp = MST.master_path(md, year, month)
    if not os.path.isfile(mp):
        return "no-master"
    tag = C.month_dir_name(year, month)

    def mlog(s):
        log(f"{tag}: {s}")
    out_dir = out_dir or out_dir_of(year, month)
    lp = os.path.join(md, "集約", LEDGER_NAME)
    xp = os.path.join(out_dir, file_name)
    if quick and not force and _quick_idle(lp, mp, year, month):
        return "skip"                 # 軽い下見（ロックも xlsx も触らない）
    al = []
    try:
        with ledger_lock(lp, wait=(0.0 if quick else LOCK_WAIT_FULL) if lock_wait is None else lock_wait):
            return _run_month_locked(md, year, month, mp, lp, out_dir, xp, force, frames, holidays, names,
                                     mlog, quick, al)
    except LedgerBusy as ex:
        if not quick:                 # すぐ載せる回は黙って次の回に
            mlog(f"✗ {ex} → この回は交代記録を更新しない")
        return "busy"
    finally:
        if alerts is not None:
            for a in al:
                a.setdefault("ym", (year, month))
            _merge_alerts(alerts, al)


def _truth_at(mp):
    """確定の正本の確定時刻（未確定は ""、確定.json が無ければ None）。壊れていれば TruthError。"""
    t = MST.read_truth(mp)
    if t is None:
        return None
    return (t.get("confirmed") or {}).get("at") or ""


def _quick_idle(lp, mp, year, month):
    """すぐ載せる回の下見: マスターが前の回のまま（更新時刻）、確定の正本も前の回のまま、xlsx に全部の記録が載っていて
    書き直し待ちも無い → True（何もしない）。"""
    try:
        lg = load_ledger(lp, year, month)
        return (lg.get("master_mtime") == os.path.getmtime(mp)
                and int(lg.get("xlsx_nrec") or 0) >= len(lg.get("records") or [])
                and not lg.get("xlsx_dirty")
                and _truth_at(mp) == lg.get("truth_at"))
    except (LedgerError, MST.TruthError, OSError, TypeError, ValueError):
        return False                  # 本番の処理でもう一度読んで判断する


def _revs_ok(lg, master):
    """台帳の state が、マスターの各科の rev まで処理済みか（未確定の台帳は True）。
    更新時刻だけで飛ばすと、読んでいる間に届いた提出を取りこぼすことがある（20260925）。"""
    if not lg.get("confirmed_at"):
        return True
    st = lg.get("state") or {}
    return all(int(e.get("rev") or 0) == int((st.get(o) or {}).get("rev") or 0)
               for o, e in (master.get("depts") or {}).items())


def _run_month_locked(md, year, month, mp, lp, out_dir, xp, force, frames, holidays, names, mlog, quick, al):
    try:
        lg = load_ledger(lp, year, month, xlsx=xp)
    except LedgerError as ex:         # 台帳にも xlsx にも触らない（人が直すまで毎回ここで止まる）
        mlog(f"✗ 台帳が読めないので触らない: {ex}")
        if not quick:
            _alert(al, f"交代記録の台帳が使えないので、この月は触らない（DTX14 も書かない）: {ex}", fail=True)
        return "ledger-error"
    frames = frames if frames is not None else C.load_frames()
    migrated = migrate(lg, mlog)
    mtime = os.path.getmtime(mp)      # 読む前に控える（読んでいる間に届いた提出を「処理済み」にしない。20260925）
    master = MST.load(mp, year, month)
    try:
        tat = _truth_at(mp)
        conf, probs = MST.effective_confirmed(mp, master)
    except MST.TruthError as ex:
        mlog(f"✗ 確定の正本が読めないので触らない: {ex}")
        return "truth-error"
    for level, text in probs:
        mlog(("⚠ " if level == "warn" else "✗ ") + text)
        if level == "alert":
            _alert(al, text)
    if quick:
        if (not force and lg.get("master_mtime") == mtime and int(lg.get("xlsx_nrec") or 0) >= len(lg["records"])
                and not lg.get("xlsx_dirty") and tat == lg.get("truth_at") and _revs_ok(lg, master)):
            return "skip"
        try:
            n = update_ledger(lg, master, md, mp, log=mlog, conf=conf, al=al)
        except LedgerNotReady:
            return "baseline-wait"    # 台帳は保存しない（15時の回が警報を出す）
        lg["master_mtime"], lg["truth_at"] = mtime, tat
        if n:
            mlog(f"交代の記録を {n} 件足した（計 {len(lg['records'])} 件・すぐ載せる回）")
        need = force or int(lg.get("xlsx_nrec") or 0) < len(lg["records"]) or bool(lg.get("xlsx_dirty"))
        kind = "none"
        if need:
            kind, _ch, _st = _sync_approvals(lg, xp, frames, mlog, al)
        _merge_alerts(lg.setdefault("alerts_pending", []), al)     # 15時の回がまとめて知らせる
        al.clear()
        _save_json(lp, lg)
        if not need:
            return "same"
        if kind == "bad":
            return "xlsx-unreadable"
        return _write_xlsx_guarded(lg, lp, xp, out_dir, year, month, frames, holidays, names, mlog, quick)
    # --- 15時の回
    if os.path.isdir(out_dir):        # 当直交代申請.xlsx の競合コピー（OneDrive。1日以内）
        cc = conflict_copies(out_dir, xp)
        if cc:
            text = (f"{FILE_NAME} の競合コピーがある（OneDrive）: {'、'.join(cc)}（{out_dir}）"
                    " → 承認欄の入力が写しの方に入っていないか確かめる")
            mlog("⚠ " + text)
            _alert(al, text, clerk=f"「{FILE_NAME}」の写し（{'、'.join(cc)}）ができています。"
                                   "2か所で同時に書いたときにできるものです。写しの方には書き込まず、管理者に知らせてください。")
    kind, ch, stale = _sync_approvals(lg, xp, frames, mlog, al)
    pending = any(r.get("approval") == REJECT and not r.get("reverted") for r in lg["records"])
    unreflected = any(r.get("approval") == APPROVE and not r.get("reflected_at") for r in lg["records"])
    if (not force and not ch and not stale and not pending and not unreflected and not migrated and kind != "bad"
            and lg.get("master_mtime") == mtime and tat == lg.get("truth_at") and _revs_ok(lg, master)
            and not lg.get("xlsx_dirty") and not lg.get("alerts_pending") and os.path.isfile(xp)
            and (not lg.get("confirmed_at") or not conf
                 or (isinstance(lg.get("approved_pending"), dict)
                     and lg.get("approved_pending_sig") == approval_sig(lg)))):
        return "skip"
    try:
        n = update_ledger(lg, master, md, mp, log=mlog, conf=conf, al=al)
    except LedgerNotReady as ex:      # 空の元の割当を焼かない（20260926）。台帳は保存せず次の回にやり直す
        text = (f"確定から交代記録の台帳を作れない（{ex}）→ 今回は作らない・この月の DTX14 は書かない。次の回にやり直す。"
                "続くなら管理者が共有マスターと 集約\\確定.json を確かめる")
        mlog("✗ " + text)
        _alert(al, text, fail=True)
        return "baseline-wait"
    lg["truth_at"] = tat
    if n:
        mlog(f"交代の記録を {n} 件足した（計 {len(lg['records'])} 件）")
    if lg.get("confirmed_at") and conf:
        revert_rejected(lg, mp, year, month, frames, mlog, al)
        held = compute_approved(lg, al, mlog)     # 保留になるのは反映していない承認だけ（反映済みの印は外さない。20260926）
        if held:
            mlog(f"⚠ 承認した申請のうち {len(held)} 件は当てはまらないので保留（同じマスの先の申請が未決など）")
            _hold_alerts(lg, held, frames, al)
    lg["master_mtime"] = mtime        # 却下の戻しで書いた分は次の回に rev で確かめる
    pend = lg.pop("alerts_pending", [])
    _merge_alerts(pend, al)
    al[:] = pend
    _save_json(lp, lg)
    if kind == "bad":
        return "xlsx-unreadable"
    if not force and not stale and xlsx_sig(lg) == lg.get("xlsx_sig") and os.path.isfile(xp) and not lg.get("xlsx_dirty"):
        return "same"
    return _write_xlsx_guarded(lg, lp, xp, out_dir, year, month, frames, holidays, names, mlog, quick)


def publish_approved(md, year, month, snap, *, out_dir=None, frames=None, holidays=None, names=None, log=print,
                     alerts=None, lock_wait=LOCK_WAIT_FULL):
    """その月の DTX14 に承認済みの割当（snap＝approved_master の kotai_snapshot）を書けたあとで呼ぶ（dtx14_daily_sync。20260925）。
    台帳の approved（手当計算 の Dr.JOY 投入が読む）を書いた割当にし、当てた承認の記録に reflected_at を付けて、
    xlsx の状態を「反映済み」に作り直す（書き直す前に承認欄を読む）。書き込みが失敗した月・書かない月では呼ばない。
    戻り: 反映済みにした記録の数。"""
    tag = C.month_dir_name(year, month)

    def mlog(s):
        log(f"{tag}: {s}")
    out_dir = out_dir or out_dir_of(year, month)
    lp = os.path.join(md, "集約", LEDGER_NAME)
    xp = os.path.join(out_dir, FILE_NAME)
    al = []
    try:
        with ledger_lock(lp, wait=lock_wait):
            lg = load_ledger(lp, year, month, xlsx=xp)
            if lg.get("confirmed_at") != snap.get("confirmed_at"):
                mlog("⚠ DTX14 に書いたあとで確定が付け直された → 反映済みの印は付けない（次の15時に）")
                return 0
            frames = frames if frames is not None else C.load_frames()
            _sync_approvals(lg, xp, frames, mlog, al)
            now = _now_iso()
            lg["approved"] = copy.deepcopy(snap["approved"])
            lg["approved_sig"] = snap.get("sig") or ""
            lg["approved_at"] = now
            applied = set(snap.get("applied") or [])
            n = 0
            for r in lg["records"]:
                if r["seq"] in applied and r.get("approval") == APPROVE and not r.get("reflected_at"):
                    r["reflected_at"] = now
                    n += 1
            lg["xlsx_dirty"] = True
            _save_json(lp, lg)
            if n:
                mlog(f"DTX14 に書けたので、承認した申請 {n} 件を反映済みにした（approved を更新）")
            _write_xlsx_guarded(lg, lp, xp, out_dir, year, month, frames, holidays, names, mlog, False)
            return n
    finally:
        if alerts is not None:
            for a in al:
                a.setdefault("ym", (year, month))
            _merge_alerts(alerts, al)


def month_dirs(root):
    out = []
    for name in sorted(os.listdir(root)):
        ym = C.parse_month_dir(name)
        if ym and os.path.isdir(os.path.join(root, name)):
            out.append((os.path.join(root, name), *ym))
    return out


def update_all(root=None, months=(), force=False, log=None, quick=False, alerts=None):
    """全部の月フォルダ（months を渡せばその月だけ）を更新する。1か月の失敗で他の月を止めない。
    quick=True = すぐ載せる回（10分おき）。本番の共有に届かなければ黙って何もしない（15時の回が警報を出す）。
    alerts = 警報を足すリスト（run_month）。
    戻り: {(年, 月): run_month の戻り（例外は 'error'）}。失敗は FAILED、知らせるものは WARNED に入っているもの。"""
    root = root or C.PROD_ROOT
    log = log or (lambda s: log_line(s, root))
    if quick and not os.path.isdir(root):
        return {}
    frames, holidays, names = C.load_frames(), C.load_holidays(), Names()
    res = {}
    for md, y, m in month_dirs(root):
        if months and (y, m) not in months:
            continue
        try:
            res[(y, m)] = run_month(md, y, m, force=force, frames=frames, holidays=holidays, names=names, log=log,
                                    quick=quick, alerts=alerts)
        except Exception as e:
            log(f"{C.month_dir_name(y, m)}: エラー {type(e).__name__}: {e}")
            res[(y, m)] = "error"
    if not months:
        # 試験用フォルダ（20261001）: 書き出しは TEST_FILE_NAME。警報は誰にも送らない（捨てる）。戻り値にも入れない
        #   （15時の反映が同じ年月の本番の月と取り違えないように）。承認しても DTX14・Dr.JOY には流れない
        #   （15時の反映は試験用フォルダを読まず、publish_approved も呼ばれないので approved は付かない）
        for md, y, m in test_month_dirs(root):
            try:
                r = run_month(md, y, m, force=force, frames=frames, holidays=holidays, names=names,
                              log=lambda s, n=os.path.basename(md): log(f"（試験用）{n}: {s}"),
                              quick=quick, alerts=[], file_name=TEST_FILE_NAME)
                if r not in ("skip", "same"):
                    log(f"（試験用）{os.path.basename(md)}: {r}")
            except Exception as e:
                log(f"（試験用）{os.path.basename(md)}: エラー {type(e).__name__}: {e}")
    return res


def test_month_dirs(root):
    """試験用の月フォルダ [(パス, 年, 月)]（同じ年月のふつうのフォルダが無いものだけ）。"""
    import set_acl as SA
    out = []
    try:
        names = sorted(os.listdir(root))
    except OSError:
        return out
    for name in names:
        ym = SA.parse_test_dir(name)
        if ym and os.path.isdir(os.path.join(root, name)) and not os.path.isdir(os.path.join(root, C.month_dir_name(*ym))):
            out.append((os.path.join(root, name), *ym))
    return out


def quick_watchdog(now=None, state_path=None, log=print):
    """すぐ載せる回（10分おき）から: 15時の反映が最後に終わってから 26 時間を超えていたら、管理者へ警報を1日1通だけ送る
    （PC が落ちていた・タスクが走らなかった日に気づくため。20260925）。戻り: 問題の一言（無ければ ""）。"""
    import dtx14_daily_sync as DS
    state_path = state_path or WATCHDOG_STATE
    problem = DS.check_last_run(now=now)
    if not problem:
        return ""
    today = (now or dt.datetime.now()).date().isoformat()
    try:
        with open(state_path, encoding="utf-8") as f:
            st = json.load(f)
    except (OSError, ValueError):
        st = {}
    if st.get("alerted") == today:
        return problem
    body = "\n".join(["当直勤務表の毎日15時の反映（タスク「Touchoku_DTX14_sync」）が、しばらく終わっていません。", "",
                      f"  {problem}", "",
                      "PC が止まっていた・ログオフしていた・タスクが無効になっている、などが考えられます。",
                      "そのあいだ DTX14 への反映、交代申請の承認の反映、却下の戻しは行われていません。",
                      f"ログ: {DS.LOCAL_LOG}", "",
                      "（このメールは10分おきのタスク「Touchoku_kotai_quick」が1日1通だけ送っています）"])
    try:
        DS.ALERT_SENDER(DS.ALERT_TO, "【警報】当直勤務表 15時の反映が動いていません", body)
    except Exception as e:
        log(f"✗ 見張りの警報メールを送れない: {type(e).__name__}: {e}")
        return problem
    try:
        os.makedirs(os.path.dirname(state_path), exist_ok=True)
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump({"alerted": today, "problem": problem}, f, ensure_ascii=False)
    except OSError:
        pass
    return problem


def main():
    ap = argparse.ArgumentParser(description="当直交代の公式記録（SharePoint の月フォルダの 当直交代申請.xlsx）を更新する")
    ap.add_argument("--month", action="append", default=[], help="YYYY-MM（省略時は全部の月フォルダ）")
    ap.add_argument("--force", action="store_true", help="変わっていなくても xlsx を作り直す")
    ap.add_argument("--quick", action="store_true",
                    help="すぐ載せる回（10分おきのタスク）: 新しい申請を載せるだけ。却下の戻し・反映は15時の回")
    ap.add_argument("--root", default=C.PROD_ROOT)
    a = ap.parse_args()
    if a.quick:
        try:
            quick_watchdog(log=lambda s: log_line(s, a.root))
        except Exception as e:        # 見張りの失敗で申請の取り込みを止めない
            log_line(f"✗ 15時の反映の見張りに失敗: {type(e).__name__}: {e}", a.root)
        try:                          # 所属の変更メール（月末と毎週水曜の9時。20260927 管理者）
            import dtx14_daily_sync as DS
            r = DS.morning_shozoku_notice()
            if r in ("sent", "gave-up"):
                log_line(f"所属の変更メール: {r}", a.root)
        except Exception as e:
            log_line(f"✗ 所属の変更メールに失敗: {type(e).__name__}: {e}", a.root)
        try:                          # 開かれていて15時に差し替えられなかった公開 PDF を、閉じられたら置く（20261007 管理者）
            import dtx14_daily_sync as DS
            DS.flush_pending_pdfs(log=lambda s: log_line(s, a.root))
        except Exception as e:
            log_line(f"✗ とっておいた PDF の置き直しに失敗: {type(e).__name__}: {e}", a.root)
    update_all(a.root, {tuple(int(x) for x in s.split("-")) for s in a.month}, a.force, quick=a.quick)


if __name__ == "__main__":
    main()
