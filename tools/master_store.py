# -*- coding: utf-8 -*-
r"""master_store.py  20261002
共有のマスター当直予定表（schema hospital.duty-master/1）の読み書き。
設計の正本 = docs/設計_当直表xlsx一本化_20260912.md §4

全科ぶんを1本の JSON に持つ。33科のアプリが同じファイルを書き換えるので、書くときは

  1. ロックを取る（排他作成のロックファイル。Node 側も同じ手順で取れる）
  2. 読み直す
  3. 自科のエントリだけ差し替える（他科は読み直した最新のまま）
  4. _履歴\ へコピーを取る（改名ではなく複製）
  5. 一時ファイル → 差し替えで上書き
  6. 読み直して自科の分が入っているか確かめる
  7. ロックを外す

の手順を通す。⚠ 2026-09-12 の実測（3プロセス×25回）で分かったこと:

  * ロック無しだと **読み手が書き換え途中の空ファイルを掴む**（JSONDecodeError が多発）
  * Windows は他プロセスが開いているファイルを os.replace できない（Permission denied）
  * `updated` の秒精度タイムスタンプは版の目印にならない（同じ秒に2回書けば区別できない）
    → 同時編集の検出は **per-dept の rev（単調増加の整数）** で行う

20260923（アプリ側 app/lib/store.cjs と揃えた）:
  * 「無い」と決めるのは os.stat が FileNotFoundError を返したときだけ（共有の一時的な不調で
    os.path.exists が False を返し、自科だけの新しいマスターで他の科を消しかねなかった）。
    書く側（put_dept など）は、マスターが無いのに _履歴\ に版が残っていれば新しく作らずに止める
  * ロックには持ち主の目印を書き、外すときは目印が自分のものの時だけ消す。落ちたロックの横取りは
    消す直前にもう一度見て、見たときと同じもの（目印と時刻）だけを消す
  * 確定（confirm）は二重に押せない（押すと baseline を焼き直して changes を消してしまう）。
    確定のときの各科の rev を confirmed.revs に残す（kotai_record が「確定後の版」を rev で決める）

20260925（本番前の見直し）:
  * 確定の正本は管理者だけが書ける `R8年N月\集約\確定.json`（confirm／unconfirm がマスターと一緒に書く）。
    マスターの confirmed は全科（Domain Users）が書き換えられるので、アプリの表示用にすぎない。
    15時の反映（dtx14_daily_sync）と交代記録（kotai_record）は effective_confirmed() で確定を決める
    （確定.json が無い月だけマスターに落ちて ⚠、食い違えば警報を出して 確定.json に従う）。
  * 一時ファイルから差し替えられないときに直接上書きへ落ちるのをやめた（app/lib/store.cjs と同じ。何度か待って駄目なら例外）。

20260926（レビューの続き）:
  * put_dept の書き直し（書いたはずの版が読み戻せなかった）は expect_rev を保ったままやり直す。以前は expect_rev を外して
    やり直していたので、その間に他の人が書いた新しい版の上に黙って重ねかねなかった（今は Conflict）。
  * 落ちたロックとみなす古さを 60 秒 → 300 秒（app/lib/store.cjs と揃える。遅い共有で書いている最中のロックを横取りしない）。
  * confirm／unconfirm も確定の正本（集約\確定.json）を見る。正本が確定なら confirm は AlreadyConfirmed。
    confirm がマスターを書いたあとで確定.json を書けなければ、マスターを確定前へ戻す（戻せなければ
    「確定.json を書けませんでした。解除→確定でやり直してください」。unconfirm はどちらの状態からでも通る）。
"""
from __future__ import annotations

import datetime as dt
import glob
import json
import os
import random
import shutil
import socket
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

SCHEMA = "hospital.duty-master/1"
DIRNAME = "master"
HISTORY = "_履歴"
MAX_HISTORY = 400

LOCK_STALE = 300.0     # これより古いロックは落ちたものとみなして横取りする（秒。20260926: 60→300、アプリ側と同じ）
LOCK_WAIT = 20.0       # ロックを待つ上限（秒）


class Conflict(Exception):
    """自科のエントリが、読み込んだときから他の誰かに更新されていた。"""

    def __init__(self, owner, rev, updated, by):
        super().__init__(f"{owner} は {updated} に {by or '別の誰か'} が更新しています（rev {rev}）")
        self.owner, self.rev, self.updated, self.by = owner, rev, updated, by


class Busy(IOError):
    """ロックが取れなかった（他の科が保存中）。"""


class MasterMissing(IOError):
    r"""マスターが見えないのに _履歴\ に版がある（消えた／共有の不調）。新しく作らずに止める。"""


class AlreadyConfirmed(Exception):
    """既に確定している月をもう一度確定しようとした。"""


class TruthError(Exception):
    r"""確定の正本（集約\確定.json）が在るのに読めない。確定しているか分からないので、決める側は止まる。"""


TRUTH_DIR = "集約"            # 管理者だけが書ける（set_acl.py）
TRUTH_NAME = "確定.json"
TRUTH_SCHEMA = "hospital.duty-confirm/1"


def _now():
    return dt.datetime.now().isoformat(timespec="seconds")


def master_dir(month_dir):
    return os.path.join(month_dir, DIRNAME)


def master_path(month_dir, year, month):
    return os.path.join(master_dir(month_dir), f"当直予定表_R{year - 2018}年{month}月.json")


def empty(year, month):
    return {"schema": SCHEMA, "year": year, "month": month, "rev": 0,
            "updated": _now(), "confirmed": None, "depts": {}}


def _missing(path):
    """無い（FileNotFoundError）ときだけ True。それ以外の失敗（SMB の一時的な切断・権限など）は投げる。"""
    try:
        os.stat(path)
        return False
    except FileNotFoundError:
        return True


def refuse_fresh(path):
    r"""マスターが無いのに _履歴\ に版が残っている＝本来あるはずのマスターが見えていない。
    ここで新しく作ると他の科の予定が全部消えるので、作らずに止める（書く側が呼ぶ）。"""
    d = os.path.join(os.path.dirname(path), HISTORY)
    try:
        names = os.listdir(d)
    except FileNotFoundError:
        return
    except OSError as e:
        raise MasterMissing(f"当直予定表の履歴が読めません（{e}）。少し待ってもう一度お試しください")
    if any(n.lower().endswith(".json") for n in names):
        raise MasterMissing(f"当直予定表（{os.path.basename(path)}）が見つかりません。"
                            f"master\\{HISTORY}\\ に過去の版があるので、新しくは作りません。"
                            "直近の版から戻してください")


class lock:
    """排他作成のロックファイル。落ちて残ったロックは古さで横取りする。
    Node 側は fs.openSync(p,'wx') で同じ約束を守ること。"""

    def __init__(self, path, wait=LOCK_WAIT):
        self.p, self.wait = path + ".lock", wait
        self.token = None

    def _info(self):
        """ロックの中身（持ち主の目印）と更新時刻。無い・読めないときは None。"""
        try:
            st = os.stat(self.p)
            with open(self.p, encoding="utf-8", errors="replace") as f:
                return f.read(), st.st_mtime
        except OSError:
            return None

    def __enter__(self):
        os.makedirs(os.path.dirname(self.p), exist_ok=True)
        t0 = time.monotonic()
        token = f"{socket.gethostname()} {os.getpid()} {_now()} {os.urandom(6).hex()}"
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
                age = time.time() - seen[1] if seen else 0.0
                if age > LOCK_STALE:
                    # 落ちた誰かの置き土産: 消す直前にもう一度見て、見たときと同じもの（目印と時刻）だけを消す
                    if self._info() == seen:
                        try:
                            os.remove(self.p)
                        except OSError:
                            pass
                    continue
                if time.monotonic() - t0 > self.wait:
                    raise Busy("他の科が保存中です。少し待ってもう一度お試しください")
                time.sleep(random.uniform(0.02, 0.12))
            except OSError as e:
                raise Busy(f"ロックを作れません: {e}")

    def __exit__(self, *a):
        # 目印が自分のものの時だけ消す（60秒を超えて横取りされたあとに他人のロックを消さない）
        cur = self._info()
        if cur and cur[0] == self.token:
            try:
                os.remove(self.p)
            except OSError:
                pass
        self.token = None
        return False


def load(path, year=None, month=None, retries=6, for_write=False):
    r"""マスターを読む。無ければ空のマスターを返す（その月の最初の1科が作る）。
    書き換えの最中を掴むことがあるので、読めなければ少し待って読み直す。
    for_write: これから書く側。無いのに _履歴\ に版があれば MasterMissing（新しく作らない）。
    「無い」は os.stat の FileNotFoundError だけ。ほかの失敗は読み直し、駄目なら IOError。"""
    last = None
    for i in range(retries):
        try:
            if _missing(path):
                if for_write:
                    refuse_fresh(path)
                return empty(year, month)
        except MasterMissing:
            raise
        except OSError as e:
            last = e
            time.sleep(0.04 * (i + 1))
            continue
        try:
            with open(path, encoding="utf-8-sig") as f:
                m = json.load(f)
            if m.get("schema") != SCHEMA:
                raise ValueError(f"schemaが違います: {m.get('schema')}")
            m.setdefault("depts", {})
            return m
        except (json.JSONDecodeError, OSError) as e:
            last = e
            time.sleep(0.04 * (i + 1))
    raise IOError(f"マスターが読めません（{last}）。master\\{HISTORY}\\ の直近の版から戻してください")


def get_dept(master, owner):
    return (master.get("depts") or {}).get(owner)


def dept_rev(master, owner):
    """同時編集の検出に使う版の目印。エントリが無ければ None。"""
    e = get_dept(master, owner)
    return None if e is None else int(e.get("rev") or 0)


def tmp_name(path):
    """差し替え用の一時ファイル名: 同じフォルダの ~<名前>_<拡張子>_<pid>.tmp。
    20261002: 「確定.json.user006」のような二重拡張子を Python で作ると、EDR が
    ランサムウェアの疑い（T1486）として SOC に通報する。拡張子を1つにする。"""
    d, b = os.path.split(path)
    stem, ext = os.path.splitext(b)
    return os.path.join(d, f"~{stem}_{ext.lstrip('.')}_{os.getpid()}.tmp")


def _write(path, master, retries=10):
    """一時ファイル → 差し替え。Windows は他プロセスが開いている間 replace できないので待って繰り返す。
    それでも駄目なら OSError（20260925: 直接上書きへ落ちるのをやめた。書いている途中を他の科が読むと
    空や途中のファイルを掴むため。app/lib/store.cjs と同じ）。一時ファイルは必ず消す。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = json.dumps(master, ensure_ascii=False)
    tmp = tmp_name(path)
    last = None
    try:
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(data)
        for i in range(retries):
            try:
                os.replace(tmp, path)
                return "replace"
            except OSError as e:
                last = e
                time.sleep(0.05 * (i + 1))
        raise OSError(f"{os.path.basename(path)} を差し替えられません（開かれている？）: {last}")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def backup(path):
    r"""上書きの前に _履歴\ へ複製する。改名ではなく複製（他人のファイルは改名できない）。"""
    if not os.path.exists(path):
        return None
    d = os.path.join(os.path.dirname(path), HISTORY)
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        return None
    base = os.path.splitext(os.path.basename(path))[0]
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(d, f"{base}_{stamp}.json")
    for i in range(2, 100):
        if not os.path.exists(dest):
            break
        dest = os.path.join(d, f"{base}_{stamp}_{i}.json")
    try:
        shutil.copyfile(path, dest)
    except OSError:
        return None
    try:   # 間引きは best-effort（_履歴\ は削除を許さない ACL にするので、普通は失敗して構わない）
        old = sorted(glob.glob(os.path.join(d, f"{base}_*.json")), key=os.path.getmtime)
        for p in old[:-MAX_HISTORY]:
            os.remove(p)
    except OSError:
        pass
    return dest


def put_dept(path, owner, *, duties=None, roster=None, changes=None, baseline=False, by="",
             expect_rev=False, year=None, month=None, retries=3):
    """自科のエントリだけ差し替えて書き戻す。戻り: (master, 保存したdept)。

    expect_rev: 読み込んだときの自科 rev（`dept_rev()` の値）。違っていたら Conflict を投げる。
                False を渡すと照合しない（同じ科を2人が開いている事故を見逃すので、
                アプリからは必ず読んだときの rev を渡すこと）。"""
    for attempt in range(max(1, retries)):
        with lock(path):
            m = load(path, year, month, for_write=True)
            cur = get_dept(m, owner) or {}
            if expect_rev is not False and cur:
                if int(cur.get("rev") or 0) != (expect_rev or 0):
                    raise Conflict(owner, cur.get("rev"), cur.get("updated"), cur.get("by"))
            ent = dict(cur)
            if duties is not None:
                ent["duties"] = list(duties)
            if roster is not None:
                ent["roster"] = list(roster)
            if changes is not None:
                ent["changes"] = list(changes)
            if baseline is not False:
                ent["baseline"] = None if baseline is None else list(baseline)
            ent.setdefault("duties", [])
            ent.setdefault("roster", [])
            ent.setdefault("changes", [])
            ent.setdefault("baseline", None)
            ent["updated"], ent["by"] = _now(), by
            ent["rev"] = int(cur.get("rev") or 0) + 1
            m.setdefault("depts", {})[owner] = ent
            m["rev"] = int(m.get("rev") or 0) + 1
            m["updated"] = ent["updated"]
            backup(path)
            _write(path, m)
            back = get_dept(load(path, year, month), owner) or {}
            if int(back.get("rev") or 0) == ent["rev"]:
                return m, ent
        # 書いたはずの版が読み戻せない（ロックを守らない書き手に上書きされた等）。expect_rev はそのまま（20260926）:
        #   読み直した自科の rev が expect_rev のままなら書き直す、他の人の新しい版になっていれば次の周で Conflict
        time.sleep(random.uniform(0.05, 0.2))
    raise Busy("マスターに書けませんでした（他の科が保存中です）")


def init(path, year, month):
    r"""月フォルダを作るときに空のマスターを置く（month_setup）。既にあれば触らず読んで返す。
    無いのに _履歴\ に版があるときは作らない（MasterMissing）。"""
    with lock(path):
        if not _missing(path):
            return load(path, year, month)
        refuse_fresh(path)
        m = empty(year, month)
        _write(path, m)
    return m


def _truth_confirmed(path, strict):
    r"""確定の正本（集約\確定.json）の confirmed（無い・未確定は None）。壊れているとき strict なら TruthError、でなければ None。"""
    try:
        t = read_truth(path)
    except TruthError:
        if strict:
            raise
        return None
    return (t or {}).get("confirmed") or None


TRUTH_FAIL_MSG = "確定.json を書けませんでした。解除→確定でやり直してください"


def confirm(path, by="", year=None, month=None):
    """この月を確定する。各科の baseline に、そのときの duties を焼き付ける。
    既に確定していれば AlreadyConfirmed（焼き直すと元の割当と changes が失われる）。解除してからなら付け直せる。
    確定しているかは、マスターの confirmed と確定の正本（集約\確定.json）の両方で見る（20260926。どちらかが確定なら断る）。
    確定.json が壊れていれば TruthError（解除すれば作り直せる）。
    確定のときの各科の rev を confirmed.revs に残す（kotai_record がこれより後の版を交代として記録する）。
    マスターを書いたあとで確定.json を書けなければ、マスターを確定前の中身へ戻して OSError（20260926）。"""
    with lock(path):
        m = load(path, year, month, for_write=True)
        tc = _truth_confirmed(path, strict=True)
        conf = m.get("confirmed")
        if tc:
            raise AlreadyConfirmed(f"この月は既に確定しています（{tc.get('at') or '?'} {tc.get('by') or ''}）。"
                                   "付け直すときは先に「確定を解除」してください")
        if conf:
            raise AlreadyConfirmed(f"共有マスターに確定の印があります（{conf.get('at') or '?'} {conf.get('by') or ''}）が、"
                                   "確定.json は未確定です。先に「確定を解除」してから、もう一度確定してください")
        before = json.loads(json.dumps(m, ensure_ascii=False))
        for ent in (m.get("depts") or {}).values():
            ent["baseline"] = list(ent.get("duties") or [])
            ent["changes"] = []
        m["confirmed"] = {"at": _now(), "by": by,
                          "revs": {o: int(e.get("rev") or 0) for o, e in (m.get("depts") or {}).items()}}
        m["rev"] = int(m.get("rev") or 0) + 1
        m["updated"] = _now()
        backup(path)
        _write(path, m)
        try:
            write_truth(path, m["confirmed"], "confirm", by)      # 確定の正本（管理者だけ。20260925）
        except OSError as e:
            # 20260926: マスターだけ確定・正本は未確定、を残さない。ロックを持っているので他の科は書いていない
            before["rev"] = int(m["rev"]) + 1
            before["updated"] = _now()
            try:
                _write(path, before)
            except OSError as e2:
                raise OSError(f"{TRUTH_FAIL_MSG}（{e}／マスターも戻せませんでした: {e2}）") from e
            raise OSError(f"{TRUTH_FAIL_MSG}（{e}。マスターは確定前に戻しました）") from e
    return m


def unconfirm(path, by="", year=None, month=None):
    """確定を解除する（押し間違いの戻し道）。confirmed を落とすだけで、
    焼いた baseline と積んだ changes はそのまま残す（消すと交代の記録が失われる）。
    マスターと確定.json のどちらか（片方だけでも）が確定なら両方を未確定にする（20260926: 確定.json を書けなかった確定・
    マスターだけ書き換えられた確定の直し道）。確定.json が壊れていても通る（作り直す）。"""
    with lock(path):
        m = load(path, year, month, for_write=True)
        _truth_confirmed(path, strict=False)          # 読めるか確かめるだけ（壊れていても write_truth が作り直す）
        m["confirmed"] = None
        m["rev"] = int(m.get("rev") or 0) + 1
        m["updated"] = _now()
        backup(path)
        _write(path, m)
        write_truth(path, None, "unconfirm", by)             # 確定の正本（管理者だけ。20260925）
    return m


# ---------------------------------------------------------------- 確定の正本（20260925）

def truth_path(path):
    r"""マスターのパス → 同じ月の 集約\確定.json。"""
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(path))), TRUTH_DIR, TRUTH_NAME)


def read_truth(path):
    r"""確定の正本を読む。戻り: None＝ファイルが無い（旧い月）／dict（{"confirmed": {at, by, revs} | None, "history": [...]}）。
    在るのに読めない・形が違うときは TruthError。"""
    tp = truth_path(path)
    try:
        with open(tp, encoding="utf-8-sig") as f:
            t = json.load(f)
    except FileNotFoundError:
        return None
    except (OSError, ValueError) as e:
        raise TruthError(f"{tp} が読めません（{type(e).__name__}: {e}）") from e
    if not isinstance(t, dict) or t.get("schema") != TRUTH_SCHEMA or "confirmed" not in t:
        raise TruthError(f"{tp} の形が違います")
    return t


def write_truth(path, confirmed, action, by=""):
    r"""確定の正本（集約\確定.json）を書く。history に誰がいつ確定・解除したかを積む（消さない）。
    一時ファイル → 差し替え（直接上書きしない）。書けなければ OSError（管理ツールに出す）。"""
    tp = truth_path(path)
    try:
        cur = read_truth(path) or {}
    except TruthError:
        cur = {}              # 壊れていたら作り直す（history は失うが、確定の状態はマスターと同じにする）
    hist = list(cur.get("history") or [])
    try:
        who = os.getlogin()
    except OSError:
        who = os.environ.get("USERNAME") or ""
    hist.append({"action": action, "at": _now(), "by": by, "user": who, "host": socket.gethostname(),
                 "confirmed": confirmed})
    t = {"schema": TRUTH_SCHEMA, "confirmed": confirmed, "updated": _now(), "history": hist}
    os.makedirs(os.path.dirname(tp), exist_ok=True)
    tmp = tmp_name(tp)
    try:
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(t, f, ensure_ascii=False, indent=1)
        for i in range(10):
            try:
                os.replace(tmp, tp)
                return t
            except OSError:
                if i == 9:
                    raise
                time.sleep(0.05 * (i + 1))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return t


def effective_confirmed(path, master):
    r"""確定しているかを決める（15時の反映・交代記録が使う。20260925）。戻り: (confirmed の dict か None, 問題の一覧)。
    問題 = [("warn"|"alert", 説明)]。
      * 集約\確定.json が無い（この仕組みより前に確定した月）→ マスターの confirmed に従い "warn"
      * 確定.json とマスターが食い違う（どちらかだけ確定・確定の時刻が違う）→ 確定.json に従い "alert"
        （マスターは全科が書けるので、書き換えられた可能性がある）
      * 確定.json が壊れている → TruthError（決める側は止まる）"""
    t = read_truth(path)
    mc = master.get("confirmed") or None
    if t is None:
        return mc, ([("warn", f"確定の正本（{TRUTH_DIR}\\{TRUTH_NAME}）が無いので、共有マスターの confirmed で決めた"
                              f"（{'確定' if mc else '未確定'}）")] if mc else [])
    tc = t.get("confirmed") or None
    if bool(tc) != bool(mc) or (tc and mc and (tc.get("at") or "") != (mc.get("at") or "")):
        def show(c):
            return f"確定（{c.get('at') or '?'} {c.get('by') or ''}）" if c else "未確定"
        return tc, [("alert", f"共有マスターの確定の印（{show(mc)}）が確定の正本（{show(tc)}）と違う"
                              " → 正本に従う（マスターが書き換えられた可能性）")]
    return tc, []


def import_submissions(path, submit_dir, year, month):
    r"""移行用: 旧 提出\ の提出JSON（hospital.duty-submission/1）をマスターへ流し込む。
    duties / roster の形は同じなので入れ替えで済む。"""
    n = 0
    with lock(path):
        m = load(path, year, month, for_write=True)
        for p in sorted(glob.glob(os.path.join(submit_dir, "*.json"))):
            try:
                sub = C.load_submission_json(p)
            except Exception:
                continue
            owner = sub.get("dept")
            if not owner or (sub.get("year"), sub.get("month")) != (year, month):
                continue
            m.setdefault("depts", {})[owner] = {
                "updated": sub.get("saved") or _now(), "by": sub.get("by") or "(移行)",
                "rev": 1, "roster": sub.get("roster") or [], "duties": sub.get("duties") or [],
                "baseline": None, "changes": []}
            n += 1
        m["rev"] = int(m.get("rev") or 0) + 1
        m["updated"] = _now()
        backup(path)
        _write(path, m)
    return m, n
