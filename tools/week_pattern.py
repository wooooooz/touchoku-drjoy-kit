# -*- coding: utf-8 -*-
r"""week_pattern.py — 区分が変わった医師の Dr.JOY の週パターンを自動で切り替える。20261031

管理者 2026-10-31 決定:
  所属の変更メール（shozoku_notice）の「週パターン」は事務担当の仕事ではない。こちらで Dr.JOY に入れて、メールには
  「変更しました」と書く。変えるのは未来の予定だけ。
  - フルタイム → パートタイム・院外研修専従: 週パターン「勤務なし（院外研修中）」。
  - → フルタイム（通常に戻す）: 兼業が無い医師は「月～金 8:30-17:15」。兼業がある医師は診療科に依頼してもらう。
    同日の追加: 兼業の申請から私たちが作った週パターン「<氏名>_自動登録」（2026-07-17 に111本）があれば、それに戻す
    （週パターンの一覧は DrJoyAutomation\drjoy-week-patterns.js で読む＝読むだけ）。
    戻す先の順 = ①「<氏名>_自動登録」 ②勤務なしにする前が平日すべて 8:30-17:15 なら「月～金 8:30-17:15」 ③診療科に依頼。
    兼業の有無は、勤務なしにする前の勤務予定の形で判断する（Dr.JOY の書き出しには週パターンの名前が載らないので、
    名前の代わりに「平日がすべて 通常勤務 8:30-17:15 だったか」を覚えておく＝STATE_PATH）。
    覚えていない医師（この仕組みより前に勤務なしにされた・初めから勤務なしで申告された）は診療科に依頼。

Dr.JOY の性質:
  取り込み（at0069 勤務予定）で D列「週パターン適用（名）」を入れた行は、その日から先をずっとその週パターンで上書きする。
  すでに入っている当直も消える。当直は毎月29日に翌月分がまとめて入る（手当計算）。
  → 変える日（D）から先に、その月の当直（オンコールでないもの）がある医師は自動で変えない（メールで管理者が対応と書く）。
     次の月以降の当直は29日の投入で上から入るので消えない。

変える日 D = その月の1日（もう過ぎていれば明日）。
流れ: 分類 → （勤務なしにする人がいれば）勤務予定を書き出して今の形を覚える → 取り込みファイル（work-schedule_v3 の形）
      → DrJoyAutomation\drjoy-at0069-import.js → もう一度書き出して、D から先の平日が思ったとおりかを確かめる。
Dr.JOY への書き込みはこの取り込みだけ（画面の CSV インポート経由）。
"""
from __future__ import annotations

import datetime as dt
import glob
import json
import os
import re
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402

PATTERN_OFF = "勤務なし（院外研修中）"
PATTERN_STD = "月～金 8:30-17:15"           # 全角の「～」（Dr.JOY の名前どおり）
DRJOY_DIR = r"C:\tools\touchoku\drjoy"
TEMPLATE = r"C:\tools\touchoku\drjoy_import\work-schedule_v3.xlsx"
DATA_ROW = 16                               # ひな形の実データの開始行（11〜15行目は記入例＝消す）
LOCAL = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku")
STATE_PATH = os.path.join(LOCAL, "week_pattern_state.json")
WORK_DIR = os.path.join(LOCAL, "week_pattern")
EXPORT_TIMEOUT = 4 * 60                     # 勤務予定の書き出しは 30〜65 秒
IMPORT_TIMEOUT = 10 * 60                    # 取り込みは受付から終了まで数分
RUNNER = None          # (args, timeout) → (returncode, 出力)。None なら node を本当に走らせる。テストは差し替える


def change_date(y, m, today):
    """変える日 = その月の1日。もう過ぎていれば明日（未来の予定だけ。管理者 20261031）。"""
    return max(dt.date(y, m, 1), today + dt.timedelta(days=1))


def load_state(path=None):
    try:
        with open(path or STATE_PATH, encoding="utf-8") as f:
            return json.load(f).get("people") or {}
    except (OSError, ValueError, AttributeError):
        return {}


def save_state(people, path=None):
    p = path or STATE_PATH
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        from master_store import tmp_name        # 二重拡張子にしない（EDR）
        tmp = tmp_name(p)
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"schema": "hospital.week-pattern/1", "people": people}, f, ensure_ascii=False, indent=1)
        os.replace(tmp, p)
        return True
    except OSError:
        return False


def duties_from(master, login, d):
    """マスターで、その人の d 以降の当直（オンコールでないもの）の日付。"""
    k = (login or "").lower()
    out = set()
    for ent in (master.get("depts") or {}).values():
        for r in ent.get("duties") or []:
            if (r.get("staff_login") or "").lower() == k and not r.get("oncall") and str(r.get("date")) >= d.isoformat():
                out.add(str(r["date"]))
    return sorted(out)


def autoreg_map(patterns):
    """週パターンの一覧 → {氏名のキー: 「<氏名>_自動登録」の名前}（同じキーが2本以上あれば使わない）。"""
    seen = {}
    for p in patterns or []:
        n = str(p.get("name") or "")
        if n.endswith("_自動登録"):
            k = nkey(n[:-len("_自動登録")])
            seen[k] = None if k in seen else n
    return {k: v for k, v in seen.items() if v}


def nkey(name):
    return C.canon(C.nfc(C.namekey(name or "")))


def classify(changes, master, y, m, today, state, code_of=None, autoreg=None):
    """changes = [{"login","name","owner","dir":"off"/"on","work"}] → 1人ずつ {…, "date", "code", "action", "pattern", "why"}。
    action = "apply"（自動で入れる）/ "duty"（当直があるので変えない）/ "dept"（診療科に依頼）/ "nocode"（職員コードが無い）"""
    code_of = code_of or C.code_of
    d = change_date(y, m, today)
    out = []
    for c in changes:
        x = dict(c, date=d.isoformat(), code=code_of(c["login"]) or "", action="apply", pattern="", why="")
        if not x["code"]:
            x.update(action="nocode", why="職員コードが引けない")
        elif c["dir"] == "off":
            x["pattern"] = PATTERN_OFF
            ds = duties_from(master, c["login"], d)
            if ds:
                x.update(action="duty", why=f"{d.month}月{d.day}日から先に当直がある（{len(ds)}日）")
        else:
            mem = state.get((c["login"] or "").lower()) or {}
            own = (autoreg or {}).get(nkey(c.get("name")))
            if own:
                x["pattern"] = own                  # 兼業の申請から作った本人の週パターン
            elif mem.get("standard") is True:
                x["pattern"] = PATTERN_STD
            else:
                x.update(action="dept", why="兼業などで科ごとの週パターンがある" if mem else "前の週パターンが分からない")
        out.append(x)
    return out


# ------------------------------------------------------------------ 勤務予定の書き出しを読む

def _hm(v):
    return str(v or "").strip()


def read_shapes(xlsx, codes, holidays, since):
    """勤務予定の書き出し（1人1シート。1行目に「職員コード」の欄）→ {code: [(date, 区分, 勤務種類, 開始, 終了)]}（since 以降の平日だけ）。
    ⚠ シート名は氏名だけ（月の書き出し）か「職員コード 氏名」（1人ずつの書き出し）。職員コードは1行目から読む。
    ⚠ read_only で開くとシートの大きさが A1 だけに見えて中身が読めない（20261031 実測）。"""
    import openpyxl
    want = set(codes)
    out = {}
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=60, values_only=True)]
            head = [str(v).strip() if v is not None else "" for v in (rows[0] if rows else [])]
            code = head[head.index("職員コード") + 1] if "職員コード" in head[:-1] else str(ws.title).split(" ")[0]
            if code not in want:
                continue
            year = None
            for v in head:
                mt = re.match(r"(\d{4})年(\d{1,2})月度", v)
                if mt:
                    year = int(mt.group(1))
            days = []
            for r in rows:
                v = list(r or ())
                mt = re.match(r"(\d{1,2})/(\d{1,2})", str(v[0] if v else "") or "")
                if not mt or year is None:
                    continue
                mo, da = int(mt.group(1)), int(mt.group(2))
                d = dt.date(year + (1 if mo < 1 else 0), mo, da)
                if d < since or C.is_holiday(d, holidays):
                    continue
                v += [None] * 8
                days.append((d.isoformat(), _hm(v[3]), _hm(v[4]), _hm(v[5]), _hm(v[6])))
            out[code] = days
    finally:
        wb.close()
    return out


def is_standard(days):
    """平日がすべて 勤務日・通常勤務 8:30-17:15 か（兼業があれば科の週間予定で平日のどこかが違う）。"""
    return bool(days) and all(k == "勤務日" and s == "08:30" and e == "17:15" for _d, k, _w, s, e in days)


def is_off(days):
    return bool(days) and all(k == "勤務しない日" for _d, k, _w, _s, _e in days)


# ------------------------------------------------------------------ 取り込みファイルと node

def write_import(path, rows):
    """rows = [(職員コード, date, 週パターン名)] → work-schedule_v3 の形の xlsx。"""
    import openpyxl
    shutil.copy2(TEMPLATE, path)
    wb = openpyxl.load_workbook(path)
    ws = wb["勤務予定"]
    for r in range(11, ws.max_row + 1):           # 記入例（2022年の架空データ）と残りを消す
        for c in range(1, 33):
            ws.cell(r, c).value = None
    for i, (code, d, pat) in enumerate(rows):
        ws.cell(DATA_ROW + i, 2).value = code
        ws.cell(DATA_ROW + i, 3).value = d.strftime("%Y/%m/%d")
        ws.cell(DATA_ROW + i, 4).value = pat
    wb.save(path)
    wb.close()


def run_node(args, timeout=IMPORT_TIMEOUT):
    if RUNNER is not None:
        return RUNNER(args, timeout)
    if "selftest" in os.path.basename(sys.argv[0] if sys.argv else ""):    # 試験から本物の Dr.JOY に触らない
        return 1, "selftest からは node を呼ばない"
    node =shutil.which("node") or r"C:\Program Files\nodejs\node.exe"
    try:
        r = subprocess.run([node] + args, cwd=DRJOY_DIR, capture_output=True, timeout=timeout,
                           stdin=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.SubprocessError) as e:
        return 1, f"{type(e).__name__}: {e}"
    text = (r.stdout or b"").decode("utf-8", "replace") + (r.stderr or b"").decode("utf-8", "replace")
    return r.returncode, text


def list_patterns(out_dir, log):
    """Dr.JOY の週パターンの一覧（読むだけ）。だめなら None。"""
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "patterns.json")
    rc, text = run_node(["drjoy-week-patterns.js", "--out", out, "--headless"], EXPORT_TIMEOUT)
    if rc != 0 or "OK: " not in text:
        log(f"  週パターン: 一覧を読めない（{rc}）{text.strip()[-300:]}")
        return None
    with open(out, encoding="utf-8") as f:
        return json.load(f)


def export_plan(month, out_dir, log):
    """勤務予定（仮登録込み）を書き出す。戻り: xlsx のパス（だめなら None）。"""
    os.makedirs(out_dir, exist_ok=True)
    rc, text = run_node(["drjoy-at0029-export.js", "--out-dir", out_dir, "--type", "working-plan",
                         "--month", month, "--headless"], EXPORT_TIMEOUT)
    mt = re.search(r"OK: (.+?\.xlsx)", text)
    if rc != 0 or not mt:
        log(f"  週パターン: 勤務予定の書き出しに失敗（{rc}）{text.strip()[-300:]}")
        return None
    return mt.group(1).strip()


def import_file(path, log):
    """戻り: (全部入ったか, 説明)。"""
    rc, text = run_node(["drjoy-at0069-import.js", "--file", path, "--headless"], IMPORT_TIMEOUT)
    mt = re.search(r"OK: .*total=(\d+) success=(\d+) error=(\d+)", text)
    if mt:
        tot, ok, err = (int(x) for x in mt.groups())
        return (rc == 0 and err == 0 and ok == tot), f"読み込み {tot} 成功 {ok} 失敗 {err}"
    return False, f"取り込みの結果が読めない（{rc}）{text.strip()[-300:]}"


# ------------------------------------------------------------------ 本体

def apply(changes, master, y, m, today=None, *, log=print, state_path=None, holidays=None, code_of=None):
    """changes を分類し、自動で入れるものを Dr.JOY に入れる。戻り: classify の行に "result"（done/failed/…）を足したもの。"""
    today = today or dt.date.today()
    holidays = holidays if holidays is not None else C.load_holidays()
    state = load_state(state_path)
    run_dir = os.path.join(WORK_DIR, dt.datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    autoreg = None
    if any(c["dir"] == "on" for c in changes):     # 戻す人がいれば「<氏名>_自動登録」を探す（読めなければ①を飛ばす）
        pats = list_patterns(run_dir, log)
        autoreg = autoreg_map(pats) if pats else None
    items = classify(changes, master, y, m, today, state, code_of, autoreg)
    for x in items:
        x["result"] = x["action"]
    check = [x for x in items if x["code"]]
    if not check:
        return items
    d = dt.date.fromisoformat(check[0]["date"])
    month = f"{d.year:04d}-{d.month:02d}"
    os.makedirs(run_dir, exist_ok=True)

    def fail(xs, why):
        for x in xs:
            x.update(result="failed", why=why)
        log(f"  ✗ 週パターン: {why}")
        return items

    # Dr.JOY のいまの勤務予定と比べる（20261031 管理者: 知らせたかどうかでなく、Dr.JOY が区分と合っているかで決める）。
    # 合っていれば何もしない（already）。勤務なしにする人は、いまの形を覚える（戻すときの判断に使う）
    before = export_plan(month, os.path.join(run_dir, "before"), log)
    if not before:
        return fail([x for x in check if x["action"] == "apply"], "勤務予定を書き出せない（Dr.JOY は変えていない）")
    shapes = read_shapes(before, [x["code"] for x in check], holidays, d)
    recorded = False
    for x in check:
        days = shapes.get(x["code"]) or []
        if x["dir"] == "off":
            if is_off(days):
                x["result"] = "already"
            elif not days:
                x.update(result="failed", why="Dr.JOY の勤務予定にこの医師のシートが無い")
            elif x["result"] == "apply":
                state[(x["login"] or "").lower()] = {
                    "name": x["name"], "standard": is_standard(days), "recorded": today.isoformat(), "from": x["date"],
                    "sample": days[:5]}
                recorded = True
        elif days and not is_off(days):           # 戻す人: もう勤務なしでない（誰かが戻した）
            x["result"] = "already"
    if recorded and not save_state(state, state_path):
        return fail([x for x in check if x["result"] == "apply"],
                    f"前の週パターンを覚えておけない（{state_path or STATE_PATH}）→ Dr.JOY は変えていない")
    todo = [x for x in items if x["result"] == "apply"]
    if not todo:
        return items
    xlsx = os.path.join(run_dir, f"週パターン_{today:%Y%m%d}.xlsx")
    write_import(xlsx, [(x["code"], dt.date.fromisoformat(x["date"]), x["pattern"]) for x in todo])
    ok, how = import_file(xlsx, log)
    log(f"  週パターン: 取り込み {how}（{len(todo)} 人。{os.path.basename(xlsx)}）")
    after = export_plan(month, os.path.join(run_dir, "after"), log)
    shapes = read_shapes(after, [x["code"] for x in todo], holidays, d) if after else {}
    for x in todo:
        days = shapes.get(x["code"]) or []
        good = (is_off(days) if x["dir"] == "off" else is_standard(days) if x["pattern"] == PATTERN_STD
                else bool(days) and not is_off(days))           # 本人の週パターン: 勤務なしでなくなったか
        if good:
            x["result"] = "done"
        else:
            x.update(result="failed", why=("取り込みで入らなかった（" + how + "）") if not ok
                     else "取り込んだが、書き出しで確かめられない" if not after
                     else f"取り込んだが、{x['date']} から先の平日が思ったとおりになっていない")
        log(f"  週パターン: {x['name']} → {x['pattern']}（{x['date']} から）{'✅' if good else '✗ ' + x['why']}")
    restored = [x for x in todo if x["result"] == "done" and x["dir"] == "on"]
    for x in restored:                             # 覚えた形は残す（やり直しの回でも同じ判断になるように）
        state.setdefault((x["login"] or "").lower(), {})["restored"] = today.isoformat()
    if restored:
        save_state(state, state_path)
    return items
