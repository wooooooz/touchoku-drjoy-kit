# -*- coding: utf-8 -*-
r"""roster_editors.py — 貴科の医師リストを直せる人（部門管理者）の一覧を作る。20261028

管理者 20261028: 貴科の医師リスト（Dr.JOY所属の申告）の × と区分の切り替えは、部門管理者だけができる。
部門管理者 = Dr.JOY のスタッフ一覧の「管理権限」が 所属管理者（その所属の秘書・部長）か 全体管理者。
毎朝3:30の名簿の更新（roster_refresh）が、Dr.JOY のエクスポート（Downloads の staffs*.csv）から
本番の 名簿\roster_editors.json を作る。アプリ（app/lib/rights.cjs）が読む。
中身はログインID（@ の前・小文字）→ {name, level（全体/所属）, depts（所属のリスト）}。職員コードは入れない。
アカウント状態が有効・仮登録の人だけ（ロック・削除は入れない）。
同日: data/roster_editors_extra.json の「追加の作成者」（科ごとに名指し。Dr.JOY の管理権限が無い医師や、
Windows のログイン名が Dr.JOY のログインIDと違う人）を owners として混ぜる（level が無い人は「追加」）。

usage: python -X utf8 tools\roster_editors.py [SRC.csv] [--out PATH]   （省略 = Downloads の最新・本番の 名簿\）
"""
import argparse
import csv
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C            # noqa: E402
import make_staff_csv as MS     # noqa: E402
import master_store as M        # noqa: E402
import shozoku as SZ            # noqa: E402

FILE_NAME = "roster_editors.json"
EXTRA = os.path.join(C.DATA, "roster_editors_extra.json")
LEVELS = {"全体管理者": "全体", "所属管理者": "所属"}
STATES = MS.STATES                # 有効・仮登録（ロック・削除は入れない。make_staff_csv と同じ）
DOMAIN = "@example.ac.jp"


def build(src, extra=None):
    """staffs*.csv（Dr.JOY のエクスポート・cp932）→ {"made", "source", "editors": {login: {...}}}。
    extra = 追加の作成者の JSON のパス（省略=data/roster_editors_extra.json。無ければ足さない）"""
    with open(src, encoding="cp932", newline="") as f:
        rows = list(csv.DictReader(f))
    if rows and "管理権限" not in rows[0]:
        raise SystemExit(f"元CSVに「管理権限」の列がありません（{src}）")
    eds = {}
    for r in rows:
        level = LEVELS.get((r.get("管理権限") or "").strip())
        login = (r.get("ログインID") or "").strip().lower()
        if not level or not login.endswith(DOMAIN) or (r.get("アカウント状態") or "").strip() not in STATES:
            continue
        name = f"{(r.get('姓') or '').strip()} {(r.get('名') or '').strip()}".strip()
        eds[login[: -len(DOMAIN)]] = {"name": name, "level": level, "depts": SZ.segments(r.get("所属"))}
    xp = extra or EXTRA
    if os.path.isfile(xp):
        with open(xp, encoding="utf-8-sig") as f:
            ex = json.load(f).get("editors") or {}
        for login, e in ex.items():
            k = str(login).strip().lower()
            cur = eds.setdefault(k, {"name": e.get("name", ""), "level": "追加", "depts": []})
            cur["owners"] = sorted(set(cur.get("owners", [])) | set(e.get("owners") or []))
    return {"made": dt.datetime.now().isoformat(timespec="seconds"), "source": os.path.basename(src),
            "editors": dict(sorted(eds.items()))}


def write(data, path):
    """一時ファイルに書いて読み戻してから差し替える（アプリが読んでいる途中に半端なファイルを見せない）。"""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = M.tmp_name(path)
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    with open(tmp, encoding="utf-8") as f:
        if json.load(f) != data:
            raise OSError(f"書き込みの確かめに失敗: {tmp}")
    os.replace(tmp, path)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?", default=None, help="Dr.JOY のスタッフ一覧のエクスポート（省略=Downloads の最新）")
    ap.add_argument("--out", default=None, help=f"書き出し先（省略=本番の 名簿\\{FILE_NAME}）")
    a = ap.parse_args()
    src = a.src or MS.newest_raw()
    if not src:
        raise SystemExit("Downloads に Dr.JOY のエクスポート（staffs*.csv）がありません")
    data = build(src)
    out = write(data, a.out or os.path.join(C.PROD_ROSTER, FILE_NAME))
    n = {lv: sum(1 for e in data["editors"].values() if e["level"] == lv) for lv in ("全体", "所属", "追加")}
    print(f"部門管理者: {os.path.basename(src)} → {out}（全体管理者 {n['全体']}・所属管理者 {n['所属']}・追加の作成者 {n['追加']}）")


if __name__ == "__main__":
    main()
