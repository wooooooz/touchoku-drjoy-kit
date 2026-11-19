# -*- coding: utf-8 -*-
"""set_acl.py — 本番共有の NTFS ACL（icacls）。20260928

ユーザー決定（2026-09-09）: 提出フォルダは誰でも見られてよい。ただし削除はされないように。
管理当直で実証した型（Domain Users に app/config=RX、提出=書込）を踏襲し、削除だけを外す:
  ルート 当直勤務表\\        Domain Users: RX（このフォルダのみ＝一覧だけ）
  R8年N月\\                  Domain Users: (OI)(CI)RX（exe を起動・config/テンプレ/HTML を読む）
  （R8年N月\\提出\\ は 20260912 の xlsx一本化で廃止。20260923 に本番からも消した）
  R8年N月\\master\\           Domain Users: (OI)(CI)(M)＝変更（アプリが上書きするので削除権が要る）
  R8年N月\\master\\_履歴\\    継承を切り CREATOR OWNER を外す。(OI)(CI)(RX,W)＝書けるが消せない
                            （20260925 に、複製を書き換え不可にする案〔OWNER RIGHTS・(RX,WD)〕を9月試験用フォルダで試したが、
                              管理者「そんなことする人いないです。想定しないでOK」で不採用。試しにかけた項目は外す）
  R8年N月\\集約\\            管理者だけ（Domain Users の項目を入れない）
  科別ツール\\ ドキュメント\\  Domain Users: (OI)(CI)RX。科別ツール\\<科>\\out\\ だけ (OI)(CI)(RX,W)（ツールの出力先）
  名簿\\                    Domain Users: (OI)(CI)RX（20260915 管理者指示: 科別ツールが共有の名簿を直接読む）
  _本体\\ 管理ツール\\        管理者だけ
管理者 = ADMINS の3人だけ（ユーザー決定 2026-09-09: 00_DrJOY から継承する十数人のうち 管理者・事務担当・事務担当2に絞る）。
  当直勤務表\\ で継承を切り、ADMINS 以外の個人アカウントを外す（SYSTEM・Domain Admins・BUILTIN\\Administrators は
  サーバ管理側なので残す）。継承を切ってある 提出\\ にも同じ削除を掛ける。

usage: python -X utf8 tools\\set_acl.py [--month 2026-10 ...] [--root] [--show PATH]
  引数なし = ルート（管理者の絞り込み込み）＋科別ツール＋ドキュメント＋存在する全月フォルダ
"""
import argparse
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

USERS = r"A\Domain Users"
# 管理者だけの月フォルダ（管理者 20260928: 9月の試験用・10月の予行は一般ユーザーから見えないように。管理者の動作検証に使う）。
#   本番の最初の月（shozoku.REAL_FROM・dtx14_daily_sync.SYNC_FROM と同じ 2026-11）より前の月と、試験用フォルダ
PRIVATE_BEFORE = (2026, 11)
# 試験用の月フォルダ（管理者 20260928: 10月も「(動作確認用テストフォルダ)」を付けて事務担当・事務担当2が混乱しないように）
TEST_SUFFIX = "(動作確認用テストフォルダ)"
PRIVATE_DIRS = ("R8年9月" + TEST_SUFFIX, "R8年10月" + TEST_SUFFIX)


def parse_test_dir(name):
    """'R8年10月(動作確認用テストフォルダ)' → (2026, 10)。試験用フォルダでなければ None。"""
    return C.parse_month_dir(name[:-len(TEST_SUFFIX)]) if name.endswith(TEST_SUFFIX) else None
DOMAIN = "A"
ADMINS = ("user003", "user004", "user005")          # 管理者・事務担当・事務担当2
KEEP = {"domain users", "domain admins", "administrator"}   # A\ 配下で個人扱いしないもの


NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # 管理ツール（--windowed）から呼ぶときに黒い窓を出さない（20260926）


def icacls(*args, check=True):
    r = subprocess.run(["icacls", *args], capture_output=True, text=True, encoding="cp932", errors="replace",
                       creationflags=NO_WINDOW)
    if check and r.returncode != 0:
        raise SystemExit(f"icacls 失敗 ({r.returncode}): icacls {' '.join(args)}\n{r.stdout}{r.stderr}")
    return r


OWNER_RIGHTS = "*S-1-3-4"                        # OWNER RIGHTS（名前は言語で変わるので SID で指す）
CREATOR_OWNER = "*S-1-3-0"                       # CREATOR OWNER


def grant(path, spec, log=print, who=USERS, label="Domain Users"):
    icacls(path, "/grant", f"{who}:{spec}")
    log(f"  {os.path.relpath(path, C.PROD_ROOT) or '.'}  ← {label} {spec}")


# 『A\名前:(』。最初の ACE はパスと同じ行に出る（『\\srv\...\集約 A\user003:(OI)(CI)(F)』）ので行頭に限らない（20260926）
ACE_RE = re.compile(r"(?:^|\s)" + re.escape(DOMAIN) + r"\\([^\\:\s][^\\:]*?):\(")


def parse_accounts(out):
    """icacls の出力から A\\xxx の個人アカウント名（小文字）を集める。
    20260926: 以前は行頭が A\\ の行だけを見ていたので、パスと同じ行に出る1つ目の ACE を読み落としていた。"""
    return {m.group(1).strip().lower() for line in out.splitlines() for m in ACE_RE.finditer(line)}


def listed_accounts(path):
    """icacls の出力から A\\xxx の個人アカウント名（小文字）を集める"""
    return parse_accounts(icacls(path, check=False).stdout)


def restrict_admins(path, log=print):
    """path の ACL から ADMINS 以外の個人アカウントを外す（継承を切ってから）"""
    icacls(path, "/inheritance:d")
    keep = KEEP | {a.lower() for a in ADMINS}
    removed = []
    for name in sorted(listed_accounts(path)):
        if name in keep:
            continue
        icacls(path, "/remove:g", f"{DOMAIN}\\{name}", check=False)
        removed.append(name)
    for a in ADMINS:                                            # 万一無ければ足す
        if a.lower() not in listed_accounts(path):
            icacls(path, "/grant", f"{DOMAIN}\\{a}:(OI)(CI)(F)")
    log(f"  {os.path.relpath(path, C.PROD_ROOT) or '.'}  管理者を {'・'.join(ADMINS)} に限定"
        + (f"（外した: {', '.join(removed)}）" if removed else "（変更なし）"))


def apply_root(log=print):
    root = C.PROD_ROOT
    restrict_admins(root, log)
    grant(root, "(RX)", log)                                   # このフォルダのみ（一覧・お読みください）
    rd = os.path.join(root, "_お読みください.txt")
    if os.path.isfile(rd):
        grant(rd, "(R)", log)
    # 名簿\ も読み取りだけ全員に（20260915 管理者指示: 科別ツールが共有の名簿を直接読む。
    # 同じ内容は各月の config\current.json（全員が読める）にも入っているので露出は増えない）
    for d in ("科別ツール", "ドキュメント", "名簿"):
        p = os.path.join(root, d)
        if os.path.isdir(p):
            grant(p, "(OI)(CI)(RX)", log)
    for dept in ("救急科", "形成外科"):
        out = os.path.join(C.PROD_TOOLS, dept, "out")
        if os.path.isdir(os.path.join(C.PROD_TOOLS, dept)):
            os.makedirs(out, exist_ok=True)
            grant(out, "(OI)(CI)(RX,W)", log)


def is_private(year, month):
    """管理者だけの月か（本番の最初の月より前）。"""
    return (year, month) < tuple(PRIVATE_BEFORE)


def private_steps(md):
    r"""管理者だけにする icacls の手順: 月フォルダ・master\・master\_履歴\ から Domain Users を外す
    （管理者3人は上のフォルダからの継承と restrict_admins で残る。集約\ は元から管理者だけ）。"""
    mst = os.path.join(md, "master")
    return [{"args": [p, "/remove:g", USERS], "check": False}
            for p in (md, mst, os.path.join(mst, "_履歴")) if os.path.isdir(p)]


def apply_private(md, log=print):
    for step in private_steps(md):
        icacls(*step["args"], check=step.get("check", True))
    for sub in ("master", os.path.join("master", "_履歴"), "集約"):
        p = os.path.join(md, sub)
        if os.path.isdir(p):
            restrict_admins(p, log)
    log(f"  {os.path.relpath(md, C.PROD_ROOT)}  管理者だけ（Domain Users を外した・試験と予行の月）")


def apply_month(year, month, log=print):
    md = C.prod_month_dir(year, month)
    if not os.path.isdir(md):
        raise SystemExit(f"月フォルダがありません: {md}")
    if is_private(year, month):                 # 9月・10月は管理者だけ（20260928）
        apply_private(md, log)
        return
    grant(md, "(OI)(CI)(RX)", log)
    # 集約\ = 管理者だけ（20260923 実測: 月フォルダの RX が継承で付いていたので、継承を切って Domain Users を外す）
    agg = os.path.join(md, "集約")
    os.makedirs(agg, exist_ok=True)
    icacls(agg, "/inheritance:d")
    icacls(agg, "/remove:g", USERS, check=False)
    restrict_admins(agg, log)
    log(f"  {os.path.relpath(agg, C.PROD_ROOT)}  管理者だけ（Domain Users なし）")

    # master\ = 全科の予定の唯一の正本（20260912）。アプリが上書きするので **変更（M）** が要る。
    # ロックファイルの作成と削除、一時ファイルからの差し替えに削除権を使うため、ここは削除を許す。
    # 代わりに master\_履歴\ を削除不可にして戻し道を守る（Modify があれば中身は元々壊せるので、
    # 削除だけ拒んでも実益が薄い。守るべきは履歴のほう）。
    for step in month_acl_steps(md):
        icacls(*step["args"], check=step.get("check", True))
        if step.get("say"):
            log(f"  {os.path.relpath(step['args'][0], C.PROD_ROOT)}  {step['say']}")
        if step.get("admins"):
            restrict_admins(step["args"][0], log)
    log(f"  {os.path.relpath(os.path.join(md, 'master', '_履歴'), C.PROD_ROOT)}  書込のみ（履歴は消せない）")


def month_acl_steps(md):
    r"""master\ と master\_履歴\ にかける icacls の手順。apply_month が順に流す（テストは中身だけ見る）。
    20260925 管理者「そんなことする人いないです。想定しないでOK」: _履歴\ の複製を書き換え不可にする案は採らず、
    従来どおり（master\ = M、_履歴\ = (OI)(CI)(RX,W)＝書けるが消せない）。試しにかけた OWNER RIGHTS の項目は外す。"""
    mst = os.path.join(md, "master")
    hist = os.path.join(mst, "_履歴")
    os.makedirs(mst, exist_ok=True)
    os.makedirs(hist, exist_ok=True)
    return [
        {"args": [mst, "/inheritance:d"]},
        {"args": [mst, "/remove:g", USERS], "check": False},
        {"args": [mst, "/remove:g", OWNER_RIGHTS], "check": False},
        {"args": [mst, "/grant", f"{USERS}:(OI)(CI)(M)"], "say": "← Domain Users (OI)(CI)(M)", "admins": True},
        {"args": [hist, "/inheritance:d"]},
        {"args": [hist, "/remove:g", CREATOR_OWNER], "check": False},
        {"args": [hist, "/remove:g", USERS], "check": False},
        {"args": [hist, "/remove:g", OWNER_RIGHTS], "check": False},
        {"args": [hist, "/grant", f"{USERS}:(OI)(CI)(RX,W)"], "say": "← Domain Users (OI)(CI)(RX,W)（書けるが消せない）",
         "admins": True},
    ]


def show(path):
    r = icacls(path, check=False)
    print(r.stdout)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", nargs="*", default=None, help="YYYY-MM …（省略=存在する全月フォルダ）")
    ap.add_argument("--root", action="store_true", help="ルート・科別ツール・ドキュメントだけ")
    ap.add_argument("--show", default=None)
    a = ap.parse_args()
    if a.show:
        show(a.show)
        return
    if not os.path.isdir(C.PROD_ROOT):
        raise SystemExit(f"本番共有に届きません: {C.PROD_ROOT}")
    print("■ ルート・共通")
    apply_root()
    if a.root:
        return
    months = []
    if a.month:
        for v in a.month:
            months.append((int(v[:4]), int(v[5:7])))
    else:
        for name in os.listdir(C.PROD_ROOT):
            ym = C.parse_month_dir(name)
            if ym:
                months.append(ym)
    for y, m in sorted(months):
        print(f"■ {C.month_dir_name(y, m)}")
        apply_month(y, m)
    if not a.month:
        for d in PRIVATE_DIRS:                  # 試験用フォルダ（月の名前ではないので上では拾わない）
            p = os.path.join(C.PROD_ROOT, d)
            if os.path.isdir(p):
                print(f"■ {d}")
                apply_private(p)


if __name__ == "__main__":
    main()
