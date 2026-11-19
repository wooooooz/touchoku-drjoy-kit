# -*- coding: utf-8 -*-
r"""wp_publish.py — 公開用の当直勤務表 PDF を院内掲示板（WordPress）の「当直勤務表」ページに載せる。20261027

管理者 2026-10-25: 掲示板には担当係（事務担当2）が Excel から書き出した PDF を手で貼っていて、チャットへのリンクが無く、
  内容も古いままのことがあった。掲示板は医療系と院外系の両方の端末から見るので、SharePoint へのリンクでは開けない
  → 公開用の PDF（公開用\R8年N月 病院当直勤務表.pdf。1本に2ページ＝ファースト・セカンド以下）を WordPress に
  アップロードし、ページには月ごとにリンクを1つ置く（埋め込まない。2本に分けない）。
毎日15時の反映（dtx14_daily_sync）の最後に呼ぶ。中身（SHA-256）が前に載せたものから変わった月だけアップロードし直し、
  ページを書き直し、前に自分が上げたファイルを消す（ほかの人が上げたファイルは消さない）。
ページは丸ごとこのツールが書く。手で直した分は次の書き換えで消える（ページの先頭に印のコメントを置く）。
  印が無いページ（人が書いたページ）は書き換えない。初めて取って代わるときだけ --takeover を付けて手で流す。
入口: https://intranet.example.ac.jp/index.php?rest_route=/wp/v2/...（/wp-json/ は 404。パーマリンクが既定のため）。
  認証はアプリケーションパスワード（Basic）。資格情報は管理者の PC の %LOCALAPPDATA%\Touchoku\wp_cred.bin に DPAPI で暗号化して置く
  （--setup で本人が打ち込む）。無ければ何もしない。

usage:
  python -X utf8 tools\wp_publish.py --setup             資格情報を入れる（本人が打つ）
  python -X utf8 tools\wp_publish.py --dry-run           書き換えず、載せる中身と差を出す
  python -X utf8 tools\wp_publish.py [--takeover]        載せる（ふだんは15時の反映が呼ぶ）
"""
import argparse
import base64
import datetime as dt
import glob
import hashlib
import html
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = "https://intranet.example.ac.jp"
PAGE_ID = None                    # 当直勤務表のページの番号（自院の WordPress で決める）
# 「当直勤務表（今月版）」＝今月の PDF に飛ぶだけのページ（20261025 管理者: スマートフォンから固定 URL で今月の当直表を開きたい）。
#   PDF の URL は上げ直すたびに変わるので、URL の変わらないこのページが、開いた端末の日付で今月の PDF に移る（JavaScript）。
#   毎月1日の 8:30 から新しい月に飛ぶ（20261026。それまでは 0:00）。JavaScript が動かない端末ではリンクを出す。None なら作らない・書かない
REDIRECT_PAGE_ID = None          # 「今月版」のページの番号（None なら作らない）
REDIRECT_MARK = "touchoku-auto-redirect"
SWITCH_MS = (8 * 60 + 30) * 60 * 1000  # 今月版が新しい月に切り替わる時刻＝毎月1日の 8:30（ミリ秒。20261026 管理者）
MAX_MONTHS = 3                         # 新しい方から何か月を載せるか（来月・今月・前月）
MARK = "touchoku-auto"
RESPONSIBLE = "文責：担当係"
LOCAL = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku")
CRED_PATH = os.path.join(LOCAL, "wp_cred.bin")
STATE_PATH = os.path.join(LOCAL, "wp_publish_state.json")
PDF_RE = re.compile(r"R(\d+)年(\d+)月 病院当直勤務表\.pdf")
TIMEOUT = 60


class WPError(Exception):
    pass


# ------------------------------------------------------------------ 資格情報（DPAPI）

def save_creds(user, password, path=CRED_PATH):
    import win32crypt
    os.makedirs(os.path.dirname(path), exist_ok=True)
    blob = win32crypt.CryptProtectData(json.dumps({"user": user, "password": password}).encode("utf-8"),
                                       "touchoku-wp", None, None, None, 0)
    with open(path, "wb") as f:
        f.write(blob)


def load_creds(path=CRED_PATH):
    if not os.path.isfile(path):
        return None
    import win32crypt
    with open(path, "rb") as f:
        _desc, data = win32crypt.CryptUnprotectData(f.read(), None, None, None, 0)
    return json.loads(data.decode("utf-8"))


# ------------------------------------------------------------------ HTTP

class Client:
    def __init__(self, creds, base=BASE):
        self.base = base
        tok = base64.b64encode(f"{creds['user']}:{creds['password']}".encode("utf-8")).decode("ascii")
        self.auth = "Basic " + tok

    def url(self, route, **q):
        u = f"{self.base}/index.php?rest_route={urllib.parse.quote(route, safe='/')}"
        return u + "".join(f"&{k}={urllib.parse.quote(str(v))}" for k, v in q.items())

    def call(self, method, route, body=None, headers=None, **q):
        h = {"Authorization": self.auth, "User-Agent": "touchoku-app"}
        data = None
        if isinstance(body, (bytes, bytearray)):
            data = bytes(body)
        elif body is not None:
            data = json.dumps(body).encode("utf-8")
            h["Content-Type"] = "application/json"
        h.update(headers or {})
        req = urllib.request.Request(self.url(route, **q), data=data, method=method, headers=h)
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                raw = r.read()
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")[:300]
            raise WPError(f"{method} {route}: HTTP {e.code} {detail}") from e
        except (urllib.error.URLError, OSError) as e:
            raise WPError(f"{method} {route}: 掲示板に届かない（{e}）") from e
        return json.loads(raw.decode("utf-8")) if raw else None

    def get_page_raw(self, page_id):
        p = self.call("GET", f"/wp/v2/pages/{page_id}", context="edit", _fields="id,content,modified")
        return (p.get("content") or {}).get("raw") or ""

    def update_page(self, page_id, content):
        return self.call("POST", f"/wp/v2/pages/{page_id}", {"content": content})

    def upload_pdf(self, path, filename, title):
        with open(path, "rb") as f:
            data = f.read()
        m = self.call("POST", "/wp/v2/media", data,
                      {"Content-Type": "application/pdf", "Content-Disposition": f'attachment; filename="{filename}"'})
        try:
            self.call("POST", f"/wp/v2/media/{m['id']}", {"title": title})
        except WPError:
            pass                        # 題が付かなくてもリンクは使える
        return m["id"], m["source_url"]

    def delete_media(self, media_id):
        self.call("DELETE", f"/wp/v2/media/{media_id}", force="true")


# ------------------------------------------------------------------ 中身

HOLD_PATH = os.path.join(LOCAL, "wp_hold.json")   # {"2026-11": "2026-10-26"}＝その月をその日まで載せない（20261026 事務担当の要望）


def held_months(today=None, path=HOLD_PATH):
    """今日まだ載せない月（{"YYYY-MM": "YYYY-MM-DD まで"} の日付が今日以降のもの）。"""
    today = today or dt.date.today()
    try:
        with open(path, encoding="utf-8") as f:
            h = json.load(f)
    except (OSError, ValueError):
        return set()
    return {k for k, until in h.items() if str(until) >= today.isoformat()}


def public_pdfs(public_dir, max_months=MAX_MONTHS, today=None, hold_path=HOLD_PATH):
    """公開用フォルダの月ごとの PDF → [(年, 月, パス)]（新しい月から max_months 件。載せない月は除く）。"""
    out = []
    hold = held_months(today, hold_path)
    for p in glob.glob(os.path.join(public_dir, "R*年*月 病院当直勤務表.pdf")):
        m = PDF_RE.fullmatch(os.path.basename(p))
        if m and ym_key(int(m.group(1)) + 2018, int(m.group(2))) not in hold:
            out.append((int(m.group(1)) + 2018, int(m.group(2)), p))
    return sorted(out, key=lambda t: (t[0], t[1]), reverse=True)[:max_months]


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def ym_key(y, m):
    return f"{y:04d}-{m:02d}"


def build_html(entries):
    """entries = [(年, 月, url, 更新日時 datetime)]（新しい月から）→ ページの本文（クラシックエディタの HTML）。"""
    # 20261025 管理者「必要最小限の表示に。Table of Contents も要らない」→ 見出し（h2）を使わない（目次のプラグインは見出しから作る）。
    #   月ごとのリンクの箇条書きと、説明と、文責だけ
    # 20261027 管理者が掲示板で説明の文言を直した（整形済みテキスト＝<pre> で2行）。その文言どおりに書く
    parts = [f"<!-- {MARK}: このページは当直勤務表アプリが毎日15時に自動で書き換えます。手で直した分は次の書き換えで消えます -->",
             "<pre>毎日15時更新。1ページ目がファースト、2ページ目がセカンド以下です。\n"
             "医師名から Teams のチャットに飛べます（Teams が使える端末に限る）</pre>",
             "<ul>"]
    for y, m, url, when in entries:
        title = f"R{y - 2018}年{m}月 当直勤務表"
        parts.append(f'<li><a href="{html.escape(url, quote=True)}" target="_blank" rel="noopener">{html.escape(title)}</a>'
                     f"（{when.month}月{when.day}日 {when.hour}時更新）</li>")
    parts += ["</ul>", f'<p style="text-align: right;">{html.escape(RESPONSIBLE)}</p>']
    return "\n".join(parts)


def build_redirect_html(entries, main_url=f"{BASE}/?page_id={PAGE_ID}"):
    """「当直勤務表（今月版）」の本文。開いた端末の日付の月の PDF にすぐ移る。その月が無ければ一覧を出したまま。
    スクリプトは1行に書く（クラシックエディタの自動整形が改行に <br> や <p> を入れないように）。"""
    urls = {f"{y:04d}-{m:02d}": url for y, m, url, _w in entries}
    # 20261026 管理者: 月の切り替えは1日の 0:00 ではなく 8:30（月末の夜の当直が明けるまでは前の月の表）。今の時刻から 8時間30分を引いた日時の月
    js = ("(function(){var m=" + json.dumps(urls, ensure_ascii=False) + ";var d=new Date(Date.now()-" + str(SWITCH_MS) + ");"
          "var k=d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2);var u=m[k];"
          "if(u){var a=document.getElementById('touchoku-now');if(a){a.href=u;}location.replace(u);}})();")
    links = "".join(f'<li><a href="{html.escape(url, quote=True)}">R{y - 2018}年{m}月 当直勤務表</a>'
                    f"（{when.month}月{when.day}日 {when.hour}時更新）</li>" for y, m, url, when in entries)
    return "\n".join([
        f"<!-- {REDIRECT_MARK}: 今月の当直勤務表の PDF に移るだけのページ。当直勤務表アプリが毎日15時に書き換えます -->",
        f'<p>今月の当直勤務表を開いています。開かないときは <a id="touchoku-now" href="{html.escape(main_url, quote=True)}">こちら</a> から選んでください。</p>',
        f"<ul>{links}</ul>",
        f"<script>{js}</script>",
    ])


def load_state(path=STATE_PATH):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"months": {}}


def save_state(st, path=STATE_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = os.path.join(os.path.dirname(path), f"~wp_state_{os.getpid()}.tmp")   # 単一拡張子（EDR）
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def publish(public_dir=None, dry=False, takeover=False, log=print, now=None, client=None,
            state_path=STATE_PATH, cred_path=CRED_PATH):
    """載せる。戻り: "nocreds" / "nochange" / "updated" / "dry"。書けないときは WPError。"""
    if public_dir is None:
        import dtx14_daily_sync as DS
        public_dir = DS.PUBLIC_DIR
    now = now or dt.datetime.now()
    pdfs = public_pdfs(public_dir, today=now.date(), hold_path=os.path.join(os.path.dirname(state_path), "wp_hold.json")
                       if state_path != STATE_PATH else HOLD_PATH)
    st = load_state(state_path)
    months = st.setdefault("months", {})
    if client is None:
        creds = load_creds(cred_path)
        if creds is None and not dry:
            log("  掲示板: 資格情報が無いので載せない（wp_publish.py --setup）")
            return "nocreds"
        client = Client(creds) if creds else None
    todo = []                               # 上げ直す月
    for y, m, p in pdfs:
        k = ym_key(y, m)
        cur = months.get(k) or {}
        if cur.get("sha") != sha256(p) or not cur.get("url"):
            todo.append((y, m, p))
    entries_preview = [(y, m, (months.get(ym_key(y, m)) or {}).get("url") or "(これから上げる)",
                        now if (y, m, p) in todo else dt.datetime.fromisoformat((months.get(ym_key(y, m)) or {}).get("when") or now.isoformat()))
                       for y, m, p in pdfs]
    if dry:
        log(f"  掲示板（dry-run）: 載せる月 {[ym_key(y, m) for y, m, _p in pdfs]}・上げ直す月 {[ym_key(y, m) for y, m, _p in todo]}")
        log(build_html(entries_preview))
        if client:
            raw = client.get_page_raw(PAGE_ID)
            log(f"  今のページに自動の印: {'あり' if MARK in raw else 'なし（--takeover が要る）'}・本文 {len(raw)} 文字")
        return "dry"
    raw = client.get_page_raw(PAGE_ID)
    if MARK not in raw and not takeover:
        raise WPError("掲示板の当直勤務表のページに自動の印が無い（人が書き直した？）→ 書き換えない。"
                      "取って代わってよければ wp_publish.py --takeover を手で流す")
    old_ids, new_ids = [], []
    try:
        return _upload_and_write(client, pdfs, todo, months, st, raw, now, old_ids, new_ids, state_path, log)
    except WPError:
        for mid in new_ids:                 # ページを書けなかった → 今回上げたファイルは消す（次の回に上げ直す）
            try:
                client.delete_media(mid)
            except WPError:
                pass
        raise


def _upload_and_write(client, pdfs, todo, months, st, raw, now, old_ids, new_ids, state_path, log):
    for y, m, p in todo:
        k = ym_key(y, m)
        fname = f"touchoku_R{y - 2018}-{m:02d}_{now:%Y%m%d-%H%M}.pdf"
        mid, url = client.upload_pdf(p, fname, f"R{y - 2018}年{m}月 病院当直勤務表（{now:%Y-%m-%d %H:%M}）")
        new_ids.append(mid)
        if (months.get(k) or {}).get("media_id"):
            old_ids.append(months[k]["media_id"])
        made = dt.datetime.fromtimestamp(os.path.getmtime(p))        # 掲示に出す「更新」＝公開用の PDF を作った時刻
        months[k] = {"sha": sha256(p), "media_id": mid, "url": url, "when": made.isoformat(timespec="minutes")}
        log(f"  掲示板: R{y - 2018}年{m}月 を上げた（{url}）")
    entries = [(y, m, months[ym_key(y, m)]["url"], dt.datetime.fromisoformat(months[ym_key(y, m)]["when"]))
               for y, m, _p in pdfs]
    body = build_html(entries)
    changed = body.strip() != raw.strip()
    if changed:
        client.update_page(PAGE_ID, body)
    rchanged = _write_redirect(client, entries, log)     # 前のファイルを消す前に、今月版のページも新しい URL に
    new_ids.clear()                          # 書けた（か書く必要が無い）ので、今回上げたファイルは残す
    if not changed and not todo and not rchanged:
        save_state(st, state_path)
        log("  掲示板: 変わりなし")
        return "nochange"
    # 載せなくなった月（古くなった月）の、自分が上げたファイルも消す
    keep = {ym_key(y, m) for y, m, _p in pdfs}
    for k in [k for k in months if k not in keep]:
        if months[k].get("media_id"):
            old_ids.append(months[k]["media_id"])
        del months[k]
    save_state(st, state_path)
    for mid in old_ids:
        try:
            client.delete_media(mid)
        except WPError as e:
            log(f"  ⚠ 掲示板: 前のファイル（{mid}）を消せない: {e}")
    log(f"  掲示板: ページを書き換えた（{len(entries)} か月・上げ直し {len(todo)}）")
    return "updated"


def _write_redirect(client, entries, log):
    """「当直勤務表（今月版）」のページを書き直す（変わったときだけ）。印の無いページ（人が書き直した）は書かずに WPError。"""
    if not REDIRECT_PAGE_ID:
        return False
    raw = client.get_page_raw(REDIRECT_PAGE_ID)
    if raw.strip() and REDIRECT_MARK not in raw:
        raise WPError(f"「当直勤務表（今月版）」のページ（{REDIRECT_PAGE_ID}）に自動の印が無い（人が書き直した？）→ 書き換えない")
    body = build_redirect_html(entries)
    if body.strip() == raw.strip():
        return False
    client.update_page(REDIRECT_PAGE_ID, body)
    log("  掲示板: 今月版のページを書き換えた")
    return True


def setup():
    import getpass
    print("院内掲示板（WordPress）の当直勤務表ページを書き換えるアカウントを入れます。")
    print("アプリケーションパスワードは、WordPress のプロフィール画面で作ったもの（空白は入れたままでよい）。")
    user = input("ユーザー名: ").strip()
    # 20261025: getpass の入力欄には貼り付けられない（Ctrl+V が文字として入る）ので、クリップボードから読む。読んだら空にする
    input("アプリケーションパスワードをコピーしてから、ここで Enter を押してください（画面には出しません）: ")
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        pw = (win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT) or "").strip()
        win32clipboard.EmptyClipboard()
    finally:
        win32clipboard.CloseClipboard()
    if not pw:
        raise SystemExit("クリップボードが空でした。パスワードをコピーしてからやり直してください")
    print(f"クリップボードから読みました（{len(pw)} 文字。クリップボードは空にしました）")
    c = Client({"user": user, "password": pw})
    try:
        me = c.call("GET", "/wp/v2/users/me", context="edit", _fields="id,name,roles,capabilities")
        raw = c.get_page_raw(PAGE_ID)
    except WPError as e:
        raise SystemExit(f"確かめられませんでした（入れていません）: {e}")
    save_creds(user, pw)
    print(f"入れました: {me.get('name')}（ページ {PAGE_ID} を読めた・{len(raw)} 文字）")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--setup", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--takeover", action="store_true", help="自動の印の無いページ（人が書いたページ）を取って代わる（初回だけ）")
    a = ap.parse_args()
    if a.setup:
        return setup()
    try:
        print(publish(dry=a.dry_run, takeover=a.takeover))
    except WPError as e:
        raise SystemExit(f"✗ {e}")


if __name__ == "__main__":
    main()
