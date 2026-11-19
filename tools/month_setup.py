# -*- coding: utf-8 -*-
"""month_setup.py — 本番共有に「月フォルダ」を作る／作り直す。20261028

ユーザー決定（2026-09-09）: 月ごとにフォルダを分ける（10月と11月が同時に活きる期間があるため）。
  当直勤務表\\
    _本体\\            当直勤務表入力(Win).exe（原本。deploy_prod --what master が置く）
                      touchoku_xlsx.exe（当直表xlsxのCLIヘルパー。無ければリポジトリの dist\\ から取る）
    名簿\\             staff_YYYYMMDD.csv（毎日更新。全部がここを見る）＋ 院外研修期間表
    科別ツール\\ 管理ツール\\
    R8年11月\\
      当直勤務表入力_R8年11月.exe   touchoku_xlsx.exe（アプリが子プロセスで呼ぶ）
      touchoku-settings.json（dataRoot="."・lockDir=00_DrJOY/当直勤務表/R8年11月）
      config\\current.json   master\\当直予定表_R8年11月.json（全科の正本）   master\\_履歴\\   集約\\
      _お読みください.txt

⚠ 2026-09-12 の xlsx一本化（docs\\設計_当直表xlsx一本化_20260912.md）で次を廃止した:
  提出\\ と提出JSON（マスターが正本）／テンプレ\\ の33科一括生成（アプリの「Excelに書き出す」で月ごとに作る）／
  Mac用フォーム HTML（Windows専用に凍結）／config\\確定.txt（マスターの confirmed に移した）

create(year, month)  … 無ければ作る。あれば config/ヘルパー/お読みくださいを作り直す（本体exe・マスター・集約は触らない）。
                       ACL は毎回かけ直す（20260923 第2陣。何度かけても同じ結果）
usage: python -X utf8 tools\\month_setup.py --year 2026 --month 11 [--deadline "10月18日(土)"] [--force-exe]
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C        # noqa: E402
import master_store as M    # noqa: E402
import make_month as MM     # noqa: E402  20260925: 子プロセスではなく同じプロセスで呼ぶ（管理ツールの exe でも動くように）

NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # 管理ツール（--windowed）から attrib を呼ぶときに黒い窓を出さない
REPLACE_WAITS = (0.5, 1, 2, 4)   # 差し替えが「使用中」で失敗したときの待ち（秒）。起動係が zip を展開している間など

EXE_MASTER = "当直勤務表入力(Win).exe"
EXE_LAUNCHER = "touchoku_launcher.exe"   # 起動係（20260925。月フォルダでは 当直勤務表入力_R8年N月.exe の名前で置く）
APP_ZIP = "touchoku_app.zip"              # 展開済みのアプリ（起動係が各 PC の %LOCALAPPDATA% へ1回だけ展開する）
APP_VER = "touchoku_app.ver"              # その版の名前（中身が同じなら同じ名前＝展開し直さない）
EXE_HELPER = "touchoku_xlsx.exe"      # 当直表xlsxの書き出し・取り込み（アプリが子プロセスで呼ぶ）


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def month_dir(year, month):
    return C.prod_month_dir(year, month)


def _hide(path):
    """隠しファイルにする（Windows。使う人が開かないファイル: xlsxヘルパー・settings）"""
    if os.name == "nt" and os.path.exists(path):
        subprocess.run(["attrib", "+h", path], capture_output=True, creationflags=NO_WINDOW)


def _unhide(path):
    if os.name == "nt" and os.path.exists(path):
        subprocess.run(["attrib", "-h", path], capture_output=True, creationflags=NO_WINDOW)


def helper_master():
    r"""ヘルパーの原本。本番の _本体\ にあればそれ、無ければリポジトリの dist\（ビルド直後）。"""
    for p in (os.path.join(C.PROD_MASTER, EXE_HELPER), os.path.join(C.ROOT, "dist", EXE_HELPER)):
        if os.path.isfile(p):
            return p
    return None


def master_json(year, month):
    return M.master_path(month_dir(year, month), year, month)


def is_closed(year, month):
    r"""確定済みか。確定の正本 集約\確定.json で決める（20260926。master_store.effective_confirmed と同じ規則）:
    正本があればそれに従う（マスターの confirmed は全科が書けるので見ない）／正本が無い旧い月はマスターの confirmed ／
    正本が壊れていれば確定とみなす（名簿の更新で作り直さない側に倒す）。"""
    p = master_json(year, month)
    try:
        t = M.read_truth(p)
    except M.TruthError:
        return True
    except OSError:
        return False
    if t is not None:
        return bool(t.get("confirmed"))
    if not os.path.isfile(p):
        return False
    try:
        return bool(M.load(p, year, month).get("confirmed"))
    except Exception:
        return False


def list_months():
    """本番にある月フォルダ → [(year, month, closed)] 昇順"""
    out = []
    if not os.path.isdir(C.PROD_ROOT):
        return out
    for name in os.listdir(C.PROD_ROOT):
        ym = C.parse_month_dir(name)
        if ym and os.path.isdir(os.path.join(C.PROD_ROOT, name, "config")):
            out.append((ym[0], ym[1], is_closed(*ym)))
    return sorted(out)


def default_target(today=None):
    """各科の締切=対象月の前月18日ごろ・翌々月の伺いは25日ごろから → 24日までは翌月、25日以降は翌々月"""
    today = today or dt.date.today()
    nxt = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
    if today.day >= 25:
        nxt = (nxt + dt.timedelta(days=32)).replace(day=1)
    return nxt.year, nxt.month


def readme_text(year, month, reiwa):
    return f"""当直勤務表 R{reiwa}年{month}月分（{year}年{month}月の当直）

■ 入力（各科）
1. このフォルダの「当直勤務表入力_R{reiwa}年{month}月.exe」をダブルクリック（その PC で初めて開くときと、アプリが新しくなったあとだけ、準備に30秒ほどかかります。2回目からは数秒で開きます。
   このフォルダから起動したときだけ動きます。コピーして別の場所から起動すると「起動できません」と出ます）
2. 画面の一番上の帯が「R{reiwa}年{month}月分」になっていることを確かめる
3. 科を選ぶ → 名簿とカレンダーに入れる → 右下の「この内容で提出する」。
   「提出」せずに閉じると、入力した内容は消えます（閉じる前に確かめる画面が出ます）。
   全科ぶんの予定は master\\当直予定表_R{reiwa}年{month}月.json 1本にまとまります（自科の分だけが書き換わります）
4. 月の途中で当直を交代したときも、アプリで直して出し直してください（交代申請書の代わりになります）。
   月が確定する前は、毎日15時にそのまま当直勤務表（DTX14）へ反映されます。
   月が確定したあと（病院長の決裁に出す直前に担当者が確定します）は、ボタンが「交代申請する」に変わり、変えた分は「交代申請」になります。
   申請は SharePoint の「当直勤務表_正本」の月のフォルダ（DTX14 と同じ所）の 当直交代申請.xlsx に載り（これが公式の交代記録です）、
   事務部がその「承認」欄で承認したものだけが、毎日15時に当直勤務表へ反映されます。
   却下された申請は、15時にアプリの予定も元に戻ります。

■ 管理当直
科の選択で「管理当直」を選ぶと入力できます（名簿は管理当直アプリの名簿と同じ並びです）。

■ Excel で作りたいとき
アプリの「Excelに書き出す」で当直表（そのまま印刷して壁に貼れます）をダウンロードし、記入して上書き保存したら「Excelから取り込む」で戻します。
取り込んだあとも「この内容で提出する」を押すまでは提出されません。
Excel のファイルを直接どこかへ置く必要はありません。科ごとのテンプレの配布は廃止しました。

■ 名簿（貴科の医師）
アプリの左の名簿は、この科の Dr.JOY の所属医師の申告です。
当直に立たない医師も含めて、この科に所属している医師を全員載せてください。
初めの顔ぶれは、Dr.JOY でその月の1日にこの科に所属している医師です。
医師ごとに、次の3つの区分から1つを選びます。
・フルタイム: 平日は毎日、当院で日中勤務する医師です（産休・育休・病休中の医師も含みます）。
・パートタイム: 当直や外来など、一部の仕事だけ当院に来る医師です（大学院生、院外研修中で外来や当直に来る医師など）。
・院外研修専従: 院外研修中で、当院の仕事はしていない医師です（この科の医師のままです）。
異動・退職・ほかの科へ移った医師と、留学・長期の派遣・研究だけの大学院生の医師は、×で名簿から外してください。
名簿と Dr.JOY の所属が違うときは、月末と毎週水曜に担当者へメールで知らせ、担当者が Dr.JOY を直します。
Dr.JOY のスタッフ一覧（医師）は毎日自動で更新されます。
Dr.JOY にまだ登録されていない医師は、アプリの検索欄から「名簿にない医師として追加」で入れられます。

■ 確定
病院長の決裁に出す当直勤務表を印刷する直前に、担当者が管理ツールの「この月を確定」を押します。
以後アプリで変えた分は「交代申請」になり（Excel では赤字になります）、承認されたものだけが当直勤務表に反映されます。
"""


def apply_acl(year, month, log=print, errors=None):
    r"""月フォルダの ACL（Domain Users は月フォルダ=RX、master\=M、master\_履歴\=書けるが消せない、集約\=管理者だけ。set_acl.py と同じ）。
    ACL は管理者権限が要る。失敗しても月フォルダ自体は使えるので、止めずに ⚠ を残す。
    set_acl は icacls の失敗を SystemExit で知らせるので、それも受け止める（20260923 第2陣: 以前は except Exception を
    すり抜けて create ごと止まっていた）。戻り: かけられたか。
    20260926: 失敗は errors（リスト）にも足す（new_month・15時の反映が警報メールに載せる）。ログの文は管理ツール①を押した人向け。"""
    try:
        import set_acl
        set_acl.apply_month(year, month, log=log)
        return True
    except (Exception, SystemExit) as e:
        log(f"⚠ ACL の設定に失敗しました（フォルダの権限。{e}）。もう一度この①を押してください。それでも出るときは管理者に連絡してください。")
        if errors is not None:
            errors.append(f"{C.month_dir_name(year, month)} のフォルダの権限（ACL）を設定できなかった（{e}）"
                          f"→ 管理者が tools\\set_acl.py --month {year}-{month:02d} を実行する")
        return False


def app_source():
    r"""起動係・アプリの zip・版の原本。本番の _本体\ に3つそろっていればそこ、無ければリポジトリの dist\（ビルド直後）。"""
    for d in (C.PROD_MASTER, os.path.join(C.ROOT, "dist")):
        if all(os.path.isfile(os.path.join(d, n)) for n in (EXE_LAUNCHER, APP_ZIP, APP_VER)):
            return d
    return None


COPIED, SAME, BUSY = "copied", "same", "busy"


def _replace_retry(tmp, dst, waits=REPLACE_WAITS):
    """os.replace を「使用中」（PermissionError）のあいだ少し待ってやり直す。戻り: 差し替えたか"""
    for w in (0,) + tuple(waits):
        if w:
            time.sleep(w)
        try:
            os.replace(tmp, dst)
            return True
        except PermissionError:
            continue
    return False


def _put_file(src, dst, hidden=False, waits=REPLACE_WAITS):
    """中身が違うときだけ、一時名で写してから差し替える（起動係が読んでいる途中の zip を壊さない）。
    戻り: COPIED＝写した／SAME＝同じだったので何もしない／BUSY＝使用中で差し替えられなかった（元のファイルのまま・一時ファイルは消す）。
    20260925: 起動係が zip を展開している最中は os.replace が PermissionError になる。以前は受け止めずに、
    配置・3:30 の名簿更新ごと止まり、148MB の .tmp が残っていた。"""
    if os.path.isfile(dst) and sha(src) == sha(dst):
        if hidden:
            _hide(dst)
        return SAME
    tmp = M.tmp_name(dst)                         # 20261002: 二重拡張子にしない（EDR がランサムウェア疑いで通報する）
    _unhide(dst)
    try:
        shutil.copy2(src, tmp)
        if sha(src) != sha(tmp):
            raise SystemExit(f"複写の検証に失敗: {dst}")
        if not _replace_retry(tmp, dst, waits):
            return BUSY
        return COPIED
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        if hidden:
            _hide(dst)


def put_app(md, reiwa, month, log=print, exe_name=None, waits=REPLACE_WAITS):
    r"""月フォルダにアプリを置く（20260925 管理者「起動速くなる？」）。
    当直勤務表入力_R8年N月.exe = 起動係（9KB）。隠しファイルで touchoku_app.zip（展開済みのアプリ）と touchoku_app.ver。
    起動係は各 PC の %LOCALAPPDATA%\Touchoku\app\<版>\ に zip を1回だけ展開し、2回目からはそれを起動する
    （起動が約19秒→約3秒。共有の exe は起動してすぐ閉じるので、アプリを開いたままでも差し替えられる）。
    zip を先に、版の名前をあとに置く（版だけ新しくて zip が古い、という瞬間を作らない）。
    zip が使用中（誰かの起動係が展開している途中）で差し替えられなければ、古い zip のまま・版の名前も変えずに ⚠ を残して続ける
    （版だけ新しくすると、起動係が古い zip を新しい版の名前で展開してしまう。次の配置か、名簿が変わった日の 3:30 の名簿更新でまた試す）。
    戻り: 起動係を置いたか（False=原本が無い）"""
    src = app_source()
    if not src:
        return False
    exe = os.path.join(md, exe_name or f"当直勤務表入力_R{reiwa}年{month}月.exe")
    zr = _put_file(os.path.join(src, APP_ZIP), os.path.join(md, APP_ZIP), hidden=True, waits=waits)
    if zr == COPIED:
        log(f"配置: {APP_ZIP}（{src} から）")
    if zr == BUSY:
        log(f"⚠ {APP_ZIP} が使用中で差し替えられない（だれかがアプリを起動した直後）→ 古い版のまま。"
            "版の名前も変えない（次の配置か、名簿が変わった日の 3:30 の名簿更新でもう一度試します。"
            "急ぐときは管理ツールの①をもう一度押してください）")
    elif _put_file(os.path.join(src, APP_VER), os.path.join(md, APP_VER), hidden=True, waits=waits) == COPIED:
        log(f"配置: {APP_VER}（版 {open(os.path.join(src, APP_VER), encoding='ascii').read().strip()}）")
    try:
        r = _put_file(os.path.join(src, EXE_LAUNCHER), exe, waits=waits)
    except PermissionError:          # 一時ファイルを作れない（フォルダの権限）など
        r = BUSY
    if r == COPIED:
        log(f"配置: {os.path.basename(exe)}（起動係）")
    elif r == BUSY:
        log(f"⚠ {os.path.basename(exe)} が使用中で差し替えられない（旧版の本体 exe を開いている人がいる。閉じたらもう一度）")
    # 20260925: xlsx ヘルパーも同じ原本から配る（deploy_prod --what months・3:30 の名簿更新で、活きている月へ行き渡る）
    hsrc = os.path.join(src, EXE_HELPER)
    if os.path.isfile(hsrc):
        try:
            hr = _put_file(hsrc, os.path.join(md, EXE_HELPER), hidden=True, waits=waits)
        except PermissionError:
            hr = BUSY
        if hr == COPIED:
            log(f"配置: {EXE_HELPER}（{src} から）")
        elif hr == BUSY:
            log(f"⚠ {EXE_HELPER} が使用中で差し替えられない（Excel の書き出し・取り込みの最中。次の配置でもう一度）")
    return True


def create(year, month, deadline=None, force_exe=False, log=print, warnings=None):
    """warnings（リスト）を渡すと、月フォルダは作れたが知らせることを足す（20260926: ACL の失敗。new_month・15時の反映が警報に載せる）。"""
    reiwa = year - 2018
    md = month_dir(year, month)
    exists = os.path.isdir(os.path.join(md, "config"))
    if not os.path.isdir(C.PROD_ROOT):
        raise SystemExit(f"本番共有に届きません: {C.PROD_ROOT}")
    master = os.path.join(C.PROD_MASTER, EXE_MASTER)
    if not app_source() and not os.path.isfile(master):
        raise SystemExit(f"本体の原本がありません: {C.PROD_MASTER}（deploy_prod.py --what master で置く）")
    helper = helper_master()
    if not helper:
        raise SystemExit(f"{EXE_HELPER} がありません（build_kyukyu_exe.bat でビルドし、_本体\\ へ置いてください）")
    staffs = C.newest_staffs()
    if not os.path.isfile(staffs):
        raise SystemExit("名簿（staff_YYYYMMDD.csv）が見つかりません")
    for d in ("config", "master", os.path.join("master", "_履歴"), "集約"):     # 提出\\ は廃止（20260923。マスターが正本）
        os.makedirs(os.path.join(md, d), exist_ok=True)
    # 締切: 既存 config を引き継ぐ
    cj = os.path.join(md, "config", "current.json")
    if deadline is None and os.path.isfile(cj):
        with open(cj, encoding="utf-8") as f:
            deadline = str(json.load(f).get("deadline") or "")
    deadline = deadline or ""
    # 月次生成（20260925: make_month を同じプロセスで呼び、月フォルダの config へ直に書く。以前は sys.executable で
    # make_month.py を子プロセスで走らせていたので、管理ツールの exe（sys.executable＝管理ツール自身・make_month.py は
    # 同梱されていない）では ① が動かなかった。make_month の ⚠（管理当直の結び付かない人など）は成功しても log に出す）
    kikan = C.find_kikan()
    try:
        cfg = MM.build_config(year, month, deadline, staffs, kikan, log=log,
                              prev_fixed=MM.read_fixed_members(cj))
    except Exception as e:
        raise SystemExit(f"月次生成に失敗: {type(e).__name__}: {e}") from e
    log(f"月次生成: R{reiwa}年{month}月（名簿 {os.path.basename(staffs)}・日数{len(cfg['days'])} 枠{len(cfg['frames'])}"
        f" 名簿{len(cfg['roster'])}名）")
    # 配置（20260912: テンプレ33科と Mac用HTML の一括配置は廃止。Excel はアプリから月ごとに作る）
    try:
        MM.write_config(cfg, cj)
    except OSError as e:
        raise SystemExit(f"配置に失敗: {cj}（{e}）") from e
    log("配置: config/current.json")
    # アプリ = 起動係＋zip（20260925）。原本が無いときだけ、従来の自己展開 exe を置く（無いときだけ複写。--force-exe で置き直し）
    exe = os.path.join(md, f"当直勤務表入力_R{reiwa}年{month}月.exe")
    put = put_app(md, reiwa, month, log)
    if put:
        pass
    elif force_exe or not os.path.isfile(exe):
        shutil.copy2(master, exe)
        if sha(master) != sha(exe):
            raise SystemExit(f"exe の複写の検証に失敗: {exe}")
        log(f"配置: {os.path.basename(exe)}（原本から複写）")
    # xlsxヘルパー（アプリが隣から呼ぶ）。中身が変わっていれば置き直す。
    # 使う人が開くものではないので隠しファイルにする（20260915 管理者指示）。
    # 20260926: put_app が同じ原本（_本体\ の起動係の隣）から置いたときは、ここで置き直さない
    # （以前は shutil.copy2 でもう一度写し、アプリが使用中だと PermissionError で create ごと止まっていた）。
    # put_app の原本にヘルパーが無いときだけ、一時名→差し替え（_put_file）で置く。使用中なら ⚠ を残して続ける
    hx = os.path.join(md, EXE_HELPER)
    src_dir = app_source() if put else None
    if not (src_dir and os.path.isfile(os.path.join(src_dir, EXE_HELPER))):
        try:
            hr = _put_file(helper, hx, hidden=True)
        except PermissionError:
            hr = BUSY
        if hr == COPIED:
            log(f"配置: {EXE_HELPER}（{helper} から複写）")
        elif hr == BUSY:
            log(f"⚠ {EXE_HELPER} が使用中で差し替えられない（Excel の書き出し・取り込みの最中。次の配置でもう一度）")
    # 全科の正本（無ければ空で作る。あるものには触らない）
    mp = M.master_path(md, year, month)
    if not os.path.isfile(mp):
        M.init(mp, year, month)
        log(f"作成: master\\{os.path.basename(mp)}（空のマスター）")
    settings = {"dataRoot": ".", "lockDir": f"00_DrJOY/当直勤務表/{C.month_dir_name(year, month)}"}
    sp = os.path.join(md, "touchoku-settings.json")
    _unhide(sp)
    with open(sp, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
    _hide(sp)                                   # 使う人が開くものではない（20260915 管理者指示）
    with open(os.path.join(md, "_お読みください.txt"), "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(readme_text(year, month, reiwa))
    # ACL: 毎回かける（20260923 第2陣。以前は初回だけ＝初回に途中で落ちると、2回目からは exists=True で永久に掛からなかった）。
    # set_acl.apply_month は何度かけても同じ結果（icacls の /grant・/remove:g・/inheritance:d だけ。/T なし）で、icacls 十数回で済む
    # （20260926: 失敗は warnings にも足す。selftest は create の中の『apply_acl(year, month, log)』の呼び出しを見る）
    apply_acl(year, month, log) if warnings is None else apply_acl(year, month, log, errors=warnings)
    log(("作り直し" if exists else "新規作成") + f": {md}")
    return md


_AUTO = object()


def refresh_roster_only(year, month, staffs=None, history_dir=None, kikan=_AUTO, log=print):
    r"""確定済みの月の config\current.json の roster（アプリが持っている Dr.JOY所属の写し）だけを、いまの名簿で作り直す。
    20261028 管理者: 決裁後も、あとから Dr.JOY に登録された先生を検索で足せるように・欠けた日の名簿で止まらないように。
    締切・日・枠・当番当直・管理当直の初期メンバー・exe・予定表には触らない（作り直すのは roster だけ）。
    戻り: 書き換えたか（変わらなければ False）。"""
    md = month_dir(year, month)
    cj = os.path.join(md, "config", "current.json")
    with open(cj, encoding="utf-8-sig") as f:
        cur = json.load(f)
    staffs = staffs or C.newest_staffs()
    if not os.path.isfile(staffs):
        raise SystemExit("名簿（staff_YYYYMMDD.csv）が見つかりません")
    new = MM.build_config(year, month, str(cur.get("deadline") or ""), staffs,
                          C.find_kikan() if kikan is _AUTO else kikan, log=lambda t: None,
                          prev_fixed=MM.read_fixed_members(cj), history_dir=history_dir)
    if new.get("roster") == cur.get("roster"):
        return False
    old_n = len(cur.get("roster") or [])
    cur["roster"] = new["roster"]
    MM.write_config(cur, cj)
    log(f"名簿の更新（確定済みの月・roster だけ）: {C.month_dir_name(year, month)}  {old_n} 名 → {len(cur['roster'])} 名")
    return True


def set_closed(year, month, closed=True, who=""):
    r"""確定する／解除する。20260912 から マスターの confirmed が正本（config\確定.txt は廃止）。
    確定では各科の baseline に そのときの duties を焼き付ける（以後の変更＝当直交代）。
    既に確定している月は確定し直さない（20260923。解除してからなら付け直せる）。"""
    p = master_json(year, month)
    if not os.path.isfile(p):
        raise SystemExit(f"マスターがありません: {p}（先に月フォルダを作ってください）")
    try:
        (M.confirm if closed else M.unconfirm)(p, by=who, year=year, month=month)
    except (M.AlreadyConfirmed, M.MasterMissing) as e:
        raise SystemExit(str(e))
    return p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=0)
    ap.add_argument("--month", type=int, default=0)
    ap.add_argument("--deadline", default=None)
    ap.add_argument("--force-exe", action="store_true")
    ap.add_argument("--list", action="store_true")
    a = ap.parse_args()
    if a.list:
        for y, m, c in list_months():
            print(f"{y}/{m}{'（確定）' if c else ''}")
        return
    if not (a.year and a.month):
        a.year, a.month = default_target()
    create(a.year, a.month, a.deadline, a.force_exe)


if __name__ == "__main__":
    main()
