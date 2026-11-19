# -*- coding: utf-8 -*-
r"""dtx14_manual_import.py — 手入力の部署（看護部・薬剤部・検査部・放射線部など）の行を、各部署の DTX14 の写しから
その月の正本 DTX14（SharePoint の同期フォルダ）へ写す。管理ツール④「手入力の部署を取り込む」の中身。20261010

20260925 追記（公開前の見直し）:
  - 1つのファイルに、その部署のほかに別の部署の行（古い写し）も入っていることがある。以前はファイルに入っている
    手入力の行を全部写し、同じ行が2つのファイルにあれば先に選んだファイルが勝った＝並べた順で他部署の行を古い値に戻せた。
    → ファイルごとに「どの部署を取り込むか」を選ぶ（choose）。選ばなかった部署の行は写さない。
      既定で選ぶのは、そのファイルでその部署の行が空でなく、しかも正本と違う部署だけ。
      同じ部署を2つ以上のファイルが既定で選びそうなときは、どれも選ばない（人が選ぶ）。
      人が同じ行を2つのファイルで選んだら、計画の conflicts に出し、その行はどちらからも写さない（管理ツールは書かせない）。
  - 14:55〜15:15 は正本へ書かない（15時の反映が別の PC から同じ正本を書くため。20261010 に 14:30〜15:30 から縮めた＝反映は2分かからない）。SharePoint の同期フォルダの正本だけが対象
    （selftest などの手元の写しには掛けない）。

管理者 2026-09-25 決定:
  - アプリが扱わない行（看護部・薬剤部・検査技師・放射線技師・管理 など）は、今までどおり各部署が DTX14 の形の Excel
    （10月までの形でも、11月からの昼/夜を2行に割った形でもよい）で送ってくる。その部署の行だけを正本へ写す。
  - アプリが書く行（frames.json の枠＝各科・管理当直）には決して触らない（15時の反映 dtx14_fill が書く）。
  - 行は行番号ではなく（A列の部署名, C列の見出し）で突き合わせる（10月の形と11月の形で行の位置が違う）。
    A列が結合されていて部署名が上のセルにしか無い行は、直前の部署の続きとして読む（dtx14_extract.locate_rows と同じ）。
    同じ見出しが2つある部署（救急外来看護師の「日勤」×2）は、上から何番目かで突き合わせる。
    送られてきた方が「全日」1行で正本が「昼」「夜」の2行（またはその逆）なら、推測で割らずに「合わない」と返す。
  - 送られてきたファイルの月（4行目 D列の日付シリアル。無ければ3行目・A1）が正本の月と違えば、そのファイルは取り込まない。
  - 日の列は4行目（無ければ3行目）の日付で対応させる（列の文字で決め打ちしない）。
  - 送られてきた行に日のマスが1つでも入っていれば、正本のその行の日のマスを丸ごと置き換える（空の日は空にする）。
    行がまるごと空なら正本の行はそのまま（まだ書いていない部署の行を消さない）。
  - 書き方は15時の反映と同じ: 正本を _上書き履歴\ へ写す → ローカルの写しに Excel COM（マクロは走らせない）で書く →
    読み戻して、写した行だけが変わり、ほかのマスは1つも変わっていないことを確かめる → 通れば正本を差し替える
    （dtx14_fill.publish）。正本が Excel で開かれている（~$）／書いている間に更新された ときは差し替えない。

使い方（GUI を通さないとき）:
  python -X utf8 tools\dtx14_manual_import.py --month 2026-11 FILE1.xlsm [FILE2.xlsx ...] [--target 正本.xlsm] [--dry-run]
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys
import tempfile
import warnings

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402

SHEETS = ("当直表 ファースト", "当直表 セカンド以下")
ROW_SERIAL = 4                # 日付シリアルの行（dtx14_fill と同じ）
ROW_DATE = 3                  # 日付の行（シリアルが無いときの予備）
COL_DAY1 = 4                  # D列
FIRST_ROW = 5                 # 表の本体
MAX_DAY_COLS = 40             # D列から右へ日付を探す幅
EXTS = (".xlsx", ".xlsm")


class ManualImportError(Exception):
    """取り込めない（正本が見つからない・開かれている・確かめが合わない など）。正本は元のまま。"""


# ------------------------------------------------------------------ 正本の場所
DTX_DIR_NAME = "当直勤務表_正本"
MONTH_DIR_RE = re.compile(r"R(\d+)\.(\d+)月 当直勤務表")
SETTINGS = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku", "admin_settings.json")


def _looks_like_root(p):
    """「当直勤務表_正本」らしいか（月のフォルダ『R8.N月 当直勤務表』が1つ以上ある）。"""
    try:
        return os.path.isdir(p) and any(MONTH_DIR_RE.fullmatch(n) for n in os.listdir(p))
    except OSError:
        return False


def remembered_root():
    try:
        import json
        with open(SETTINGS, encoding="utf-8") as f:
            return json.load(f).get("dtx_root") or None
    except (OSError, ValueError):
        return None


def remember_root(p):
    """管理ツール④で人が選んだ「当直勤務表_正本」を覚える（%LOCALAPPDATA%\\Touchoku\\admin_settings.json）。"""
    import json
    d = {}
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        pass
    d["dtx_root"] = p
    os.makedirs(os.path.dirname(SETTINGS), exist_ok=True)
    with open(SETTINGS, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)


def _sync_mount_points():
    """OneDrive が同期している場所（HKCU\\Software\\SyncEngines\\Providers\\OneDrive\\*\\MountPoint）。"""
    out = []
    try:
        import winreg
        base = r"Software\SyncEngines\Providers\OneDrive"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base) as k:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                try:
                    with winreg.OpenKey(k, sub) as s:
                        mp = winreg.QueryValueEx(s, "MountPoint")[0]
                        if mp and mp not in out:
                            out.append(mp)
                except OSError:
                    pass
    except Exception:
        pass
    return out


def _find_under(base, depth):
    """base の下 depth 段までで「当直勤務表_正本」を探す（ファイルの中身は読まない＝オンデマンドのまま）。"""
    level = [base]
    for _ in range(depth + 1):
        nxt = []
        for d in level:
            if os.path.basename(d) == DTX_DIR_NAME and _looks_like_root(d):
                return d
            try:
                with os.scandir(d) as it:
                    for e in it:
                        if e.is_dir(follow_symlinks=False) and not e.name.startswith((".", "~")):
                            nxt.append(e.path)
            except OSError:
                pass
        level = nxt
        if len(level) > 5000:
            break
    return None


def dtx_roots():
    """SharePoint の「当直勤務表_正本」の候補。15時の反映は管理者の PC（user003）の同期フォルダを見るが、
    管理ツールは事務担当の PC で動くので、使っている人のホームの下の同じ場所も探す。
    20261009: 事務担当の PC で④の月が出なかった（同期のしかたが違うと場所が違う＝ライブラリごと同期・
    OneDrive のショートカットなど）。人が選んだ場所（remember_root）と、OneDrive が同期している場所の下も探す。"""
    import dtx14_daily_sync as S
    out = []
    r = remembered_root()
    if r:
        out.append(r)
    parts = os.path.normpath(S.DTX_ROOT).split(os.sep)
    if len(parts) > 3 and parts[1].lower() == "users":
        mine = os.path.join(os.path.expanduser("~"), *parts[3:])
        if os.path.normcase(mine) != os.path.normcase(S.DTX_ROOT):
            out.append(mine)
    out.append(S.DTX_ROOT)
    return out


_SEARCH = {}          # この起動中に一度だけ探す（20261009 レビュー: 見つからないとタブを替えるたびに探し直していた）


def search_roots():
    """OneDrive の同期先とホームの OneDrive* の下にある「当直勤務表_正本」を全部（重複なし）。"""
    home = os.path.expanduser("~")
    bases = _sync_mount_points()
    try:
        bases += [os.path.join(home, n) for n in os.listdir(home)
                  if n.startswith("OneDrive") or n == "Contoso"]
    except OSError:
        pass
    hits = []
    for b in bases:
        hit = _find_under(b, 4)
        if hit and os.path.normcase(os.path.realpath(hit)) not in {os.path.normcase(os.path.realpath(h)) for h in hits}:
            hits.append(hit)
    return hits


def dtx_root():
    """使える「当直勤務表_正本」のフォルダ（無ければ None）。決まった候補に無ければ、OneDrive の同期先の下を探す。
    見つかったのが1か所だけなら覚える。2か所以上なら覚えず None（④で人に選んでもらう。20261009 レビュー:
    古い写しやショートカットを先に見つけて、同期していない場所を覚え続けるおそれがあった）。"""
    for p in dtx_roots():
        if _looks_like_root(p):
            return p
    if "hits" not in _SEARCH:
        _SEARCH["hits"] = search_roots()
        if len(_SEARCH["hits"]) == 1:
            try:
                remember_root(_SEARCH["hits"][0])
            except OSError:
                pass
    hits = _SEARCH["hits"]
    return hits[0] if len(hits) == 1 else None


def month_folder(y, m, root=None):
    import dtx14_daily_sync as S
    return os.path.join(root or dtx_root() or S.DTX_ROOT, f"R{y - 2018}.{m}月 当直勤務表")


def official_path(y, m, root=None):
    """その月の正本 DTX14（名前の規則 dtx14_daily_sync.is_master_name に合う .xlsm がちょうど1本）。
    戻り: パス。無い・2本以上 → ManualImportError（dtx_file と同じ規則。ログは共有へ書かない）。"""
    import dtx14_daily_sync as S
    d = month_folder(y, m, root)
    if not os.path.isdir(d):
        raise ManualImportError(f"R{y - 2018}.{m}月の当直勤務表のフォルダがありません: {d}")
    hits = sorted(n for n in os.listdir(d)
                  if n.lower().endswith(".xlsm") and not n.startswith("~$") and S.is_master_name(n, y, m))
    if not hits:
        raise ManualImportError(f"R{y - 2018}.{m}月の正本（★R{y - 2018}.{m}月用 DTX14….xlsm）がありません: {d}")
    if len(hits) > 1:
        raise ManualImportError(f"正本の名前のファイルが2本以上あるので、どれに書くか決められません: {'、'.join(hits)}（{d}）")
    return os.path.join(d, hits[0])


def list_months(root=None):
    """「当直勤務表_正本」にある月 [(年, 月)] 昇順。"""
    root = root or dtx_root()
    out = []
    if not root or not os.path.isdir(root):
        return out
    for n in os.listdir(root):
        mt = MONTH_DIR_RE.fullmatch(n)
        if mt and os.path.isdir(os.path.join(root, n)):
            out.append((int(mt.group(1)) + 2018, int(mt.group(2))))
    return sorted(out)


def locked(path):
    return os.path.exists(os.path.join(os.path.dirname(path), "~$" + os.path.basename(path)))


QUIET_FROM, QUIET_TO = dt.time(14, 55), dt.time(15, 15)   # 15時の反映（管理者の PC・2分かからない）＋ OneDrive の行き来の余裕（20261010）


def in_quiet_window(now=None):
    """14:55〜15:15 か（この間は正本へ書かない）。"""
    t = (now or dt.datetime.now()).time()
    return QUIET_FROM <= t < QUIET_TO


QUIET_MESSAGE = ("いまは 14時55分〜15時15分 なので取り込めません。\n"
                 "この時間は、15時の自動処理が同じ当直勤務表（DTX14）に書き込んでいます。\n"
                 "15時15分を過ぎてから、もう一度「取り込む」を押してください。")


def is_official(path):
    """SharePoint の「当直勤務表_正本」の下のファイルか（正本。14:55〜15:15 の書き込み止めの対象）。"""
    p = os.path.normcase(os.path.abspath(path))
    for r in dtx_roots():
        r = os.path.normcase(os.path.abspath(r))
        if p.startswith(r.rstrip("\\/") + os.sep):
            return True
    return False


# ------------------------------------------------------------------ 読む（openpyxl・読み取りだけ）
def load_grids(path):
    """{シート名: 値の2次元リスト（1始まりの行 r は grid[r-1]）}。当直表の2シートのうちあるものだけ。"""
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")           # 「Conditional Formatting extension is not supported」（読むときに出る）
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            out = {}
            for sh in SHEETS:
                if sh in wb.sheetnames:
                    ws = wb[sh]
                    try:
                        ws.reset_dimensions()
                    except Exception:
                        pass
                    out[sh] = [list(r) for r in ws.iter_rows(values_only=True)]
            return out
        finally:
            wb.close()


def cellv(grid, r, c):
    if 0 < r <= len(grid):
        row = grid[r - 1]
        if 0 < c <= len(row):
            return row[c - 1]
    return None


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 < v < 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    return None


def book_month(grids):
    """その表の月 (年, 月)。4行目 D列の日付シリアル → 3行目 D列の日付 → A1 の順に見る（dtx14_fill は D4 を確かめる）。
    シートごとに違えば None（壊れている）。"""
    seen = set()
    for grid in grids.values():
        for r, c in ((ROW_SERIAL, COL_DAY1), (ROW_DATE, COL_DAY1), (1, 1)):
            d = _as_date(cellv(grid, r, c))
            if d:
                seen.add((d.year, d.month))
                break
    return seen.pop() if len(seen) == 1 else None


def day_cols(grid, y, m):
    """{日: 列}。4行目（無ければ3行目）の日付のうち、その月のものだけ（30日の月の AH列＝翌月1日は入れない）。"""
    for r in (ROW_SERIAL, ROW_DATE):
        out = {}
        for c in range(COL_DAY1, COL_DAY1 + MAX_DAY_COLS):
            d = _as_date(cellv(grid, r, c))
            if d and (d.year, d.month) == (y, m) and d.day not in out:
                out[d.day] = c
        if out:
            return out
    return {}


def _key(s):
    return C.namekey(C.nfc(str(s or "").strip())).replace("\n", "")


def row_blocks(grid, start=FIRST_ROW):
    """表の行 → [{row, dept, label, k, lk, occ}]。C列に見出し（昼・夜・全日・深夜・日勤…）がある行だけ。
    A列の部署名が空の行は直前の部署の続き（A列の結合は上のセルにだけ値がある）。A列もC列も空の行で続きは切れる
    （dtx14_extract.locate_rows と同じ読み方）。同じ部署に同じ見出しが2つあれば occ=1,2…（上から何番目か）。"""
    out, cur, seen = [], None, {}
    for r in range(start, len(grid) + 1):
        a = str(cellv(grid, r, 1) or "").strip()
        c = str(cellv(grid, r, 3) or "").strip()
        if a:
            cur = a
        elif not c:
            cur = None
            continue
        if not cur or not c:
            continue
        k, lk = _key(cur), _key(c)
        seen[(k, lk)] = seen.get((k, lk), 0) + 1
        out.append({"row": r, "dept": cur.replace("\n", " "), "label": c, "k": k, "lk": lk, "occ": seen[(k, lk)]})
    return out


def app_keys(frames=None):
    """アプリが書く部署 {シート: {部署名のキー}}（frames.json の枠。消滅した科・管理当直も含む）。"""
    frames = C.load_frames() if frames is None else frames
    out = {sh: set() for sh in SHEETS}
    for f in frames:
        sh = f.get("sheet")
        if sh in out and f.get("row"):
            out[sh].add(_key(f.get("dept_dtx14") or f.get("dept") or ""))
    return out


def app_rows(frames, grids):
    """アプリが書く行 {(シート, 行)}（dtx14_fill.frame_rows で見出しから引いた行。昼夜の2行とも）。部署名での判定の念押し。"""
    import dtx14_extract as X
    import dtx14_fill as F
    loc = {sh: X.locate_rows(g, sh) for sh, g in grids.items()}
    out = set()
    for f in frames:
        if f.get("sheet") in grids and f.get("row"):
            rr = F.frame_rows(f, loc)
            if rr:
                out.add((f["sheet"], rr[0]))
                out.add((f["sheet"], rr[1]))
    return out


def manual_rows(grids, frames=None):
    """手入力の行 [{sheet, row, dept, label, k, lk, occ}]（アプリが書く行を除いた、C列に見出しのある部署の行）。"""
    frames = C.load_frames() if frames is None else frames
    keys = app_keys(frames)
    rows = app_rows(frames, grids)
    out = []
    for sh in SHEETS:
        if sh not in grids:
            continue
        for b in row_blocks(grids[sh]):
            if b["k"] in keys[sh] or (sh, b["row"]) in rows:
                continue
            out.append(dict(b, sheet=sh))
    return out


def _blank(v):
    return v is None or (isinstance(v, str) and not v.strip())


def _norm(v):
    if _blank(v):
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        return v.replace("\r\n", "\n")
    return v


# ------------------------------------------------------------------ 計画
def _choice_key(p):
    return os.path.normcase(os.path.abspath(p))


def plan_import(src_paths, target_path, frames=None, year=None, month=None, choose=None):
    """送られてきたファイル → 正本へ写す計画。正本には触らない（読むだけ）。
    choose = {ファイルのパス: {部署のキー, …}}（どの部署を取り込むか。20260925）。None／そのファイルが無ければ既定
    （行が空でなく、正本と違う部署。2つ以上のファイルが既定で同じ部署を選ぶときは、どれも選ばない）。
    戻り: {target, year, month, ndays, writes:[{sheet,row,dept,label,cells:{列: 値}, filled, src}],
           conflicts:[(シート, 部署, 見出し, [ファイル名…])],
           files:[{path, name, ok, reason, copied:[…], unmatched:[…], skipped:[…], notchosen:[…], app:[部署名…],
                   depts:[{key, dept, labels, filled, differs, default, chosen, shared:[ほかのファイル名…]}]}]}"""
    frames = C.load_frames() if frames is None else frames
    try:
        tg = load_grids(target_path)
    except Exception as e:
        raise ManualImportError(f"正本を読めません: {target_path}（{type(e).__name__}: {e}）") from e
    missing = [sh for sh in SHEETS if sh not in tg]
    if missing:
        raise ManualImportError(f"正本に「{'」「'.join(missing)}」のシートがありません: {target_path}")
    ym = book_month(tg)
    if not ym:
        raise ManualImportError(f"正本の月が読めません（4行目 D列の日付）: {target_path}")
    if year and month and ym != (year, month):
        raise ManualImportError(f"正本の月が R{ym[0] - 2018}.{ym[1]}月で、選んだ月（R{year - 2018}.{month}月）と違います: {target_path}")
    y, m = ym
    import calendar
    ndays = calendar.monthrange(y, m)[1]
    keys = app_keys(frames)
    t_cols = {sh: day_cols(tg[sh], y, m) for sh in SHEETS}
    for sh in SHEETS:
        if sorted(t_cols[sh]) != list(range(1, ndays + 1)):
            raise ManualImportError(f"正本の「{sh}」の4行目の日付が1日〜{ndays}日そろっていません: {target_path}")
    t_manual = manual_rows(tg, frames)
    t_index = {(b["sheet"], b["k"], b["lk"], b["occ"]): b for b in t_manual}
    t_labels = {}
    for b in t_manual:
        t_labels.setdefault((b["sheet"], b["k"]), []).append(b["label"])
    plan = {"target": target_path, "year": y, "month": m, "ndays": ndays, "writes": [], "files": [], "conflicts": [],
            "manual": [(b["sheet"], b["row"], b["dept"], b["label"]) for b in t_manual]}
    cand = []                                # 1つ目の段: ファイルごとに、正本と合う空でない行を集める（まだ写すと決めない）
    for p in src_paths:
        name = os.path.basename(p)
        res = {"path": p, "name": name, "ok": False, "reason": "", "copied": [], "unmatched": [], "skipped": [],
               "notchosen": [], "app": [], "depts": []}
        plan["files"].append(res)
        rows_of = []
        cand.append((res, rows_of))
        if os.path.normcase(os.path.abspath(p)) == os.path.normcase(os.path.abspath(target_path)):
            res["reason"] = "正本そのものです（各部署から送られてきたファイルを選んでください）"
            continue
        if os.path.splitext(p)[1].lower() not in EXTS:
            res["reason"] = "xlsx か xlsm ではありません（Excel で「Excel マクロ有効ブック」か「Excel ブック」で保存し直してください）"
            continue
        try:
            sg = load_grids(p)
        except Exception as e:
            res["reason"] = f"読めません（{type(e).__name__}: {e}）"
            continue
        if not sg:
            res["reason"] = "「当直表 ファースト」「当直表 セカンド以下」のシートがありません（DTX14 の形ではありません）"
            continue
        sym = book_month(sg)
        if not sym:
            res["reason"] = "何月の表か読めません（4行目 D列の日付が空です）"
            continue
        if sym != (y, m):
            res["reason"] = f"R{sym[0] - 2018}.{sym[1]}月の表です（正本は R{y - 2018}.{m}月）→ 取り込みません"
            continue
        res["ok"] = True
        app_seen = []
        for sh, grid in sg.items():
            s_cols = day_cols(grid, y, m)
            if not s_cols:
                res["unmatched"].append((sh, "", "", "4行目の日付が読めないのでこのシートは取り込みません"))
                continue
            miss = [d for d in range(1, ndays + 1) if d not in s_cols]
            if miss:
                # 20261009 レビュー: 4行目に無い日を「空」とみなすと、正本のその日を消してしまう
                res["unmatched"].append((sh, "", "", f"4行目の日付が欠けています（{'・'.join(str(d) for d in miss[:5])}日"
                                                     f"{'ほか' if len(miss) > 5 else ''}）→ このシートは取り込みません"))
                continue
            for b in row_blocks(grid):
                where = f"{b['dept']}（{b['label']}）"
                if b["k"] in keys[sh]:
                    if b["dept"] not in app_seen:
                        app_seen.append(b["dept"])
                    continue
                vals = {d: _norm(cellv(grid, b["row"], s_cols[d])) if d in s_cols else None
                        for d in range(1, ndays + 1)}
                filled = sum(1 for v in vals.values() if v is not None)
                if not filled:
                    res["skipped"].append((sh, b["dept"], b["label"], "空の行（正本はそのまま）"))
                    continue
                tb = t_index.get((sh, b["k"], b["lk"], b["occ"]))
                if tb is None:
                    have = t_labels.get((sh, b["k"]))
                    if not have:
                        why = f"正本の「{sh}」に「{b['dept']}」の行がありません"
                    elif b["occ"] > 1 and b["label"] in have:
                        why = f"「{b['label']}」の行が正本より多い（正本は {have.count(b['label'])} 行）"
                    else:
                        why = (f"見出しが合いません（送られてきた方は「{b['label']}」、正本は「{'／'.join(have)}」）"
                               "→ 推測で割り振らないので、正本と同じ見出しの形で送ってもらってください")
                    res["unmatched"].append((sh, b["dept"], b["label"], why))
                    continue
                cells = {t_cols[sh][d]: v for d, v in vals.items()}
                differs = any(_norm(cellv(tg[sh], tb["row"], c)) != v for c, v in cells.items())
                rows_of.append({"key": b["k"], "sheet": sh, "tb": tb, "cells": cells, "filled": filled,
                                "differs": differs, "where": where})
        res["app"] = app_seen
        if not (rows_of or res["unmatched"] or res["skipped"]):
            res["reason"] = "手入力の部署の行が見つかりません（アプリが書く科の行だけでした）"

    # 2つ目の段: ファイルごとの部署の一覧と既定の選び方
    for res, rows_of in cand:
        by = {}
        for x in rows_of:
            d = by.setdefault(x["key"], {"key": x["key"], "dept": x["tb"]["dept"], "labels": [], "filled": 0,
                                         "differs": False, "default": False, "chosen": False, "shared": []})
            d["labels"].append(x["tb"]["label"])
            d["filled"] += x["filled"]
            d["differs"] = d["differs"] or x["differs"]
        for d in by.values():
            d["default"] = d["differs"]
        res["depts"] = list(by.values())
    owners = {}                              # 部署のキー → それを既定で選ぶファイル
    for res, _r in cand:
        for d in res["depts"]:
            if d["default"]:
                owners.setdefault(d["key"], []).append(res)
    for key, rs in owners.items():
        if len(rs) > 1:                      # 同じ部署が2つ以上のファイルで正本と違う → どれも既定では選ばない
            for res in rs:
                for d in res["depts"]:
                    if d["key"] == key:
                        d["default"] = False
                        d["shared"] = [o["name"] for o in rs if o is not res]
    picks = {_choice_key(k): set(v) for k, v in (choose or {}).items()}

    # 3つ目の段: 選んだ部署の行だけを写す。同じ行を2つのファイルで選んでいたら、どちらからも写さない
    chosen_rows = []
    for res, rows_of in cand:
        pick = picks.get(_choice_key(res["path"]))
        for d in res["depts"]:
            d["chosen"] = (d["key"] in pick) if pick is not None else d["default"]
        on = {d["key"] for d in res["depts"] if d["chosen"]}
        for x in rows_of:
            if x["key"] in on:
                chosen_rows.append((res, x))
            else:
                res["notchosen"].append((x["sheet"], x["tb"]["dept"], x["tb"]["label"], x["filled"], x["differs"]))
    who = {}
    for res, x in chosen_rows:
        who.setdefault((x["sheet"], x["tb"]["row"]), []).append(res["name"])
    for (sh, _row), names in who.items():
        if len(names) > 1:
            tb = next(x["tb"] for _res, x in chosen_rows if (x["sheet"], x["tb"]["row"]) == (sh, _row))
            plan["conflicts"].append((sh, tb["dept"], tb["label"], names))
    for res, x in chosen_rows:
        sh, tb = x["sheet"], x["tb"]
        if len(who[(sh, tb["row"])]) > 1:
            others = [n for n in who[(sh, tb["row"])] if n != res["name"]]
            res["unmatched"].append((sh, tb["dept"], tb["label"],
                                     f"ほかのファイル（{'、'.join(others)}）でも選んでいるので写しません"
                                     "（どちらか1つだけで選んでください）"))
            continue
        plan["writes"].append({"sheet": sh, "row": tb["row"], "dept": tb["dept"], "label": tb["label"],
                               "cells": x["cells"], "filled": x["filled"], "src": res["name"]})
        res["copied"].append((sh, tb["dept"], tb["label"], x["filled"], x["where"]))
    return plan


def plan_text(plan):
    """計画 → 人が読む結果の一覧（ログ欄に出す）。"""
    out = [f"正本: {plan['target']}（R{plan['year'] - 2018}.{plan['month']}月）"]
    groups = {}
    for _sh, _r, dept, label in plan.get("manual") or []:
        groups.setdefault(dept, []).append(label)
    if groups:
        out.append("正本の手入力の行: " + "、".join(f"{d}（{'・'.join(ls)}）" for d, ls in groups.items()))
    out.append("")
    for f in plan["files"]:
        out.append(f"■ {f['name']}")
        if not f["ok"]:
            out.append(f"  ✗ 取り込みません: {f['reason']}")
            out.append("")
            continue
        if f["reason"]:
            out.append(f"  ⚠ {f['reason']}")
        if f.get("depts"):
            out.append("  取り込む部署: " + ("、".join(d["dept"] for d in f["depts"] if d["chosen"]) or "（なし）"))
        for sh, dept, label, n, _w in f["copied"]:
            out.append(f"  ○ 写す: {dept}（{label}）{n}日分  ［{sh}］")
        for sh, dept, label, n, differs in f.get("notchosen") or []:
            out.append(f"  － 選んでいないので写さない: {dept}（{label}）"
                       + ("" if differs else "  ※正本と同じ") + f"  ［{sh}］")
        for sh, dept, label, why in f["unmatched"]:
            out.append(f"  ✗ 合わない: {dept}（{label}）→ {why}  ［{sh}］")
        for sh, dept, label, why in f["skipped"]:
            out.append(f"  － 空なので飛ばす: {dept}（{label}）  ［{sh}］")
        for d in f.get("depts") or []:
            if d["shared"] and not d["chosen"]:
                out.append(f"  ⚠ {d['dept']} は、ほかのファイル（{'、'.join(d['shared'])}）にも正本と違う内容があるので、"
                           "最初は選んでいません。どちらのファイルから取り込むか選んでください")
        if f["app"]:
            out.append(f"  （アプリが書く科の行 {len(f['app'])} 部署は見ていません。管理当直もアプリが書きます）")
        out.append("")
    for sh, dept, label, names in plan.get("conflicts") or []:
        out.append(f"✗ 同じ行を2つ以上のファイルで選んでいます: {dept}（{label}）← {'、'.join(names)}  ［{sh}］"
                   " → どれか1つのファイルだけで選び直してください")
    out.append(f"写す行: {len(plan['writes'])} 行")
    return "\n".join(out)


# ------------------------------------------------------------------ 書く
def _compare(before, after, writes, ndays_cols):
    """読み戻しの確かめ。写した行の日のマスは計画どおり、それ以外（両シートの5行目から下・A〜日付の最後の列）は元のまま。
    戻り: 食い違いの一覧（空なら OK）。"""
    want = {(w["sheet"], w["row"], c): v for w in writes for c, v in w["cells"].items()}
    bad = []
    for sh in SHEETS:
        b, a = before.get(sh) or [], after.get(sh) or []
        last_c = ndays_cols[sh]
        for r in range(FIRST_ROW, max(len(b), len(a)) + 1):
            for c in range(1, last_c + 1):
                exp = want[(sh, r, c)] if (sh, r, c) in want else _norm(cellv(b, r, c))
                got = _norm(cellv(a, r, c))
                if exp != got:
                    bad.append((sh, r, c, exp, got))
    return bad


def write_com(path, writes, log=print):
    """ローカルの写し path に Excel COM で値だけ書く（マクロ・書式・入力規則は触らない。マクロは走らせない）。"""
    import win32com.client
    import dtx14_fill as F
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3          # msoAutomationSecurityForceDisable
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=False)
        n = 0
        for w in writes:
            ws = wb.Worksheets(w["sheet"])
            for c, v in sorted(w["cells"].items()):
                cell = ws.Cells(w["row"], c)
                if v is None:
                    cell.ClearContents()
                else:
                    cell.Value = F.xl_text(v)   # = + - @ で始まる文字列を数式にしない（20260926）
                n += 1
        wb.Save()
        wb.Close(SaveChanges=False)
        return n
    finally:
        xl.Quit()


def apply_import(plan, target_path, log=print, backup=True, frames=None, now=None):
    """計画どおりに正本へ写す。正本を _上書き履歴\\ へ写す → ローカルの写しに書く → 読み戻して確かめる → 正本を差し替える。
    戻り: {rows, cells, backup}。何もできないときは ManualImportError（正本は元のまま）。
    SharePoint の正本には 14:55〜15:15 は書かない（15時の反映と重ねない。20260925・20261010 に縮めた）。"""
    import dtx14_daily_sync as S
    import dtx14_fill as F
    writes = plan["writes"]
    if plan.get("conflicts"):
        raise ManualImportError("同じ行を2つ以上のファイルで選んでいます。どれか1つのファイルだけで選び直してください")
    if not writes:
        raise ManualImportError("写す行がありません")
    if in_quiet_window(now) and is_official(target_path):
        raise ManualImportError(QUIET_MESSAGE)
    if os.path.normcase(os.path.abspath(plan["target"])) != os.path.normcase(os.path.abspath(target_path)):
        raise ManualImportError("計画を作った正本と書く先が違います")
    if locked(target_path):
        raise ManualImportError(f"正本がいま Excel で開かれています（閉じてからもう一度押してください）: {target_path}")
    S.hydrate(target_path)
    st0 = os.stat(target_path)
    work_dir = tempfile.mkdtemp(prefix="dtx14_manual_")
    work = os.path.join(work_dir, "work" + os.path.splitext(target_path)[1])
    try:
        shutil.copy2(target_path, work)
        before = load_grids(work)
        # 計画のあとで正本が変わっていないか（15時の反映が走った など）: 手入力の行の位置が同じか
        if [(b["sheet"], b["row"], b["dept"], b["label"]) for b in manual_rows(before, frames)] != plan["manual"]:
            raise ManualImportError("計画を作ったあとで正本の行の形が変わりました。もう一度「取り込む」を押してください")
        bak = S.backup(target_path) if backup else ""
        if bak:
            log(f"退避: {bak}")
        n = write_com(work, writes, log)
        after = load_grids(work)
        last = {sh: max(day_cols(before[sh], plan["year"], plan["month"]).values()) for sh in SHEETS}
        bad = _compare(before, after, writes, last)
        if bad:
            raise ManualImportError(f"書いたあとの確かめが合いません（{len(bad)} マス。例 {bad[:3]}）→ 正本は差し替えていません")
        log(f"確かめ: 写した {len(writes)} 行のほかは1マスも変わっていない ✅")
        if locked(target_path):
            raise ManualImportError("書いている間に正本が Excel で開かれました → 差し替えていません（閉じてからもう一度）")
        st = os.stat(target_path)
        if (st.st_mtime, st.st_size) != (st0.st_mtime, st0.st_size):
            raise ManualImportError("書いている間に正本が更新されました → 差し替えていません（もう一度押してください）")
        if in_quiet_window() and is_official(target_path):      # 書いている間に 14:55 を回った
            raise ManualImportError(QUIET_MESSAGE)
        try:
            F.publish(work, target_path, log)
        except F.FillError as e:
            raise ManualImportError(str(e)) from e
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    log(f"正本を差し替えました: {target_path}（{len(writes)} 行 / {n} マス）")
    return {"rows": len(writes), "cells": n, "backup": bak}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM")
    ap.add_argument("--target", default=None, help="正本（省略=SharePoint の同期フォルダの その月の正本）")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--dept", action="append", default=None,
                    help="取り込む部署（A列の名前。何度でも。どのファイルにも同じに効く）。省略=既定（空でなく正本と違う部署）")
    ap.add_argument("files", nargs="+")
    a = ap.parse_args()
    y, m = (int(x) for x in a.month.split("-"))
    try:
        target = a.target or official_path(y, m)
        choose = {f: {_key(d) for d in a.dept} for f in a.files} if a.dept else None
        plan = plan_import(a.files, target, year=y, month=m, choose=choose)
        print(plan_text(plan))
        if not a.dry_run and plan["writes"]:
            apply_import(plan, target)
    except ManualImportError as e:
        raise SystemExit(str(e))
    return 0


if __name__ == "__main__":
    sys.exit(main())
