# -*- coding: utf-8 -*-
r"""make_month.py  20261031
月次生成器: staffs.csv ＋ 枠マスタ ＋ 祝日表 → 対象月の一式を作る。

20260927（管理者: 名簿 = Dr.JOY の所属医師の申告）:
  roster[].dept は、対象月の1日の Dr.JOY の所属（shozoku.dept_on）。初期研は所属履歴、それ以外の人は今日の所属
  （staff CSV）に、今日より後に始まる所属履歴の組（日付付きの予定。他院を除く）を重ねる（20261028）。
  院外研修期間表は計画で実際と違う人が多いので、初期値（ingai_week5 / ingai_note）には使わない（--kikan は受けて無視する）。
  config に "shozoku"（所属を決める提出単位・外す提出単位・当直をしない部署・外す職位）を載せる（アプリが同じ決まりで見る）。

  python -X utf8 tools/make_month.py --year 2026 --month 10 --deadline "9月20日(金) 13:00"
      [--staffs <staffs.csv>] [--out <出力先>] [--legacy-templates]

生成物（--out 既定 out\R{reiwa}年{m}月\）:
  config\current.json     … アプリが読む当月設定（年月・締切・日一覧・枠・名簿・科マップ）

⚠ 2026-09-12 の xlsx一本化で、**33科の Excel テンプレと Mac用フォーム HTML の一括生成は既定で止めた**
  （設計 = docs\設計_当直表xlsx一本化_20260912.md §8）。
  Excel は各科がアプリの「Excelに書き出す／取り込む」で必要な月ぶんだけ作る（ヘルパー tools\touchoku_xlsx.py）。
  Mac は凍結。旧来の一括生成が要るときだけ `--legacy-templates` を付ける
  （関数 write_template_xlsx / build_mac_html は selftest が使うので残してある）。

※共有(UNC)へは書かない。生成物を検分してから人手で配置する。
  （管理ツール①・15時の自動作成・3:30 の名簿更新は month_setup が build_config / write_config を同じプロセスで呼び、
   月フォルダの config へ直に書く。20260925: 以前は sys.executable で make_month.py を子プロセスで走らせていたが、
   管理ツールの exe では sys.executable が管理ツール自身で make_month.py も同梱されていないので ① が動かなかった）
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402
import bucho_roster  # noqa: E402
import shozoku as SZ  # noqa: E402

WD_JP = ["月", "火", "水", "木", "金", "土", "日"]


def read_fixed_members(config_path):
    """いまの config の fixedMembers（無い・読めなければ None）。管理当直の名簿が読めないときに引き継ぐため（20260925）。"""
    try:
        with open(config_path, encoding="utf-8-sig") as f:
            fx = json.load(f).get("fixedMembers")
        return fx if isinstance(fx, dict) else None
    except (OSError, ValueError, AttributeError):
        return None


NAIKA_IPPAN = "内科"
NAIKA_IPPAN_FILE = os.path.join(C.DATA, "naika_ippan_members.json")


def ingai_logins(year, month):
    """その月（無ければ前の月）の共有マスターで、どこかの「貴科リスト」で院外研修専従になっている医師のログイン（小文字）。"""
    import master_store as MS
    out = set()
    for y, m in ((year, month), (year - 1, 12) if month == 1 else (year, month - 1)):
        try:
            mst = MS.load(MS.master_path(C.prod_month_dir(y, m), y, m), y, m)
        except Exception:
            continue
        got = {str(x.get("login") or "").lower() for e in (mst.get("depts") or {}).values()
               for x in e.get("roster") or [] if C.work_of(x) == "ingai"}
        if any((e.get("roster") or []) for e in (mst.get("depts") or {}).values()):
            return got                            # その月に名簿があれば、その月の区分で決める
        out |= got
    return out


def naika_ippan_members(roster, dept_map, year, month, log=print, named=None, ingai=None):
    """内科の初期メンバー [{login, name}]（管理者 20261031）。
    専攻医 = 内科の所属（dept_map の4科）に当たる専攻医の全員から、院外研修専従を除く。
    教員・大学院生 = data/naika_ippan_members.json に名指しした人（R8.9〜11 に内科の当直に入った人）だけ。
    初期研とその他の職位は入れない。"""
    pats, _ranks = C.dept_map_entry(dept_map.get(NAIKA_IPPAN))
    if named is None:
        with open(NAIKA_IPPAN_FILE, encoding="utf-8-sig") as f:
            named = {str(x["login"]).lower() for x in json.load(f).get("members") or []}
    ingai = ingai_logins(year, month) if ingai is None else ingai
    out, n_sen, n_named, n_ingai = [], 0, 0, 0
    for r in roster:
        if not r.get("doctor", True) or not any(t and t in (r.get("dept") or "") for t in pats):
            continue
        lg = str(r.get("login") or "").lower()
        rank = C.rank_of(r.get("title"))
        if rank == "専":
            if lg in ingai:
                n_ingai += 1
                continue
            n_sen += 1
        elif rank in ("教", "研") and lg in named:
            n_named += 1
        else:
            continue
        out.append({"login": r["login"], "name": r.get("name") or ""})
    log(f"  内科の初期メンバー: 専攻医 {n_sen} 名（院外研修専従 {n_ingai} 名を除く）＋ 教員・大学院生 {n_named} 名")
    return out


def build_config(year, month, deadline, staffs_path, kikan_path=None, log=print, prev_fixed=None, history_dir=None):
    """当月設定（config/current.json の中身）を作って返す。書かない（書くのは write_config）。
    log = 経過と ⚠ の出し先（month_setup は自分のログへ流す＝成功したときも ⚠ が見える。20260925）。
    prev_fixed = いまの config の fixedMembers。管理当直の名簿（roster.csv）が読めないときは空にせずこれを引き継ぐ。
    history_dir = 所属履歴_YYYY.xlsx の置き場（省略=本番の 名簿フォルダ。shozoku.HISTORY_DIR）。"""
    frames = C.load_frames()
    holidays = C.load_holidays()
    oncall_only = C.daytime_oncall_only()
    people, exact, canon_map = C.load_staffs(staffs_path)
    with open(os.path.join(C.DATA, "dept_map.json"), encoding="utf-8-sig") as f:
        dept_map = json.load(f)["map"]

    days = []
    d = dt.date(year, month, 1)
    while d.month == month:
        days.append({
            "date": d.isoformat(), "day": d.day, "weekday": WD_JP[d.weekday()],
            "holiday": C.is_holiday(d, holidays),
            "holidayName": holidays.get(d, ""),
            # 休診日だが宿日直手当が出ない日（開院記念日）。昼は当番当直も含めて全員オンコール（管理者 20260922）
            **({"daytimeOncallOnly": True} if d in oncall_only else {}),
        })
        d += dt.timedelta(days=1)

    # dept = 初期研だけ対象月の1日の所属（所属履歴）。それ以外は今日の所属（staff CSV）。20260927 管理者
    # 20261028 管理者: 月のうち1日でも所属していれば、その科の貴科の医師リストに載せる → dept はその月のうちに就く
    #   所属をすべて並べる（/内科(腎臓・内分泌)/産科婦人科/ など。11/20 からローテーションで移る初期研は両方の科に出る）
    # kikan_path（院外研修期間表）は初期値に使わない（20260927。計画で実際と違う人が多い）。引数は互換のために残す
    hist = SZ.load_history(history_dir, staffs=people, log=log)
    first = dt.date(year, month, 1)
    last = (first + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    docs = [p for p in people if p["login"]]

    def month_dept(p):
        segs = []
        for d in hist.depts_during(p["login"], first, last, p.get("title")) or [p["dept"]]:
            segs += [x for x in SZ.segments(d) if x not in segs]
        return "/" + "/".join(segs) + "/" if segs else p["dept"]
    # 20261015: config は全職員が読むので職員コードを載せない（管理側は contract.code_of で 管理ツール\名簿\ から引く）
    roster = [{"login": p["login"], "code": "", "name": p["name"],
               "kana": p["kana"], "dept": month_dept(p), "title": p["title"],
               "doctor": p["is_doctor"]} for p in docs]
    n_moved = sum(1 for p, r in zip(docs, roster) if r["dept"] != p["dept"])
    log(f"  所属: {first:%Y-%m}月のうちの Dr.JOY の所属で名簿を作った（今日の所属と違う人 {n_moved} 名"
        + (f"・所属履歴 {', '.join(n for n, _t in hist.files)}" if hist.files else "・所属履歴なし＝今日の所属") + "）")
    for x in hist.problems:
        log(f"  ⚠ {x}")
    # 管理当直（管理者 20260923）: 管理当直アプリの名簿を初期メンバーにする。oncall=その人を選んだときの当/オの初期値
    fixed = {}
    try:
        fixed[bucho_roster.OWNER] = bucho_roster.resolve(people, log=log)
    except OSError as e:
        keep = (prev_fixed or {}).get(bucho_roster.OWNER)
        if keep:      # 20260925: 読めないだけで空にしない（前の config の初期メンバーをそのまま使う）
            fixed[bucho_roster.OWNER] = keep
            log(f"  ⚠ 管理当直の名簿が読めない（{e}）→ 前の設定の初期メンバー {len(keep)} 人をそのまま使う")
        else:
            log(f"  ⚠ 管理当直の名簿が読めない（{e}）→ 管理当直の初期メンバーなし（①の検索で足せる）")

    # 内科（管理者 20261031）: 初期メンバー = 4科の専攻医（院外研修専従を除く）＋ R8.9〜11 に内科の当直に入った教員・大学院生
    #   （data/naika_ippan_members.json）。以前は4科の医師を職位を問わず全員入れていた（59人）。ほかの人は診療科が検索欄から足す
    try:
        fixed[NAIKA_IPPAN] = naika_ippan_members(roster, dept_map, year, month, log)
    except Exception as e:                       # 作れなくても月は作る（従来どおり所属パターンの全員が初期メンバー）
        log(f"  ⚠ 内科の初期メンバーを作れない（{type(e).__name__}: {e}）→ 4科の医師全員のまま")

    return {
        "version": "1", "schema": "hospital.duty-config/1",
        "year": year, "month": month, "reiwa": year - 2018,
        "deadline": deadline,
        "days": days,
        # 当番当直（20260922 管理者）: {date, band, dept}。renderer が科×役割=ファーストの欄に赤字で出し、オンコールを弾く
        "toban": [{"date": d, "band": b, "dept": dp} for d, b, dp in C.load_toban()
                  if d.startswith(f"{year:04d}-{month:02d}-")],
        "frames": frames,
        "deptMap": dept_map,
        "roster": roster,
        # 提出単位ごとの決まった初期メンバー（dept_map の所属パターンより優先）: {owner: [{login, oncall}]}
        "fixedMembers": fixed,
        # 連続当直の警告を出す提出単位（枠マスタの consecutive_warn_owners。20260914 管理者決定: 救急科と小児科）
        "consecutiveWarnOwners": C.consecutive_warn_owners(),
        # 所属の申告の決まり（20260927 管理者）: owners = 名簿が Dr.JOY の所属を決める提出単位／exclude_owners = 外す提出単位／
        # non_duty_depts = 当直をしない部署（所属の部分一致）／exclude_ranks = 照合から外す職位（contract.rank_of の印）
        "shozoku": SZ.config_block(frames, dept_map),
    }


def write_config(cfg, path):
    """config/current.json を書く。BOM無しUTF-8（Node/JSの JSON.parse 用）。一時ファイルに書き、読み戻して確かめてから
    差し替える（アプリが読んでいる途中に半端なファイルを見せない）。戻り: path"""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    from master_store import tmp_name                # 20261002: 二重拡張子にしない（EDR がランサムウェア疑いで通報する）
    tmp = tmp_name(path)
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    with open(tmp, encoding="utf-8") as f:
        if json.load(f) != cfg:
            raise OSError(f"config の書き込みの確かめに失敗: {tmp}")
    os.replace(tmp, path)
    return path


def default_members(cfg, owner):
    """①今月の当科の医師の初期候補（dept_mapの部分一致・職位条件つき・医師のみ）。"""
    entry = cfg["deptMap"].get(owner, [])
    return [p for p in cfg["roster"] if p["doctor"] and C.matches_dept_map(entry, p)]


N_DUTY = 3      # 当直1〜3（連名当直）
N_ONCALL = 6    # オンコール1〜6（複数オンコール）
BAND_ORDER = {"昼": 0, "夜": 1}   # 同じ日付内は 昼→夜（バンドは常に昼/夜の2区分）


def template_rows(cfg, frames_own):
    """テンプレの行順 = 日付昇順 → 昼→夜 → 枠の定義順。
    20260831 ユーザー決定: 「全日」行は廃止し、全部を昼/夜の行に分ける
    （全日枠は 平日=夜1行 / 土日祝=昼+夜の2行。contract.expand_bands と同じ規則）。
    yield (day_dict, band_str, frame)。"""
    for d in cfg["days"]:
        hol = d["holiday"]
        rows = []
        for idx, f in enumerate(frames_own):
            if f["shape"] == "昼" and not hol:
                continue  # 昼枠は土日祝だけ
            if f["shape"] == "全日":
                bands = ["昼", "夜"] if hol else ["夜"]
            else:
                bands = [f["shape"]]
            for band in bands:
                rows.append((BAND_ORDER[band], idx, band, f))
        for _, _, band, f in sorted(rows, key=lambda t: (t[0], t[1])):
            yield d, band, f


def write_template_xlsx(cfg, owner, frames_own, path):
    """入力規則付きExcelテンプレ（主経路）。1行 = 1枠 × 1日（複数人は列で並べる）。
    当直1〜3・オンコール1〜6の9列とも隠しシート「名簿」を参照するドロップダウン。
    空欄=その人数は居ない（正常）。
    ⚠入力規則はコピペで破れるので、取込側の検証が本丸（テンプレは入力補助）。"""
    import openpyxl
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "当直入力"

    # 隠しシート: 名簿（自科の初期候補 → 続けて全医師）
    lst = wb.create_sheet("名簿")
    members = default_members(cfg, owner)
    seen = {m["login"] for m in members}
    names = [m["name"].replace(" ", "") for m in members]
    others = sorted((p["name"].replace(" ", "") for p in cfg["roster"]
                     if p["doctor"] and p["login"] not in seen), key=str)
    for i, n in enumerate(names + others, start=1):
        lst.cell(row=i, column=1, value=n)
    lst.sheet_state = "hidden"
    n_names = len(names) + len(others)

    duty_cols = [f"当直{i}" for i in range(1, N_DUTY + 1)]
    on_cols = [f"オンコール{i}" for i in range(1, N_ONCALL + 1)]
    ws.append(["日付", "曜日", "昼夜", "枠"] + duty_cols + on_cols + ["備考"])   # 20260831 C/D入替（昼夜が先）
    ws.freeze_panes = "A2"
    widths = {"A": 12, "B": 5, "C": 6, "D": 22}
    name_first = 5                                  # E列=当直1
    name_last = 4 + N_DUTY + N_ONCALL               # M列=オンコール6
    for i in range(name_first, name_last + 1):
        widths[get_column_letter(i)] = 14
    widths[get_column_letter(name_last + 1)] = 18   # 備考
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    r = 2
    for d, band, f in template_rows(cfg, frames_own):
        c = ws.cell(row=r, column=1, value=dt.datetime.fromisoformat(d["date"]))
        c.number_format = "yyyy/mm/dd"
        ws.cell(row=r, column=2, value=d["weekday"] + ("祝" if d["holidayName"] else ""))
        ws.cell(row=r, column=3, value=band)
        ws.cell(row=r, column=4, value=C.frame_label(f, frames_own))
        r += 1
    last = r - 1
    if last >= 2:
        dv_name = DataValidation(type="list", formula1=f"=名簿!$A$1:$A${n_names}",
                                 allow_blank=True, showErrorMessage=True,
                                 error="名簿にある氏名を選んでください（手入力は取込時に弾かれます）")
        dv_band = DataValidation(type="list", formula1='"昼,夜"', allow_blank=True)
        ws.add_data_validation(dv_name)
        ws.add_data_validation(dv_band)
        dv_name.add(f"{get_column_letter(name_first)}2:{get_column_letter(name_last)}{last}")
        dv_band.add(f"C2:C{last}")
    info = wb.create_sheet("説明")
    info.append([f"当直入力テンプレ  R{cfg['reiwa']}年{cfg['month']}月  {owner}"])
    info.append(["・行は 日付順 → 同じ日付は昼→夜 の順。1行 = 1枠 × 1日 × 昼/夜"])
    info.append(["・当直に立つ人は 当直1〜3、オンコール（泊まらない・呼び出し待機）は オンコール1〜6 の列に、ドロップダウンから選ぶ（下にスクロールすると全医師）"])
    info.append(["・空欄は「その人数は居ない」の意味で正常（連名当直・複数オンコールのときだけ2人目以降を使う）"])
    info.append(["・昼と夜の両方に立つ枠は土日祝が昼・夜の2行になります。同じ人が昼夜通しで立つときは両方の行に同じ名前を書く"])
    info.append(["・立つ人がいない日は行ごと空欄のまま残してよい（その行は取り込まれません）"])
    info.append(["・書けたらこのファイルをそのまま提出フォルダへ（ファイル名は変えない）"])
    wb.save(path)
    return last - 1


def build_mac_html(cfg, template_path, out_path):
    with open(template_path, encoding="utf-8") as f:
        tpl = f.read()
    js = json.dumps(cfg, ensure_ascii=False).replace("</script>", "<\\/script>")
    html = tpl.replace("__CONFIG_JSON__", js)
    with open(out_path, "w", encoding="utf-8-sig") as f:
        f.write(html)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=0)
    ap.add_argument("--month", type=int, default=0)
    ap.add_argument("--deadline", default="")
    ap.add_argument("--staffs", default=None, help="名簿CSV（省略=Downloadsの最新 staffs*.csv・500行以上）")
    ap.add_argument("--kikan", default=None,
                    help="（20260927 から使わない。互換のために受けるだけ）院外研修期間表")
    ap.add_argument("--out", default=None)
    ap.add_argument("--legacy-templates", action="store_true",
                    help="旧来の 33科 Excel テンプレと Mac用フォーム HTML も作る（既定=作らない・20260912）")
    ap.add_argument("--interactive", action="store_true")
    a = ap.parse_args()
    a.staffs = a.staffs or C.newest_staffs()
    if a.interactive or not (a.year and a.month):
        print("=== 当直勤務表 月次設定 ===")
        if not a.year:
            a.year = int(input("対象の年（西暦・例 2026）: "))
        if not a.month:
            a.month = int(input("対象の月（1-12）: "))
        if not a.deadline:
            a.deadline = input("締切の表示（例 9月20日(金) 13:00・空欄可）: ").strip()

    out = a.out or os.path.join(C.ROOT, "out", f"R{a.year - 2018}年{a.month}月")
    cfg_path = os.path.join(out, "config", "current.json")
    cfg = build_config(a.year, a.month, a.deadline, a.staffs, a.kikan, prev_fixed=read_fixed_members(cfg_path))
    reiwa = cfg["reiwa"]
    write_config(cfg, cfg_path)       # current.json は BOM無しUTF-8（Node/JSの JSON.parse 用）

    n_tpl = 0
    if a.legacy_templates:      # 旧来の一括生成（20260912 に既定オフ。要るときだけ明示）
        tpl_dir = os.path.join(out, "テンプレ")
        mac_dir = os.path.join(out, "Mac用フォーム")
        for d in (tpl_dir, mac_dir):
            os.makedirs(d, exist_ok=True)
        owners = C.owners_of(cfg["frames"])
        for o in owners:
            own = [f for f in C.active_frames(cfg["frames"]) if f.get("owner", f["dept"]) == o]
            safe = o.replace("/", "・").replace("\\", "・").replace(" ", "").replace("　", "")
            write_template_xlsx(cfg, o, own, os.path.join(
                tpl_dir, f"当直入力_R{reiwa}年{a.month}月_{safe}.xlsx"))
            n_tpl += 1
        # Mac用HTMLの原本 = Electronのrendererそのもの（window.td不在ならブラウザモードで動く）
        tpl_html = os.path.join(C.ROOT, "app", "renderer", "index.html")
        if os.path.exists(tpl_html):
            build_mac_html(cfg, tpl_html, os.path.join(
                mac_dir, f"当直入力_R{reiwa}年{a.month}月.html"))

    print(f"生成: R{reiwa}年{a.month}月  日数{len(cfg['days'])} 枠{len(cfg['frames'])}"
          f" 名簿{len(cfg['roster'])}名"
          + (f" テンプレ{n_tpl}科（--legacy-templates）" if a.legacy_templates else ""))
    print(f"  {out}")


if __name__ == "__main__":
    main()
