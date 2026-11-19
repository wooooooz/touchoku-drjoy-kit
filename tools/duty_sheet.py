# -*- coding: utf-8 -*-
"""duty_sheet.py  20261031
新形式の当直表xlsx（schema hospital.duty-sheet/1）を書く・読む。
設計の正本 = docs/設計_当直表xlsx一本化_20260912.md

A列の**レーンキー**（行を挿入されても列を動かされても読める機械用の目印）:
  `frame_id|band|kind|no`        … ふつうのレーン。band は 昼 / 夜 / 全日
  `夜frame_id+昼frame_id|昼夜|kind|no`
      … 20260922 管理者決定「全日タイプ／昼夜タイプは月ごとにデータから決める」で足した形。
        同じ 科×役割 の 昼枠＋夜枠 が、その月の土日祝で顔ぶれ（人＋当/オ）が毎日同じときに
        1レーンへまとめたもの。枠の部分が `+` で繋がっていたらこの形。
        取込では 土日祝＝昼と夜の両方へ、平日＝夜だけへ展開する（往復で原子レコードが一致する）。
        平日に見えている顔ぶれは夜枠のものなので、夜の frame_id を先に書く。

幾何（当直表シート）:
  行1 タイトル（右に「○月○日○時○分現在」） / 行2 空行（高さだけ残す。20260913） / 行3 凡例
  行4 列キー（上記）（非表示）
  行5 種別 当直/オンコール（非表示）
  行6 枠名（横結合） 行7 昼/夜/昼夜（横結合） 行8 スロット名
  行9〜 日付行（当月の日数ぶん）
  A列 ISO日付（非表示） B列 日 C列 曜日 D列〜 氏名 … 備考 … _祝(非表示) … 集計(印刷範囲外)

シート: 当直表 / 名簿 / 説明 ＋ _ID / _原本 / _メタ（_ は veryHidden）

⚠ 表示名（姓だけ・同姓は名を足す）は xlsx の中だけの話。
   原子レコードの name / name_disp はフルネームのまま作る（互換3シートB列の全行一致を守る）。
"""
from __future__ import annotations

import datetime as dt
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

SCHEMA = "hospital.duty-sheet/1"

SH_MAIN, SH_ROSTER, SH_HELP = "当直表", "名簿", "説明"
SH_ID, SH_BASE, SH_META = "_ID", "_原本", "_メタ"
ID_LIST_COL = 6          # _ID の F列 = ドロップダウンの一覧（A〜D は取り込みが読む名簿なので触らない）

# --- カレンダー（月曜はじまり・週=行・月〜日=列。アプリの②と同じ形。20260912ユーザー指示）---
ROW_TITLE, ROW_SUB = 1, 2
ROW_W0 = 3                       # 最初の週ブロックの先頭行（曜日見出し行は廃止。日付の右に「(水)」）
# 凡例は表の下（20260912ユーザー指示）。集計はさらにその下（印刷範囲は列で外している）
COL_KEY, COL_LABEL, COL_KIND, COL_MON = 1, 2, 3, 4   # A=機械用キー / B=枠 / C=当オ / D〜J=月〜日
COL_SUN = COL_MON + 6
COL_SUM = COL_SUN + 3            # 集計の左端（印刷範囲外。あいだは空けておく）
# 週ブロックの行構成: 0=ISO日付(隠) 1=土日祝フラグ(隠) 2=日付(見える) 3..=レーン 末尾=備考
OFF_ISO, OFF_FLAG, OFF_DATE, OFF_LANE = 0, 1, 2, 3
KEY_ISO, KEY_FLAG, KEY_DATE, KEY_NOTE = "dates", "flags", "date", "note"

KIND_DUTY, KIND_ONCALL = "当直", "オンコール"
KIND_SLOT = "枠"          # 20260912ユーザー決定: 当直/オンコールは行ではなくセルで決める
                          # （現物DTX14と同じ「オンコールは (氏名) の括弧書き」）


def wrapped_parens(s):
    """現物DTX14と同じ判定（dtx14_extract.wrapped_parens）。括弧で包まれていればオンコール。"""
    t = str(s or "").strip()
    return len(t) >= 2 and t[0] in "(（" and t[-1] in ")）"


def strip_parens(s):
    t = str(s or "").strip()
    return t[1:-1].strip() if wrapped_parens(t) else t
# 昼 → 夜 → 昼夜（全日）の順。昼と夜を隣同士にするため、両方にまたがる全日は最後
# 「昼夜」= 昼枠＋夜枠を1レーンにまとめたもの（20260922）。全日と同じまとまりに置く
BAND_RANK = {"昼": 0, "夜": 1, "全日": 2, "昼夜": 2}
BAND_BOTH = ("全日", "昼夜")            # 1レーンで昼と夜の両方を受け持つバンド
WD_JP = ["月", "火", "水", "木", "金", "土", "日"]


# ---------------------------------------------------------------- 表示名
def display_names(members):
    """メンバー → [(表示名, 姓, 足した部分)]。
    姓だけ。同姓（canon(nfc())で判定）がいる組だけ、区別がつくまで名を1文字ずつ足す。
    名を使い切っても衝突するときは (職名の先頭2文字) → 連番 の順に落ちる。"""
    sei = [str(m.get("sei") or "") for m in members]
    mei = [str(m.get("mei") or "") for m in members]
    out = list(sei)
    groups = defaultdict(list)
    for i, s in enumerate(sei):
        groups[C.canon(C.nfc(s))].append(i)

    def distinct(cands):
        return len({C.canon(C.nfc(c)) for c in cands}) == len(cands)

    for idxs in groups.values():
        if len(idxs) == 1:
            continue
        done = False
        for k in range(1, max((len(mei[i]) for i in idxs), default=0) + 1):
            cand = [sei[i] + mei[i][:k] for i in idxs]
            if distinct(cand):
                for i, c in zip(idxs, cand):
                    out[i] = c
                done = True
                break
        if done:
            continue
        cand = [sei[i] + mei[i] + "(" + str(members[i].get("title") or "")[:2] + ")" for i in idxs]
        if distinct(cand):
            for i, c in zip(idxs, cand):
                out[i] = c
            continue
        for n, i in enumerate(idxs, start=1):
            out[i] = sei[i] + mei[i] + ("" if n == 1 else f"({n})")
    return [(out[i], sei[i], out[i][len(sei[i]):]) for i in range(len(members))]


# ---------------------------------------------------------------- 列
class Col:
    __slots__ = ("frame", "label", "band", "kind", "no", "idx", "seq", "span", "pair")

    def __init__(self, frame, label, band, kind, no, pair=None):
        self.frame, self.label, self.band, self.kind, self.no = frame, label, band, kind, no
        # pair = 「昼夜」1レーンにまとめた **昼枠**（このとき frame は夜枠）。20260922
        self.pair = pair
        self.idx = self.seq = 0
        self.span = 1

    @property
    def fid_key(self):
        """A列キーの枠の部分。昼夜レーンは `夜frame_id+昼frame_id`（docstring 参照）。"""
        if self.pair:
            return f"{self.frame['frame_id']}+{self.pair['frame_id']}"
        return self.frame["frame_id"]

    @property
    def key(self):
        return f"{self.fid_key}|{self.band}|{self.kind}|{self.no}"

    def takes(self, fid, band):
        """原子レコード（枠・バンド）がこのレーンのものか。"""
        if self.pair:
            return fid in (self.frame["frame_id"], self.pair["frame_id"])
        return self.frame["frame_id"] == fid and self.band in (band, "全日")

    def stands_on(self, d, holidays):
        """その日にこの列が立つか（立たない日はグレーの斜線にする）。
        規則は contract.band_stands（取込の弾きと同じ。20260926）:
        昼は土日祝だけ・開院記念日の昼は管理当直を置かない（院長がいる。管理者 20260923）。"""
        return C.band_stands(self.frame, self.band, d, holidays, _oncall_only_days())


_ONCALL_ONLY = None


def _oncall_only_days():
    global _ONCALL_ONLY
    if _ONCALL_ONLY is None:
        _ONCALL_ONLY = C.daytime_oncall_only()
    return _ONCALL_ONLY


def _iso_date(iso):
    try:
        return dt.date.fromisoformat(str(iso)[:10])
    except (TypeError, ValueError):
        return None


def toban_index(toban=None):
    """当番当直（管理者 20260922）→ {(iso, 昼/夜, 科名の namekey)}。
    toban = [(date, band, dept)]（contract.load_toban）か [{date, band, dept}]（config の toban）。None なら data/toban.csv。"""
    if toban is None:
        toban = C.load_toban()
    out = set()
    for t in toban or ():
        d, b, dp = (t.get("date"), t.get("band"), t.get("dept")) if isinstance(t, dict) else tuple(t)[:3]
        if d and b and dp:
            out.add((str(d)[:10], str(b), C.namekey(dp)))
    return out


def lane_is_toban(c, d, holidays, tb):
    """レーン c の d 日のマスが当番当直か（20260925。アプリの isToban と同じ規則）:
    ファーストの枠で、科名（dept_dtx14、無ければ dept）とバンドが当番当直の表に載っている。
    昼夜・全日のレーンは、その日に立つバンド（土日祝=昼と夜・平日=夜）のどちらかが当番なら当番。"""
    if not tb or d is None:
        return False
    if c.band in ("昼", "夜"):
        bands = (c.band,)
    else:
        bands = ("昼", "夜") if C.is_holiday(d, holidays) else ("夜",)
    iso = d.isoformat()
    for f in [c.frame] + ([c.pair] if c.pair else []):
        if f.get("role") != "first":
            continue
        k = C.namekey(f.get("dept_dtx14") or f.get("dept") or "")
        if any((iso, b, k) in tb for b in bands):
            return True
    return False


def group_key(f):
    """全日タイプ／昼夜タイプを決める単位 = 科×役割（アプリの「昼夜共通」スイッチと同じ）。"""
    return C.namekey(f.get("dept") or "") + "|" + str(f.get("role") or "")


def merge_pairs(frames_own, duties, holidays, prev_types=None):
    """20260922 管理者決定（全日タイプ／昼夜タイプは月ごとにデータから決める）。
    同じ 科×役割 の 昼枠＋夜枠 の組のうち、その月の**土日祝のすべての日**で
    顔ぶれ（人＋当直/オンコールの別）が同じものを「昼夜」1レーンにまとめる。
    返り値 {夜frame_id: 昼frame}。
    まだ何も入っていない組は、prev_types（前月のタイプ {科|役割: bool}）があればそれに従い、
    無ければまとめない（＝従来どおり昼・夜の2レーン）。"""
    groups = defaultdict(list)
    for f in frames_own:
        groups[group_key(f)].append(f)
    by = defaultdict(lambda: defaultdict(set))     # frame_id -> iso -> {(login, oncall)}
    for r in duties:
        by[r.get("frame_id")][r["date"]].add((r.get("staff_login"), bool(r.get("oncall"))))
    out = {}
    for key, fs in groups.items():
        day = [f for f in fs if f["shape"] == "昼"]
        night = [f for f in fs if f["shape"] == "夜"]
        if len(fs) != 2 or len(day) != 1 or len(night) != 1:
            continue
        dmap, nmap = by.get(day[0]["frame_id"], {}), by.get(night[0]["frame_id"], {})
        if not dmap and not nmap:                  # 白紙 → 前月のタイプを引き継ぐ
            if prev_types and prev_types.get(key):
                out[night[0]["frame_id"]] = day[0]
            continue
        same = True
        for iso in set(dmap) | set(nmap):
            d = _iso_date(iso)
            if d is None or not C.is_holiday(d, holidays):
                continue                           # 平日は夜しか立たないので判定に使わない
            if dmap.get(iso, set()) != nmap.get(iso, set()):
                same = False
                break
        if same:
            out[night[0]["frame_id"]] = day[0]
    return out


def _bands_of_frame(f, duties, holidays):
    """枠 → 列にするバンド。全日かつ share_day_night は「全日」1本。
    ただし昼と夜で人が食い違っている月は、黙って潰さず 昼/夜 の2本に割る。"""
    shape = f["shape"]
    if shape != "全日":
        return [shape]
    if not f.get("share_day_night"):
        return ["昼", "夜"]
    by = defaultdict(lambda: defaultdict(set))
    for r in duties:
        if r.get("frame_id") == f["frame_id"]:
            by[r["date"]][r["band"]].add((r["staff_login"], bool(r.get("oncall"))))
    for d, bb in by.items():
        if "昼" in bb and bb["昼"] != bb.get("夜", set()):
            return ["昼", "夜"]
    return ["全日"]


def build_columns(frames_own, duties, holidays, oncall_slots=None, max_lanes=6, prev_types=None):
    """レーンの並び。予備のオンコール行は、レーンが max_lanes 本までの科にだけ足す
    （レーンが多い科は1枚に収まらなくなる方が痛いので、実データぶんに詰める）。"""
    for headroom in (1, 0):
        cols = _build(frames_own, duties, holidays, oncall_slots, headroom, prev_types)
        if len(cols) <= max_lanes or headroom == 0:
            return cols
    return cols


def _slots_of(f):
    return ((f.get("duty_slots") or 0) + (f.get("oncall_slots") or 0)) or f.get("capacity", 1)


def _build(frames_own, duties, holidays, oncall_slots=None, headroom=1, prev_types=None):
    """枠マスタ＋実データ → 列の並び。枠名でまとめ、枠内は 昼→全日→夜、
    各バンドで 当直1..n → オンコール1..m。人数は実データが既定を超えたら広げる。"""
    used = defaultdict(int)   # (frame_id, band, KIND_SLOT) -> その日の人数の最大
    cnt = defaultdict(set)
    for r in duties:
        cnt[(r.get("frame_id"), r["band"], r["date"])].add(r["staff_login"])
    for (fid, band, _d), s in cnt.items():
        k = (fid, band, KIND_SLOT)
        used[k] = max(used[k], len(s))

    filled = bool(duties)
    cols = []
    # 20260912ユーザー指示: ファースト/セカンドではなく **昼・夜でまとめる**。
    # 昼夜共通（全日）の枠はどちらにも属さないので最後のまとまりにする
    # 20260922: 全日タイプの 昼枠＋夜枠 は「昼夜」1レーンにまとめる（merge_pairs）
    pairs = merge_pairs(frames_own, duties, holidays, prev_types)
    folded = {f["frame_id"] for f in pairs.values()}      # 夜枠のレーンに畳んだ昼枠
    items = []
    for f in sorted(frames_own, key=lambda x: x.get("order", 0)):
        if f["frame_id"] in folded:
            continue
        if f["frame_id"] in pairs:
            items.append((BAND_RANK["昼夜"], f.get("order", 0), "昼夜", f, pairs[f["frame_id"]]))
            continue
        for band in _bands_of_frame(f, duties, holidays):
            items.append((BAND_RANK.get(band, 9), f.get("order", 0), band, f, None))
    if True:
        for _, _, band, f, pair in sorted(items, key=lambda t: (t[0], t[1])):
            label = "昼夜" if band in BAND_BOTH else band
            probe = ["昼", "夜"] if band in BAND_BOTH else [band]
            fs = [f] + ([pair] if pair else [])
            # レーンの本数 = 枠マスタの duty_slots + oncall_slots（現物DTX14の実績由来。
            # アプリ②の既定行数と同じ）。当直かオンコールかは行ではなくセルの括弧で決まる
            dflt = oncall_slots
            if dflt is None:
                dflt = max(_slots_of(x) for x in fs)
            used_n = max([used.get((x["frame_id"], b, KIND_SLOT), 0)
                          for x in fs for b in probe] or [0])
            # 中身がある月は「実際の最大＋予備1本」に詰める（白紙の月は枠マスタのぶん空けておく）。
            # 20260925 管理者: 実際の人数が枠マスタの本数を超えた日があれば、その人数まで欄を増やす
            # （以前は枠マスタの本数で打ち切り、はみ出た人を週の下の空き行に書いていた＝取り込み直すと落ちた）
            n_slot = dflt if not filled else max(used_n, min(max(1, dflt), max(1, used_n + headroom)))
            made = [Col(f, label, band, KIND_SLOT, i, pair) for i in range(1, n_slot + 1)]
            for k, c in enumerate(made, start=1):
                c.seq, c.span = k, len(made)
            cols.extend(made)
    for i, c in enumerate(cols):
        c.idx = i                 # レーンの並び順（カレンダーでは行になる）
    return cols


def _group_by_label(frames_own):
    """枠名（contract.frame_label）でまとめる。定義順は保つ。"""
    order, buckets = [], defaultdict(list)
    for f in sorted(frames_own, key=lambda x: x.get("order", 0)):
        lab = C.frame_label(f, frames_own)
        if lab not in buckets:
            order.append(lab)
        buckets[lab].append(f)
    return [(lab, buckets[lab]) for lab in order]


def member_from_duty(r):
    """割当（原子レコード）の氏名から、メンバー1人分を作る（名簿に居ない人の書き出し用。20261009）。"""
    lg = r.get("staff_login") or ""
    name = C.namekey_display(r.get("name") or r.get("name_disp") or "")
    if not name and lg.startswith("unreg:"):
        name = C.namekey_display(lg[len("unreg:"):])
    name = name or lg
    sei, _, mei = name.partition(" ")
    return {"login": lg, "code": r.get("staff_code") or "", "sei": sei, "mei": mei, "name": name,
            "kana": "", "title": r.get("title") or "", "in_house": True, "work": "full",
            "unregistered": bool(r.get("unregistered")) or lg.startswith("unreg:")}


def members_with_duties(members, *duty_lists):
    """割当に出てくるのにメンバーに居ない人を、割当の氏名から足す（20261009）。
    メンバーに居ない人の割当は、以前は表に出ずに黙って消えていた（別の月の書き出しで今月の名簿を使っていたとき）。
    その Excel を取り込むと、その人の当直がその月の予定から消える。足しておけば表に出て、取り込みでも残る。"""
    out = list(members or [])
    have = {m.get("login") for m in out}
    for rows in duty_lists:
        for r in rows or []:
            lg = r.get("staff_login")
            if lg and lg not in have:
                have.add(lg)
                out.append(member_from_duty(r))
    return out


def place(cols, duties, members, holidays):
    """原子レコード → {(iso, レーン番号): (メンバー番号, オンコールか)} と、
    レーンに収まらなかった分 {iso: [文言]}。
    当直を先、オンコールを後にして詰める（現物DTX14の並びと同じ）。
    メンバーに居ない人の割当も黙って落とさず、収まらなかった分と同じ行（各週の下の空き行）に氏名を出す（20261009。
    write_book は先に members_with_duties で足すので、ふつうはここに来ない）。"""
    idx_of = {m["login"]: i for i, m in enumerate(members)}
    slots = defaultdict(list)      # (iso, frame_id, band) -> [(当直が先, メンバー番号, オンコールか)]
    seen = set()
    over = defaultdict(list)
    for r in sorted(duties, key=lambda x: (x["date"], x.get("frame_id") or "", x["band"])):
        mi = idx_of.get(r.get("staff_login"))
        if mi is None:
            nm = member_from_duty(r)["name"]
            t = f"({nm})" if r.get("oncall") else nm
            if t not in over[r["date"]]:
                over[r["date"]].append(t)
            continue
        band, oc = r["band"], bool(r.get("oncall"))
        # 「全日」「昼夜」1本にまとめたレーンは、昼と夜の同じ人を1回だけ置く
        for c in cols:
            if not c.takes(r.get("frame_id"), band):
                continue
            k = (r["date"], c.frame["frame_id"], c.band)
            if (k, mi, oc) in seen:
                break
            seen.add((k, mi, oc))
            slots[k].append((1 if oc else 0, mi, oc))
            break
    cell = {}
    by_key = defaultdict(list)
    for c in cols:
        by_key[(c.frame["frame_id"], c.band)].append(c)
    for (iso, fid, band), items in slots.items():
        lane = by_key.get((fid, band), [])
        for n, (_o, mi, oc) in enumerate(sorted(items)):
            if n < len(lane):
                cell[(iso, lane[n].idx)] = (mi, oc)
            else:
                nm = members[mi].get("name") or ""
                over[iso].append(f"({nm})" if oc else nm)
    return cell, over


# ---------------------------------------------------------------- 意匠
def _style():
    """意匠は sheet_style.py が正本。無い/壊れているときだけ素の見た目に落ちる（警告を出す）。"""
    try:
        import sheet_style
        return sheet_style
    except Exception as e:                                   # pragma: no cover
        print(f"[duty_sheet] sheet_style を読めません（素の見た目で作ります）: {e}", file=sys.stderr)
        return _PlainStyle()


class _PlainStyle:
    """sheet_style が無いときの最低限。意匠ではなく、動かすためだけのもの。"""
    FONT_BODY = FONT_TITLE = "游ゴシック"
    BASE_SIZE, SUFFIX_DELTA = 11.0, -2.0
    COL_W = {"day": 4, "weekday": 5, "name": 9, "note": 16, "gap": 2,
             "summary_name": 10, "summary_num": 5}
    ROW_H = {"title": 24, "subtitle": 15, "legend": 14, "header_frame": 18,
             "header_band": 15, "header_slot": 15, "day": 18}
    LEGEND = "░=この日は使いません ／ 空欄=該当者なし ／ 赤字=当初から交代"
    PAGE = {"orientation_default": "portrait", "orientation_wide": "landscape",
            "wide_threshold": 8, "fit_to_width": 1, "fit_to_height": 1,
            "margins": {"left": .4, "right": .4, "top": .5, "bottom": .4,
                        "header": .2, "footer": .2},
            "header_footer": None}

    def font(self, role):
        from openpyxl.styles import Font
        if role == "title":
            return Font(name=self.FONT_TITLE, size=16, bold=True)
        if role in ("header_frame",):
            return Font(name=self.FONT_BODY, size=11, bold=True)
        if role == "blocked":
            return Font(name=self.FONT_BODY, size=9, color="BBBBBB")
        if role.startswith("name_toban"):
            return Font(name=self.FONT_BODY, size=10, color="D9D9D9" if role.endswith("oncall") else "FFFFFF")
        return Font(name=self.FONT_BODY, size=10)

    def fill(self, role):
        from openpyxl.styles import PatternFill
        m = {"header_frame": "DDDDDD", "header_band": "EEEEEE", "header_slot": "F5F5F5",
             "row_sat": "EAF1F8", "row_sun": "FBECEC", "blocked": "F0F0F0",
             "name_toban": "595959", "name_toban_oncall": "595959"}
        c = m.get(role)
        return PatternFill("solid", fgColor=c) if c else None

    def border(self, role):
        from openpyxl.styles import Border, Side
        if role == "none":
            return Border()
        t = Side(style="thin", color="AAAAAA")
        return Border(left=t, right=t, top=t, bottom=t)

    def align(self, role):
        from openpyxl.styles import Alignment
        if role == "title":
            return Alignment(horizontal="left", vertical="center")
        return Alignment(horizontal="center", vertical="center")

    @property
    def DXF(self):
        from openpyxl.styles import Font as F, PatternFill as P
        from openpyxl.styles.differential import DifferentialStyle
        return {"changed": DifferentialStyle(font=F(color="CC0000", b=True)),
                "crowded": DifferentialStyle(fill=P("solid", bgColor="FFE0B2")),
                "missing": DifferentialStyle(fill=P("solid", bgColor="FFE5E5")),
                "toban_oncall": DifferentialStyle(font=F(color="D9D9D9")),
                "toban_changed": DifferentialStyle(font=F(color="FFFFFF", b=True), fill=P("solid", bgColor="C00000")),
                "toban_missing": DifferentialStyle(fill=P("solid", bgColor="595959"))}


# 職位の印（20260912ユーザー指示）。判定はアプリの rankOf() と同じ順（職名の部分一致）。
# 色は意匠側で上書きできる（sheet_style に RANK_COLOR があればそちらを使う）。
RANK_COLOR = {"教": "000000", "研": "7030A0", "専": "0070C0", "初": "00A050"}
RANK_DELTA = -4.0          # 氏名より何pt小さくするか（上付きなので実際はさらに小さく出る）
RANK_GAP = " "             # 印と氏名のあいだ


rank_of = C.rank_of        # 20260913: dept_map の職位条件と共用するため contract へ移した


RANK_ORDER = {"教": 0, "研": 1, "専": 2, "初": 3, "": 4}


def sort_members(members):
    """名簿の並び = 職位ランク（教→研→専→初→無印）。
    教員はログインID昇順（ID番号=採用順の序列）、それ以外は かな→氏名 の順。
    アプリの①の並び（名簿の規則）と同じ考え方。"""
    def key(m):
        r = rank_of(m.get("title"))
        return (m.get("order") if m.get("order") is not None else 10 ** 9,   # 管理当直: 名簿の並び（管理者 20260923）
                RANK_ORDER.get(r, 4),
                (m.get("login") or "") if r == "教" else "",
                str(m.get("kana") or ""), str(m.get("name") or ""))
    return sorted(members, key=key)


def _inline(S, role, delta=0.0):
    from openpyxl.cell.text import InlineFont
    if hasattr(S, "inline"):
        try:
            return S.inline(role)
        except Exception:
            pass
    return InlineFont(rFont=S.FONT_BODY, sz=S.BASE_SIZE + delta)


def _name_cell(disp, sei, suffix, S, role="name", rank="", oncall=False):
    """表示名のセル値。
    ・職位の印（教/研/専/初）を氏名の左上に小さく上付きで置く
    ・姓のうしろに足した部分（同姓の区別）は2pt小さくする"""
    if not suffix and not rank:
        return f"({disp})" if oncall else disp
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    runs = []
    if rank:
        col = getattr(S, "RANK_COLOR", RANK_COLOR).get(rank, "000000")
        if role.startswith("name_toban"):       # 当番当直の灰地では職位の色が沈むので氏名と同じ色にする（20260925）
            rgb = getattr(getattr(S.font(role), "color", None), "rgb", None)
            col = rgb[-6:] if isinstance(rgb, str) and len(rgb) >= 6 else "FFFFFF"
        sz = S.BASE_SIZE + getattr(S, "RANK_DELTA", RANK_DELTA)
        runs.append(TextBlock(InlineFont(rFont=S.FONT_BODY, sz=sz,
                                         color=col, vertAlign="superscript"),
                              rank + getattr(S, "RANK_GAP", RANK_GAP)))
    bold = not oncall                      # 当直は太字。オンコールは灰字のまま
    if oncall:
        runs.insert(0, TextBlock(_font_inline(S, role, 0.0, bold), "("))
    runs.append(TextBlock(_font_inline(S, role, 0.0, bold), sei))
    if suffix:
        runs.append(TextBlock(_font_inline(S, role, S.SUFFIX_DELTA, bold), suffix))
    if oncall:
        runs.append(TextBlock(_font_inline(S, role, 0.0, bold), ")"))
    return CellRichText(*runs)


def _bold(f):
    """同じ書式のまま太字にした Font（当直のセル用）。"""
    from openpyxl.styles import Font
    return Font(name=f.name, size=f.sz, bold=True, italic=f.i, color=f.color)


def _font_inline(S, role, delta=0.0, bold=None):
    """意匠の font(role) と同じ色で、大きさだけ変えた InlineFont を作る。"""
    from openpyxl.cell.text import InlineFont
    f = S.font(role)
    rgb = getattr(getattr(f, "color", None), "rgb", None)
    col = rgb[-6:] if isinstance(rgb, str) and len(rgb) >= 6 else None
    return InlineFont(rFont=f.name or S.FONT_BODY, sz=(f.sz or S.BASE_SIZE) + delta,
                      color=col, b=f.b if bold is None else bold)


def _date_cell(S, day, wd, hname, role):
    """日付セル。「1 (木)」の形。曜日は少し小さく、祝日名はさらに小さく次の行に。"""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    runs = [TextBlock(_font_inline(S, role), str(day)),
            TextBlock(_font_inline(S, role, -4.0, bold=False), f" ({wd})")]
    if hname:
        runs.append(TextBlock(_inline(S, "holiday_name", -3.0), chr(10) + hname))
    return CellRichText(*runs)


def people_index(members):
    """_ID のメンバー → contract.import_rows が使う (people, exact, canon_map)。
    突合先を全医師名簿ではなくこの科のメンバーに閉じるための入口。

    ⚠ **表示名でも引けるようにする。** 完全同姓同名の2人はフルネームでは一意に決まらず
    （resolve_name が ambiguous を返す）、_ID の表示名（壬野太郎(専攻) など）だけが
    手がかりになる。_rows_from は突合できたセルを表示名で返すので、ここで両方を登録する。"""
    people, exact, canon_map = [], defaultdict(list), defaultdict(list)
    for m in members:
        name = (str(m.get("sei") or "") + " " + str(m.get("mei") or "")).strip() or str(m.get("name") or "")
        p = {"login": m.get("login") or "", "code": m.get("code") or "", "name": name,
             "kana": m.get("kana") or "", "dept": m.get("dept") or "",
             "title": m.get("title") or "", "job": "医師", "is_doctor": True}
        i = len(people)
        people.append(p)
        for k in dict.fromkeys([C.namekey(name), C.namekey(str(m.get("disp") or ""))]):
            if not k:
                continue
            exact[k].append(i)
            canon_map[C.canon(C.nfc(k))].append(i)
    return people, exact, canon_map


# ---------------------------------------------------------------- カレンダーの幾何
def month_days(year, month):
    d, out = dt.date(year, month, 1), []
    while d.month == month:
        out.append(d)
        d += dt.timedelta(days=1)
    return out


def weeks_of(year, month):
    """月曜はじまりの週へ分割。前後月のマスは None（灰色の空マス）。
    アプリの renderGrid と同じ規則（2026-08-30フィードバック）。"""
    weeks, wk = [], [None] * 7
    for d in month_days(year, month):
        wk[d.weekday()] = d
        if d.weekday() == 6:
            weeks.append(wk)
            wk = [None] * 7
    if any(wk):
        weeks.append(wk)
    return weeks


def block_height(lanes):
    return OFF_LANE + len(lanes) + 1          # 末尾に備考行


def unit_label(c, owner):
    """B列の見出し = バンドのまとまり（昼 / 夜 / 昼夜）。20260912ユーザー指示。
    20260913: 別の科の枠を含む提出単位（内科のセカンド=総合診療科・血液…、
    内科消化器・呼吸器のセカンド=消化器・感染症呼吸器）では、枠マスタの short（無ければ科名）を前に付ける。"""
    dept = c.frame.get("dept", "")
    if owner and dept != owner and f"{dept}当直" != owner:      # 管理当直（枠の科も「管理当直」）は「管理当直　昼」にしない（20261031）
        return f"{c.frame.get('short') or dept}　{c.label}"
    return c.label


def kind_cell(S, c):
    """C列のセル値。序数の数字は大きく、st / nd は小さく（20260912ユーザー指示）。"""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    t = kind_label(c)
    if len(t) < 2 or not t[0].isdigit():
        return t
    return CellRichText(TextBlock(_font_inline(S, "header_slot"), t[0]),
                        TextBlock(_font_inline(S, "header_slot", -3.0), t[1:]))


def kind_label(c):
    """C列 = 枠。**1st / 2nd だけ**（20260912ユーザー指示）。
    同じ枠が複数行あっても通し番号は付けない。当かオかは氏名の色で見せる（灰字=オンコール）。
    ⚠ C列は幅4で「縮小して全体を表示する」が効くので、4文字を超えるラベルを入れると
       その行だけ字が小さくなる（`2nd-1` で実際に起きた）。ここは3文字までに保つこと。"""
    role = c.frame.get("role")
    return "1st" if role == "first" else "2nd" if role == "second" else "当直"


def _lighten(fill, factor):
    """塗りを白へ寄せて薄くする（factor=3 なら 1/3 の濃さ）。土日祝の名前セル用。"""
    from openpyxl.styles import PatternFill
    rgb = getattr(getattr(fill, "fgColor", None), "rgb", None)
    if not isinstance(rgb, str) or len(rgb) < 6:
        return None
    h = rgb[-6:]
    try:
        v = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    except ValueError:
        return None
    out = "".join(f"{255 - int((255 - x) / factor):02X}" for x in v)
    return PatternFill("solid", fgColor=out, bgColor=out)


def _no_shrink(a):
    """「縮小して全体を表示する」を外した同じ配置。
    ⚠ これが効いていると、同じ字の大きさを指定していても**幅が入り切らない文字列だけ**
       Excel が勝手に縮めて刷る。プロポーショナル書体では「2nd」は「1st」より広いので、
       同じ3文字でも 2nd だけ小さくなる（20260912 に実際に起きた）。見出しには使わない。"""
    from openpyxl.styles import Alignment
    return Alignment(horizontal=a.horizontal, vertical=a.vertical, wrap_text=a.wrapText,
                     shrink_to_fit=False, indent=a.indent or 0, text_rotation=a.textRotation or 0)


# 横罫線の濃さは4段階。週の区切り(595959) > 昼夜の境目(8C8C8C) > レーンの境目(BFBFBF) > マス(D9D9D9)
BAND_SIDE = ("thin", "808080")     # 昼/夜のまとまりの境目
LANE_SIDE = ("thin", "A6A6A6")     # 枠が変わる境目（1st と 2nd の間）
#   同じ枠の中（2nd と 2nd の間）はマスの罫線のまま。ここにも線を引くと、
#   どの段も同じ濃さに見えて構造が読めなくなる（20260912ユーザー指摘）


def _side(S, name, dflt):
    from openpyxl.styles import Side
    st, col = getattr(S, name, dflt)
    return Side(style=st, color=col)


def _band_side(S):
    return _side(S, "BAND_SIDE", BAND_SIDE)


def _lane_side(S):
    return _side(S, "LANE_SIDE", LANE_SIDE)


def _edge(ws, row, col, which, side):
    """既にある罫線を保ったまま、上辺か下辺だけ差し替える。"""
    from openpyxl.styles import Border
    b = ws.cell(row, col).border
    kw = {"left": b.left, "right": b.right, "top": b.top, "bottom": b.bottom}
    kw[which] = side
    ws.cell(row, col).border = Border(**kw)


def _hrule(S):
    """見出し列用の、薄い横罫線だけの枠（マスの罫線と同じ細さ）。
    昼/夜 と 1st/2nd の下に入れて、見出しから右のマスへ目で追えるようにする。"""
    from openpyxl.styles import Border
    return Border(bottom=S.border("name").bottom)


def _box(ws, S, row, col, text, role, r1=None, c1=None, align_role="header", border=None):
    cell = ws.cell(row, col, text)
    cell.font = S.font(role)
    cell.border = border if border is not None else S.border("header")
    cell.alignment = _no_shrink(S.align(align_role))
    f = S.fill(role)
    if f:
        cell.fill = f
    if (r1 or row) > row or (c1 or col) > col:
        ws.merge_cells(start_row=row, start_column=col, end_row=r1 or row, end_column=c1 or col)
    return cell


# ---------------------------------------------------------------- 書き出し
def write_book(path, *, year, month, owner, frames_own, holidays, members,
               duties=(), baseline=None, deadline="", source="blank",
               master_rev=None, tool="duty_sheet/2", orient=None, asof=None,
               prev_types=None, toban=None):
    """当直表ブックを書く（月曜はじまりのカレンダー・A4横）。
    members = [{login, code, sei, mei, name, title, in_house, work, unregistered}]（work が無ければ in_house で読む＝contract.work_of）
    asof = マスターから読み込んだ時刻（見出しの右に「○月○日○時○分現在」と小さく出す。None なら出さない）
    prev_types = 前月のタイプ {科|役割: bool}（20260922）。白紙の月のレーンの形だけに効く
    toban = 当番当直 [(date, band, dept)] か [{date, band, dept}]（None なら data/toban.csv）。
            そのマスはグレー地に白字（20260925 管理者: DTX14・アプリの画面と同じ見た目）"""
    import openpyxl
    from openpyxl.utils import get_column_letter as L
    from openpyxl.cell.rich_text import CellRichText, TextBlock

    S = _style()
    duties = list(duties)
    # 割当・元の割当に出てくる人は必ずメンバーに入れる（黙って表から消さない。20261009）
    members = members_with_duties(members, duties, baseline)
    members = sort_members(members)      # ドロップダウンと名簿の並びを職位順に固定
    disp = display_names(members)
    ranks = [rank_of(m.get("title")) for m in members]
    # 職位の印はセルの「文字」の一部になる。ドロップダウンの元（_ID!D）・_原本・集計も
    # 同じ文字にしないと、取込で人が引けず、交代の赤字も全マスで誤爆する
    gap = getattr(S, "RANK_GAP", RANK_GAP)
    full = [(r + gap if r else "") + d[0] for r, d in zip(ranks, disp)]
    lanes = build_columns(frames_own, duties, holidays, prev_types=prev_types)
    weeks = weeks_of(year, month)
    cell, over = place(lanes, duties, members, holidays)
    base_cell = place(lanes, list(baseline or []), members, holidays)[0] if baseline is not None else {}
    H = block_height(lanes)
    last_row = ROW_W0 + len(weeks) * H - 1
    tb = toban_index(toban)
    toban_cells = []                     # [(行, 列, レーン番号, 週番号)] 当番当直のマス（条件付き書式でも使う）

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SH_MAIN
    ws.sheet_view.showGridLines = False
    wsb = wb.create_sheet(SH_BASE)

    # 行1〜3 見出し（20260913 ユーザー指示: 2行目の文言は出さない。行は空のまま高さを残す。
    # 題名の右に、マスターから読み込んだ時刻を小さく「○月○日○時○分現在」）
    title = f"R{year - 2018}年{month}月　{owner}　当直予定表"
    if asof is not None:
        stamp = f"{asof.month}月{asof.day}日 {asof.hour}時{asof.minute:02d}分現在"
        ws.cell(ROW_TITLE, COL_LABEL).value = CellRichText([
            TextBlock(_font_inline(S, "title"), title),
            TextBlock(_font_inline(S, "subtitle", -1.0, bold=False), "　　" + stamp)])
    else:
        ws.cell(ROW_TITLE, COL_LABEL, title)
    ws.cell(ROW_SUB, COL_LABEL).value = None
    for r, role in ((ROW_TITLE, "title"), (ROW_SUB, "subtitle")):
        c = ws.cell(r, COL_LABEL)
        c.font, c.alignment = S.font(role), S.align(role)
        ws.row_dimensions[r].height = S.ROW_H[role]
        ws.merge_cells(start_row=r, start_column=COL_LABEL, end_row=r, end_column=COL_SUN)

    # 週ブロック
    for wi, wk in enumerate(weeks):
        r0 = ROW_W0 + wi * H
        ws.cell(r0 + OFF_ISO, COL_KEY, KEY_ISO)
        ws.cell(r0 + OFF_FLAG, COL_KEY, KEY_FLAG)
        ws.cell(r0 + OFF_DATE, COL_KEY, KEY_DATE)
        ws.row_dimensions[r0 + OFF_ISO].hidden = True
        ws.row_dimensions[r0 + OFF_FLAG].hidden = True
        ws.row_dimensions[r0 + OFF_DATE].height = S.ROW_H.get("cal_date", S.ROW_H["header_slot"])
        for li, c_ in enumerate(lanes):
            ws.cell(r0 + OFF_LANE + li, COL_KEY, c_.key)
            ws.row_dimensions[r0 + OFF_LANE + li].height = S.ROW_H.get("cal_lane", S.ROW_H["day"])
        note_row = r0 + OFF_LANE + len(lanes)
        ws.cell(note_row, COL_KEY, KEY_NOTE)
        ws.row_dimensions[note_row].height = S.ROW_H.get("cal_note", S.ROW_H["legend"])

        # B列=バンド（同じバンドのレーンは縦に結合）／C列=1st/2nd。
        # どちらも薄い横罫線を入れて、見出しから右のマスへ目で追えるようにする
        rule = _hrule(S)
        for a, b, _k in _runs(lanes, lambda c: (unit_label(c, owner), c.band)):
            _box(ws, S, r0 + OFF_LANE + a, COL_LABEL, unit_label(lanes[a], owner),
                 "header_frame", r1=r0 + OFF_LANE + b, align_role="header", border=rule)
            for li in range(a, b + 1):        # 結合の下辺を出すため範囲の各セルにも入れる
                ws.cell(r0 + OFF_LANE + li, COL_LABEL).border = rule
        # 1st/2nd は対等の見出しなので同じ字。当かオかは氏名の色で見せる。
        # 同じバンドの中で 2nd が縦に並ぶときは結合して1つだけ書く（20260913ユーザー指示）
        for a, b, _k in _runs(lanes, lambda c: (c.label, c.band, kind_label(c))):
            _box(ws, S, r0 + OFF_LANE + a, COL_KIND, kind_cell(S, lanes[a]), "header_slot",
                 r1=r0 + OFF_LANE + b, border=rule)
            for li in range(a, b + 1):        # 結合の下辺を出すため範囲の各セルにも入れる
                ws.cell(r0 + OFF_LANE + li, COL_KIND).border = rule
        # 備考行は見出しを出さない。週と週のあいだの「空き」として働かせる（20260912ユーザー指示）

        for di, d in enumerate(wk):
            col = COL_MON + di
            if d is None:                                   # 前後月の空マス。白のまま（20260912ユーザー指示）
                for r in range(r0 + OFF_DATE, note_row):
                    ws.cell(r, col).border = S.border("day" if r == r0 + OFF_DATE else "name")
                continue
            iso, hname = d.isoformat(), holidays.get(d, "")
            hol = C.is_holiday(d, holidays)
            sun, sat = bool(hname) or d.weekday() == 6, d.weekday() == 5 and not hname
            ws.cell(r0 + OFF_ISO, col, iso)
            ws.cell(r0 + OFF_FLAG, col, 1 if hol else 0)
            drole = "weekday_sun" if sun else "weekday_sat" if sat else "day"
            dc = ws.cell(r0 + OFF_DATE, col,
                         _date_cell(S, d.day, WD_JP[d.weekday()], hname, drole))
            dc.font, dc.alignment, dc.border = S.font(drole), S.align("day"), S.border("day")
            rf = S.fill("row_sun" if sun else "row_sat" if sat else "row_weekday")
            if rf:
                dc.fill = rf
            for li, c_ in enumerate(lanes):
                cl = ws.cell(r0 + OFF_LANE + li, col)
                cl.border, cl.alignment = S.border("name"), S.align("name")
                if not c_.stands_on(d, holidays):
                    cl.font = S.font("blocked")
                    bf = S.fill("blocked")
                    if bf:
                        cl.fill = bf
                    continue
                got = cell.get((iso, li))
                oc = bool(got[1]) if got else False
                is_tb = lane_is_toban(c_, d, holidays, tb)
                if is_tb:                        # 当番当直: グレー地に白字（土日の薄い色より優先）
                    role = "name_toban_oncall" if oc else "name_toban"
                    toban_cells.append((r0 + OFF_LANE + li, col, li, wi))
                else:
                    role = "name_oncall" if oc else "name"
                cl.font = _bold(S.font(role)) if not oc else S.font(role)
                nf = (S.fill(role) if is_tb else None) \
                    or (_lighten(rf, 3.0) if (sun or sat) else None) or S.fill(role) or rf
                if nf:
                    cl.fill = nf
                if got:
                    cl.value = _name_cell(*disp[got[0]], S, role, ranks[got[0]], oc)
                bgot = base_cell.get((iso, li))
                if bgot:
                    t = full[bgot[0]]
                    wsb.cell(r0 + OFF_LANE + li, col, f"({t})" if bgot[1] else t)
            # 罫線も塗りも付けない。週ブロックどうしが縦線でつながって見えないようにする
            nc = ws.cell(note_row, col, " / ".join(over.get(iso, [])) or None)
            nc.font, nc.alignment = S.font("note"), S.align("note")
        # 昼／夜のまとまりの境目（昼の上・昼と夜の間・夜の下）を少し濃い罫線にする
        # 横罫線の濃さ。⚠ 隣り合うセルは「上の下辺」と「下の上辺」が同じ線になるので、
        #   片方だけ濃くしても、もう片方の薄い線に上書きされて効かない。必ず両方に入れる
        bs, ls = _band_side(S), _lane_side(S)
        runs = _runs(lanes, lambda c: (c.label, c.band))
        starts = {a for a, _b, _k in runs}
        edges = {}                               # 境目の上の行 -> 線
        for li in range(len(lanes)):
            above = (r0 + OFF_DATE) if li == 0 else (r0 + OFF_LANE + li - 1)
            if li in starts:                     # 昼/夜 が変わる
                edges[above] = bs
            elif lanes[li].frame["frame_id"] != lanes[li - 1].frame["frame_id"]:
                edges[above] = ls                # 枠が変わる（1st → 2nd）
            # 同じ枠の続き（2nd → 2nd）はマスの罫線のまま
        edges[r0 + OFF_LANE + len(lanes) - 1] = bs      # 夜の下
        for above, side in edges.items():
            for col in range(COL_LABEL, COL_SUN + 1):
                _edge(ws, above, col, "bottom", side)
                _edge(ws, above + 1, col, "top", side)
    wsb.sheet_state = "veryHidden"

    # 凡例は表の下。印刷範囲の最後の行になる
    legend_row = last_row + 1
    lg = ws.cell(legend_row, COL_LABEL, S.LEGEND + (getattr(S, "LEGEND_TOBAN", "") if toban_cells else ""))
    lg.font, lg.alignment = S.font("legend"), S.align("legend")
    ws.row_dimensions[legend_row].height = S.ROW_H["legend"]
    ws.merge_cells(start_row=legend_row, start_column=COL_LABEL,
                   end_row=legend_row, end_column=COL_SUN)

    ws.column_dimensions[L(COL_KEY)].hidden = True
    # B列の幅: 既定 15。別の科の枠を含む単位（「膠原病　昼夜」「感染症呼吸器　昼夜」）は入る幅まで広げる
    lbl_w = max([len(unit_label(c, owner)) for c in lanes] or [0])
    # 全角1字 ≈ 2.4 幅（11pt・見出し書体の実測。「総合診療　昼夜」7字が 15 では右端が切れた）
    ws.column_dimensions[L(COL_LABEL)].width = max(S.COL_W.get("lane_label", 15), lbl_w * 2.4 + 2)
    # 「2nd」が 11pt のまま入り切る幅。狭いと縮小が効いて 1st と大きさが揃わない
    ws.column_dimensions[L(COL_KIND)].width = max(S.COL_W.get("lane_kind", 5.5), 5.5)
    # 既定は A4縦。カレンダーは縦に長い形なので、横だと高さが足りず77%に縮む
    # （実測 2026-09-12: 中身 464x685pt / 縦の刷れる面 530x774pt・横は 777x527pt）
    land = (orient or S.PAGE.get("cal_orientation") or "portrait") == "landscape"
    # 刷り幅を使い切る（横=7列+見出し2列≒112文字幅／縦=同≒77文字幅）
    day_w = S.COL_W.get("cal_day" if land else "cal_day_portrait", 13.5 if land else 9.0)
    for i in range(7):
        ws.column_dimensions[L(COL_MON + i)].width = day_w
    ws.print_area = f"{L(COL_LABEL)}1:{L(COL_SUN)}{legend_row}"
    if hasattr(S, "apply_page"):
        S.apply_page(ws, land)
    ws.page_setup.orientation = (S.PAGE["orientation_wide"] if land
                                 else S.PAGE["orientation_default"])
    # カレンダーは1枚なので見出しの繰り返しは要らない。
    # ⚠ openpyxl の print_title_rows は None を無視する（setterがNoneを弾く）ので内部を直に消す
    ws._print_rows = None
    ws._print_cols = None

    _summary(ws, S, lanes, weeks, members, full, legend_row + 2, H)
    _validation(ws, members, lanes, weeks, H)
    _cond_formats(ws, S, lanes, weeks, baseline, H, toban_cells)
    _side_sheets(wb, S, owner, year, month, deadline, members, disp, source, master_rev, tool,
                 ranks, full)
    wb.save(path)
    return {"lanes": len(lanes), "weeks": len(weeks), "placed": len(cell),
            "overflow": sum(len(v) for v in over.values()), "toban": len(toban_cells)}


def _lane_ranges(lanes, weeks, H, kinds=None):
    """(レーン, [そのレーンが各週で占める D〜J の範囲]) を返す。"""
    from openpyxl.utils import get_column_letter as L
    out = []
    for li, c_ in enumerate(lanes):
        if kinds and c_.kind not in kinds:
            continue
        rs = [ROW_W0 + w * H + OFF_LANE + li for w in range(len(weeks))]
        out.append((li, c_, [f"{L(COL_MON)}{r}:{L(COL_SUN)}{r}" for r in rs], rs))
    return out


def _summary(ws, S, lanes, weeks, members, full, start_row, H):
    """印刷範囲外の人別集計。週ごとの 1×7 の範囲に COUNTIF/COUNTIFS を足すだけにする。"""
    from openpyxl.utils import get_column_letter as L
    allr = _lane_ranges(lanes, weeks, H)
    flag = {w: f"${L(COL_MON)}${ROW_W0 + w * H + OFF_FLAG}:${L(COL_SUN)}${ROW_W0 + w * H + OFF_FLAG}"
            for w in range(len(weeks))}
    _box(ws, S, start_row, COL_SUM, "【今月の回数】", "summary_head", c1=COL_SUM + 3)
    for j, t in enumerate(("氏名", "当直", "オ", "土日祝")):
        _box(ws, S, start_row + 1, COL_SUM + j, t, "summary_head")
    ws.column_dimensions[L(COL_SUM)].width = S.COL_W["summary_name"]
    for j in range(1, 4):
        ws.column_dimensions[L(COL_SUM + j)].width = S.COL_W["summary_num"]

    def abs_(rng):
        a, b = rng.split(":")
        return f"${a[0]}${a[1:]}:${b[0]}${b[1:]}"

    for i, m in enumerate(members):
        r = start_row + 2 + i
        who = f"${L(COL_SUM)}{r}"
        nc = ws.cell(r, COL_SUM, full[i])
        # 当直=素の氏名、オンコール=括弧つき。セルの中身で見分ける（行では決まらない）
        vals = ["+".join(f"COUNTIF({abs_(rg)},{who})" for _li, _c, rgs, _rs in allr for rg in rgs) or "0",
                "+".join(f'COUNTIF({abs_(rg)},"("&{who}&")")' for _li, _c, rgs, _rs in allr for rg in rgs) or "0",
                "+".join(f"COUNTIFS({abs_(rgs[w])},{who},{flag[w]},1)"
                         for _li, _c, rgs, _rs in allr for w in range(len(weeks))) or "0"]
        for c0 in [nc] + [ws.cell(r, COL_SUM + 1 + j, "=" + v) for j, v in enumerate(vals)]:
            c0.font, c0.border, c0.alignment = S.font("summary_body"), S.border("summary"), S.align("summary")
            f = S.fill("summary_body")
            if f:
                c0.fill = f


def _validation(ws, members, lanes, weeks, H):
    """氏名セルのドロップダウン。参照元は veryHidden の _ID!F 列（氏名と (氏名) が人ごとに並ぶ）。"""
    from openpyxl.worksheet.datavalidation import DataValidation
    if not members or not lanes:
        return
    col = chr(64 + ID_LIST_COL)
    dv = DataValidation(
        type="list", formula1=f"='{SH_ID}'!${col}$2:${col}${2 * len(members) + 1}",
        allow_blank=True, showErrorMessage=True,
        error="この科のメンバーから選んでください（オンコールは (氏名) の方を選びます）。"
              "増やすときはアプリの①で足してから書き出し直します。",
        errorTitle="メンバーにありません")
    dv.errorStyle = "warning"     # 一覧に無い書き方も止めずに注意だけにする（取り込みで名簿と突き合わせる）
    ws.add_data_validation(dv)
    for _li, _c, rgs, _rs in _lane_ranges(lanes, weeks, H):
        for rg in rgs:
            dv.add(rg)


def _cond_formats(ws, S, lanes, weeks, baseline, H, toban_cells=()):
    """交代の赤字・連続当直の橙・未入力の薄赤。週ブロックごと・レーンごとに1本ずつ掛ける。
    当番当直のマス（toban_cells）には、そのマスだけの規則を先に掛ける（先に足した規則ほど優先。20260925）。"""
    from openpyxl.utils import get_column_letter as L
    from openpyxl.formatting.rule import Rule
    M = L(COL_MON)

    def add(ref, formula, dxf):
        ws.conditional_formatting.add(ref, Rule(type="expression", dxf=dxf,
                                                formula=[formula], stopIfTrue=False))

    def required(c_):                   # 必ず1人要るマス（ファーストの1人目。アプリの runChecks と同じ条件）
        return (c_.no == 1 and not c_.frame.get("oncall_default")
                and str(c_.frame.get("frame_id", "")).startswith("f-")
                and (c_.frame.get("used_2026_09") or 0) > 0)

    def missing_cond(c_, here, colL, r0):
        return (f"AND({here}=\"\",{colL}${r0 + OFF_FLAG}=1)" if c_.band == "昼"
                else f"AND({here}=\"\",{colL}${r0 + OFF_ISO}<>\"\")")
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles import Font
    _oc = S.font("name_oncall")
    oc_dxf = DifferentialStyle(font=Font(
        b=False, color=getattr(getattr(_oc, "color", None), "rgb", None) or "595959"))
    # 当番当直（グレー地に白字）: 下の「オンコールは灰字」は灰地に沈むので薄い灰で上書きし、
    # 未入力の薄赤も灰地のまま（赤の点線枠だけ）、交代は赤地に白字（20260929）
    dxfs = getattr(S, "DXF", {}) or {}
    for (r, col, li, w) in toban_cells:
        here = f"{L(col)}{r}"
        if baseline is not None and dxfs.get("toban_changed") is not None:
            add(here, f"AND({here}<>\"\",{here}<>'{SH_BASE}'!{here})", dxfs["toban_changed"])
        if dxfs.get("toban_oncall") is not None:
            add(here, f'LEFT({here},1)="("', dxfs["toban_oncall"])
        if dxfs.get("toban_missing") is not None and required(lanes[li]):
            add(here, missing_cond(lanes[li], here, L(col), ROW_W0 + w * H), dxfs["toban_missing"])
    for li, c_, rgs, rs in _lane_ranges(lanes, weeks, H):
        for w, (rg, r) in enumerate(zip(rgs, rs)):
            here = f"{M}{r}"
            add(rg, f'LEFT({here},1)="("', oc_dxf)      # オンコール（括弧書き）は灰字
            if baseline is not None:
                add(rg, f"AND({here}<>\"\",{here}<>'{SH_BASE}'!{here})", S.DXF["changed"])
            if required(c_):
                add(rg, missing_cond(c_, here, M, ROW_W0 + w * H), S.DXF["missing"])


def _runs(items, keyf):
    """隣り合う同じキーをまとめる → [(先頭, 末尾, キー)]（見出しの結合用）。"""
    out = []
    for i, c in enumerate(items):
        k = keyf(c)
        if out and out[-1][2] == k and out[-1][1] == i - 1:
            out[-1][1] = i
        else:
            out.append([i, i, k])
    return [(a, b, k) for a, b, k in out]


def _side_sheets(wb, S, owner, year, month, deadline, members, disp, source, master_rev, tool,
                 ranks=None, full=None):
    """名簿（人が見るだけ）・説明（可視）と、_ID・_メタ（veryHidden）。"""
    ro = wb.create_sheet(SH_ROSTER)
    ro.cell(1, 1, "アプリで確定したメンバーです。ここを直しても反映されません。"
                  "増やすときはアプリの①で足してから書き出し直してください。").font = S.font("legend")
    for j, t in enumerate(("表示名", "氏名", "職名", "勤務区分"), start=1):
        _box(ro, S, 2, j, t, "header_slot")
    for i, m in enumerate(members):
        r = 3 + i
        ro.cell(r, 1, _name_cell(*disp[i], S, "name", (ranks or [""] * len(members))[i]))
        ro.cell(r, 2, m.get("name") or "")
        ro.cell(r, 3, m.get("title") or "")
        ro.cell(r, 4, C.WORK_LABEL[C.work_of(m)])      # 20260927: フルタイム／パートタイム／院外研修専従
    for j, w in enumerate((S.COL_W["summary_name"], 16, 18, 14), start=1):
        ro.column_dimensions[chr(64 + j)].width = w

    idw = wb.create_sheet(SH_ID)
    for j, t in enumerate(("ログインID", "姓", "名", "表示名"), start=1):
        idw.cell(1, j, t)
    for i, m in enumerate(members):
        idw.cell(2 + i, 1, m.get("login") or "")
        idw.cell(2 + i, 2, m.get("sei") or "")
        idw.cell(2 + i, 3, m.get("mei") or "")
        idw.cell(2 + i, 4, (full or [d[0] for d in disp])[i])
    # ドロップダウンの一覧（F列）: 人ごとに 氏名（当直）→ (氏名)（オンコール）の順に並べる（管理者 20260922。
    # 括弧を手で打つと「メンバーにありません」の警告が出ていた）。取り込みは A〜D 列しか読まない
    idw.cell(1, ID_LIST_COL, "一覧")
    for i in range(len(members)):
        name = (full or [d[0] for d in disp])[i]
        idw.cell(2 + 2 * i, ID_LIST_COL, name)
        idw.cell(3 + 2 * i, ID_LIST_COL, f"({name})")
    idw.sheet_state = "veryHidden"

    hp = wb.create_sheet(SH_HELP)
    for line in (f"当直予定表  R{year - 2018}年{month}月  {owner}",
                 f"提出期限: {deadline}" if deadline else "", "",
                 "■ 書き方",
                 "・月曜はじまりのカレンダーです。アプリの②と同じ並びです",
                 "・左の列がその日の枠（「夜 ファースト」など）と 当＝当直／オ＝オンコール です",
                 "・氏名のセルはドロップダウンです。この科のメンバーから選んでください",
                 "・オンコール（泊まらない・呼び出し待機）は、一覧の括弧つきの方を選びます → (管理者)",
                 "  当直（泊まる）は括弧なし。旧来の当直表と同じ書き方です",
                 "・氏名は姓だけで出ます。同姓の人がいる場合だけ、区別がつくまで名が足されます",
                 "・空欄は「その人数は居ない」という意味で正常です",
                 "・斜線のマスは、その日にその枠が立たないところです。書き込まないでください",
                 "・欄が足りないときは、先にアプリで入力して書き出し直すと、人数分の欄が出ます", "",
                 "■ 取り込み",
                 "・書けたらアプリの「Excelから取り込む」でこのファイルを読み込ませてください",
                 "・メンバーを増やすときは、先にアプリの①で足してから、もう一度書き出してください", "",
                 "■ 当直交代",
                 "・確定した月の表を書き換えると、変わったところが赤字になります",
                 "・そのまま取り込むと、当直交代として記録されます"):
        hp.append([line])
    hp.column_dimensions["A"].width = 80

    mt = wb.create_sheet(SH_META)
    for k, v in (("schema", SCHEMA), ("year", year), ("month", month), ("owner", owner),
                 ("source", source), ("master_rev", master_rev), ("tool", tool),
                 ("generated", dt.datetime.now().isoformat(timespec="seconds"))):
        mt.append([k, v])
    mt.sheet_state = "veryHidden"


# ---------------------------------------------------------------- 読み取り
def _day_of(v):
    """見える日付セル（「3 (火)」＋改行＋祝日名）→ 日の数。読めなければ None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    n = ""
    for ch in str(v or "").strip():
        if not ch.isdigit():
            break
        n += ch
    return int(n) if n else None


def _has_value(ws, r):
    """その行の日の欄（月〜日）に何か書かれているか。"""
    for c in range(COL_MON, COL_SUN + 1):
        v = ws.cell(r, c).value
        if v is not None and str(v).strip():
            return True
    return False


def _geometry(ws, year=None, month=None, rejects=None):
    """A列のキーだけで (行, 列, 枠, バンド, 種別, ISO日付) を組み立てる。
    行を挿入されても列を動かされても、キーのある行だけを拾うので壊れない。

    rejects（リスト）を渡すと、黙って読み落とす代わりに理由を積む（20260926）:
      ・週ブロックの中（隠しISO行〜備考行）に挿入された行（A列にキーが無い）に文字がある
        → その行は取り込めない（キーが無い行はどの枠の欄か分からない）
      ・隠しISO行と見える日付行（と、このファイルの年月）が食い違う
        → 日付がずれて読まれるので、その週は取り込まない
      ・隠しISO行が失われた週に氏名がある → 日付が分からないので取り込まない"""
    cells, notes, isos, flags = [], [], {}, {}
    in_block = checked = iso_seen = False
    told = set()

    def reject(r, why):
        if rejects is not None and (r, why) not in told:
            told.add((r, why))
            rejects.append({"line": r, "row": {}, "why": why})

    def check_week(r, date_row):
        """隠しISO行の日付を、見える日付（あれば）とこのファイルの年月で確かめる。
        食い違えばその週を捨てて（isos を空に）理由を積む。"""
        nonlocal isos, flags, checked
        checked = True
        if rejects is None:
            return
        bad = None
        for c, iso in sorted(isos.items()):
            d = _iso_date(iso)
            if d is None:
                continue                    # 読めない日付は contract.import_rows が「日付が読めません」で弾く
            if year and month and (d.year, d.month) != (int(year), int(month)):
                bad = (f"この週の隠し行の日付（{d.isoformat()}）がこのファイルの年月（{year}年{month}月）と合いません。"
                       "この週は取り込めません。アプリから書き出し直してください")
                break
            if date_row is not None:
                vis = _day_of(ws.cell(date_row, c).value)
                if vis is not None and vis != d.day:
                    bad = (f"この週の日付が食い違っています（見える日付は {vis}日、隠し行は {d.month}月{d.day}日）。"
                           "この週は取り込めません。アプリから書き出し直してください")
                    break
        if bad:
            reject(date_row or r, bad)
            isos, flags = {}, {}

    for r in range(1, ws.max_row + 1):
        key = str(ws.cell(r, COL_KEY).value or "").strip()
        if key == KEY_ISO:
            isos, flags = {}, {}
            in_block, checked, iso_seen = True, False, True
            for c in range(COL_MON, COL_SUN + 1):
                v = ws.cell(r, c).value
                if v not in (None, ""):
                    isos[c] = v
        elif key == KEY_FLAG:               # 土日祝フラグ（昼夜レーンの展開に使う。20260922）
            in_block = True
            flags = {}
            for c in range(COL_MON, COL_SUN + 1):
                v = ws.cell(r, c).value
                if v is None or v == "":
                    continue
                flags[c] = bool(v) if isinstance(v, (int, float)) else str(v).strip() not in ("0", "")
        elif key == KEY_DATE:
            in_block = True
            if iso_seen:
                check_week(r, r)
        elif key == KEY_NOTE:
            notes.append((r, dict(isos)))
            in_block, checked, iso_seen = False, False, False
        elif key.count("|") >= 3:
            in_block = True
            if rejects is not None and not iso_seen:
                # 隠しISO行が失われた週（前の週の日付で読んでしまう）→ 読まずに理由を積む
                if _has_value(ws, r):
                    reject(r, "この週の日付の隠し行が失われていて、日付が分かりません。"
                              "この週は取り込めません。アプリから書き出し直してください")
                continue
            if not checked:
                check_week(r, None)
            fid, band, kind, _no = key.split("|")[:4]
            for c, iso in isos.items():
                cells.append((r, c, fid, band, kind, iso, flags.get(c)))
        elif in_block and _has_value(ws, r):
            # 週の中に挿入された行（A列にキーが無い）。どの枠の欄か分からないので読めない（20260926）
            reject(r, "挿入された行は取り込めません。アプリで人数を増やしてから書き出し直してください")
    return cells, notes


def _geometry_fallback(ws, frames_own, holidays, year, month):
    """A列のキーが失われたとき。見える日付行から週ブロックを割り出し、
    枠マスタから組んだレーンの並びに当てる（生成規則が決まっているので位置で一致する）。"""
    lanes = build_columns(frames_own, [], holidays)
    H = block_height(lanes)

    def is_date_row(r):
        n = 0
        for c in range(COL_MON, COL_SUN + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and 1 <= int(v) <= 31:
                n += 1
            elif isinstance(v, str) and v.strip()[:2].rstrip().isdigit():
                n += 1
        return n >= 3

    first = next((r for r in range(1, ws.max_row + 1) if is_date_row(r)), None)
    if first is None:
        return [], []
    cells, notes, r0 = [], [], first - OFF_DATE
    while r0 + OFF_DATE <= ws.max_row:
        isos = {}
        for c in range(COL_MON, COL_SUN + 1):
            v = ws.cell(r0 + OFF_DATE, c).value
            s = str(v).strip() if v is not None else ""
            day = "".join(ch for ch in s[:2] if ch.isdigit())
            if day:
                try:
                    isos[c] = dt.date(year, month, int(day)).isoformat()
                except ValueError:
                    pass
        if not isos:
            break
        for li, c_ in enumerate(lanes):
            for c, iso in isos.items():
                d0 = _iso_date(iso)
                cells.append((r0 + OFF_LANE + li, c, c_.fid_key, c_.band, c_.kind, iso,
                              C.is_holiday(d0, holidays) if d0 else None))
        notes.append((r0 + OFF_LANE + len(lanes), dict(isos)))
        r0 += H
    return cells, notes


def _rows_from(ws, cells, by_disp, holidays=None):
    out = []
    for r, c, fid, band, kind, iso, hol in cells:
        v = ws.cell(r, c).value
        raw = str(v).strip() if v is not None else ""
        if not raw:
            continue
        oc = wrapped_parens(raw) or kind == KIND_ONCALL   # 括弧書き＝オンコール（旧形式のキーも受ける）
        s = strip_parens(raw)
        m = (by_disp.get(s) or by_disp.get(C.canon(C.nfc(s)))
             or by_disp.get(C.namekey(s)) or by_disp.get(C.canon(C.nfc(C.namekey(s)))))
        # 「昼夜」1レーン（A列キーが `夜fid+昼fid`。20260922）は、土日祝なら昼と夜の両方へ、
        # 平日は夜だけへ展開する。これで原子レコードが元どおりに戻る
        if "+" in fid:
            night, day = (fid.split("+") + [""])[:2]
            if hol is None:                               # 隠しフラグ行が失われたブック
                d0 = _iso_date(iso)
                hol = C.is_holiday(d0, holidays or {}) if d0 else False
            targets = [(night, "夜")] + ([(day, "昼")] if hol else [])
        else:
            targets = [(fid, band)]
        # 突合できたセルは **このブックでの表示名** で返す。フルネームに戻してしまうと、
        # 完全同姓同名の2人が区別できなくなる（people_index が表示名も引けるようにしてある）
        for fid2, band2 in targets:
            out.append({"日付": iso, "枠": fid2, "昼夜": band2,
                        "氏名": (m["disp"] or m["name"]) if m else s,
                        "オンコール": "○" if oc else "",
                        "_line": r, "_col": c, "_disp": raw,
                        "_lane": "昼夜" if "+" in fid else band})     # 灰色のマスの判定用（split_gray）
    return out


def split_gray(rows, frames, holidays):
    """_rows_from の行 → (取り込む行, 弾き)。その日にその枠が立たないマス（灰色の斜線）に書かれた氏名を
    理由つきで弾く（20260926。以前は平日の昼枠などに書いても黙って取り込まれていた）。
    規則は Col.stands_on と同じ（contract.band_stands）。
    「昼夜」1レーン（夜枠＋昼枠）の昼は、昼枠が立たない日（管理当直の開院記念日など）だけ黙って落とす
    （そのマスは夜のマスとして見えている＝灰色ではない）。"""
    fmap = {f["frame_id"]: f for f in frames}
    oo = _oncall_only_days()
    kept, rejects = [], []
    for row in rows:
        f = fmap.get(str(row.get("枠") or ""))
        d = _iso_date(row.get("日付"))
        if f is None or d is None or row.get("昼夜") != "昼" or C.band_stands(f, "昼", d, holidays, oo):
            kept.append(row)
            continue
        if row.get("_lane") == "昼夜":
            continue
        who = row.get("_disp") or row.get("氏名")
        rejects.append({"line": row.get("_line"),
                        "row": {"日付": row.get("日付"), "枠": row.get("枠"), "昼夜": "昼", "氏名": who},
                        "why": f"{C.md_label(d)} 昼: この日はこの枠が立ちません（灰色のマス）。『{who}』を消してください"})
    return kept, rejects


def is_duty_sheet(path):
    """新形式かどうか（旧フラット行テンプレと見分ける）。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        try:
            return SH_META in wb.sheetnames or SH_ID in wb.sheetnames
        finally:
            wb.close()
    except Exception:
        return False


def read_book(path, frames_own=None, holidays=None):
    """当直表ブック → {meta, owner, year, month, members, rows, base_rows, notes}。
    rows / base_rows は contract.import_rows の入力形。
    rejects = 行の挿入・隠し日付行の食い違いなど、ブックの形が壊れていて読めなかったところ（20260926）。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    meta = {}
    if SH_META in wb.sheetnames:
        for row in wb[SH_META].iter_rows(values_only=True):
            if row and row[0]:
                meta[str(row[0])] = row[1] if len(row) > 1 else None
    ws = wb[SH_MAIN] if SH_MAIN in wb.sheetnames else wb[wb.sheetnames[0]]

    members = []
    if SH_ID in wb.sheetnames:
        for row in wb[SH_ID].iter_rows(min_row=2, values_only=True):
            login, sei, mei, d = (list(row) + [None] * 4)[:4]
            if not (login or sei):
                continue
            sei, mei = str(sei or ""), str(mei or "")
            members.append({"login": str(login or ""), "sei": sei, "mei": mei,
                            "disp": str(d or ""), "name": (sei + " " + mei).strip(),
                            "title": "", "code": ""})
    # セルの文字 → メンバー。
    #   1) ドロップダウンの表示名（教 管理者 など。作りから一意）
    #   2) フルネーム・職位の印を外した表示名（「管理者」）は、**1人にしか当たらないときだけ**（20260926）。
    #      完全同姓同名の2人のフルネームを先着の1人に寄せない＝contract.import_rows が
    #      「同姓同名が2人います。ドロップダウンから選んでください」で弾く
    by_disp = {}
    for m in members:
        for k in (m["disp"], C.canon(C.nfc(m["disp"]))):
            if k:
                by_disp.setdefault(k, m)
    loose = defaultdict(set)
    for i, m in enumerate(members):
        ks = [C.namekey(m["name"]), C.canon(C.nfc(C.namekey(m["name"])))]
        bare = _strip_rank(m["disp"])
        if bare and bare != m["disp"]:
            ks += [bare, C.canon(C.nfc(bare))]
        for k in ks:
            if k:
                loose[k].add(i)
    for k, idx in loose.items():
        if k not in by_disp and len(idx) == 1:
            by_disp[k] = members[next(iter(idx))]

    try:
        y0, m0 = int(meta.get("year") or 0), int(meta.get("month") or 0)
    except (TypeError, ValueError):
        y0 = m0 = 0
    rejects = []
    cells, note_rows = _geometry(ws, y0, m0, rejects)
    if not cells and frames_own:
        cells, note_rows = _geometry_fallback(ws, frames_own, holidays or {},
                                              int(meta.get("year") or 0), int(meta.get("month") or 0))
    rows = _rows_from(ws, cells, by_disp, holidays)
    base_rows = _rows_from(wb[SH_BASE], cells, by_disp, holidays) if SH_BASE in wb.sheetnames else []
    notes = []
    for r, isos in note_rows:
        for c, iso in isos.items():
            v = ws.cell(r, c).value
            if v not in (None, ""):
                notes.append({"date": iso, "text": str(v).strip(), "_line": r})
    wb.close()
    return {"meta": meta, "owner": meta.get("owner"), "year": meta.get("year"),
            "month": meta.get("month"), "members": members, "rows": rows,
            "base_rows": base_rows, "notes": notes, "rejects": rejects}


_RANKS = ("教", "研", "専", "初")


def _strip_rank(disp):
    """表示名から職位の印を外す（「教 管理者」→「管理者」）。印が無ければそのまま。"""
    t = str(disp or "").strip()
    if len(t) >= 3 and t[0] in _RANKS and t[1] in (" ", "\u3000"):
        return t[2:].strip()
    return t
