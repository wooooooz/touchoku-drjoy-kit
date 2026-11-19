# -*- coding: utf-8 -*-
"""bucho_roster.py — 管理当直アプリの名簿 → 当直勤務表アプリの「管理当直」の初期メンバー。20261028

管理者 2026-09-23: 管理当直は当直勤務表アプリで事務担当が「管理当直」を選んで入力する。名簿は管理当直アプリの名簿を使う。
管理当直アプリの名簿（本番 `00_DrJOY\\管理当直\\tools\\data\\roster.csv`。列 no,oncall,dept,name,reading）は
氏名が「甲野教授」「丙川智准教授」のような 姓（＋名の頭）＋職位 で、ログインIDを持たない。
ここで Dr.JOY の名簿（staffs）と結び付ける:
  1) 職位（教授・准教授・副部長…）を外した文字が、姓と一致する人（無ければ 姓名の先頭が一致する人。「丙川智」など）
  2) 読み（reading）が かな の先頭と一致する人（ず/づ・じ/ぢ は同じとみなす。「てずか」＝「てづか」）
  3) 職位が教員（contract.rank_of = 教）の人
  4) 名簿の科が Dr.JOY の所属に含まれる人（内科A＝循環器 のように所属の言葉に読み替える。
     救命センターなど Dr.JOY に所属の無い部署の部長は、その人の本来の所属を DEPT_BY_NAME / DEPT_BY_NO に書く）
  20260925: どの段も、絞った結果が0人なら「絞らずに前の候補に戻る」ことはしない＝結び付けずに飛ばして log に出す
  （以前は0人だと前の候補に戻していたので、44人中12人が科の確かめ無しで結ばれていた。部長が Dr.JOY の名簿から
   消えたときに同姓の若手や他科の教員に結ばれるおそれがあった）。最後に1人に決まらなかった人も飛ばす
  （事務担当がアプリの①の検索で足せる）。
oncall 列が TRUE の部長は、その人を選んだときに最初から「オ」にする（R8.9 の現物で印と括弧書きが全件一致）。
20261028 管理者: 読みの段で0人になっても、同じ姓・教員・所属が合う人がちょうど1人なら結び付ける（Dr.JOY の読みの入力ミスで
  外れていた＝庚田 一郎「からじま」）。log に「読みが Dr.JOY と違う」と残す。それでも結び付かない部長は unresolved() で返し、
  毎日15時の反映が事務担当へ知らせる（dtx14_daily_sync.bucho_check。アプリで足したあとは知らせない）。
"""
import csv
import os
import re
import unicodedata

import contract as C

ROSTER = os.path.join(os.path.dirname(C.PROD_ROOT), "管理当直", "tools", "data", "roster.csv")
OWNER = "管理当直"
TITLE = re.compile(r"(副院長|病院長|院長|准教授|教授|講師|助教|副部長|部長|科長|先生|医師)$")
DEPT_WORDS = {            # 見本: 管理当直アプリの科名（NFKC 後）→ Dr.JOY の所属に出てくる言葉。自院の科に合わせて書く
    "内科A": ("循環器", "腎臓"),
    "外科A": ("消化管",),
}
# Dr.JOY に所属の無いセンター・部の部長（兼任）→ その人の Dr.JOY の所属の言葉。20260925（R8.9 の Dr.JOY の所属で確かめた）。
# 名簿の氏名（職位つき）で引く＝部長が替われば効かなくなり、所属の確かめで0人→結び付けずに知らせる（黙って他人に結ばない）。
DEPT_BY_NAME = {          # 見本: 名簿の氏名（職位つき）→ その人の Dr.JOY の所属の言葉
    "見本教授": ("形成外科",),        # 救命センター
}
# 名簿の no → Dr.JOY の所属の言葉（見本: {"37": ("リハビリテーション部",)}）
DEPT_BY_NO = {}
_YOTSUGANA = str.maketrans({"づ": "ず", "ぢ": "じ"})
WHY = {"氏名": "Dr.JOY の名簿に同じ姓の人がいない", "読み": "読みが合う人がいない",
       "教員": "教員（職名＝医師）がいない", "所属": "所属が合う人がいない"}


def _flat(s):
    return C.canon(C.nfc((s or "").replace(" ", "").replace("　", "")))


def _hira(s):
    s = unicodedata.normalize("NFKC", s or "").replace(" ", "").replace("　", "")
    return "".join(chr(ord(c) - 0x60) if "ァ" <= c <= "ヶ" else c for c in s).translate(_YOTSUGANA)


def _surname(name):
    return _flat(unicodedata.normalize("NFKC", name or "").split(" ")[0])


def dept_words(r):
    """名簿の1行 → Dr.JOY の所属で確かめる言葉（氏名 → no → 科名 の順に引く）。"""
    name = (r.get("name") or "").strip()
    dept = unicodedata.normalize("NFKC", r.get("dept") or "").strip()
    return (DEPT_BY_NAME.get(name) or DEPT_BY_NO.get(str(r.get("no") or "").strip())
            or DEPT_WORDS.get(dept) or (dept,))


def surname_key(r):
    """名簿の氏名（庚田副部長 など）→ 職位を外した姓の照合キー。"""
    return _flat(TITLE.sub("", (r.get("name") or "").strip()))


def match(r, staff):
    """名簿の1行 → (候補の一覧, 止まった段 or "")。どの段も0人になったらそこで止める（前の候補に戻さない。20260925）。
    20261028: 読みの段で0人になったときだけ、読みを外して 教員・所属 で絞り、ちょうど1人なら結び付ける（段は "読み違い"）。"""
    key = surname_key(r)
    cands = ([p for p in staff if _surname(p["name"]) == key]
             or [p for p in staff if _flat(p["name"]).startswith(key)]) if key else []
    if not cands:
        return [], "氏名"
    reading = _hira(r.get("reading"))
    words = dept_words(r)
    is_kyo = lambda p: C.rank_of(p.get("title")) == "教"            # noqa: E731
    in_dept = lambda p: any(w and w in (p.get("dept") or "") for w in words)   # noqa: E731
    by_name = cands
    for step, pred in (("読み", lambda p: _hira(p.get("kana")).startswith(reading)), ("教員", is_kyo), ("所属", in_dept)):
        cands = [p for p in cands if pred(p)]
        if not cands:
            if step == "読み":
                alt = [p for p in by_name if is_kyo(p) and in_dept(p)]
                if len(alt) == 1:
                    return alt, "読み違い"
            return [], step
    return cands, ""


def resolve(people, path=None, log=print):
    """戻り: [{"login", "oncall", "order"}]。order = 名簿の no（管理当直アプリの Excel の並び。管理者 20260923）。読めなければ OSError。
    結び付かない部長は log に ⚠ で出す（month_setup は成功したときもそのままログに出す）。"""
    path = path or ROSTER
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.DictReader(f) if (r.get("name") or "").strip()]
    staff = [p for p in people if p.get("login")]
    out, miss, kana = [], [], []
    for r in rows:
        cands, stop = match(r, staff)
        if stop == "読み違い":
            kana.append(f"{r['name']}（名簿の読み {r.get('reading')}・Dr.JOY {cands[0].get('kana')}）")
            stop = ""
        if stop:
            miss.append(f"{r['name']}（{r.get('dept')}・{WHY[stop]}）")
        elif len(cands) == 1:
            out.append({"login": cands[0]["login"],
                        "oncall": str(r.get("oncall") or "").strip().upper() == "TRUE",
                        "order": int(r.get("no") or 0) if str(r.get("no") or "").strip().isdigit() else 999})
        else:
            miss.append(f"{r['name']}（{r.get('dept')}・候補が{len(cands)}人で決められない）")
    log(f"  管理当直の名簿: {len(out)}/{len(rows)} 人を Dr.JOY の名簿と結び付けた"
        + (f"（オンコールの印 {sum(1 for x in out if x['oncall'])} 人）" if out else ""))
    out.sort(key=lambda x: x["order"])
    if kana:
        log(f"  注意: 読みが Dr.JOY と違うが、姓・教員・所属で1人に決まったので結び付けた: {'、'.join(kana)}")
    if miss:
        log(f"  ⚠ 結び付かない部長 {len(miss)} 人（アプリの①の検索で足してください）: {'、'.join(miss)}")
    return out


def unresolved(people, path=None):
    """結び付かない部長 [(名簿の氏名, 科, 理由, 姓の照合キー)]。毎日15時の反映が事務担当へ知らせるのに使う（20261028）。"""
    path = path or ROSTER
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.DictReader(f) if (r.get("name") or "").strip()]
    staff = [p for p in people if p.get("login")]
    out = []
    for r in rows:
        cands, stop = match(r, staff)
        if stop == "読み違い":
            continue
        if stop:
            out.append((r["name"].strip(), r.get("dept") or "", WHY[stop], surname_key(r)))
        elif len(cands) != 1:
            out.append((r["name"].strip(), r.get("dept") or "", f"候補が{len(cands)}人で決められない", surname_key(r)))
    return out
