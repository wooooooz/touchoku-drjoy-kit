# -*- coding: utf-8 -*-
r"""unreg_rebind.py — 名簿外（手入力）で入れた医師を、Dr.JOY に登録された本人に付け替える。20261031

管理者 20261028: 名簿外で入れた医師があとで Dr.JOY に登録されたら、自動で本人に置き換える。
  同姓同名で所属（初期研修医はその月の1日のローテーション先）が同じなら、機械的に結び付けてよい（承認は要らない）。

決まり:
  名簿外 = ログインIDが unreg:氏名（または unregistered）。氏名（空白・NFC・異体字を寄せる＝contract.canon）が同じ
  Dr.JOY の医師がちょうど1人いて、その人のその月の1日の所属（shozoku.History.dept_on・初期研は所属履歴）が、
  名簿外として載っている提出単位のどれかの所属（dept_map）に当たれば（月の途中に始まる予定の組も見る＝初期研のローテーション）、その名簿外のキーを本人に付け替える。
  同じキーがほかの提出単位にも載っていれば、そこも同じ本人に付け替える（同じ人なので）。
  当たらないもの（同姓同名が2人以上・所属が違う・Dr.JOY にいない）は付け替えず、理由を返す（ログに残す）。
書き換える場所: その月の共有マスターの各科の roster・duties・baseline（決裁のときの割当）を同時に。
  書き換えは master_store.put_dept（ロック・控え・rev の照合）。by = "unreg_rebind"。
  交代記録（kotai_record）は名簿外と本人を同じ人として比べる（_compact が norm_login を通す）ので、交代申請にはならない。
毎日15時の反映（dtx14_daily_sync）が、交代記録の前に前月〜来月について呼ぶ。

usage: python -X utf8 tools\unreg_rebind.py --month 2026-11 [--apply]   （--apply なしは下見だけ）
"""
import argparse
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C            # noqa: E402
import master_store as MS       # noqa: E402
import shozoku as SZ            # noqa: E402

BY = "unreg_rebind"


def is_unreg(login, m=None):
    return str(login or "").startswith("unreg:") or bool((m or {}).get("unregistered"))


def nkey(name):
    return C.canon(C.nfc(C.namekey(name or "")))


def doctors_by_name(people):
    out = {}
    for p in people:
        if p.get("is_doctor", True) and p.get("login"):
            out.setdefault(nkey(p.get("name")), []).append(p)
    return out


def plan(master, people, hist, dept_map, y, m):
    """戻り: (付け替え {名簿外のログイン: Dr.JOY の人}, 付け替えない [(提出単位, 氏名, 理由)])"""
    d1 = dt.date(y, m, 1)
    dend = (d1 + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    by_name = doctors_by_name(people)
    seen = {}                                     # 名簿外のログイン → (氏名, [提出単位])
    for owner, ent in (master.get("depts") or {}).items():
        for mem in ent.get("roster") or []:
            if is_unreg(mem.get("login"), mem):
                seen.setdefault(mem["login"], (mem.get("name") or "", []))[1].append(owner)
        for r in ent.get("duties") or []:
            if is_unreg(r.get("staff_login"), r):
                s = seen.setdefault(r["staff_login"], (r.get("name") or "", []))
                if owner not in s[1]:
                    s[1].append(owner)
    mapping, skipped = {}, []
    for login, (name, owners) in seen.items():
        cands = by_name.get(nkey(name), [])
        if not cands:
            skipped += [(o, name, "Dr.JOY の名簿にいない") for o in owners]
            continue
        if len(cands) > 1:
            skipped += [(o, name, f"同姓同名が {len(cands)} 人") for o in owners]
            continue
        p = cands[0]
        # その月のうちに就く所属（1日の所属＋月の途中に始まる予定の組。初期研のローテーションは 11/20 などに替わる）
        deps = hist.depts_during(p["login"], d1, dend, p.get("title"))

        def fits(o):
            # 所属に加えて職位の条件も見る（内科C 初期研は初期研だけ。コードレビュー 20261031: 同姓同名の教授に付け替えるおそれ）
            pats, ranks = C.dept_map_entry(dept_map.get(o))
            return (any(SZ.matches(dep, pats) for dep in deps)
                    and (not ranks or C.rank_of(p.get("title")) in ranks))
        hit = [o for o in owners if fits(o)]
        if hit:
            mapping[login] = p
        else:
            shown = "・".join(SZ.dept_display(x) for x in deps) or "なし"
            for o in owners:
                pats, ranks = C.dept_map_entry(dept_map.get(o))
                why = ("この提出単位には所属の決まりが無い（管理当直など）" if not pats
                       else f"職位が違う（{p.get('title') or '職名なし'}）" if any(SZ.matches(dep, pats) for dep in deps)
                       else f"所属が違う（Dr.JOY の {m}月の所属: {shown}）")
                skipped.append((o, name, why))
    return mapping, skipped


def _rebind_entry(ent, mapping):
    """1科分の roster・duties・baseline を書き換えた (roster, duties, baseline, 変えたか)。"""
    changed = False

    def swap_rec(r):
        nonlocal changed
        p = mapping.get(r.get("staff_login"))
        if not p:
            return r
        changed = True
        r2 = {k: v for k, v in r.items() if k != "unregistered"}
        r2["staff_login"], r2["name"] = p["login"], p.get("name") or r.get("name")
        return r2

    duties = [swap_rec(dict(r)) for r in ent.get("duties") or []]
    base = ent.get("baseline")
    baseline = None if base is None else [swap_rec(dict(r)) for r in base]
    roster, have = [], {str(x.get("login") or "").lower() for x in ent.get("roster") or [] if not is_unreg(x.get("login"), x)}
    for mem in ent.get("roster") or []:
        p = mapping.get(mem.get("login"))
        if not p:
            roster.append(mem)
            continue
        changed = True
        if p["login"].lower() in have:           # 本人がもう載っていれば、名簿外の行は落とす（区分は本人の行のまま）
            continue
        m2 = {k: v for k, v in mem.items() if k != "unregistered"}
        m2.update(login=p["login"], name=p.get("name") or mem.get("name"), title=p.get("title") or mem.get("title"), code="")
        roster.append(m2)
        have.add(p["login"].lower())
    return roster, duties, baseline, changed


def apply(path, y, m, mapping, log=print):
    """マスターに書く。戻り: 書き換えた提出単位の一覧。"""
    done = []
    master = MS.load(path, y, m)
    for owner, ent in (master.get("depts") or {}).items():
        roster, duties, baseline, changed = _rebind_entry(ent, mapping)
        if not changed:
            continue
        try:
            MS.put_dept(path, owner, duties=duties, roster=roster,
                        baseline=baseline if ent.get("baseline") is not None else False,
                        by=BY, expect_rev=int(ent.get("rev") or 0), year=y, month=m)
            done.append(owner)
        except MS.Conflict:
            log(f"  {owner}: 今ちょうど保存された → 名簿外の付け替えは次の回に")
    return done


def run_month(y, m, *, do_apply=True, log=print, people=None, hist=None, dept_map=None):
    """1か月分。戻り: (付け替えたもの [(提出単位, 氏名, ログインID)], 付け替えないもの)。"""
    path = MS.master_path(C.prod_month_dir(y, m), y, m)
    if not os.path.isfile(path):
        return [], []
    master = MS.load(path, y, m)
    people = people if people is not None else C.load_staffs(C.newest_staffs())[0]
    hist = hist or SZ.load_history()
    dept_map = dept_map if dept_map is not None else SZ.load_dept_map()
    mapping, skipped = plan(master, people, hist, dept_map, y, m)
    if not mapping:
        return [], skipped
    owners = apply(path, y, m, mapping, log=log) if do_apply else sorted(master.get("depts") or {})
    rows = []
    for owner in owners:
        ent = (master.get("depts") or {}).get(owner) or {}
        names = {x.get("login"): x.get("name") for x in ent.get("roster") or []}
        names.update({r.get("staff_login"): r.get("name") for r in ent.get("duties") or []})
        for lg, p in mapping.items():
            if lg in names:
                rows.append((owner, p.get("name"), p["login"]))
    return rows, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="YYYY-MM")
    ap.add_argument("--apply", action="store_true", help="共有マスターを書き換える（なしは下見）")
    a = ap.parse_args()
    y, m = (int(x) for x in a.month.split("-"))
    rows, skipped = run_month(y, m, do_apply=a.apply)
    for o, n, lg in rows:
        print(f"{'付け替えた' if a.apply else '付け替える'}: {o} {n} → {lg.split('@')[0]}")
    for o, n, why in skipped:
        print(f"付け替えない: {o} {n}（{why}）")
    if not rows and not skipped:
        print("名簿外の医師はいない")


if __name__ == "__main__":
    main()
