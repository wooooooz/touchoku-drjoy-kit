# -*- coding: utf-8 -*-
r"""touchoku_xlsx.py  20261009
当直表xlsx の書き出し・取り込みだけをやる CLI（GUI なし）。Electron アプリが子プロセスで呼ぶ。
設計の正本 = docs/設計_当直表xlsx一本化_20260912.md §5

  touchoku_xlsx.py export --owner "循環器内科" --ym 2026-10 --out <path.xlsx>
                          [--from <draft.json>] [--master <master.json>]
                          [--config <current.json>] [--root <月フォルダの親>]
  touchoku_xlsx.py import --in <path.xlsx> --out <result.json>
                          [--config <current.json>]

終了コード 0 = 成功 / 1 = 失敗（理由を stderr に1行）。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402
import duty_sheet as D        # noqa: E402
import master_store as M      # noqa: E402


def _here():
    return os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))


def load_config(path=None):
    """アプリと同じ current.json（枠・名簿・祝日つき）。省略時はヘルパーの隣の config\。"""
    path = path or os.path.join(_here(), "config", "current.json")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def split_name(name):
    """『姓 名』→ (姓, 名)。空白が無ければ全部を姓として扱う（手入力の名簿外医師など）。"""
    s = C.namekey_display(name)
    sei, _, mei = s.partition(" ")
    return sei, mei


def members_from_roster(roster):
    """提出/マスターの roster → duty_sheet が食うメンバー。姓名は『姓 名』を割って作る。"""
    out = []
    for p in roster or []:
        sei, mei = split_name(p.get("name") or "")
        out.append({"login": p.get("login") or "", "code": p.get("code") or "",
                    "sei": p.get("sei") or sei, "mei": p.get("mei") or mei,
                    "name": p.get("name") or f"{sei} {mei}".strip(),
                    "kana": p.get("kana") or "", "title": p.get("title") or "",
                    "in_house": p.get("in_house", True), "work": C.work_of(p),   # 20260927: 区分3つ
                    "unregistered": bool(p.get("unregistered")),
                    **({"order": p["order"]} if p.get("order") is not None else {})})
    return out


def attach_fixed_order(members, cfg, owner):
    """管理当直: config.fixedMembers の order（管理当直アプリの名簿の並び）をメンバーに付ける（管理者 20260923）。"""
    fo = {x.get("login"): x.get("order") for x in ((cfg.get("fixedMembers") or {}).get(owner) or [])}
    for m in members:
        if m.get("login") in fo and fo[m["login"]] is not None:
            m["order"] = fo[m["login"]]
    return members


def default_members(cfg, owner):
    """--from もマスターも無いときの仮のメンバー（アプリ①の初期候補と同じ規則）。
    決まった初期メンバー（管理当直 = 管理当直アプリの名簿。管理者 20260923）があれば、所属パターンより優先。"""
    fixed = [x.get("login") for x in ((cfg.get("fixedMembers") or {}).get(owner) or [])]
    if fixed:
        by = {p.get("login"): p for p in cfg.get("roster") or []}
        return members_from_roster([by[lg] for lg in fixed if lg in by])
    entry = (cfg.get("deptMap") or {}).get(owner, [])
    return members_from_roster([p for p in cfg.get("roster") or []
                                if p.get("doctor") and C.matches_dept_map(entry, p)])


def add_duty_people(members, cfg, *duty_lists):
    """割当に出てくるのに members に居ない人を足す（20261009）。
    config の名簿（Dr.JOY）に居ればそこから職名・かなも引き、居なければ割当の氏名から作る。"""
    have = {m.get("login") for m in members}
    by = {p.get("login"): p for p in cfg.get("roster") or []}
    add = []
    for rows in duty_lists:
        for r in rows or []:
            lg = r.get("staff_login")
            if lg and lg not in have and lg in by:
                have.add(lg)
                add.append(by[lg])
    return D.members_with_duties(list(members) + members_from_roster(add), *duty_lists)


def frames_of(cfg, owner):
    frames = cfg.get("frames") or C.load_frames()
    own = [f for f in C.active_frames(frames) if f.get("owner", f["dept"]) == owner]
    if not own:
        raise SystemExit(f"枠マスタに『{owner}』の枠がありません")
    return frames, own


def resolve_master(args, year, month):
    if args.master:
        return args.master
    root = args.root or C.PROD_ROOT
    return M.master_path(os.path.join(root, C.month_dir_name(year, month)), year, month)


def cmd_export(args):
    year, month = [int(x) for x in args.ym.split("-")[:2]]
    cfg = load_config(args.config)
    frames, own = frames_of(cfg, args.owner)
    holidays = C.load_holidays()

    duties, members, baseline, source = [], [], None, "blank"
    # 対象月（config の月）でない月か。その月の名簿はその月のマスターのもの（20261009）
    other_month = bool(cfg.get("year")) and (int(cfg.get("year")), int(cfg.get("month") or 0)) != (year, month)
    if args.__dict__.get("from_"):
        with open(args.from_, encoding="utf-8-sig") as f:
            dr = json.load(f)
        duties = dr.get("duties") or []
        # 20261009: 違う月の書き出しでは画面（対象月）の名簿を使わない。以前はアプリが対象月の名簿を送り、
        #   その名簿に居ない人の当直が表から黙って消えていた（その Excel を取り込むと、その月の予定からも消える）
        if not other_month:
            members = members_from_roster(dr.get("roster"))
        source = "draft"
    mpath = resolve_master(args, year, month)
    ent = None
    if os.path.exists(mpath):
        try:
            ent = M.get_dept(M.load(mpath, year, month), args.owner)
        except Exception as e:
            print(f"マスターが読めませんでした（無視して続けます）: {e}", file=sys.stderr)
    if ent:
        baseline = ent.get("baseline")
        # 画面の割当が無い（キーが無い／空）ときはマスターの割当で書く（20260926: 対象月でない月の書き出しで
        # アプリが duties:[] を送り、白紙の当直表が出ていた。キーを省く形に変わってもどちらでも同じ）
        if not duties and ent.get("duties"):
            duties, source = ent["duties"], "master"
        if not members:
            members = members_from_roster(ent.get("roster"))
    # 割当・元の割当に出てくるのに名簿に居ない人は足す（20261009）。名簿（config の roster）に居ればそこから、
    #   居なければ割当の氏名から（duty_sheet.members_with_duties）。黙って表から消さない
    members = add_duty_people(members, cfg, duties, baseline)
    if not members:
        members = default_members(cfg, args.owner)
        if not members:
            raise SystemExit(f"『{args.owner}』のメンバーが分かりません。アプリで①を確定してから書き出してください")

    attach_fixed_order(members, cfg, args.owner)
    info = D.write_book(args.out, year=year, month=month, owner=args.owner, frames_own=own,
                        holidays=holidays, members=members, duties=duties, baseline=baseline,
                        deadline=(cfg.get("deadline") or "") if (cfg.get("year"), cfg.get("month")) == (year, month) else "",
                        source=source, master_rev=(ent or {}).get("rev"),
                        asof=dt.datetime.now(),        # マスターを読んだ時点（見出し右の「○月○日○時○分現在」）
                        # 当番当直のマスはグレー地に白字（20260925）。月の config の表を使う（exe に同梱した data/ は古いことがある）。
                        # config に toban が無い（古い config）ときだけ data/toban.csv
                        toban=cfg.get("toban"))
    print(json.dumps({"ok": True, "out": args.out, "owner": args.owner, "year": year,
                      "month": month, "source": source, "members": len(members), **info},
                     ensure_ascii=False))
    return 0


def diff_changes(base_records, records):
    """_原本 と現在値の差 → 交代一覧。**受理された原子レコード同士**で比べる。

    セル同士で比べると、名簿に無い氏名など「取り込まれなかったもの」まで交代として数えてしまい、
    duties と changes が食い違う。枠×日×種別×バンドの集合で比べ、同じ人・同じ枠で
    昼と夜が揃って変わったものは 1件（昼夜）にまとめる。"""
    def group(rr):
        g = {}
        for r in rr:
            k = (r["date"], r.get("frame_id"), "オンコール" if r.get("oncall") else "当直", r["band"])
            g.setdefault(k, set()).add((r.get("staff_login") or "", r.get("name") or ""))
        return g

    a, b = group(base_records), group(records)
    raw = []
    for k in sorted(set(a) | set(b)):
        gone = sorted(a.get(k, set()) - b.get(k, set()))
        came = sorted(b.get(k, set()) - a.get(k, set()))
        date, fid, kind, band = k
        for i in range(max(len(gone), len(came))):
            f_ = gone[i] if i < len(gone) else ("", "")
            t_ = came[i] if i < len(came) else ("", "")
            raw.append({"date": date, "frame_id": fid, "kind": kind, "band": band,
                        "from": f_[1], "to": t_[1],
                        "from_login": f_[0], "to_login": t_[0]})
    merged, seen = [], {}
    for c in raw:
        k = (c["date"], c["frame_id"], c["kind"], c["from_login"], c["to_login"])
        if k in seen and {seen[k]["band"], c["band"]} == {"昼", "夜"}:
            seen[k]["band"] = "昼夜"
            continue
        seen[k] = dict(c)
        merged.append(seen[k])
    return merged


def canonical_owner(name, frames):
    """科名 → 枠マスタの提出単位の正式名。空白の有無は問わない（ファイル名は「循環器内科」、
    枠マスタは「循環器内科」。20260926）。当たらなければそのまま返す。"""
    k = C.namekey(name)
    for f in frames:
        o = f.get("owner", f["dept"])
        if C.namekey(o) == k:
            return o
    return name


def cmd_import(args):
    holidays = C.load_holidays()
    # 祝日表も渡す: 隠しの土日祝フラグ行が失われたブックでも、昼夜レーンの昼（祝日）を落とさない（20260926）
    bk = D.read_book(args.in_, holidays=holidays)
    cfg = load_config(args.config)
    owner = bk["owner"] or ""
    year, month = bk["year"], bk["month"]
    if not (owner and year and month):     # _メタ が無い/壊れている場合はファイル名から
        base = os.path.splitext(os.path.basename(args.in_))[0]
        parts = base.split("_")
        owner = owner or (canonical_owner(parts[-1], cfg.get("frames") or C.load_frames())
                          if len(parts) >= 3 else "")
        if not (year and month) and len(parts) >= 2:
            import re
            m = re.fullmatch(r"R(\d+)年(\d+)月", parts[-2])
            if m:
                year, month = 2018 + int(m.group(1)), int(m.group(2))
    if not (owner and year and month):
        raise SystemExit("このファイルの科名・年月が分かりません（_メタ が失われています）")
    year, month = int(year), int(month)

    frames, _own = frames_of(cfg, owner)
    code_by_login = {p.get("login"): p.get("code") for p in cfg.get("roster") or []}
    members = [dict(m, code=code_by_login.get(m["login"], "")) for m in bk["members"]]
    pidx = D.people_index(members)

    def take(rows, warnings):
        # 灰色のマス（その日にその枠が立たない）に書かれた氏名は理由つきで弾く（20260926）
        kept, gray = D.split_gray(rows, frames, holidays)
        recs, rejs = C.import_rows(kept, owner, year, month, frames, holidays, *pidx,
                                   roster_label="この科のメンバー",
                                   not_found_hint="。アプリの①で追加してから書き出し直してください",
                                   warnings=warnings)
        return recs, gray + rejs

    # 定員を超える「当」は弾かずに残し warnings で知らせる（連名の当直もある。20260926）
    warnings = []
    records, rejects = take(bk["rows"], warnings)
    rejects = sorted((bk.get("rejects") or []) + rejects,
                     key=lambda x: (x.get("line") is None, x.get("line") or 0))
    # _原本 も同じ検証を通してからレコード同士で比べる（弾いた行が交代に紛れ込まないように）
    changes = diff_changes(take(bk["base_rows"], [])[0], records) if bk["base_rows"] else []
    res = {"ok": True, "owner": owner, "year": year, "month": month,
           "duties": records, "changes": changes, "rejects": rejects, "warnings": warnings,
           "notes": bk["notes"], "meta": bk["meta"]}
    with open(args.out, "w", encoding="utf-8", newline="\n") as f:
        json.dump(res, f, ensure_ascii=False)
    print(json.dumps({"ok": True, "owner": owner, "year": year, "month": month,
                      "duties": len(records), "changes": len(changes),
                      "rejects": len(rejects), "warnings": len(warnings),
                      "notes": len(bk["notes"])}, ensure_ascii=False))
    return 0


def main(argv=None):
    # 20260912: PyInstaller の exe は -X utf8 が効かず stdout/stderr が CP932 になる。
    # アプリ（main.js）は utf8 で読むので、ここで固定する（--out の JSON は元から UTF-8）。
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ap = argparse.ArgumentParser(prog="touchoku_xlsx")
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("export", help="当直表xlsxを書き出す")
    e.add_argument("--owner", required=True)
    e.add_argument("--ym", required=True, help="2026-10 の形")
    e.add_argument("--out", required=True)
    e.add_argument("--from", dest="from_", default=None, help="アプリの画面にある割当（draft.json）")
    e.add_argument("--master", default=None)
    e.add_argument("--root", default=None, help="月フォルダの親（既定=本番共有）")
    e.add_argument("--config", default=None)
    e.set_defaults(func=cmd_export)
    i = sub.add_parser("import", help="記入済みの当直表xlsxを読む")
    i.add_argument("--in", dest="in_", required=True)
    i.add_argument("--out", required=True)
    i.add_argument("--config", default=None)
    i.set_defaults(func=cmd_import)
    a = ap.parse_args(argv)
    try:
        return a.func(a)
    except SystemExit:
        raise
    except Exception as ex:
        print(f"{type(ex).__name__}: {ex}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
