# -*- coding: utf-8 -*-
"""toban_plan.py — 1年分の当番当直（病院全体の当直を各科が持ち回る）の割当案を作る。20261008

管理者 2026-10-08 の申し送り「当番当直割当：2027年版」に従う。優先順:
  1. 1人当たり合計回数（合計 ÷ 人員）を揃える（最大/最小 ≤ 1.10）
  2. 各科の 休日-昼 / 休日-夜 / 平日-夜 の構成比を全体の構成比に近づける
  3. GW（4/29〜5/5）と年末年始（12/28〜1/3）は同じ科が2回以上当たらない（昼＋夜を合わせて。必ず守る）
  4. 同じ科の連日・短い間隔（休日6日・平日4日が目安）を減らす（罰点）
  ほかに必ず守ること: 同じ休日の昼と夜は別の科。月ごとにも人員比に近づける（2026年版で前半に大きい科が偏った反省）。

作り方: 1・2 は先に整数の割当数（科×区分）で決め打ちする（最大剰余法）。
  そのあと同じ区分の枠どうしで担当科を入れ替える焼きなましで 3・4・月ごとの偏りを詰める。
  入れ替えは科×区分の回数を変えないので、1・2 は最後まで崩れない。

休日 = 土日 + data/holidays.csv（祝日・年末年始 12/29〜1/3・開院記念日。管理者 20260922: 12/28 は開院日、
  記念日は土日なら翌月曜）。
人員 = 名簿\\staff_*.csv（Dr.JOY）の所属で数えた医師から初期研を除いた数（教員・大学院生・専攻医。管理者 20261008）。

usage: python -X utf8 tools\\toban_plan.py --year 2027 [--staff staff.csv] [--prev data\\toban.csv] [--out-dir DIR] [--seed N]
出力: toban_<year>.csv（date,band,dept・UTF-8 BOM・CRLF＝data/toban.csv と同じ形）と toban_<year>_検証.xlsx
"""
import argparse
import collections
import csv
import datetime as dt
import glob
import math
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Toban group (DTX14 first-sheet name, same spelling as data/toban.csv) -> Dr.JOY affiliations
GROUPS = {
    "内科": ["内科(腎臓・内分泌)", "内科(血液・膠原病)", "内科(神経)", "総合診療科"],
    "外科": ["外科(消化管)", "外科(肝胆膵・乳腺)"],
    "内科 消化器・呼吸器": ["内科(消化器)", "内科(感染症・呼吸器)"],
    "救急科": ["救急部"],
    "小児科": ["小児科"],
    "精神科": ["精神科"],
    "耳鼻咽喉科": ["耳鼻咽喉科"],
    "産科婦人科": ["産科婦人科"],
    "循環器内科": ["循環器内科"],
    "泌尿器科": ["泌尿器科"],
    "整形外科": ["整形外科"],
    "形成外科": ["形成外科"],
    "皮膚科": ["皮膚科"],
    "脳神経外科": ["脳神経外科"],
    "眼科": ["眼科"],
    "外科 心臓・呼吸器": ["外科(心臓血管・呼吸器)"],
}
CATS = ("休日-昼", "休日-夜", "平日-夜")
HARD = 1e6


def latest_staff():
    files = sorted(glob.glob(os.path.join(C.PROD_ROSTER, "staff_*.csv")))
    if not files:
        raise SystemExit("名簿の staff_*.csv が見つかりません")
    return files[-1]


def count_staff(path):
    with open(path, encoding="cp932", newline="") as f:
        rows = list(csv.DictReader(f))
    n = collections.Counter()
    for r in rows:
        if C.rank_of(r.get("職名")) == "初":
            continue  # 初期研 rotate between depts; not counted (管理者 20261008)
        segs = [s for s in (r.get("所属") or "").split("/") if s]
        hit = {g for g, pats in GROUPS.items() for s in segs if s in pats}
        for g in hit:
            n[g] += 1
    missing = [g for g in GROUPS if not n[g]]
    if missing:
        raise SystemExit(f"人員0の科があります: {missing}")
    return {g: n[g] for g in GROUPS}


def load_holidays():
    out = set()
    with open(os.path.join(ROOT, "data", "holidays.csv"), encoding="utf-8-sig") as f:
        for r in csv.reader(f):
            if r and r[0][:1].isdigit():
                out.add(dt.date.fromisoformat(r[0]))
    return out


def build_slots(year, holidays):
    """[(date, cat)] in calendar order (day before night)."""
    d, end, slots = dt.date(year, 1, 1), dt.date(year, 12, 31), []
    while d <= end:
        if d.weekday() >= 5 or d in holidays:
            slots += [(d, "休日-昼"), (d, "休日-夜")]
        else:
            slots.append((d, "平日-夜"))
        d += dt.timedelta(days=1)
    return slots


def largest_remainder(total, weights):
    s = sum(weights.values())
    raw = {k: total * w / s for k, w in weights.items()}
    q = {k: math.floor(v) for k, v in raw.items()}
    for k in sorted(raw, key=lambda k: raw[k] - q[k], reverse=True)[: total - sum(q.values())]:
        q[k] += 1
    return q


def quotas(staff, cat_total):
    """dept -> cat -> count. Row sums = per-capita fair totals, column sums = slot counts."""
    total = sum(cat_total.values())
    tot = largest_remainder(total, staff)
    frac = {d: {c: tot[d] * cat_total[c] / total for c in CATS} for d in staff}
    q = {d: {c: math.floor(frac[d][c]) for c in CATS} for d in staff}
    # Fill remainders cell by cell, largest fraction first, respecting both margins
    cells = sorted(((frac[d][c] - q[d][c], d, c) for d in staff for c in CATS), reverse=True)
    for _, d, c in cells:
        if sum(q[d].values()) < tot[d] and sum(q[x][c] for x in staff) < cat_total[c]:
            q[d][c] += 1
    assert all(sum(q[d].values()) == tot[d] for d in staff)
    assert all(sum(q[d][c] for d in staff) == cat_total[c] for c in CATS)
    return q


def windows(year):
    """Periods where a dept may appear at most once (day+night counted together)."""
    return [
        ("年末年始(前年から)", dt.date(year - 1, 12, 28), dt.date(year, 1, 3)),
        ("GW", dt.date(year, 4, 29), dt.date(year, 5, 5)),
        ("年末年始", dt.date(year, 12, 28), dt.date(year + 1, 1, 3)),
    ]


class Plan:
    def __init__(self, year, slots, staff, q, prev, seed):
        self.year, self.slots, self.staff, self.q = year, slots, staff, q
        self.depts = list(staff)
        self.prev = prev  # [(date, dept)] from the previous year (tail only matters)
        self.rng = random.Random(seed)
        self.wins = windows(year)
        tot_staff = sum(staff.values())
        per_month = collections.Counter(d.month for d, _ in slots)
        self.exp_m = {(g, m): staff[g] / tot_staff * per_month[m] for g in staff for m in range(1, 13)}
        self.cum_slots = {m: sum(per_month[k] for k in range(1, m + 1)) for m in range(1, 13)}
        self.share = {g: staff[g] / tot_staff for g in staff}

    # ---------- initial ----------
    def initial(self):
        assign = [None] * len(self.slots)
        for c in CATS:
            idx = [i for i, (_, cc) in enumerate(self.slots) if cc == c]
            # Spread each dept evenly over the year: stride ordering by target positions
            marks = []
            for g in self.depts:
                k = self.q[g][c]
                marks += [((j + 0.5) / k + self.rng.random() * 1e-3, g) for j in range(k)]
            marks.sort()
            for i, (_, g) in zip(idx, marks):
                assign[i] = g
        self.a = assign
        self.index()

    # ---------- cost ----------
    def dept_dates(self, g):
        out = [(d, None) for d, x in self.prev if x == g]
        out += [self.slots[i] for i in self.by[g]]
        return out

    def index(self):
        self.by = {g: set() for g in self.depts}
        for i, g in enumerate(self.a):
            self.by[g].add(i)

    def dept_cost(self, g):
        items = self.dept_dates(g)
        cost = 0.0
        # Hard: window at most once
        for _, s, e in self.wins:
            k = sum(1 for d, _ in items if s <= d <= e)
            if k > 1:
                cost += HARD * (k - 1)
        # Spacing
        items.sort(key=lambda t: (t[0], t[1] != "休日-昼"))
        for (d1, c1), (d2, c2) in zip(items, items[1:]):
            gap = (d2 - d1).days
            if gap == 0:
                cost += HARD  # same holiday day and night
            elif gap == 1:
                cost += 400
            hol = (c1 or "").startswith("休日") and (c2 or "").startswith("休日")
            if hol and gap < 6:
                cost += 25 * (6 - gap) ** 2
            elif not hol and gap < 4:
                cost += 25 * (4 - gap) ** 2
        # Monthly balance and cumulative balance (no front/back loading)
        mc = collections.Counter(d.month for d, c in items if c and d.year == self.year)
        cum = 0
        for m in range(1, 13):
            cost += 30 * (mc[m] - self.exp_m[(g, m)]) ** 2
            cum += mc[m]
            cost += 30 * (cum - self.share[g] * self.cum_slots[m]) ** 2
        return cost

    def total_cost(self):
        return sum(self.dept_cost(g) for g in self.depts)

    # ---------- anneal ----------
    def anneal(self, iters, t0=200.0, t1=0.05):
        by_cat = {c: [i for i, (_, cc) in enumerate(self.slots) if cc == c] for c in CATS}
        dc = {g: self.dept_cost(g) for g in self.depts}
        cur = sum(dc.values())
        best, best_a = cur, list(self.a)
        weights = [len(by_cat[c]) for c in CATS]
        for it in range(iters):
            t = t0 * (t1 / t0) ** (it / iters)
            c = self.rng.choices(CATS, weights)[0]
            i, j = self.rng.sample(by_cat[c], 2)
            gi, gj = self.a[i], self.a[j]
            if gi == gj:
                continue
            self.a[i], self.a[j] = gj, gi
            self.by[gi].discard(i); self.by[gi].add(j); self.by[gj].discard(j); self.by[gj].add(i)
            ni, nj = self.dept_cost(gi), self.dept_cost(gj)
            delta = ni + nj - dc[gi] - dc[gj]
            if delta <= 0 or self.rng.random() < math.exp(-delta / t):
                dc[gi], dc[gj] = ni, nj
                cur += delta
                if cur < best:
                    best, best_a = cur, list(self.a)
            else:
                self.a[i], self.a[j] = gi, gj
                self.by[gi].discard(j); self.by[gi].add(i); self.by[gj].discard(i); self.by[gj].add(j)
        self.a = best_a
        self.index()
        return best


def load_prev(path, year):
    out = []
    if path and os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                d = dt.date.fromisoformat(r["date"])
                if d.year == year - 1 and d.month == 12:
                    out.append((d, r["dept"]))
    return out


def check(plan):
    """Return (report rows, checks dict) for the verification sheet."""
    slots, a, staff, year = plan.slots, plan.a, plan.staff, plan.year
    cnt = {g: collections.Counter() for g in staff}
    for (d, c), g in zip(slots, a):
        cnt[g][c] += 1
    total = len(slots)
    rows = []
    for g in staff:
        s = sum(cnt[g].values())
        hol = cnt[g]["休日-昼"] + cnt[g]["休日-夜"]
        rows.append({"科": g, "休日-昼": cnt[g]["休日-昼"], "休日-夜": cnt[g]["休日-夜"], "平日-夜": cnt[g]["平日-夜"],
                     "合計": s, "人員": staff[g], "1人当たり": s / staff[g], "休日割合": hol / s,
                     "休日昼割合": cnt[g]["休日-昼"] / s, "休日夜割合": cnt[g]["休日-夜"] / s, "平日夜割合": cnt[g]["平日-夜"] / s})
    per = [r["1人当たり"] for r in rows]
    chk = {"枠の数": total, "最大/最小(1人当たり)": max(per) / min(per),
           "全体の休日割合": sum(1 for _, c in slots if c != "平日-夜") / total}
    # Same holiday day/night
    by_day = collections.defaultdict(dict)
    for (d, c), g in zip(slots, a):
        by_day[d][c] = g
    chk["同日昼夜が同じ科"] = [d for d, v in by_day.items() if v.get("休日-昼") and v.get("休日-昼") == v.get("休日-夜")]
    viol = []
    for name, s, e in plan.wins:
        k = collections.Counter(g for (d, _), g in zip(slots, a) if s <= d <= e)
        k.update(g for d, g in plan.prev if s <= d <= e)
        viol += [f"{name}: {g} {n}回" for g, n in k.items() if n > 1]
    chk["GW・年末年始の重複"] = viol
    # Spacing summary
    consec, short_h, short_w = [], 0, 0
    for g in staff:
        items = sorted(plan.dept_dates(g), key=lambda t: (t[0], t[1] != "休日-昼"))
        for (d1, c1), (d2, c2) in zip(items, items[1:]):
            if d2.year != year:
                continue  # both in the previous year: not this plan's doing
            gap = (d2 - d1).days
            if gap == 1:
                consec.append(f"{g} {d1:%m/%d}-{d2:%m/%d}")
            hol = (c1 or "").startswith("休日") and (c2 or "").startswith("休日")
            if hol and gap < 6:
                short_h += 1
            elif not hol and gap < 4:
                short_w += 1
    chk["連日"] = consec
    chk["休日どうし6日未満"] = short_h
    chk["平日を含む4日未満"] = short_w
    # Monthly
    month = {g: collections.Counter() for g in staff}
    for (d, _), g in zip(slots, a):
        month[g][d.month] += 1
    return rows, chk, month


def write_outputs(plan, out_dir, staff_path):
    os.makedirs(out_dir, exist_ok=True)
    y = plan.year
    csv_path = os.path.join(out_dir, f"toban_{y}.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(["date", "band", "dept"])
        for (d, c), g in zip(plan.slots, plan.a):
            w.writerow([d.isoformat(), c[-1], g])
    rows, chk, month = check(plan)

    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = openpyxl.Workbook()
    font = "Yu Gothic UI"
    bold = Font(name=font, bold=True)
    norm = Font(name=font)
    head_fill = PatternFill("solid", fgColor="D9E1F2")

    ws = wb.active
    ws.title = "科別集計"
    cols = ["科", "休日-昼", "休日-夜", "平日-夜", "合計", "人員", "1人当たり", "休日割合"]
    ws.append(cols)
    for r in sorted(rows, key=lambda r: -r["人員"]):
        ws.append([r[c] for c in cols])
    tot = [sum(r[c] for r in rows) for c in ("休日-昼", "休日-夜", "平日-夜", "合計", "人員")]
    ws.append(["合計", *tot, tot[3] / tot[4], (tot[0] + tot[1]) / tot[3]])
    ws.append([])
    ws.append(["最大/最小（1人当たり）", chk["最大/最小(1人当たり)"]])
    ws.append(["全体の休日割合", chk["全体の休日割合"]])
    ws.append(["人員の出典", os.path.basename(staff_path) + "（Dr.JOY の所属医師から初期研を除く）"])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = norm
            if isinstance(cell.value, float):
                cell.number_format = "0.0%" if cell.column in (8,) or (cell.column == 2 and cell.row > len(rows) + 3) else "0.000"
    ws.cell(len(rows) + 5, 2).number_format = "0.000"
    ws.cell(len(rows) + 6, 2).number_format = "0.0%"
    for c in ws[1]:
        c.font, c.fill = bold, head_fill
    ws.column_dimensions["A"].width = 26
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 11

    ws = wb.create_sheet("点検")
    items = [
        ("A 最大/最小（1人当たり）≦ 1.10", f"{chk['最大/最小(1人当たり)']:.3f}", chk["最大/最小(1人当たり)"] <= 1.10),
        ("E 同じ休日の昼と夜が同じ科", "なし" if not chk["同日昼夜が同じ科"] else ", ".join(map(str, chk["同日昼夜が同じ科"])), not chk["同日昼夜が同じ科"]),
        ("C・D GW・年末年始に同じ科が2回以上", "なし" if not chk["GW・年末年始の重複"] else " / ".join(chk["GW・年末年始の重複"]), not chk["GW・年末年始の重複"]),
        ("F 同じ科の連日", f"{len(chk['連日'])} 件" + ("（" + "、".join(chk["連日"]) + "）" if chk["連日"] else ""), not chk["連日"]),
        ("F 休日どうしの間隔が6日未満", f"{chk['休日どうし6日未満']} 件", None),
        ("F 平日を含む間隔が4日未満", f"{chk['平日を含む4日未満']} 件", None),
    ]
    ws.append(["項目", "結果", "判定"])
    for name, val, ok in items:
        ws.append([name, val, "" if ok is None else ("OK" if ok else "NG")])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = norm
    for c in ws[1]:
        c.font, c.fill = bold, head_fill
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 70

    ws = wb.create_sheet("月別")
    ws.append(["科", *[f"{m}月" for m in range(1, 13)], "合計"])
    tot_staff = sum(plan.staff.values())
    per_month = collections.Counter(d.month for d, _ in plan.slots)
    for g in sorted(plan.staff, key=lambda g: -plan.staff[g]):
        ws.append([g, *[month[g][m] for m in range(1, 13)], sum(month[g].values())])
    ws.append(["（枠の数）", *[per_month[m] for m in range(1, 13)], len(plan.slots)])
    ws.append([])
    ws.append(["期待値との差（実際 − 人員比）"])
    for g in sorted(plan.staff, key=lambda g: -plan.staff[g]):
        ws.append([g, *[round(month[g][m] - plan.staff[g] / tot_staff * per_month[m], 1) for m in range(1, 13)]])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = norm
    for c in ws[1]:
        c.font, c.fill = bold, head_fill
    ws.column_dimensions["A"].width = 26

    ws = wb.create_sheet("当番当直")
    wd = "月火水木金土日"
    ws.append(["日付", "曜日", "区分", "担当科"])
    for (d, c), g in zip(plan.slots, plan.a):
        ws.append([d, wd[d.weekday()], c, g])
    for row in ws.iter_rows(min_row=2):
        row[0].number_format = "yyyy/mm/dd"
        for cell in row:
            cell.font = norm
    for c in ws[1]:
        c.font, c.fill = bold, head_fill
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 24
    for ws_ in wb.worksheets:
        for row in ws_.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")
    xlsx_path = os.path.join(out_dir, f"toban_{y}_検証.xlsx")
    wb.save(xlsx_path)
    return csv_path, xlsx_path, rows, chk


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--staff")
    ap.add_argument("--prev", default=os.path.join(ROOT, "data", "toban.csv"))
    ap.add_argument("--out-dir", default=os.path.join(ROOT, "out"))
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--iters", type=int, default=400000)
    ap.add_argument("--restarts", type=int, default=4)
    args = ap.parse_args()

    staff_path = args.staff or latest_staff()
    staff = count_staff(staff_path)
    slots = build_slots(args.year, load_holidays())
    cat_total = collections.Counter(c for _, c in slots)
    q = quotas(staff, cat_total)
    prev = load_prev(args.prev, args.year)

    best = None
    for k in range(args.restarts):
        p = Plan(args.year, slots, staff, q, prev, args.seed + k)
        p.initial()
        cost = p.anneal(args.iters)
        print(f"restart {k}: cost {cost:.0f}")
        if best is None or cost < best[0]:
            best = (cost, p)
    plan = best[1]
    csv_path, xlsx_path, rows, chk = write_outputs(plan, args.out_dir, staff_path)

    print(f"人員: {os.path.basename(staff_path)}  枠: {len(slots)}（休日-昼 {cat_total['休日-昼']} / 休日-夜 {cat_total['休日-夜']} / 平日-夜 {cat_total['平日-夜']}）")
    print(f"{'科':<14}{'休昼':>4}{'休夜':>4}{'平夜':>4}{'合計':>4}{'人員':>4}{'1人当':>7}{'休日%':>7}")
    for r in sorted(rows, key=lambda r: -r["人員"]):
        print(f"{r['科']:<14}{r['休日-昼']:>4}{r['休日-夜']:>4}{r['平日-夜']:>4}{r['合計']:>4}{r['人員']:>4}{r['1人当たり']:>7.3f}{r['休日割合']:>7.1%}")
    for k, v in chk.items():
        print(k, ":", v)
    print(csv_path)
    print(xlsx_path)


if __name__ == "__main__":
    main()
