# -*- coding: utf-8 -*-
r"""pdf_links.py — 公開用の当直勤務表 PDF の氏名に、Teams のチャットへのリンクを重ねる。20261007

管理者 20261007「名前をクリックしたら Teams のチャットが開くようにしたい。書式はまったく変えずに」。
PDF の見た目は変えない。氏名の文字の上に、枠線の無いリンクの注釈（/Link・/URI）を重ねるだけ。
Excel（DTX14）には付けない（1マスに1リンクしか付かず、何人も入るマスで分けられない・字が青い下線に変わる・
クリックで Teams が開いて事務担当が直しにくい）。

どの氏名が誰か: dtx14_extract.extract（DTX14 の当直表と人別シートの E列の氏名 → Dr.JOY の名簿でログインID）。
  手作業の月（10月まで）もアプリが書く月（11月から）も同じ道で引ける。引けない人にはリンクを付けない。
PDF のどこがどのマスか: Excel のマスの位置（COM の Left/Top/Width/Height。ポイント）と PDF の文字の位置を、
  その頁に1回だけ出てくる氏名を目印に1次式で合わせる（印刷の縮小・余白・中央寄せはすべて1次式に入る）。
  合わせきれない頁（目印が少ない・ずれが大きい）にはリンクを付けない。PDF の公開は止めない。

usage（試し）: python -X utf8 tools\pdf_links.py <DTX14.xlsm> <出力.pdf> <年> <月>
"""
import calendar
import os
import re
import statistics
import sys
import urllib.parse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

TEAMS_CHAT = "https://teams.microsoft.com/l/chat/0/0?users="
COL_DAY1 = 4                     # D列 = 1日（dtx14_fill と同じ）
MIN_ANCHORS = 8                  # 1頁の目印の最低数
MAX_RESID = 2.5                  # 目印のずれ（中央値・pt）の許容
LOGIN_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+$")


def teams_url(login):
    return TEAMS_CHAT + urllib.parse.quote(login.strip().lower(), safe="@.")


def core_name(text):
    """マスの1行 → 氏名の部分（(姓)・昼:姓・夜:姓 の飾りを外す）。"""
    t = str(text or "").strip().replace("（", "(").replace("）", ")").replace("：", ":")
    if t.startswith("(") and t.endswith(")"):
        t = t[1:-1].strip()
    for p in ("昼:", "夜:"):
        if t.startswith(p):
            t = t[len(p):].strip()
    return t


def cell_logins(records, sheet_first, sheet_second):
    """extract の records → {(シート名, 行, 列): {氏名: ログインID}}。同じマスで同じ氏名が別の人に引けたら外す。"""
    out, bad = {}, set()
    for r in records:
        login = str(r.get("staff_login") or "").strip()
        if not LOGIN_RE.match(login):
            continue
        sheet = sheet_first if r.get("role") == "first" else sheet_second
        key = (sheet, int(r["src_row"]), COL_DAY1 + int(str(r["date"])[8:10]) - 1)
        nm = core_name(r.get("name_disp"))
        if not nm:
            continue
        cur = out.setdefault(key, {})
        if nm in cur and cur[nm] != login:
            bad.add((key, nm))
        cur[nm] = login
    for key, nm in bad:
        out[key].pop(nm, None)
    return out


def sheet_geometry(ws, ndays):
    """COM のシート → {cols: {列: (left, width)}, rows: {行: (top, height)}, text: {(行, 列): 文字}}（日付の列だけ）。"""
    ur = ws.UsedRange
    last_row = ur.Row + ur.Rows.Count - 1
    c1, c2 = COL_DAY1, COL_DAY1 + ndays - 1
    vals = ws.Range(ws.Cells(1, c1), ws.Cells(last_row, c2)).Value or ()
    text = {}
    for i, row in enumerate(vals, start=1):
        for j, v in enumerate(row, start=c1):
            if isinstance(v, str) and v.strip():
                text[(i, j)] = v
    cols = {c: (float(ws.Columns(c).Left), float(ws.Columns(c).Width)) for c in range(c1, c2 + 1)}
    rows = {}
    for r in range(1, last_row + 1):
        h = float(ws.Rows(r).Height)
        if h > 0:
            rows[r] = (float(ws.Rows(r).Top), h)
    return {"cols": cols, "rows": rows, "text": text}


def _fit(pairs):
    """[(excel, pdf)] → (傾き, 切片)。最小二乗。"""
    n = len(pairs)
    mx = sum(a for a, _ in pairs) / n
    my = sum(b for _, b in pairs) / n
    sxx = sum((a - mx) ** 2 for a, _ in pairs)
    if sxx <= 0:
        return None
    k = sum((a - mx) * (b - my) for a, b in pairs) / sxx
    return k, my - k * mx


def _spans(page):
    """頁の文字のかたまり [(文字, (x0, y0, x1, y1))]。
    ⚠ PyMuPDF は同じ高さに並ぶ隣のマスの文字を1つのかたまりにまとめることがある（合成の PDF で「癸田甲0509」）。
    文字の間に字の大きさの3割より広い隙間があれば、そこで分ける。"""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for ln in b.get("lines", []):
            for s in ln["spans"]:
                gap = 0.3 * float(s.get("size") or 5)
                cur = []
                for ch in s["chars"]:
                    if cur and ch["bbox"][0] - cur[-1]["bbox"][2] > gap:
                        out.append(cur)
                        cur = []
                    cur.append(ch)
                if cur:
                    out.append(cur)
    res = []
    for chars in out:
        while chars and not chars[0]["c"].strip():
            chars = chars[1:]
        while chars and not chars[-1]["c"].strip():
            chars = chars[:-1]
        if chars:
            res.append(("".join(c["c"] for c in chars),
                        (min(c["bbox"][0] for c in chars), min(c["bbox"][1] for c in chars),
                         max(c["bbox"][2] for c in chars), max(c["bbox"][3] for c in chars))))
    return res


def _nz(t):
    return str(t).strip().replace("（", "(").replace("）", ")").replace("：", ":")


def _resid(pairs, f):
    return [abs(b - (f[0] * a + f[1])) for a, b in pairs]


def calibrate(geo, spans):
    """Excel の点 → PDF の点の1次式 (kx, bx, ky, by)。合わなければ None と理由。
    ⚠ 横と縦で縮み方が違う（列の幅は印刷では文字の幅で決まり、画面のポイントと同じ比にならない。10月で 0.289 と 0.268）。
    ⚠ セカンド以下は同じ先生が隣の日にも並ぶので、隣のマスの同じ名前に当てると1列ずれたまま「合って」しまう（10月で実測）。
    1) シート全体で1回だけ出る行（マスの中の1行）の文字で大まかに合わせる（縦はマスの中央で近似）。
    2) その式で全部の文字をマスに当て、そのマスにあり左右のマスには無い名前だけを目印に合わせ直す
       （横はマスの中央、縦はマスに入っている文字の上下の中央）。"""
    lines_x = {}
    for (r, c), t in geo["text"].items():
        for s in t.split("\n"):
            if _nz(s):
                lines_x.setdefault(_nz(s), []).append((r, c))
    count_p = {}
    for t, _bb in spans:
        count_p[_nz(t)] = count_p.get(_nz(t), 0) + 1
    px, py = [], []
    for t, (x0, y0, x1, y1) in spans:
        cells = lines_x.get(_nz(t)) or []
        if count_p[_nz(t)] != 1 or len(cells) != 1 or cells[0][0] not in geo["rows"]:
            continue
        r, c = cells[0]
        left, w = geo["cols"][c]
        top, h = geo["rows"][r]
        px.append((left + w / 2, (x0 + x1) / 2))
        py.append((top + h / 2, (y0 + y1) / 2))
    if len(px) < 4:
        return None, f"目印が {len(px)} 個しかない"
    fx, fy = _fit(px), _fit(py)
    for _ in range(3):
        if not fx or not fy or fx[0] <= 0 or fy[0] <= 0:
            return None, "合わせられない"
        px, py, ys = [], [], {}
        for t, (x0, y0, x1, y1) in spans:
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            row, col = _locate(geo, (cx - fx[1]) / fx[0], (cy - fy[1]) / fy[0])
            cell = geo["text"].get((row, col))
            nt = _nz(t)
            if cell is None or nt not in [_nz(s) for s in cell.split("\n")]:
                continue
            if any(nt in [_nz(s) for s in geo["text"].get((row, col + d), "").split("\n")] for d in (-1, 1)):
                continue                    # 隣の日にも同じ名前 → 目印にしない
            left, w = geo["cols"][col]
            px.append((left + w / 2, cx))
            ys.setdefault((row, col), []).append((y0, y1))
        for (row, col), bbs in ys.items():
            top, h = geo["rows"][row]
            py.append((top + h / 2, (min(a for a, _ in bbs) + max(b for _, b in bbs)) / 2))
        if len(px) < MIN_ANCHORS or len(py) < 4:
            return None, f"合う目印が {len(px)} 個しかない"
        fx, fy = _fit(px), _fit(py)
        if not fx or not fy:
            return None, "合わせられない"
        mx, my = statistics.median(_resid(px, fx)), statistics.median(_resid(py, fy))
        px = [p for p, e in zip(px, _resid(px, fx)) if e <= 3 * mx + 1]
        py = [p for p, e in zip(py, _resid(py, fy)) if e <= 3 * my + 1]
        fx, fy = _fit(px), _fit(py)
    if not fx or not fy or fx[0] <= 0 or fy[0] <= 0:
        return None, "合わせられない"
    rx, ry = statistics.median(_resid(px, fx)), statistics.median(_resid(py, fy))
    if rx > MAX_RESID or ry > MAX_RESID:
        return None, f"ずれが大きい（横 {rx:.1f}pt・縦 {ry:.1f}pt）"
    return (fx[0], fx[1], fy[0], fy[1]), f"目印 {len(px)} 個・ずれ 横 {rx:.1f}pt 縦 {ry:.1f}pt"

def _locate(geo, x, y):
    col = next((c for c, (l, w) in geo["cols"].items() if l <= x < l + w), None)
    row = next((r for r, (t, h) in geo["rows"].items() if t <= y < t + h), None)
    return row, col


def page_links(page, sheet, geo, logins):
    """1頁ぶんのリンク [(fitz.Rect, url)] と (付けた数, 氏名らしい文字の数, 説明)。"""
    import fitz
    spans = _spans(page)
    tr, why = calibrate(geo, spans)
    if tr is None:
        return [], 0, 0, why
    kx, bx, ky, by = tr
    links, names = [], 0
    for t, bb in spans:
        x0, y0, x1, y1 = bb
        row, col = _locate(geo, ((x0 + x1) / 2 - bx) / kx, ((y0 + y1) / 2 - by) / ky)
        if row is None or col is None or (row, col) not in geo["text"]:
            continue
        nm = core_name(t)
        if not nm or nm in ("(", ")"):
            continue                        # 括弧だけの切れ端（色の違う氏名は PDF で「(」「姓」「)」に分かれる）
        names += 1
        login = logins["cells"].get((sheet, row, col), {}).get(nm)
        if not login and nm in {core_name(s) for s in geo["text"][(row, col)].split("\n")}:
            # 人別シートで引けないときの科の名簿からの引き当ては、そのマスに実際に書いてある名前だけ（20261009 レビュー:
            #   位置合わせが1行ずれると、隣の行の科の同姓の先生に付き得た）
            login = by_dept(logins["rows"].get((sheet, row)), nm)
        if login:
            links.append((fitz.Rect(bb), teams_url(login)))
    return links, len(links), names, why


def _pdf_str(s):
    return "(" + s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)") + ")"


def add_links(doc, page, links):
    """リンクの注釈をまとめて足す（insert_link を1つずつ呼ぶと 3,000 個で30秒以上かかるため）。枠線なし。"""
    if not links:
        return
    m = page.transformation_matrix
    refs = []
    for rect, url in links:
        r = rect * m
        x0, x1 = sorted((r.x0, r.x1))
        y0, y1 = sorted((r.y0, r.y1))
        xref = doc.get_new_xref()
        doc.update_object(xref, f"<</Type/Annot/Subtype/Link/Rect[{x0:.2f} {y0:.2f} {x1:.2f} {y1:.2f}]"
                                f"/Border[0 0 0]/F 4/A<</S/URI/URI{_pdf_str(url)}>>>>")
        refs.append(f"{xref} 0 R")
    kind, val = doc.xref_get_key(page.xref, "Annots")
    old = ""
    if kind == "array":
        old = val.strip()[1:-1]
    elif kind == "xref":
        old = doc.xref_object(int(val.split()[0])).strip()[1:-1]
    doc.xref_set_key(page.xref, "Annots", "[" + (old + " " if old else "") + " ".join(refs) + "]")


def link_pdf(pdf_in, pdf_out, sheets, geos, logins, log=print):
    """pdf_in の頁 i を sheets[i] とみなしてリンクを足し、pdf_out に保存する。戻り: 付けた数（頁の数が合わなければ None）。"""
    import fitz
    doc = fitz.open(pdf_in)
    try:
        if len(doc) != len(sheets):
            log(f"  リンク: PDF が {len(doc)} 頁（{len(sheets)} 頁のはず）→ 付けない")
            return None
        total = 0
        for i, sheet in enumerate(sheets):
            page = doc[i]
            links, n, names, why = page_links(page, sheet, geos[sheet], logins)
            add_links(doc, page, links)
            total += n
            log(f"  リンク: {sheet} {n}/{names}（{why}）")
        doc.save(pdf_out, garbage=3, deflate=True)
        return total
    finally:
        doc.close()


def collect_geometry(wb, sheets, year, month):
    ndays = calendar.monthrange(year, month)[1]
    return {sn: sheet_geometry(wb.Worksheets(sn), ndays) for sn in sheets}


def by_dept(cands, nm):
    """人別シートで引けなかった氏名を、その行の科の医師から引く（姓が1人にだけ合うときだけ。当て推量はしない）。
    nm は姓か、同姓の区別に名の頭を足したもの（例「春野健」）。"""
    if not cands:
        return None
    import contract as C
    k = C.canon(C.namekey(nm))
    hits = {p["login"] for p in cands
            if k.startswith(C.canon(p["sei"])) and C.canon(C.namekey(p["name"])).startswith(k)}
    return hits.pop() if len(hits) == 1 else None


def collect_logins(xlsm, staffs, year, month):
    """{"cells": {(シート, 行, 列): {氏名: ログインID}}（人別シートから）, "rows": {(シート, 行): [その科の医師]}}"""
    import json
    import dtx14_extract as X
    import contract as C
    recs, _unres = X.extract(xlsm, staffs, year, month)
    people, _e, _c = C.load_staffs(staffs)
    doctors = [dict(p, sei=str(p["name"]).split(" ")[0]) for p in people
               if p.get("is_doctor") and LOGIN_RE.match(str(p.get("login") or ""))]
    with open(os.path.join(C.DATA, "dept_map.json"), encoding="utf-8-sig") as f:
        dmap = json.load(f)["map"]
    frames = C.load_frames()
    rows = {}
    for sheet in (X.SHEET_FIRST, X.SHEET_SECOND):
        grid, merged_a = X.load_sheet_grid(xlsm, sheet)
        for r in range(5, len(grid) + 1):
            label = X.cellv(grid, merged_a.get(r, r), 1)
            if not label:
                continue
            label = C.namekey(label)
            owners = {f.get("owner") or f["dept"] for f in frames
                      if C.namekey(f.get("dept_dtx14") or f["dept"]) == label}
            ents = [dmap[o] for o in owners if o in dmap]
            if ents:
                rows[(sheet, r)] = [p for p in doctors if any(C.matches_dept_map(e, p) for e in ents)]
    return {"cells": cell_logins(recs, X.SHEET_FIRST, X.SHEET_SECOND), "rows": rows}


def main():
    """試し: DTX14 を Excel で PDF にし、リンクを足して出力する（公開用フォルダには置かない）。"""
    import shutil
    import tempfile
    import contract as C
    import win32com.client
    src, out, year, month = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4])
    sheets = ("当直表 ファースト", "当直表 セカンド以下")
    paper = {"当直表 ファースト": 9, "当直表 セカンド以下": 8}
    work_dir = tempfile.mkdtemp(prefix="pdflinks_")
    work = os.path.join(work_dir, "_work.xlsm")
    shutil.copy2(src, work)
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.AutomationSecurity = 3
        wb = xl.Workbooks.Open(work, UpdateLinks=0, ReadOnly=True)
        for sn in sheets:
            ps = wb.Worksheets(sn).PageSetup
            ps.PaperSize = paper[sn]
            ps.Orientation = 2
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
        wb.Worksheets(sheets).Select()
        plain = os.path.join(work_dir, "_out.pdf")
        xl.ActiveSheet.ExportAsFixedFormat(0, plain)
        geos = collect_geometry(wb, sheets, year, month)
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    logins = collect_logins(work, C.newest_staffs(), year, month)
    n = link_pdf(plain, out, sheets, geos, logins)
    print("links:", n, os.path.getsize(plain), "->", os.path.getsize(out))
    shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
