# -*- coding: utf-8 -*-
"""dtx14_check_borders.py — 書き出した DTX14 の罫線・A列結合・行の高さを雛形と突き合わせる（開発用）。20260922

usage: python -X utf8 tools\dtx14_check_borders.py <雛形.xlsm> <出力.xlsm> [日数]
畳んだ／展開したブロックごとに、見えている下端の線・昼夜の間の線・A列の結合・行の高さが雛形と合うかを COM で読む。
"""
#   For every dept block (rows 5..end with a label in C): the visible bottom edge of the block must have the
#   same style as the template's block bottom; inside an expanded 2-row block the 昼/夜 separator must exist.
import sys, os
import win32com.client

tpl, out = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
XL_NONE = -4142
LAST = 4 + (int(sys.argv[3]) if len(sys.argv) > 3 else 30) - 1


def style(b):
    ls = b.LineStyle
    return None if ls in (None, XL_NONE) else (ls, b.Weight)


def blocks(ws, last_row):
    """[(a_row, rows[])] dept blocks by A column (merged or value)."""
    out, cur = [], None
    for r in range(5, last_row + 1):
        c = str(ws.Cells(r, 3).Value or "").strip()
        if not c:
            cur = None
            continue
        a = str(ws.Cells(r, 1).Value or "").strip()
        top = ws.Cells(r, 1).MergeArea.Row == r
        if cur is None or (a and top and not ws.Rows(r).Hidden):
            cur = [r, []]
            out.append(cur)
        cur[1].append(r)
    return out


xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False; xl.EnableEvents = False; xl.AutomationSecurity = 3
try:
    wt = xl.Workbooks.Open(tpl, UpdateLinks=0, ReadOnly=True)
    wo = xl.Workbooks.Open(out, UpdateLinks=0, ReadOnly=True)
    bad = 0
    for sn, last_row in (("当直表 ファースト", 64), ("当直表 セカンド以下", 62)):
        t, o = wt.Worksheets(sn), wo.Worksheets(sn)
        tb, ob = blocks(t, last_row), blocks(o, last_row)
        print(f"== {sn}: template blocks {len(tb)} / output blocks {len(ob)}")
        for (ta, trows), (oa, orows) in zip(tb, ob):
            name = str(t.Cells(ta, 1).Value or "").strip()
            if str(o.Cells(oa, 1).Value or "").strip() != name:
                print(f"  ✗ block mismatch {name} vs {o.Cells(oa, 1).Value}"); bad += 1; continue
            vis = [r for r in orows if not o.Rows(r).Hidden]
            if not vis:
                print(f"  ✗ {name}: no visible row"); bad += 1; continue
            # bottom edge of block: template = last row bottom (or next row top); output = last visible row bottom (or next visible row top)
            def edge_bottom(ws, r, nxt):
                sty = []
                for c in range(1, LAST + 1):
                    s = style(ws.Cells(r, c).Borders(9)) or (style(ws.Cells(nxt, c).Borders(8)) if nxt else None)
                    sty.append(s)
                return sty
            tnext = trows[-1] + 1
            onext = vis[-1] + 1
            while o.Rows(onext).Hidden and onext < last_row + 3:
                onext += 1
            te, oe = edge_bottom(t, trows[-1], tnext), edge_bottom(o, vis[-1], onext)
            diff = [c + 1 for c in range(LAST) if te[c] != oe[c]]
            if diff:
                print(f"  ✗ {name}: bottom edge differs at cols {diff[:6]} template={te[3]} output={oe[3]}"); bad += 1
            # inner separator in expanded 2-row blocks
            if len(vis) == 2:
                inner = [style(o.Cells(vis[0], c).Borders(9)) or style(o.Cells(vis[1], c).Borders(8)) for c in range(2, LAST + 1)]
                tin = [style(t.Cells(trows[0], c).Borders(9)) or style(t.Cells(trows[1], c).Borders(8)) for c in range(2, LAST + 1)]
                if inner != tin:
                    print(f"  ✗ {name}: 昼/夜 separator differs template={tin[2]} output={inner[2]}"); bad += 1
            # A-column: visible rows should not be merged beyond visible rows
            ma = o.Cells(vis[0], 1).MergeArea
            if ma.Rows.Count > 1 and len(vis) == 1:
                print(f"  ✗ {name}: A still merged while collapsed"); bad += 1
            # row height check: text lines must fit
            for r in vis:
                vals = o.Range(o.Cells(r, 2), o.Cells(r, LAST)).Value
                lines = max([v.count("\n") + 1 for v in vals[0] if isinstance(v, str)] or [1])
                need = lines * 11 * 1.3
                if o.Rows(r).RowHeight < need:
                    print(f"  ✗ {name} row {r}: height {o.Rows(r).RowHeight:.1f} < needed {need:.1f} ({lines} lines)"); bad += 1
    print("問題", bad, "件")
    wt.Close(False); wo.Close(False)
finally:
    xl.Quit()
