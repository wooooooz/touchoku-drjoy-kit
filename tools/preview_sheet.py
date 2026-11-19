# -*- coding: utf-8 -*-
r"""preview_sheet.py  20260912
当直表xlsx を「紙に刷ったとおり」の画像にする開発用ツール。意匠を目で見て直すために使う。

  python -X utf8 tools/preview_sheet.py <book.xlsx> [--out <png>] [--dpi 120] [--sheet 当直表]

Excel COM で PDF に出し（= 印刷と同じ縮小がかかった状態）、PyMuPDF で PNG にする。
⚠ Range.CopyPicture 経由の画像書き出しは不安定で、黙って白紙を返すことがある。必ずこちらを使う。
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys

PS = r"""
$ErrorActionPreference='Stop'
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowTitle -eq ''} | Out-Null
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false
try {
  $wb = $xl.Workbooks.Open('%(xlsx)s')
  $ws = $wb.Worksheets.Item('%(sheet)s')
  $ws.Activate()
  Write-Output ("PAGES=" + $ws.PageSetup.Pages.Count)
  $ws.ExportAsFixedFormat(0, '%(pdf)s')
  $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
"""


def to_pdf(xlsx, pdf, sheet="当直表"):
    if os.path.exists(pdf):
        os.remove(pdf)
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                        PS % {"xlsx": os.path.abspath(xlsx), "pdf": os.path.abspath(pdf), "sheet": sheet}],
                       capture_output=True, text=True, encoding="utf-8")
    if not os.path.exists(pdf):
        raise SystemExit(f"PDFに出せませんでした:\n{r.stdout}\n{r.stderr}")
    pages = next((ln.split("=")[1].strip() for ln in (r.stdout or "").splitlines()
                  if ln.startswith("PAGES=")), "?")
    return pages


def to_png(pdf, png, dpi=120):
    import fitz
    d = fitz.open(pdf)
    outs = []
    for i, page in enumerate(d):
        p = png if d.page_count == 1 else f"{os.path.splitext(png)[0]}_p{i + 1}.png"
        page.get_pixmap(dpi=dpi).save(p)
        outs.append((p, round(page.rect.width), round(page.rect.height)))
    d.close()
    return outs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--out", default=None)
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--sheet", default="当直表")
    a = ap.parse_args()
    base = os.path.splitext(a.xlsx)[0]
    pdf = base + ".pdf"
    png = a.out or (base + ".png")
    pages = to_pdf(a.xlsx, pdf, a.sheet)
    outs = to_png(pdf, png, a.dpi)
    print(f"ページ数={pages}")
    for p, w, h in outs:
        print(f"  {p}  ({w}x{h}pt {'横' if w > h else '縦'})")


if __name__ == "__main__":
    sys.exit(main())
