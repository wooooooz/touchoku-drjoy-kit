# -*- coding: utf-8 -*-
r"""pack_app.py — 展開済みのアプリ（app\release\win-unpacked）を dist\touchoku_app.zip にまとめ、版の名前を dist\touchoku_app.ver に書く。20261002

起動係（app\launcher\TouchokuLauncher.cs）は、月フォルダの touchoku_app.ver を見て、各 PC の
%LOCALAPPDATA%\Touchoku\app\<版>\ に zip を1回だけ展開し、2回目からはそれを起動する（管理者 20260925「起動速くなる？」）。
版の名前 = 中身の当直勤務表入力.exe と resources\app.asar の SHA-256 の頭12桁（中身が同じなら同じ版＝展開し直さない）。

usage: python -X utf8 tools\pack_app.py
"""
import hashlib
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "app", "release", "win-unpacked")
DIST = os.path.join(ROOT, "dist")
ZIP = os.path.join(DIST, "touchoku_app.zip")
VER = os.path.join(DIST, "touchoku_app.ver")


def version():
    h = hashlib.sha256()
    for rel in ("当直勤務表入力.exe", os.path.join("resources", "app.asar")):
        with open(os.path.join(SRC, rel), "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
    return h.hexdigest()[:12]


def main():
    if not os.path.isfile(os.path.join(SRC, "当直勤務表入力.exe")):
        raise SystemExit(f"展開済みのアプリがありません（先に npx electron-builder --win）: {SRC}")
    os.makedirs(DIST, exist_ok=True)
    ver = version()
    tmp = os.path.join(DIST, f"~touchoku_app_zip_{os.getpid()}.tmp")   # 20261002: 二重拡張子にしない（EDR がランサムウェア疑いで通報する）
    n = 0
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for base, _dirs, files in os.walk(SRC):
            for fn in files:
                p = os.path.join(base, fn)
                z.write(p, os.path.relpath(p, SRC))       # 日本語の名前は UTF-8 の印つきで入る（.NET の ZipFile が読める）
                n += 1
    os.replace(tmp, ZIP)
    with open(VER, "w", encoding="ascii", newline="\r\n") as f:
        f.write(ver + "\n")
    print(f"packed {n} files -> {ZIP} ({os.path.getsize(ZIP):,} bytes)  ver={ver}")


if __name__ == "__main__":
    sys.exit(main())
