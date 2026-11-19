# -*- coding: utf-8 -*-
"""frame_defaults.py — 枠ごとの既定行数（当直の人数・オンコールの人数）を現物DTX14から求めて
data/frames.json に書き込む。20260902

ユーザー指示（2026-09-02）: アプリの各マスの初期行数は「今までの当直勤務表に倣う」。
日によって当直数・オンコール数に揺れがあれば多い方（最大値）を採用する。
（例: 救急科のセカンド以下は 昼=当直2+オンコール3 / 夜=当直3+オンコール3。
  空欄の行は「その人数は居ない」の意味なので多めに出しても害はない）

  duty_slots   … 既定で「当」になる行数（= その枠で同じ日・同じバンドに立った当直の最大人数）
  oncall_slots … 既定で「オ」になる行数（= 同じくオンコールの最大人数）
ファーストは常に1行（追加なし）なので書かず、oncall_default（多数派）だけ検算して報告する。

usage: python -X utf8 tools/frame_defaults.py [--xlsm PATH] [--staffs PATH] [--year Y --month M]
"""
import argparse, collections, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dtx14_extract as X
import contract as C

DEFAULT_XLSM = r"C:\tools\touchoku\sample\DTX14_sample.xlsm"
FRAMES = os.path.join(C.ROOT, "data", "frames.json")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsm", default=DEFAULT_XLSM)
    ap.add_argument("--staffs", default=r"C:\Users\user003\Downloads\staffs (65).csv")
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--month", type=int, default=9)
    a = ap.parse_args()

    recs, _ = X.extract(a.xlsm, a.staffs, a.year, a.month)
    cnt = collections.Counter()
    for r in recs:
        cnt[(r["frame_id"], r["date"], r["band"], bool(r["oncall"]))] += 1
    mx = collections.defaultdict(lambda: [0, 0])      # frame_id -> [max duty, max oncall]
    days = collections.defaultdict(lambda: [0, 0])    # frame_id -> [days with duty, days with oncall]
    for (fid, d, b, oc), c in cnt.items():
        i = 1 if oc else 0
        mx[fid][i] = max(mx[fid][i], c)
        days[fid][i] += 1

    with open(FRAMES, encoding="utf-8") as f:
        doc = json.load(f)
    frames = doc["frames"] if isinstance(doc, dict) else doc
    changed = []
    for fr in frames:
        m = mx.get(fr["frame_id"])
        if fr["role"] == "first":
            if m:
                want = days[fr["frame_id"]][1] > days[fr["frame_id"]][0]
                if bool(fr.get("oncall_default")) != want:
                    print(f"  ⚠ {fr['frame_id']} {fr['dept']}: oncall_default={fr.get('oncall_default')} だが"
                          f" 実績の多数派は {'オンコール' if want else '当直'}（当直{days[fr['frame_id']][0]}日/"
                          f"オンコール{days[fr['frame_id']][1]}日）")
            continue
        if not m:
            continue   # 今月の実績が無い枠（管理者枠など）は触らない
        d_, o_ = m
        if d_ == 0 and o_ == 0:
            continue
        if fr.get("duty_slots") != d_ or fr.get("oncall_slots") != o_:
            changed.append((fr["frame_id"], fr["dept"], fr["shape"], fr.get("duty_slots"), fr.get("oncall_slots"), d_, o_))
        fr["duty_slots"] = d_
        fr["oncall_slots"] = o_

    with open(FRAMES, "w", encoding="utf-8", newline="\n") as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(f"frames.json 更新: {len(changed)} 枠")
    for fid, dept, shape, od, oo, d_, o_ in changed:
        print(f"  {fid:6} {dept:14} {shape:2}  当直{od}→{d_} オンコール{oo}→{o_}")


if __name__ == "__main__":
    main()
