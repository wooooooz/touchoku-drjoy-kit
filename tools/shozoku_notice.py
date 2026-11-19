# -*- coding: utf-8 -*-
r"""shozoku_notice.py — 所属の変更メール（アプリの名簿 ⇔ Dr.JOY の所属・週パターン）。20261031

管理者 2026-09-27 決定:
  アプリの左の名簿（提出単位ごと・月ごと）は、その科の Dr.JOY 所属医師の申告。
  名簿と Dr.JOY の所属が食い違えば、月末と毎週水曜に事務担当へメールで知らせる（cc 管理者・甲野・己野・事務担当2）。

対象の月 (y, m) ごとに、共有マスター R8年N月\master\当直予定表_*.json（読むだけ）と、あれば前の月のマスターを読む。
所属を決める提出単位（shozoku.defining_owners）で、名簿の出ている科ごとに:
  加わる   … 名簿にいるが、Dr.JOY の今日の所属（staff CSV）がその科に合わない人
               （所属が当直をしない部署なら「当直をしない部署の医師（所属はそのまま）」の見出しへ）
  外れる   … Dr.JOY の今日の所属がその科に合うが、名簿にいない人
               （所属が同じく合うほかの提出単位の名簿にいれば出さない）
  週パターン … 区分（フルタイム ⇔ パートタイム・院外研修専従）が変わった人。
               比べる相手 = この月について前に知らせた区分（STATE_PATH）→ 無ければ前の月のマスターの区分（どの科でも）。
               どちらも無ければ新しく申告された人で、フルタイム以外だけを出す。
               ⚠ 週パターンは Dr.JOY から読めないので、知らせた区分を覚えておく（同じ知らせを毎週繰り返さない。
                  名簿を元に戻したときは、逆向きの知らせが出る）。所属のほうは Dr.JOY を直すまで毎回出る。
  ⚠ 同日の訂正（管理者）: 所属履歴は正本ではない（初期研のローテーションを入れる道具）。初期研はここでは出さないので、
     比べるのは今日の所属だけ（所属履歴の未来の組＝消し残しの「他院」などは見ない。shozoku.History.dept_on）。
  ⚠ 20261028 管理者: 内科全部・外科のローテーションなどは、事務担当が department_history.csv で日付付きで
     入れている。そこで「今日より後に始まる組」（予定）は見る＝対象月の1日の所属で比べる（shozoku.History.dept_on）。
     月の途中に予定どおり移る人（その月のうちに名簿の科の所属に就く／名簿の科から離れる）は出さない（depts_during）。
初期研（職位「初」）と名簿外は、どれにも出さない。外す提出単位（集中治療室・内科・管理当直・内科C 初期研）は見ない。
知らせることが無ければメールしない。

いつ（dtx14_daily_sync の毎日15時の本番の回だけ。--month の回・kotai_record のすぐ載せる回では呼ばない）:
  月末の日と毎週水曜 → 来月だけ（20260927 管理者: 月曜ではなく水曜。 月が始まってからの所属の変更は、その月は触らず次の月の下書きに活かす）。SYNC_FROM 以降で共有マスターがある月だけ（targets_for）。

usage: python -X utf8 tools\shozoku_notice.py --month 2026-11 [--dry]
         [--master <当直予定表.json>] [--prev-master <前の月の当直予定表.json>]
  本文を表示する。--dry なら送らない（覚えた区分も書かない）。付けなければ本当に送る。
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import re
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import contract as C          # noqa: E402
import shozoku as SZ          # noqa: E402

TO = "user004@example.ac.jp"                                        # 事務担当
CC = ["user003@example.ac.jp", "user007@example.ac.jp", "user008@example.ac.jp", "user005@example.ac.jp"]   # 管理者・甲野・己野・事務担当2
WEEK_APPLY = None      # week_pattern.apply の代わり（テストは差し替える）。None なら本物
SENDER = None          # 送る関数 (to, subject, body, cc) → None。None なら Outlook（dtx14_daily_sync.outlook_send）。テストは差し替える
STATE_PATH = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Touchoku",
                          "shozoku_notice_state.json")         # 知らせた区分（テストは差し替える）
KEEP_MONTHS = 6                                                  # 覚えた区分を残す月数
HEAD = "所属が変わったなら「Dr.JOY所属」を、直し忘れなら診療科に「貴科リスト」の整備を頼んでください。"   # 20261031 短く・決めつけない
NON_DUTY_HEAD = "当直をしない部署の医師（「Dr.JOY所属」はそのまま）"
INGAI_HEAD = "初期研以外で他院にいる医師（「Dr.JOY所属」を各科に戻してください）"


def reiwa_label(y, m):
    return f"R{y - 2018}年{m}月"


def prev_month(y, m):
    return (y - 1, 12) if m == 1 else (y, m - 1)


def next_month(y, m):
    return (y + 1, 1) if m == 12 else (y, m + 1)


def targets_for(today, sync_from=(2026, 11), exists=None):
    """その日に知らせる月: 月末の日と毎週水曜 → 来月だけ（両方の日でも1通）。
    今月は比べない（20260927 管理者: 月の途中で Dr.JOY の所属を正しく直しても、今月の名簿は月初めの申告のままなので、
    比べると直した向きを戻すような知らせになる。月の途中の変更は、次の月の名簿の初期の顔ぶれに自然に入る）。
    sync_from より前の月・exists(y, m) が False の月は除く。"""
    out = []
    this = (today.year, today.month)
    if (today + dt.timedelta(days=1)).month != today.month or today.weekday() == 2:
        out.append(next_month(*this))
    out = sorted(ym for ym in out if ym >= tuple(sync_from))
    if exists is not None:
        out = [ym for ym in out if exists(*ym)]
    return out


def master_path(y, m):
    import master_store as MS
    return MS.master_path(C.prod_month_dir(y, m), y, m)


def read_master(path, y, m):
    """共有マスターを読む（読むだけ）。無ければ None。"""
    import master_store as MS
    if not path or not os.path.isfile(path):
        return None
    return MS.load(path, y, m)


def _key(login):
    return str(login or "").strip().lower()


def _nk(name):
    """氏名の照合キー（空白を除き NFC・異体字を寄せる。contract.resolve_name と同じ規則）。"""
    return C.canon(C.nfc(C.namekey(name or "")))


def _cls(work):
    return "full" if work == "full" else "nonfull"


def rosters_of(master, owners):
    """{owner: [名簿の人]}（名簿が空でない提出単位だけ）。"""
    out = {}
    for o in owners:
        ros = ((master.get("depts") or {}).get(o) or {}).get("roster") or []
        if ros:
            out[o] = ros
    return out


def prev_works(prev_master, owners):
    """前の月の区分 {login: work}。所属を決める提出単位の名簿を先に、無ければほかの提出単位でも。"""
    out = {}
    if not prev_master:
        return out
    depts = prev_master.get("depts") or {}
    order = [o for o in owners if o in depts] + sorted(o for o in depts if o not in owners)
    for o in order:
        for mem in (depts.get(o) or {}).get("roster") or []:
            k = _key(mem.get("login"))
            if k and k not in out:
                out[k] = C.work_of(mem)
    return out


def month_items(y, m, master, prev_master, hist, frames=None, dept_map=None, told=None, recorded=None):
    """1か月分の知らせ。戻り: {"owners": {owner: [行]}, "non_duty": [行], "told": {key: 区分}}
    told = これまでに知らせた区分 {"YYYY-MM|login": "full"/"nonfull"}（読むだけ。新しく知らせる分を戻りの told に入れる）。"""
    frames = frames if frames is not None else C.load_frames()
    dept_map = dept_map if dept_map is not None else SZ.load_dept_map()
    told = told or {}
    d1 = dt.date(y, m, 1)
    dend = dt.date(*next_month(y, m), 1) - dt.timedelta(days=1)
    owners = SZ.defining_owners(frames, dept_map)
    pats = {o: SZ.owner_patterns(o, dept_map) for o in owners}
    rosters = rosters_of(master, owners)
    listed = {o: {_key(x.get("login")) for x in ros} for o, ros in rosters.items()}
    # 集合当直で庶務が選んだ「Dr.JOY所属」（行の seg。20261031 コードレビュー③）: {owner: {login: seg}}
    chosen = {o: {_key(x.get("login")): s for x in ros if (s := SZ.chosen_seg(x, pats[o]))} for o, ros in rosters.items()}
    # 名簿外（unreg:）で入れた医師が、あとから Dr.JOY に登録されると、ログインIDが違うので「名簿にいない」に見える。
    # 氏名（異体字・空白を寄せて）が同じなら、名簿にいるものとして扱う（20261028）
    listed_names = {o: {_nk(x.get("name")) for x in ros if SZ.is_unregistered(x)} for o, ros in rosters.items()}
    by_owner = {o: [] for o in rosters}
    non_duty, new_told = [], {}

    def title_of(mem):
        p = hist.people.get(_key(mem.get("login")))
        return (p or {}).get("title") or mem.get("title") or ""

    # 加わる（名簿にいるが所属が合わない）
    for o, ros in rosters.items():
        either = f"（{'／'.join(pats[o])} のどちらか）" if len(pats[o]) >= 2 else ""
        for mem in ros:
            if SZ.is_excluded(mem, title_of(mem)):
                continue
            # 集合当直で所属を選んだ医師は、その所属に合うかで見る（20261031: 消化器 → 感染症・呼吸器 への移りも知らせる）
            seg = SZ.chosen_seg(mem, pats[o])
            want = [seg] if seg else pats[o]
            tag = f"{DEPT_OPEN}{seg}{DEPT_CLOSE}" if seg else ""
            dep = hist.dept_on(mem.get("login"), d1, title_of(mem))
            if SZ.matches(dep, want):
                continue
            if any(SZ.matches(x, want) for x in hist.depts_during(mem.get("login"), d1, dend, title_of(mem))):
                continue                          # 月の途中に予定どおりこの科へ移る（日付付きで入っている）
            name = mem.get("name") or mem.get("login") or ""
            if not dep:
                by_owner[o].append(f"加わる：{name}（今の所属：Dr.JOY の名簿に見つからない）{'' if seg else either}{tag}")
            elif SZ.is_non_duty(dep):
                non_duty.append(f"{o}：{name}（今の所属：{SZ.dept_display(dep)}）")
            else:
                by_owner[o].append(f"加わる：{name}（今の所属：{SZ.dept_display(dep)}）{'' if seg else either}{tag}")

    # 外れる（所属が合うが名簿にいない）
    for k, p in sorted(hist.people.items(), key=lambda kv: (kv[1].get("kana") or "", kv[0])):
        if not p.get("is_doctor", True) or SZ.is_excluded(p, p.get("title")):
            continue
        # 20261028 管理者: 月のうち1日でも所属していれば、その科の名簿に載せる。所属ごとに、合う提出単位のどれかの名簿にいるかを見る
        hit = []
        for dep in hist.depts_during(k, d1, dend):
            hd = [o for o in owners if SZ.matches(dep, pats[o])]
            # 所属が合う提出単位のどれかの名簿にいる（集合当直で別の所属を選ばれていれば、この所属からは外れる。20261031）
            if any(k in listed.get(o, ()) and (k not in chosen.get(o, {}) or SZ.matches(dep, [chosen[o][k]])) for o in hd):
                continue
            if any(_nk(p.get("name")) in listed_names.get(o, ()) for o in hd):
                continue                          # 名簿外で入れた医師が Dr.JOY に登録された（氏名で同じ人）
            for o in hd:
                if o not in [h for h, _s in hit]:
                    # その医師の Dr.JOY所属（提出単位に当たる区切り）。集合当直でも診療科で見出しを分けるため（20261031）
                    seg = next((x for x in SZ.segments(dep) if SZ.matches(x, pats[o])), dep)
                    hit.append((o, seg))
        for o, seg in hit:
            if o in rosters:
                by_owner[o].append(f"外れる：{p.get('name') or k}{DEPT_OPEN}{seg}{DEPT_CLOSE}")

    # 週パターン（区分の変化）
    pw = prev_works(prev_master, owners)
    seen, week = set(), []
    works = {}                                    # ログイン → [(提出単位, 区分)]（2つ以上の「貴科リスト」に載る医師。20261031）
    for o, ros in rosters.items():
        for mem in ros:
            k = _key(mem.get("login"))
            if k:
                works.setdefault(k, []).append((o, C.work_of(mem)))
    for o, ros in rosters.items():
        for mem in ros:
            k = _key(mem.get("login"))
            if not k or k in seen or SZ.is_excluded(mem, title_of(mem)):
                continue
            seen.add(k)
            cur = C.work_of(mem)
            # 集合当直では、行をその医師の「Dr.JOY所属」の見出しへ（選んだ所属 → Dr.JOY の所属。20261031 コードレビュー③）
            seg = SZ.member_seg(mem, hist.dept_on(mem.get("login"), d1, title_of(mem)), pats[o])
            tag = f"{DEPT_OPEN}{seg}{DEPT_CLOSE}" if seg else ""
            if len({_cls(w) for _o, w in works.get(k, [])}) > 1:
                # 科によってフルタイムとそれ以外が食い違う → 週パターンは変えない（コードレビュー 20261031: 枠マスタで先の科の区分だけで
                #   勤務なしにしていた）。管理者が確かめる
                both = "・".join(f"{x}：{C.WORK_LABEL[w]}" for x, w in works[k])
                by_owner[o].append(f"{WEEK_PREFIX['conflict']}{mem.get('name') or k}（{both}）{tag}")
                continue
            tk = f"{y:04d}-{m:02d}|{k}"
            before = told.get(tk) or (_cls(pw[k]) if k in pw else None)
            name = mem.get("name") or k
            line = None
            if before is None:
                if cur != "full":
                    line = f"週パターンを勤務なしに（新しく申告）：{name}（{C.WORK_LABEL[cur]}）"
                    new_told[tk] = "nonfull"
            elif before != _cls(cur):
                line = (f"週パターンを勤務なしに：{name}（{C.WORK_LABEL[cur]}）" if cur != "full"
                        else f"週パターンを通常に：{name}")
                new_told[tk] = _cls(cur)
            if line:
                line += tag
                by_owner[o].append(line)
            # 20261031 管理者: 週パターンは知らせたかどうかでなく、Dr.JOY が区分と合っているかで直す（week_pattern が比べる）。
            #   区分が変わった人（line あり）に加えて、フルタイム以外の全員と、こちらで勤務なしにしてフルタイムに戻った人を渡す
            if line or cur != "full" or k in (recorded or {}):
                week.append({"owner": o, "line": line, "tk": tk, "login": mem.get("login"), "name": name,
                             "work": cur, "dir": "on" if cur == "full" else "off", "seg": seg})
    return {"owners": {o: xs for o, xs in by_owner.items() if xs}, "non_duty": non_duty, "told": new_told, "week": week,
            "ingai": [f"{n}（{t}）：{w}" for n, t, w in hist.ingai_violations(dt.date(y + 1, 12, 31))]}   # 先の月の予定も早めに知らせる


def has_items(res):
    return bool(res["owners"] or res["non_duty"] or res.get("ingai"))


SUBJECT_MARK = " ※自動送信"      # 20261031 管理者: 管理者が手で送ったように見えないように（自動送信なので末尾）
CONTACT_NAME = "当直勤務表"
NOT_FOUND = "（今の所属：Dr.JOY の名簿に見つからない）"
# 科ごとの行を、することごとにまとめる（20261031 管理者: 「以下を Dr.JOY 所属に追加して」＋列挙の形に）
GROUPS = [
    # 20261031 管理者: 食い違いの多くは診療科がリストを整えていないだけ → 「追加して」と決めつけず、食い違いとして短く並べる
    ("add", "「貴科リスト」にあるが、「Dr.JOY所属」は別の科："),
    ("remove", "「Dr.JOY所属」はこの科だが、「貴科リスト」にない："),
    ("missing", "Dr.JOY にアカウントがない："),
    ("week_off", "週パターンを「勤務なし（院外研修中）」に："),
    ("week_on", "週パターンを通常の勤務に："),
    # 20261031 管理者: 週パターンは事務担当の仕事ではない → こちらで Dr.JOY に入れて「変更しました」と書く（week_pattern）
    ("week_off_done", "週パターンを「勤務なし（院外研修中）」にしました（作業不要）："),
    ("week_on_done", "週パターンを元に戻しました（作業不要）："),
    ("week_dept", "フルタイムに戻った。週パターンの設定を診療科に依頼してください："),
    ("week_duty", "区分が変わったが、当直があるので週パターンは変えていない（管理者が対応）："),
    ("week_failed", "週パターンを自動で変えられなかった（管理者が対応）："),
    ("week_conflict", "「貴科リスト」の区分が科によって違うので、週パターンは変えていない（管理者が対応）："),
]
UNTIDY_MARK = " ※まだ「貴科リスト」を直していない"
DEPT_OPEN, DEPT_CLOSE = "《", "》"           # 行に付けた Dr.JOY所属（本文では見出しにする）
WEEK_PREFIX = {"done_off": "週パターン済_勤務なし：", "done_on": "週パターン済_通常：", "dept": "週パターン_診療科：",
               "duty": "週パターン_当直あり：", "failed": "週パターン_失敗：", "conflict": "週パターン_区分食い違い："}
WEEK_GROUP = {"done_off": "week_off_done", "done_on": "week_on_done", "dept": "week_dept", "duty": "week_duty",
              "failed": "week_failed", "conflict": "week_conflict"}


def _split_dept(line):
    """行の末尾の《Dr.JOY所属》を外す → (行, 所属 or None)"""
    if line.endswith(DEPT_CLOSE) and DEPT_OPEN in line:
        i = line.rindex(DEPT_OPEN)
        return line[:i], line[i + 1:-1]
    return line, None


def _group_of(line):
    """month_items の1行 → (グループ, 名前の部分)"""
    line, _dept = _split_dept(line)
    for k, pre in WEEK_PREFIX.items():
        if line.startswith(pre):
            return WEEK_GROUP[k], line[len(pre):]
    if line.startswith("加わる："):
        rest = re.sub(r"（[^（）]*のどちらか）", "", line[len("加わる："):])
        return ("missing", rest.replace(NOT_FOUND, "")) if NOT_FOUND in rest else ("add", re.sub(r"（今の所属：(.*?)）", r"（「Dr.JOY所属」は\1）", rest))
    if line.startswith("外れる："):
        return "remove", line[len("外れる："):]
    if line.startswith("週パターンを通常に："):
        return "week_on", line.split("：", 1)[1]
    if line.startswith("週パターンを勤務なしに"):
        return "week_off", line.split("：", 1)[1] + ("（新しく申告）" if "新しく申告" in line.split("：", 1)[0] else "")
    return "add", line


def mail_subject(months):
    """months = 知らせることがある月 [(y, m)]"""
    y, m = months[0]
    return f"【当直勤務表】「Dr.JOY所属」の変更（{reiwa_label(y, m)}分{'ほか' if len(months) > 1 else ''}）{SUBJECT_MARK}"


def owner_head(o, dept_map):
    """提出単位 → 見出しにする Dr.JOY所属（集合当直は「A／B」）。引けなければ提出単位の名前。"""
    pats = SZ.owner_patterns(o, dept_map) if dept_map else []
    return "／".join(pats) if pats else o


def mail_body(results, hist, compared, dept_map=None):
    """results = [((y, m), month_items の戻り)]、compared = [(y, m, 前の月のマスターがあったか)]"""
    if dept_map is None:
        try:
            dept_map = SZ.load_dept_map()
        except Exception:
            dept_map = {}
    out = ["事務担当", "",
           "当直勤務表アプリからの自動送信です。",
           "「貴科リスト」と「Dr.JOY所属」が食い違っています。",
           HEAD, ""]
    for (y, m), res in results:
        if not has_items(res):
            continue
        out.append(f"■ {reiwa_label(y, m)}分")
        # 見出しは診療科（Dr.JOY所属）。集合当直（内科 消化器・呼吸器など）は、所属の分かる行（外れる）をその科の見出しへ、
        # 行き先の決まらない行（加わる・週パターン）を「A／B」の見出しへ（20261031 管理者: 集合当直は診療科ではない）
        heads, owners_of = {}, {}
        for o, xs in res["owners"].items():
            for x in xs:
                body_line, dept = _split_dept(x)
                h = dept or owner_head(o, dept_map)
                g, txt = _group_of(body_line)
                heads.setdefault(h, {}).setdefault(g, []).append(txt)
                if o not in owners_of.setdefault(h, []):
                    owners_of[h].append(o)
        for h, grouped in heads.items():
            mark = [o for o in owners_of[h] if o in (res.get("untidy") or ())]
            out.append(f"【{h}】" + (f"{UNTIDY_MARK}（{'・'.join(mark)}）" if mark and h != mark[0] else UNTIDY_MARK if mark else ""))
            for g, head in GROUPS:
                if grouped.get(g):
                    out.append(head)
                    out += [f"・{t}" for t in grouped[g]]
            out.append("")
        if res["non_duty"]:
            out.append(f"【{NON_DUTY_HEAD}】")
            out += [f"・{x}" for x in res["non_duty"]]
        if res.get("ingai"):
            out.append(f"【{INGAI_HEAD}】")
            out += [f"・{x}" for x in res["ingai"]]
        if res["non_duty"] or res.get("ingai"):
            out.append("")
    out.append("――")
    seen = []                                      # 比べたもの（1行に。20261031 短く）
    if hist.staffs_name:
        seen.append(f"Dr.JOY のスタッフ一覧 {hist.staffs_name}")
    for name, t in hist.files:                    # 日付付きの予定（今日より後に始まる組）を見るために読んだ版（20261028）
        seen.append(f"所属履歴 {name}（{t.month}/{t.day} {t:%H:%M}）")
    seen += [f"{reiwa_label(y, m)}分の「貴科リスト」" for y, m, _p in compared]
    out.append("比べたもの：" + "・".join(seen))
    out += ["", "（管理者の PC が月末と毎週水曜の朝9時に自動で送っています）",
            f"問い合わせ先: {CONTACT_NAME}の担当"]
    return "\n".join(out)


def load_state(path=None):
    p = path or STATE_PATH
    try:
        with open(p, encoding="utf-8") as f:
            st = json.load(f)
        return st.get("told") or {}
    except (OSError, ValueError, AttributeError):
        return {}


def save_state(told, today, path=None):
    """覚えた区分を書く（古い月は捨てる）。書けなくても落ちない（戻り: 書けたか）。"""
    p = path or STATE_PATH
    lim = today.year * 12 + today.month - KEEP_MONTHS
    keep = {k: v for k, v in told.items()
            if k[:7].replace("-", "").isdigit() and int(k[:4]) * 12 + int(k[5:7]) >= lim}
    try:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        from master_store import tmp_name            # 20261002: 二重拡張子にしない（EDR がランサムウェア疑いで通報する）
        tmp = tmp_name(p)
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"schema": "hospital.shozoku-notice/1", "told": keep}, f, ensure_ascii=False, indent=1)
        os.replace(tmp, p)
        return True
    except OSError:
        return False


def send(to, subject, body, cc):
    if SENDER is not None:
        return SENDER(to, subject, body, cc)
    import dtx14_daily_sync as DS
    return DS.outlook_send(to, subject, body, cc="; ".join(cc))


SYSTEM_BY = {"unreg_rebind"}                    # 人でない保存（名簿外の付け替え）
ADMINS_FALLBACK = {"user003", "user004"}


def load_admins():
    """全体管理者（管理者・事務担当など）のログイン名。その人の保存は「科が整えた」に数えない。"""
    try:
        with open(os.path.join(C.PROD_ROSTER, "roster_editors.json"), encoding="utf-8") as f:
            eds = json.load(f).get("editors") or {}
        got = {k.lower() for k, v in eds.items() if isinstance(v, dict) and v.get("level") == "全体"}
        return got or ADMINS_FALLBACK
    except (OSError, ValueError, AttributeError):
        return ADMINS_FALLBACK


def _roster_key(ros):
    return sorted((str(x.get("login") or "").lower(), C.work_of(x)) for x in ros or [])


def tidied_owners(path, master, admins=None):
    """その月のリストを科（管理者・自動処理でない人）が一度でも直した提出単位（20261031 管理者）。
    直した = 共有マスターの版（_履歴 と今の版）を順に見て、リスト（人と区分）が変わった版の保存者が科の人、
    または、いまのリストにフルタイム以外か手入力の医師がいる（最初の提出の前に直した分は版に残らないため）。"""
    admins = admins if admins is not None else load_admins()
    tidy = {o for o, e in (master.get("depts") or {}).items()
            if any(C.work_of(x) != "full" or SZ.is_unregistered(x) for x in e.get("roster") or [])}
    prev = {}
    for h in sorted(glob.glob(os.path.join(os.path.dirname(path), "_履歴", "*.json"))) + [path]:
        try:
            with open(h, encoding="utf-8") as f:
                d = json.load(f)
        except (OSError, ValueError):
            continue
        for o, e in (d.get("depts") or {}).items():
            k = _roster_key(e.get("roster"))
            who = str(e.get("by") or "").split(" ")[0].strip().lower()
            if o in prev and prev[o] != k and who and who not in SYSTEM_BY and who not in admins:
                tidy.add(o)
            prev[o] = k
    return tidy


def _week_recorded():
    """こちらで勤務なしにした医師（week_pattern の覚え書き）{login: …}。読めなければ空。"""
    try:
        import week_pattern as WP
        return WP.load_state()
    except Exception:
        return {}


def apply_week(res, y, m, master, today, log, week_apply=None):
    """週パターンの知らせを Dr.JOY に入れ（week_pattern.apply）、名簿の行を結果の行に差し替える。戻り: 問題の文。
    week_apply = week_pattern.apply の代わり（テスト）。入らなかった医師は覚えた区分から外す＝次の回にやり直す。"""
    import week_pattern as WP
    fn = week_apply or WEEK_APPLY or WP.apply
    try:
        items = fn([dict(w) for w in res["week"]], master, y, m, today, log=log)
    except Exception as e:                        # 週パターンが入らなくても、所属のメールは送る
        items = [dict(w, result="failed", why=f"{type(e).__name__}: {e}", date="") for w in res["week"]]
    probs = []
    for x in items:
        r = x.get("result")
        lines = res["owners"].get(x["owner"]) or []
        if r in ("already", "nocode"):            # Dr.JOY がもう区分と合っている／アカウントが無い（別の見出しで出る）→ 書かない
            if x.get("line") in lines:
                lines.remove(x["line"])
                if not lines:
                    res["owners"].pop(x["owner"], None)
            continue
        key = ("done_on" if x["dir"] == "on" else "done_off") if r == "done" else r if r in ("dept", "duty") else "failed"
        when = ""
        if x.get("date"):
            d = dt.date.fromisoformat(x["date"])
            when = f"{d.month}月{d.day}日から"
        label = C.WORK_LABEL.get(x.get("work"), "")
        pat = x.get("pattern") or ""
        pat = f"兼業の申請から作った「{pat}」" if pat.endswith("_自動登録") else f"「{pat}」"
        info = {"done_off": f"（{label}・{when}）", "done_on": f"（{pat}・{when}）",
                "dept": f"（{x.get('why')}）", "duty": f"（{label}・{x.get('why')}）",
                "failed": f"（{x.get('why')}）"}[key]
        new = f"{WEEK_PREFIX[key]}{x['name']}{info}" + (f"{DEPT_OPEN}{x['seg']}{DEPT_CLOSE}" if x.get("seg") else "")
        if x.get("line") in lines:
            lines[lines.index(x["line"])] = new
        else:                                     # 区分は前から同じだが、Dr.JOY が合っていなかった医師（20261031）
            res["owners"].setdefault(x["owner"], []).append(new)
        if key == "failed" and x.get("line"):
            res["told"].pop(x["tk"], None)
            probs.append(f"週パターンを入れられない: {x['name']}（{x.get('why')}）")
    return probs


def run(targets, *, dry=False, master_file=None, prev_master_file=None, hist=None, sender=None,
        log=print, today=None, frames=None, dept_map=None, week_apply=None):
    """targets の月を比べて、知らせることがあれば1通送る。
    週パターンの知らせは、送る前に Dr.JOY に入れる（week_pattern。dry では入れない。week_apply=False で入れない）。
    戻り: {"sent": bool, "subject", "body", "months": [(y, m)], "n": 行数, "problems": [文]}"""
    today = today or dt.date.today()
    master_file = master_file or master_path
    prev_master_file = prev_master_file or master_file
    hist = hist or SZ.load_history(log=log)
    frames = frames if frames is not None else C.load_frames()
    dept_map = dept_map if dept_map is not None else SZ.load_dept_map()
    told = load_state()
    results, compared, new_told, problems = [], [], {}, list(hist.problems)
    for (y, m) in targets:
        mst = read_master(master_file(y, m), y, m)
        if mst is None:
            continue
        py, pm = prev_month(y, m)
        pmst = None
        if (py, pm) >= tuple(SZ.REAL_FROM):          # 予行の月（R8.10 など）とは比べない（20260927）
            try:
                pmst = read_master(prev_master_file(py, pm), py, pm)
            except Exception as e:                # 前の月が読めなくても、所属の照合はする
                problems.append(f"{reiwa_label(py, pm)}分のマスターが読めない（{e}）")
                pmst = None
        res = month_items(y, m, mst, pmst, hist, frames, dept_map, told, recorded=_week_recorded())
        try:
            res["untidy"] = set(res["owners"]) - tidied_owners(master_file(y, m), mst)
        except Exception as e:                    # 印が付けられなくてもメールは送る
            log(f"  所属の変更: リストを直した科を確かめられない（{type(e).__name__}: {e}）")
            res["untidy"] = set()
        if res.get("week") and not dry and week_apply is not False:
            problems += apply_week(res, y, m, mst, today, log, week_apply)
        compared.append((y, m, pmst is not None))
        results.append(((y, m), res))
        new_told.update(res["told"])
    months = [ym for ym, res in results if has_items(res)]
    n = sum(len(xs) for _ym, res in results for xs in list(res["owners"].values()) + [res["non_duty"], res.get("ingai") or []])
    out = {"sent": False, "subject": "", "body": "", "months": months, "n": n, "problems": problems}
    if not months:
        log("  所属の変更: 知らせることなし" + (f"（{', '.join(reiwa_label(*ym) for ym, _r in results)}）" if results else ""))
        return out
    out["subject"] = mail_subject(months)
    out["body"] = mail_body(results, hist, compared)
    if dry:
        log(f"  (dry-run) 所属の変更メールは送らない: {out['subject']}（{n} 行）")
        return out
    (sender or send)(TO, out["subject"], out["body"], CC)
    out["sent"] = True
    log(f"  所属の変更メールを送った: {TO}（cc {'; '.join(CC)}）『{out['subject']}』{n} 行")
    if new_told:
        merged = dict(told)
        merged.update(new_told)
        if not save_state(merged, today):
            problems.append(f"知らせた区分を覚えておけない（{STATE_PATH}）→ 次の回も同じ週パターンの知らせが出る")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="対象の月 YYYY-MM")
    ap.add_argument("--dry", action="store_true", help="送らない（本文を表示するだけ）")
    ap.add_argument("--master", default=None, help="当直予定表の JSON（省略=本番の月フォルダ）")
    ap.add_argument("--prev-master", default=None, help="前の月の当直予定表の JSON（省略=本番の月フォルダ）")
    a = ap.parse_args()
    y, m = (int(x) for x in a.month.split("-"))
    mf = (lambda yy, mm: a.master) if a.master else master_path
    pf = (lambda yy, mm: a.prev_master) if a.prev_master else (
        (lambda yy, mm: None) if a.master else master_path)
    res = run([(y, m)], dry=a.dry, master_file=mf, prev_master_file=pf)
    for x in res["problems"]:
        print(f"⚠ {x}")
    if res["body"]:
        print(f"件名: {res['subject']}\n宛先: {TO}\ncc: {'; '.join(CC)}\n")
        print(res["body"])


if __name__ == "__main__":
    main()
