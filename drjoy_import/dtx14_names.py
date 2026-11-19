# -*- coding: utf-8 -*-
"""DTX14のE列(フルネーム)が空の当直行を、氏名まで解決する。20260828

★DTX14そのものは書き換えない（E列は「所属医師」シートからマクロで生成される列で、
  事務担当が埋める作業領域。直接書いても次にマクロを回すと消える）。
  Dr.JOYへの投入に必要なぶんだけ、ここで解決する。

解決の材料（この順に当てる）
  1. 前月のDTX14のE列 … 同じ枠でその略称が誰だったか。最も確か
  2. 科の分かる名簿（DTX14「勤怠名簿」＋ staffs.csv）… 枠の科で絞れる
  3. 医師だけの名簿（統合ファイル＋臨床研修ローテーション表）
     … 姓が医師の中で一意なら確定。看護師・事務は当直に立たないので落とす
  4. 統合ファイルの `Entra部署` `dtx14_entra.py` … ★一番効く。医師のほぼ全員に実際の診療科が入っている
  5. 臨床研修ローテーション表 `dtx14_rota.py` … 初期研が**その日**どの科にいるか
  6. 専攻医の院外研修期間表 `dtx14_senko.py` … 専攻医の診療科＋院外に出ている期間

R8.8月分の実績＝E列未記入 435行を**全部**機械解決。
★正しさは「全部埋まったか」ではなく、**E列が既に埋まっている1,202行で当ててみて測る**
  （検証用のスクリプト）。現状 **1,202/1,202＝誤り0**。
  規則を足すたびにこれで測ること。実際「残った1人を採る」規則は 5/13 → 0/6 と外しまくっていた。

★**無理に当てはめない**。まれに「当院の職員でも研修中でもない医師が、よそから当直だけ
  しに来る」ことがある（ユーザー説明 2026-08-26）。どの名簿にも載っていないので、
  候補が1人も残らないのは**異常ではなく正しい答え**。そういう人は職員コードも人別シートも
  無いため、Dr.JOYへの投入では自然に見送られる。

★B列の略称は **括弧付きのことがある**（`(瑠璃)` 等＝オンコール表記）。必ず外して照合する。
★初期研は1.5〜3か月で所属が変わる（専攻医は基本変わらない。外科の肝胆膵・上部・下部だけ数か月で動く）
  ので、初期研は所属で絞らず「その日どの科に回っているか」で見る。
★**ローテーションは月で引いてはいけない**。ローテーションの区切りは月の途中にあり、
  8月の途中で科が替わる。日付で引くまで銀杏・桜庭・柊木の3人が解けなかった。
★**消去法で1人に絞れても、その人の所属が枠と違うなら採らない**
  （一度「耳鼻科の医師が産婦人科の当直」という答えを出した）。
"""
from __future__ import annotations

import csv
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, r"C:\tools\touchoku\drjoy_import")
import openpyxl

# 当直の枠の科 → 名簿の部署
FRAME_TO_DEPT = {
    "内科": None,                        # 内科全体からの持ち回り。科で絞れない
    "循環器内科": "循環器内科",
    "内科 消化器・呼吸器": ["内科(消化器)", "内科(感染症・呼吸器)"],
    "外科 心臓・呼吸器": "外科(心臓血管・呼吸器)",
    "外科": ["外科(消化管)", "外科(肝胆膵・乳腺)"],
    "小児科": "小児科",
    "形成外科": "形成外科",
    "救急科": "救急部",
    "整形外科": "整形外科",
    "歯科口腔外科": "歯科口腔外科",
    "泌尿器科": "泌尿器科",
    "産科婦人科": "産科婦人科",
    "皮膚科": "皮膚科",
    "眼科": "眼科",
    "精神科": "精神科",
    "耳鼻咽喉科": "耳鼻咽喉科",
    "脳神経外科": "脳神経外科",
    "管理当直": None,                             # 全科から
    "集中治療室": None,
    "麻酔科": "麻酔科",
}

# ★当直の用語は**2つの軸**（ユーザー説明 2026-08-26）。混同しないこと。
#   軸① ファースト／セカンド以下 … 病棟・救急外来が最初にコールする順。
#        呼ばれてすぐ判断するので **基本 専攻医以上**。
#   軸② 当直(括弧なし)／オンコール(括弧あり) … **病院に泊まるかどうか**だけ。
#        経験年数とは相関しない（初期研でも泊まるし、教員でもオンコールはある）。
#   → 下は軸①の帰結。**ファーストの行に限って**初期研を外す。
#     セカンド以下・オンコールには初期研が普通に入る
#     （産科婦人科のセカンド＝銀杏恵子・桜庭健一・柊木健三はいずれも初期研）。
#   ⚠「ファースト＝専攻医以上」は傾向であって絶対ではない。R8.8月のファーストに
#     初期研が5行あり、白砂誠二(N期)は循環器内科のファーストに立っている。
#     → **例外0件を実測できた精神科・産科婦人科だけ**に効かせる。全科に広げると取り違える。
#   実測(E列のフルネームで): R8.7月ファースト748行中 初期研0行／R8.8月 536行中5行。
NO_RESIDENT_FIRST = ("精神科", "産科婦人科")

# ★どの規則でも解けなかった (枠の科, 略称) → 氏名。resolve() が最初に見る。20260828
#   R8.9月でDTX14のE列が空だった54行を、staffs.csv・人事給与システム・所属医師シートで確定させたもの。
#   ⚠**DTX14には触っていない**（E列はマクロが再生成するので直接書いても消えるため。
#     ここで解けば投入には十分で、事務担当のファイルを不在中に書き換えずに済む）。
#   救急科の8名は全員「初期研修医N期がローテーションで救急科の当直に入っている」形。
#   所属医師シートの救急科の行には載らず、外科系の行に載っているのでE列が空になっていた。
MANUAL_RESOLVE = {
    # 同姓が1名だけ
    ("救急科", "冬野"): "冬野五郎",      # 初期研N期
    ("救急科", "雪村"): "雪村秋子",    # 初期研N期
    ("救急科", "霞野"): "霞野次郎",      # 初期研N期
    ("救急科", "青葉"): "青葉裕子",      # 初期研N期
    ("救急科", "椿井"): "椿井明美",      # 専攻医N期・他院
    ("耳鼻咽喉科", "楓川"): "楓川健二",
    # 同姓が複数いるので、同じ日に本人名義の当直がある方を外して確定した
    ("救急科", "夕凪"): "夕凪春子",    # 夕凪太郎は9/13・19・22に本人名義の当直あり
    ("救急科", "若葉"): "若葉洋子",        # 若葉直子は9/4に本人名義の当直あり
    # 枠と所属が一致する方
    ("耳鼻咽喉科", "紅林"): "紅林誠一",    # 紅林健一は泌尿器科
    ("救急科", "癸田"): "癸田浩二",      # 救急科の空欄は全員初期研N期。癸田四郎は循環器内科で別に7日
    # 異体字（髙＝高）
    ("内科", "髙杉"): "高杉見本",     # 内科(腎臓・内分泌)・大学院生N期
    ("内科", "高杉"): "高杉見本",
    # ⚠ここだけ決め手が弱い（ユーザー判断 2026-08-28）。黒潮姓は8名。
    #   黒潮浩一は9/12に本人名義の当直があり除外。内科系は黒潮真二だけなので採った。
    ("内科", "黒潮"): "黒潮真二",       # 内科(神経)・専攻医N期
}

# ★DTX14の「勤怠名簿」は更新されているか怪しいので使わない（ユーザー指示 2026-08-29）。20260829
#   代わりに staffs.csv（Dr.JOYから毎回取り直す最新版）を正とする。
#   ⚠ただし staffs.csv だけにすると**別人に当てる誤りが増える**（下の VAGUE_DEPTS を参照）。
#     必ず STRICT_NARROW とセットで使うこと。
#   R8.5〜9月の答え合わせ（E列が埋まっている6,636行）
#     勤怠名簿あり・従来   : 誤り24（うち本当の別人2）／決められず  672
#     勤怠名簿なし・従来   : 誤り48（うち本当の別人28）／決められず 1675  ←★これは危ない
#     勤怠名簿なし・厳しく : 誤り20（うち本当の別人**0**）／決められず 2372  ←採用
USE_ROSTER = False

# ★所属としては書いてあるが「その科ではない」とは言えないもの。20260829
#   staffs.csv は**当院で当直に立っている医師を「他院」と書くことがある**
#   （ある医師＝勤怠名簿では耳鼻咽喉科／staffs.csvでは他院）。
#   これを「耳鼻咽喉科ではない」と読むと、同じ枠の同姓の別の医師に化ける。
VAGUE_DEPTS = ("", "他院", "専門研修医", "初期研修医", "初期研修", "総合診療科")

# ★枠の科で絞るとき、所属のはっきりしない候補が残っていたら決めない。20260829
#   False にすると従来どおり（絞れた1人を採る）。
STRICT_NARROW = True

# ★全科から出る枠。診療科での絞り込みをしてはいけない
#   （実データで「循環器内科の医師が集中治療室の枠」206件。ほぼ全部これ）
ALL_DEPT_FRAMES = ("管理当直", "集中治療室")

# 統合ファイルの職種名のうち、当直に立ちうるもの（看護師・事務・委託は落とす）
DOCTOR_KINDS = ("医師", "専門研修医", "初期研修医", "歯科医師", "教員", "大学院生")

# 臨床研修ローテーション表の科目名 → 当直の枠の科
# ⚠外部病院・クリニックに出ている月は当院の当直に入らない → どの枠にも当たらない（None のまま）
_ROTA = {
    "総合診療科": "内科", "内科A": "内科", "内科B": "内科", "内科C": "内科",
    "救急": "救急科", "麻酔科": "麻酔科", "整形外科": "整形外科", "脳神経外科": "脳神経外科",
    "精神科": "精神科", "小児科": "小児科", "産婦人科": "産科婦人科", "泌尿器科": "泌尿器科",
    "皮膚科": "皮膚科", "眼科": "眼科", "耳鼻咽喉科": "耳鼻咽喉科", "形成外科": "形成外科",
    "外科A・C": "外科", "外科B": "外科 心臓・呼吸器",
}


def ROTA_TO_FRAME(subject):
    """ローテ表の科目名を当直の枠の科に寄せる。`（必修＋選択）` や全角数字のゆれを吸収。"""
    s = re.sub(r"[（(].*?[）)]|[/／\s　]", "", str(subject or ""))
    s = s.translate(str.maketrans("１２３", "123"))
    return _ROTA.get(s)

_PAREN = re.compile(r"[（(【\[]?([^）)】\]]+)[）)】\]]?")
_NAME = re.compile(r"[一-龥ヶ々]{1,4}[\s]?[一-龥ぁ-んァ-ヶ々]{1,5}")


# 異体字。★DTX14と名簿で字が違うことがある（R8.9月の「髙」＝名簿は「高」）。
# ⚠同じ字の書き方違いだけを入れる。**斎/斉・沢/澤 のような別の字は併合しない**。
_ITAIJI = str.maketrans({
    "髙": "高", "﨑": "崎", "桒": "桑", "濵": "浜", "栁": "柳", "德": "徳",
    "眞": "真", "淺": "浅", "靑": "青", "賴": "頼", "曻": "昇", "硏": "研",
    "瀨": "瀬", "圡": "土", "祐": "祐", "凰": "凰",
    # 20260906 ある医師。正しいのは旧字の「惠」（LDAP・人事給与システム・統合ファイル・DTX14すべて惠。
    # 一時期Dr.JOYだけ「恵」で結べず、当直6人日が「職員コードが分からない」で見送られていた）。
    # これは比較用の寄せなので惠・恵のどちらで来ても一致する。名前を書くときは使わない。
    "惠": "恵",
})


def clean_abbr(s):
    """B列の略称。`(瑠璃)` のような括弧付き（オンコール表記）を外す。"""
    t = str(s or "").replace("　", " ").strip()
    m = _PAREN.fullmatch(t)
    return (m.group(1) if m else t).strip().translate(_ITAIJI)


def _sur(name):
    """姓。★索引の鍵になるので異体字も寄せる（「髙」で引いて「高」に当てる）。"""
    return str(name or "").replace("　", " ").strip().split(" ")[0].translate(_ITAIJI)


def canon(name):
    """氏名の正規化。★同じ人が「銀杏 恵子」と「銀杏恵子」の2通りで名簿に入っており、
    そのままだと1人を2人と数えて「候補が複数」になり解決できなくなる。
    あわせて異体字も寄せる（「髙」と「高」）。"""
    return str(name or "").replace(" ", "").replace("　", "").translate(_ITAIJI)


def _cell(r, i):
    return str(r[i]).strip() if len(r) > i and r[i] is not None else ""


def _sheet(wb, name):
    ws = wb[name]
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    return ws


def load_roster(dtx14_path):
    """DTX14の「勤怠名簿」→ [(部署, 氏名)]"""
    out = []
    if not dtx14_path or not os.path.exists(dtx14_path):
        return out
    wb = openpyxl.load_workbook(dtx14_path, read_only=True, data_only=True)
    if "勤怠名簿" in wb.sheetnames:
        for i, r in enumerate(_sheet(wb, "勤怠名簿").iter_rows(values_only=True)):
            if i == 0:
                continue
            dep, nm = _cell(r, 0), _cell(r, 2)
            if dep and nm:
                out.append((dep, nm))
    wb.close()
    return out


def load_prev(prev_dtx14):
    """前月のDTX14 → {(枠の科, 略称): {フルネーム}}"""
    m = defaultdict(set)
    if not prev_dtx14 or not os.path.exists(prev_dtx14):
        return m
    wb = openpyxl.load_workbook(prev_dtx14, read_only=True, data_only=True)
    for sn in ("人別当直日", "人別ファースト"):
        if sn not in wb.sheetnames:
            continue
        for i, r in enumerate(_sheet(wb, sn).iter_rows(values_only=True)):
            if i == 0 or not r or not _cell(r, 0):
                continue
            full, ab = _cell(r, 4), clean_abbr(_cell(r, 1))
            if full and ab:
                m[(_cell(r, 0), ab)].add(full.replace(" ", "").replace("　", ""))
    wb.close()
    return m


def load_staffs(staffs_csv):
    """Dr.JOYの staffs.csv → [(所属, 氏名)]"""
    out = []
    if not staffs_csv or not os.path.exists(staffs_csv):
        return out
    rows = list(csv.reader(open(staffs_csv, encoding="cp932", errors="replace")))
    h = rows[0]
    iS, iM, iD = h.index("姓"), h.index("名"), h.index("所属")
    for r in rows[1:]:
        if len(r) <= iD:
            continue
        s, m2 = (r[iS] or "").strip(), (r[iM] or "").strip()
        if s:
            out.append(((r[iD] or "").strip("/"), f"{s} {m2}".strip()))
    return out


def load_tougou(tougou_csv):
    """統合ファイル抽出CSV → ([(所属部署, 氏名)], {職種名: {氏名}})（★医師系の職種だけ）

    ⚠専攻医・初期研の所属部署は「専門研修医」「初期研修医」としか入っておらず
      科では絞れない。姓が医師の中で一意かどうかを見るためだけに使う。
    """
    out, kinds = [], defaultdict(set)
    if not tougou_csv or not os.path.exists(tougou_csv):
        return out, kinds
    for r in csv.DictReader(open(tougou_csv, encoding="utf-8-sig")):
        if (r.get("就業状況") or "").strip() == "無効":
            continue
        k = (r.get("職種名") or "").strip()
        if k not in DOCTOR_KINDS:
            continue
        s, m2 = (r.get("姓") or "").strip(), (r.get("名") or "").strip()
        if s:
            nm = f"{s} {m2}".strip()
            out.append(((r.get("所属部署") or "").strip(), nm))
            kinds[k].add(nm.replace(" ", ""))
    return out, kinds


def load_rinken(paths):
    """臨床研修ローテーション表（個人別シート）→ [("", 氏名)]

    初期研の氏名を拾うためのもの。科はローテーションで動くので持たない。
    """
    out, seen = [], set()
    for p in paths or []:
        if not os.path.exists(p):
            continue
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for sn in wb.sheetnames:
            if "個人別" not in sn:
                continue
            for r in _sheet(wb, sn).iter_rows(values_only=True):
                for c in (r or [])[:6]:
                    s = str(c or "").replace("　", " ").strip()
                    if _NAME.fullmatch(s) and len(s.replace(" ", "")) >= 3 and s not in seen:
                        seen.add(s)
                        out.append(("", s))
        wb.close()
    return out


def build_resolver(dtx14, prev_dtx14=None, staffs_csv=None,
                   tougou_csv=None, rinken=None, rota=None, senko=None, entra=None, verbose=False):
    """→ resolve(枠の科, 略称, 当直日) = (氏名 or None, 理由, 候補list)

    rota = {氏名(スペース無し): [(開始日, 終了日, 科目名)]}。初期研を日付で絞る。
    ★月単位ではだめ。月の途中でローテーションが替わり、同じ月でも科が違う。
    """
    # ★勤怠名簿は既定で使わない（USE_ROSTER）。staffs.csv が正。
    roster = load_roster(dtx14) if USE_ROSTER else []
    staffs = load_staffs(staffs_csv)
    named = roster + staffs                                     # 科が分かる名簿
    tou, kinds = load_tougou(tougou_csv)
    docs = tou + load_rinken(rinken)                            # 医師だけ（科は当てにしない）
    prev = load_prev(prev_dtx14)
    dentists, residents = kinds.get("歯科医師", set()), kinds.get("初期研修医", set())
    rota = rota or {}
    # senko = (frames_of(氏名)->set|None, is_away(氏名, 日付)->bool)
    # entra  = frames_of(氏名)->set|None（統合ファイルの Entra部署。一番広く科が入っている）
    # ★どちらか一方でも「その枠に立ちうる」と言えば残す（union）。
    #   実データで1件だけ食い違った（ある専攻医＝専攻医表と Entra で所属が違った）ので、
    #   片方が「立たない」と言うだけでは落とさない。
    sen_frames, sen_away = senko or ((lambda _n: None), (lambda _n, _d: False))
    ent_frames = entra or (lambda _n: None)

    def known_frames(name):
        """その人が立ちうる枠。どの名簿も知らなければ None（＝否定しない）。"""
        parts = [f for f in (sen_frames(name), ent_frames(name)) if f is not None]
        if not parts:
            return None
        return set().union(*parts)

    def index(pairs):
        d = defaultdict(dict)          # 姓 → {氏名(正規化): 部署}
        for dep, nm in pairs:
            d[_sur(nm)].setdefault(canon(nm), dep)
        return d

    by_named, by_doc = index(named), index(named + docs)

    def by_prefix(ab):
        """★略称が**姓の頭だけ**のことがある（例: 姓の頭の1字だけ）。前方一致でも候補を拾う。"""
        out = dict(by_doc.get(ab, {}))
        for sur, names in by_doc.items():
            if sur != ab and sur.startswith(ab):
                out.update(names)
        return out
    # 氏名 → 部署。★DTX14の勤怠名簿が最優先（当直表と同じ土俵）。
    # ⚠staffs.csv は出向者を「他院」と書いており、名簿の実所属を潰してしまう。
    by_dept = {}
    for src in (roster, staffs, docs):
        for dep, nm in src:
            if dep:
                by_dept.setdefault(canon(nm), dep)
    if verbose:
        print(f"  名簿(科あり) {sum(len(v) for v in by_named.values())}名 / "
              f"医師プール {sum(len(v) for v in by_doc.values())}名 / 前月 {len(prev)}件")

    def rota_frames(name, date):
        """その日その人がローテーションしている科を、当直の枠の科に直して返す。"""
        key = str(name).replace(" ", "").replace("　", "")
        sp = rota.get(key)
        if not sp or date is None:
            return None                    # 材料が無い（不明であって否定ではない）
        subj = {su for s, e, su in sp if s <= date <= e}
        return {ROTA_TO_FRAME(s) for s in subj} if subj else None

    def keepable(n, frame, date=None, sole=False):
        """その人がその枠・その日に立ちうるか。1人ずつの判定。

        ①歯科医師は歯科口腔外科の枠にしか立たない
        ②初期研。⚠ローテ先以外の枠にも立つ（実データで 白砂誠二＝総合診療科の月に
          循環器内科の当直）。**外部病院に出ている月だけ**落とす。
        ③所属が「他院」の人は当院にいない
        ④「内科」の枠は内科系・総合診療科から出る
        ⑤所属がはっきり別の科の人は立たない
        ⑥診療科（Entra部署／専攻医の院外研修期間表）がその枠と違う
        （④〜⑥は全科から出る枠＝管理当直・集中治療室では見ない）
        """
        key = canon(n)
        if key in dentists and frame != "歯科口腔外科":
            return False
        if key in residents and rota_away(n, date):
            return False
        if frame in ALL_DEPT_FRAMES:
            return True
        # ★診療科(Entra部署／専攻医の院外研修期間表)が分かるなら、それだけで決める。
        #   ⚠名簿の「他院」は staffs.csv 由来で**古いことがある**。
        #     琥珀和子（救急部の専攻医）は staffs.csv が他院のままだが、
        #     実際は8月に救急科のファーストを16回やっている。Entra部署の方が正しい。
        kf = known_frames(n)
        if kf is not None:
            return frame in kf
        # 診療科が分からないときだけ、名簿の所属で判断する
        dep = by_dept.get(key, "")
        if dep == "他院" and not sole:
            # ⚠この印は staffs.csv 由来で古いことがある。**他に候補が居るときだけ**効かせる。
            #   院内に同姓の医師が1人しかいないなら、その人しかありえない（月岡夏子）。
            return False
        if frame == "内科" and dep and not (dep.startswith("内科") or dep == "総合診療科"):
            return False
        return compatible(n, frame)

    def plausible(names, frame, date=None):
        """★どの経路でも必ずこれを通す。「名簿で姓が一意」だけ素通しにしていたため、
        他院に出ている初期研が救急科のファーストに採られていた（翡翠＝6行）。"""
        out = [n for n in names if keepable(n, frame, date)]
        return out or list(names)          # 全部落ちたら絞らない（材料が古い可能性）

    def rotating_other(ab, chosen, date):
        """同じ姓（頭一致を含む）の**研修医**がその日ローテーション中なら、決めない。

        ⚠常勤医と研修医が同姓だと取り違える。実データ＝循環器内科の「白砂」は
          常勤の白砂真一ではなく、総合診療科を回っていた初期研の白砂誠二だった。
          ローテ先と違う枠にも立つので枠では弾けない。**人に聞くのが正しい。**
        """
        if date is None:
            return False
        for n in by_prefix(ab):
            if canon(n) == canon(chosen) or canon(n) not in residents:
                continue
            subj = {u for s, e, u in rota.get(canon(n), []) if s <= date <= e}
            # その日ローテ先が分かっていて、かつ**院内**なら危ない（院外に出ていれば当直に立たない）
            if subj and not rota_away(n, date):
                return True
        return False

    def rota_away(name, date):
        """その日、外部の病院・クリニックに出ているか（＝当院の当直に立たない）。"""
        key = str(name).replace(" ", "").replace("　", "")
        sp = rota.get(key)
        if not sp or date is None:
            return False
        subj = {su for s, e, su in sp if s <= date <= e}
        return bool(subj) and all(
            ROTA_TO_FRAME(s) is None and any(w in s for w in ("病院", "クリニック", "医院"))
            for s in subj)

    def compatible(name, frame):
        """消去法で1人に絞れたとき、その人が本当にその枠に立ちうるかを念のため見る。

        ⚠「他が全員消えたから残った」だけでは決めない。実際に
          「耳鼻科の医師が産婦人科の枠」に化けたことがある。
          所属が別の科だとはっきり分かる人は採らない（所属不明・研修医は許す）。
        """
        key = canon(name)
        dep = by_dept.get(key, "")
        if frame in ALL_DEPT_FRAMES:
            return True
        # ★初期研は1.5〜3か月で所属が変わるので、**名簿の所属で絞ってはいけない**。
        #   R8.9月の雪村秋子（初期研）は名簿が「外科(消化管)」だが救急科のセカンドに入る。
        #   どの科に居るかは dtx14_rota（その日のローテ）でだけ見る。
        if key in residents:
            return True
        if not dep or dep in ("専門研修医", "初期研修医", "初期研修", "総合診療科"):
            return True                    # 材料が無い／どの科にも回りうる
        want = FRAME_TO_DEPT.get(frame)
        if want is None:                   # 内科・管理当直・集中治療室は科を問わない
            return True
        wl = want if isinstance(want, list) else [want]
        return dep in wl

    def resolve(frame, abbr, date=None, first=False):
        ab = clean_abbr(abbr)
        # ★どの規則でも解けなかった (枠, 略称) を人手で確定したもの。20260828
        #   規則より先に見る（規則が別人を当てるのを防ぐため）。根拠は MANUAL_RESOLVE のコメント。
        hit = MANUAL_RESOLVE.get((frame, ab))
        if hit:
            return hit, "手動の対応表", []
        no_res = first and frame in NO_RESIDENT_FIRST
        # ★その日その科に回っている研修医が1人だけなら、それが答え。
        #   ⚠**前月照合より先に見る**。ローテで入れ替わるので前月とは別人になる
        #   （実データ＝救急科の「朝霧」は7月が朝霧花子・8月は朝霧冬子）。
        on = [n for n in sorted(by_prefix(ab))
              if frame in (rota_frames(n, date) or set())]
        if len(on) == 1:
            return on[0], "その日その科に回っている", []
        p = prev.get((frame, ab))
        if p and len(p) == 1:
            return next(iter(p)), "前月の同じ枠", []

        cands = {n: d for n, d in by_named.get(ab, {}).items()
                 if not (no_res and n in residents)}
        if (len(cands) == 1 and keepable(next(iter(cands)), frame, date)
                and (no_res or not rotating_other(ab, next(iter(cands)), date))):
            return next(iter(cands)), "名簿で姓が一意", []
        if len(cands) > 1:
            want = FRAME_TO_DEPT.get(frame)
            if want:
                wl = want if isinstance(want, list) else [want]
                nar = [n for n, dep in cands.items() if dep in wl]
                vague = [n for n, dep in cands.items() if dep in VAGUE_DEPTS]
                if STRICT_NARROW and vague:
                    return None, "枠の科で絞れるが所属不明の候補が残る", sorted(cands)
                if len(nar) == 1 and (no_res or not rotating_other(ab, nar[0], date)):
                    return nar[0], "名簿＋枠の科", []
                if len(nar) > 1:
                    return None, "枠の科で絞っても複数", sorted(nar)

        pool = {n: d for n, d in by_doc.get(ab, {}).items()
                if not (no_res and n in residents)}
        if (len(pool) == 1 and keepable(next(iter(pool)), frame, date, sole=True)
                and (no_res or not rotating_other(ab, next(iter(pool)), date))):
            return next(iter(pool)), "医師の中で姓が一意", []
        if len(pool) > 1:
            # ★その日その科に回っている研修医が1人だけなら、それが答え
            # ★診療科がその枠と一致する人が1人だけなら、それが答え
            if frame not in ALL_DEPT_FRAMES:
                sp = [n for n in sorted(pool) if frame in (known_frames(n) or set())]
                if len(sp) == 1 and (no_res or not rotating_other(ab, sp[0], date)):
                    return sp[0], "診療科が枠と一致", []
            want = FRAME_TO_DEPT.get(frame)
            if want:
                wl = want if isinstance(want, list) else [want]
                nar = [n for n, dep in pool.items() if dep in wl]
                vague = [n for n, dep in pool.items() if dep in VAGUE_DEPTS]
                if STRICT_NARROW and vague:
                    return None, "枠の科で絞れるが所属不明の候補が残る", sorted(pool)
                if len(nar) == 1 and (no_res or not rotating_other(ab, nar[0], date)):
                    return nar[0], "医師＋枠の科", []
            # ⚠**純粋な**消去法（他が全員消えたから残った）は実測 5/13 しか当たらないので採らない。
            #   残った1人に「診療科がその枠と一致する」という積極的な裏づけがあるときだけ採る。
            #   ★「残った1人＋診療科が一致」も試したが実測 0/6 で全部外れた（姓の頭の字を、
            #     同じ字で始まる別の医師と読む）。**姓が一意に決まらない行は人に聞く**のが正しい。
            nar = plausible(sorted(pool), frame, date)
            if (len(nar) == 1 and frame in (known_frames(nar[0]) or set())
                    and (no_res or not rotating_other(ab, nar[0], date))):
                return nar[0], "絞り込みの結果1人＋診療科が一致", []
            return None, "医師の中で姓が複数", nar
        # ★ここに来たら**無理に当てはめない**。
        #   まれに「当院の職員でも研修中でもない医師が、よそから当直だけしに来る」
        #   ことがある（ユーザー説明 2026-08-26）。その人はどの名簿にも載っていない。
        #   職員コードも人別シートも無いので、Dr.JOYへの投入では自然に見送られる。
        return None, ("名簿に該当者がいない（外部から当直だけに来た医師の可能性）"
                      if not by_doc.get(ab) else
                      "候補はいるが枠に立ちえない（外部の医師の可能性）"), sorted(by_doc.get(ab, {}))

    return resolve
