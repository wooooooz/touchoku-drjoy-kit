# -*- coding: utf-8 -*-
"""dtx14_trim_vba.py — DTX14 から、アプリに置き換わったマクロとシートとボタンを取り除く。20260922

管理者 2026-09-22: 人別3シートはアプリが書く（F=教員・G=ログインID・H=職員コード）ので、
所属医師の転記／漏れチェック／名前チェック／フルネーム記載／教員チェック は要らない。
保存時のフック（Workbook_BeforeSave＝所属医師の科名選択と赤字のチェック）も要らない。
DTX14 のマクロは 97% 管理者が書いたものなので削ってよい。
同日の追加指示「ランキング系（top20 など）も消して、ボタンも整理して」:
  働き方レポート（dtx18_stats）は回数を Dr.JOY の出退勤入力情報から数えていて DTX14 の top20 を読まない。
  人別3シートを作る旧マクロは「セカンド以下 5〜38行目」を決め打ちしていて、2行に割った雛形では行を読み漏らし、
  しかも ClearContents でアプリが書いた G/H（ID）まで消す → アプリが書く以上、残すと害になるので消す。
  行の高さ +7pt も同じ決め打ちで、押すたびに 7pt ずつ積み上がる（アプリが毎回 +7 済みの高さを書く）ので消す。

やること（Excel COM・VBProject を触るので「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要）:
  1) 標準モジュールを削除（REMOVE_MODULES）
  2) ThisWorkbook の Workbook_BeforeSave を丸ごと削除
  3) 残ったコードに無いマクロを指すボタン（図形の OnAction）を全シートから削除
  4) 人別シートの飾りの括弧（Right Brace）と I:K の回数表と A:H の文字色を消し、残ったボタン（当直情報転記）をデータの右（J列）へ寄せる
  5) シート「所属医師」「勤怠名簿」を削除
  何度走らせてもよい（無いものは飛ばす）。

残るマクロ: Module1 灰色のカッコ／Module2 カッコを付ける・外す／Module7 文字の整理／Module22 全角かっこ→半角／
  Module25・26 当番当直の赤枠／Module17 当直情報転記（Dr.JOY の出退勤入力情報へ）。

usage: python -X utf8 tools\\dtx14_trim_vba.py --in <DTX14.xlsm> [--out <出力.xlsm>]   （--out 省略＝上書き）
"""
import argparse
import os
import re
import shutil

REMOVE_MODULES = [
    # 所属医師・勤怠名簿まわり（20260922 第1弾）
    "Module4",    # 逆転記
    "Module5",    # 当直勤務表の転記ほか
    "Module13",   # 所属医師の漏れチェック
    "Module14",   # 所属医師の転記
    "Module15",   # 人別当直日にフルネーム記載
    "Module16",   # 人別オンコール日にフルネーム記載
    "Module19",   # 名前チェックとハイライト
    "Module20",   # 所属医師漏れ間違いチェックコンボ
    "Module21",   # 教員チェック
    # ランキングと人別3シートの旧作成マクロ（20260922 第2弾）
    "Module3",    # top20（CalculateFrequencyAndOutputTop20）
    "Module11",   # 科ごとの回数（CountDepartmentsInActiveSheet）
    "Module6",    # 人別オンコール日を作る（ExtractOncallData）
    "Module9",    # 人別当直日を作る（ExtractDutyData・IsFaculty）
    "Module18",   # 人別ファーストを作る
    "Module23",   # 昼夜行処理
    "Module10",   # 初期研未定の削除
    "Module24",   # 人別シート一括処理マクロ
    "Module12",   # CheckAndMarkDuty（無いシート「ここに当直勤務表を貼る」を読む死んだコード）
    "Module8",    # 行の高さを7ポイント増やす_非表示行除外（アプリが +7 済みの高さを書く）
]
REMOVE_SHEETS = ["所属医師", "勤怠名簿"]
JINBETSU_SHEETS = ["人別当直日", "人別オンコール日", "人別ファースト"]
RANKING_COLS = "I:K"
PARK = {"当直情報転記": ("人別当直日", "J", 2, 150, 48)}   # 残すボタン → (シート, 列, 行, 幅, 高さ)
PROC_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+([^\s(]+)", re.M | re.I)


def _procs(code_module):
    n = code_module.CountOfLines
    text = code_module.Lines(1, n) if n else ""
    return {m.group(1).lower() for m in PROC_RE.finditer(text)}


def _target(onaction):
    """OnAction『'ブック名'!マクロ名』→ マクロ名（小文字）。"""
    s = str(onaction or "").strip()
    if not s:
        return ""
    s = s.split("!")[-1].strip("'\" ")
    return s.split(".")[-1].lower()


def trim(src, out=None, log=print):
    import win32com.client
    src = os.path.abspath(src)
    out = os.path.abspath(out) if out else src
    if out != src:
        shutil.copy2(src, out)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(out, UpdateLinks=0, ReadOnly=False)
        vbp = wb.VBProject
        # 1) モジュール削除
        comps = {c.Name: c for c in vbp.VBComponents}
        for name in REMOVE_MODULES:
            if name in comps:
                gone = sorted(_procs(comps[name].CodeModule))
                vbp.VBComponents.Remove(comps[name])
                log(f"  モジュール削除: {name} {gone}")
        # 2) Workbook_BeforeSave 削除
        tw = vbp.VBComponents("ThisWorkbook").CodeModule
        try:
            start = tw.ProcStartLine("Workbook_BeforeSave", 0)
            count = tw.ProcCountLines("Workbook_BeforeSave", 0)
            tw.DeleteLines(start, count)
            log(f"  ThisWorkbook: Workbook_BeforeSave を削除（{count} 行）")
        except Exception:
            log("  ThisWorkbook: Workbook_BeforeSave は無い")
        # 3) 行き先の無いボタンを削除
        alive = set()
        for c in vbp.VBComponents:
            alive |= _procs(c.CodeModule)
        for ws in wb.Worksheets:
            for sh in list(ws.Shapes):
                try:
                    proc = _target(sh.OnAction)
                except Exception:
                    proc = ""
                if proc and proc not in alive:
                    log(f"  ボタン削除: {ws.Name} / {sh.Name} → {proc}")
                    sh.Delete()
        # 4) 人別シートの括弧を消し、残ったボタンを J 列へ
        for ws in wb.Worksheets:
            if ws.Name not in JINBETSU_SHEETS:
                continue
            for sh in list(ws.Shapes):
                if sh.Name.startswith("Right Brace"):
                    log(f"  図形削除: {ws.Name} / {sh.Name}")
                    sh.Delete()
            ws.Range("A:H").Font.ColorIndex = -4105      # 氏名の色分けは不要（管理者 20260922）
            if ws.Application.WorksheetFunction.CountA(ws.Range(RANKING_COLS)) > 0:
                ws.Range(RANKING_COLS).Clear()           # top20／科ごとの回数 が書いた表（マクロごと消したので古いまま残る）
                log(f"  {ws.Name}: {RANKING_COLS} の回数表を消去")
        for ws in wb.Worksheets:
            for sh in ws.Shapes:
                try:
                    proc = _target(sh.OnAction)
                except Exception:
                    proc = ""
                for name, (sheet, col, row, w, h) in PARK.items():
                    if proc == name.lower() and ws.Name == sheet:
                        sh.Left = ws.Columns(col).Left + 6       # 境目ちょうどだと左上のセルが隣になる
                        sh.Top = ws.Rows(row).Top + 3
                        sh.Width = w
                        sh.Height = h
                        log(f"  ボタン移動: {ws.Name} / {sh.Name}（{name}）→ {col}{row}")
        # 5) シート削除
        for name in REMOVE_SHEETS:
            try:
                wb.Worksheets(name).Delete()
                log(f"  シート削除: {name}")
            except Exception:
                pass
        left = sorted(c.Name for c in vbp.VBComponents if c.Type == 1)
        log(f"  残った標準モジュール: {left}")
        wb.Save()
        wb.Close(SaveChanges=False)
        log(f"保存: {out}")
    finally:
        xl.Quit()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_", required=True)
    ap.add_argument("--out")
    a = ap.parse_args()
    trim(a.in_, a.out)


if __name__ == "__main__":
    main()
