# -*- coding: utf-8 -*-
r"""dtx14_autoclose_vba.py — 正本 DTX14 に「14:55 に保存して閉じる」マクロを入れる。20261025

20261025 管理者: 正本からの印刷・PDF への書き出しも止める（掲示板にチャットへのリンクの無い、古い PDF が貼られていた）。
  Workbook_BeforePrint は印刷と PDF への書き出し（ファイル > エクスポート・ExportAsFixedFormat）の両方で通るので、ここで取り消し、
  公開用の PDF（公開用\R8年N月 病院当直勤務表.pdf）を案内する。止めないのは、見えない Excel（自動処理）と管理者（PRINT_OK_USERS）。
  マクロを有効にせずに開いたとき・ブラウザの Excel では止まらない（IRM は 15時の処理が openpyxl で読めなくなるので見送った）。

管理者 2026-10-09: 15時の自動書き込み（dtx14_daily_sync）は、ほかの PC（事務担当の PC）で正本が Excel で開かれていても
気づけない（OneDrive は ~$ を同期しない）。開いていた人があとで保存すると「…-PC名.xlsm」の写しになり、手直しが正本に入らない。
→ 正本そのものに、次のマクロを入れる（「保存して閉じる」でよい。誤った保存は SharePoint のバージョン履歴と _上書き履歴 で戻せる）。
  - 開いたとき（人が Excel で開いたときだけ。読み取り専用・見えない Excel・別のマクロからの読み取り専用の open では何もしない）:
      14:55〜15:15 なら「15時の自動書き込みの時間」と知らせて、保存せずに閉じる。
      それ以外なら、次の 14:55 に「保存して閉じる」を予約する（Application.OnTime）。
  - 閉じるとき: 予約を取り消す（取り消さないと、14:55 に Excel が正本を開き直してマクロを走らせる）。
  管理者の PC の自動処理（15時の反映・PDF・④・new_month・toban_fix）は AutomationSecurity=3 / EnableEvents=False で開くので、このマクロは動かない。
  20261010 管理者: 時間を 14:50 閉じる・14:30〜15:30 から、14:55 閉じる・14:55〜15:15 に縮めた。15時の反映はふだん2分かからない
  （9/23〜10/9 の毎日の回で 40秒〜1分32秒）。14:55 は、閉じるときの保存が OneDrive で管理者の PC に届くまでの余裕、
  15:15 は、反映が書いた正本がほかの PC に届くまでの余裕。
入れるのは ★の正本だけ（【診療部科名】の雛形には入れない＝各部署の写しが勝手に閉じないように）。翌月の正本は new_month が前月の正本から
作るので、一度入れれば以後の月にも入る。古い版が入っていれば、今の版に入れ替える（VERSION の印で見分ける）。

書き方は15時の反映と同じ: _上書き履歴\ へ控え → ローカルの写しに Excel COM（マクロは走らせない）で書く → 読み戻してマクロとシートの値が
変わっていないことを確かめる → 書いている間に正本が変わっていなければ一時名で写して差し替える。14:55〜15:15 は断る。
VBProject を触るので「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要。

usage: python -X utf8 tools\dtx14_autoclose_vba.py --file <★R8.N月用 DTX14….xlsm> [--dry-run]
"""
import argparse
import datetime as dt
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

MODULE = "modTouchokuAutoClose"
MARK = "Touchoku_ScheduleClose"
VERSION = "TouchokuAutoCloseVer 20261025a"   # 印が変われば入れ替える（「〜」は VBA の CP932 に無く「?」に化けた）
# 印刷・PDF への書き出しを止めてよい人の Windows のログイン名（管理者 20261025: 管理者は正本から刷れる）
PRINT_OK_USERS = ("user003",)
CLOSE_AT = "14:55:00"
QUIET_FROM, QUIET_TO = "14:55:00", "15:15:00"

MODULE_CODE = f'''Option Explicit
' 当直勤務表アプリの15時の自動書き込みと重ならないように、{CLOSE_AT[:5]} に保存して閉じる。20261010（管理者）
' 開いたとき ThisWorkbook.Workbook_Open が予約し、閉じるとき Workbook_BeforeClose が取り消す。
' {VERSION}
Public TouchokuNextClose As Date

Private Function TouchokuProc(ByVal name As String) As String
    TouchokuProc = "'" & ThisWorkbook.Name & "'!" & name
End Function

Public Sub Touchoku_ScheduleClose()
    Dim d As Date
    Touchoku_CancelClose
    d = Date + TimeValue("{CLOSE_AT}")
    If Now >= d Then d = d + 1
    TouchokuNextClose = d
    On Error Resume Next
    Application.OnTime TouchokuNextClose, TouchokuProc("Touchoku_SaveAndClose")
End Sub

Public Sub Touchoku_CancelClose()
    On Error Resume Next
    If TouchokuNextClose <> 0 Then
        Application.OnTime TouchokuNextClose, TouchokuProc("Touchoku_SaveAndClose"), , False
    End If
    TouchokuNextClose = 0
End Sub

Public Sub Touchoku_SaveAndClose()
    TouchokuNextClose = 0
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    ThisWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub

Public Sub Touchoku_CloseNoSave()
    On Error Resume Next
    ThisWorkbook.Close SaveChanges:=False
End Sub

Public Function Touchoku_IsQuiet() As Boolean
    Touchoku_IsQuiet = (Time >= TimeValue("{QUIET_FROM}") And Time < TimeValue("{QUIET_TO}"))
End Function

Public Function Touchoku_ByHand() As Boolean
    ' 人が Excel で開いたときだけ（読み取り専用・見えない Excel・別のマクロからの読み取りでは何もしない）
    Touchoku_ByHand = (Not ThisWorkbook.ReadOnly) And Application.Visible And Application.UserControl
End Function

Public Function Touchoku_PrintAllowed() As Boolean
    ' 印刷・PDF への書き出しを通すとき: 自動処理（見えない Excel）か、管理者（20261025 管理者）
    Dim u As String
    Touchoku_PrintAllowed = True
    If Not (Application.Visible And Application.UserControl) Then Exit Function
    u = LCase(Environ("USERNAME"))
    {" ".join(f'If u = "{x}" Then Exit Function' for x in PRINT_OK_USERS)}
    Touchoku_PrintAllowed = False
End Function

Public Sub Touchoku_PrintBlocked()
    ' 正本からの印刷・PDF への書き出しは止め、公開用の PDF を案内する（20261025 管理者: 掲示板にリンクの無い PDF が貼られていた）
    Dim d As Variant, ym As String
    ym = ""
    On Error Resume Next
    d = ThisWorkbook.Worksheets(1).Range("A1").Value
    If IsDate(d) Then ym = "R" & (Year(d) - 2018) & "年" & Month(d) & "月 "
    On Error GoTo 0
    MsgBox "この当直勤務表（Excel の正本）からは、印刷と PDF への書き出しはできません。" & vbCrLf & vbCrLf & _
           "印刷・掲示板・配布には、公開用の PDF を使ってください。" & vbCrLf & _
           "  場所: 当直表チーム > 公開用 > " & ym & "病院当直勤務表.pdf" & vbCrLf & _
           "  （毎日15時に最新に書き直されます。氏名から Teams のチャットに飛べます）", vbExclamation, "当直勤務表"
End Sub
'''

THISWB_CODE = f'''' {MARK}: 当直勤務表アプリの15時の自動書き込みと重ならないように（20261010 管理者）{VERSION}
Private Sub Workbook_Open()
    If Not Touchoku_ByHand() Then Exit Sub
    If Touchoku_IsQuiet() Then
        MsgBox "いまは15時の自動書き込み（当直勤務表アプリ）の時間です（{QUIET_FROM[:5]} から {QUIET_TO[:5]} まで）。" & vbCrLf & _
               "{QUIET_TO[:5]} を過ぎてから開いてください。見るだけなら公開用の PDF を開いてください。" & vbCrLf & vbCrLf & _
               "このファイルは保存せずに閉じます。", vbInformation, "当直勤務表"
        Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!Touchoku_CloseNoSave"
    Else
        Touchoku_ScheduleClose
    End If
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    ' 印刷と PDF への書き出し（ファイル > エクスポート・名前を付けて保存の PDF）の両方でここを通る（20261025）
    If Touchoku_PrintAllowed() Then Exit Sub
    Cancel = True
    Touchoku_PrintBlocked
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Touchoku_CancelClose
End Sub
'''


def _hm(s):
    h, m = s.split(":")[:2]
    return dt.time(int(h), int(m))


def _quiet(now=None):
    t = (now or dt.datetime.now()).time()
    return _hm(QUIET_FROM) <= t < _hm(QUIET_TO)


def _code(cm):
    n = cm.CountOfLines
    return cm.Lines(1, n) if n else ""


def _our_block(cm):
    """ThisWorkbook の中の、このツールが入れた範囲（印の行〜Workbook_BeforeClose の End Sub）。(開始行, 行数) か None。"""
    n = cm.CountOfLines
    lines = [cm.Lines(i, 1) for i in range(1, n + 1)]
    start = next((i for i, s in enumerate(lines) if s.lstrip().startswith(f"' {MARK}:")), None)
    if start is None:
        return None
    bc = next((i for i in range(start, n) if lines[i].strip().startswith("Private Sub Workbook_BeforeClose")), None)
    if bc is None:
        return None
    end = next((i for i in range(bc, n) if lines[i].strip() == "End Sub"), None)
    if end is None:
        return None
    return start + 1, end - start + 1


def install(path, dry=False, log=print):
    import win32com.client
    import dtx14_fill as F
    if _quiet():
        raise SystemExit(f"{QUIET_FROM[:5]}〜{QUIET_TO[:5]} は正本に書きません（15時の反映と重なるため）")
    path = os.path.abspath(path)
    base = os.path.basename(path)
    if not base.startswith("★"):
        raise SystemExit(f"★の正本だけに入れます（【診療部科名】の雛形などには入れない）: {base}")
    if F.locked(path) if hasattr(F, "locked") else os.path.exists(os.path.join(os.path.dirname(path), "~$" + base)):
        raise SystemExit("正本が Excel で開かれています。閉じてからやり直してください")
    st0 = os.stat(path)
    work_dir = tempfile.mkdtemp(prefix="dtx14_vba_")
    work = os.path.join(work_dir, "_work.xlsm")
    shutil.copy2(path, work)
    what = "追加"
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.AutomationSecurity = 3
        wb = xl.Workbooks.Open(work, UpdateLinks=0, ReadOnly=False)
        vbp = wb.VBProject
        comps = {c.Name: c for c in vbp.VBComponents}
        tw = vbp.VBComponents("ThisWorkbook").CodeModule
        twc = _code(tw)
        mod_code = _code(comps[MODULE].CodeModule) if MODULE in comps else ""
        if MARK in twc and VERSION in twc and VERSION in mod_code:
            wb.Close(SaveChanges=False)
            log(f"今の版が入っている（変更なし）: {base}")
            return False
        if MARK in twc:                               # 古い版 → 入れ替える
            blk = _our_block(tw)
            if not blk:
                raise SystemExit("ThisWorkbook の古いマクロの範囲が見分けられない。手で直してください（自動では入れ替えない）")
            tw.DeleteLines(*blk)
            at = blk[0]
            what = "更新"
        else:
            for proc in ("Workbook_Open", "Workbook_BeforeClose", "Workbook_BeforePrint"):
                if f"Sub {proc}(" in twc:
                    raise SystemExit(f"ThisWorkbook に {proc} がもうあります。手で合わせてください（自動では入れない）")
            at = tw.CountOfLines + 1
            if tw.CountOfLines:
                tw.InsertLines(at, "")
                at += 1
        before_mods = sorted(c.Name for c in vbp.VBComponents)
        if MODULE in comps:
            vbp.VBComponents.Remove(comps[MODULE])
        m = vbp.VBComponents.Add(1)                # vbext_ct_StdModule
        m.Name = MODULE
        m.CodeModule.AddFromString(MODULE_CODE)
        tw.InsertLines(at, THISWB_CODE.rstrip("\n"))
        if dry:
            wb.Close(SaveChanges=False)
            log(f"(dry-run) {what}できることを確かめた: {base}（モジュール {before_mods}）")
            return False
        wb.Save()
        wb.Close(SaveChanges=False)
        # 読み戻し
        wb = xl.Workbooks.Open(work, UpdateLinks=0, ReadOnly=True)
        names = [c.Name for c in wb.VBProject.VBComponents]
        twc2 = _code(wb.VBProject.VBComponents("ThisWorkbook").CodeModule)
        ok = (MODULE in names and VERSION in _code(wb.VBProject.VBComponents(MODULE).CodeModule)
              and VERSION in twc2 and twc2.count("Sub Workbook_Open(") == 1 and twc2.count("Sub Workbook_BeforeClose(") == 1
              and twc2.count("Sub Workbook_BeforePrint(") == 1)
        wb.Close(SaveChanges=False)
        if not ok:
            raise SystemExit("読み戻しでマクロが確かめられない（正本は変えていない）")
    finally:
        xl.Quit()
    # シートの値が変わっていないか（openpyxl で両方を読んで比べる）
    import openpyxl
    a = openpyxl.load_workbook(path, data_only=True, read_only=True, keep_vba=False)
    b = openpyxl.load_workbook(work, data_only=True, read_only=True, keep_vba=False)
    for name in a.sheetnames:
        ra = [tuple(r) for r in a[name].iter_rows(values_only=True)]
        rb = [tuple(r) for r in b[name].iter_rows(values_only=True)]
        if ra != rb:
            raise SystemExit(f"シート「{name}」の値が変わった（正本は変えていない）")
    a.close()
    b.close()
    st1 = os.stat(path)
    if (st1.st_mtime, st1.st_size) != (st0.st_mtime, st0.st_size):
        raise SystemExit("書いている間に正本が変わった（正本は変えていない）。あとでやり直してください")
    if _quiet():
        raise SystemExit(f"書いている間に {QUIET_FROM[:5]} を回った（正本は変えていない）")
    bak_dir = os.path.join(os.path.dirname(path), "_上書き履歴")
    os.makedirs(bak_dir, exist_ok=True)
    stem, ext = os.path.splitext(base)
    bak = os.path.join(bak_dir, f"{stem}_{dt.datetime.now():%Y%m%d_%H%M}自動で閉じるマクロ{what}前{ext}")
    shutil.copy2(path, bak)
    log(f"控え: {bak}")
    F.publish(work, path, log=log)
    shutil.rmtree(work_dir, ignore_errors=True)
    log(f"{what}した: {base}（{CLOSE_AT[:5]} に保存して閉じる・{QUIET_FROM[:5]}〜{QUIET_TO[:5]} は開いたらすぐ閉じる・"
        "印刷と PDF への書き出しは止めて公開用の PDF を案内）")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True, action="append")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    for f in a.file:
        install(f, a.dry_run)


if __name__ == "__main__":
    main()
