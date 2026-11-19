# Run-TouchokuRosterDaily.ps1  20261020
# Daily roster update for the 当直勤務表 app.
# Run by the scheduled task
# \Dtx\DrJoyAutomation-TouchokuRoster at 3:30.
#
#   1. Export the Dr.JOY staff list (Export-DrJoyStaff.ps1 -Headless)
#      -> Downloads\staffs (N).csv. Tool 20 (7:30) and PortableEMR use it too,
#      so this step is always attempted. Also copied to the shared folder
#      00_DrJOY\Dr.JOY管理ファイル\staffs.csv (overwritten daily, 2026-09-28).
#   2. スタッフ所属履歴 -> 管理ツール\所属履歴\所属履歴_YYYY.xlsx (see below). Before step 3,
#      because roster_refresh.py reads it for the month folders (2026-09-27).
#   3. python -X utf8 tools\roster_refresh.py in C:\tools\touchoku
#      (doctors-only staff_YYYYMMDD.csv to the production share, and rebuild
#      the live month folders when the roster changed). Runs even when step 1
#      failed, using the newest staffs CSV already in Downloads.
#   Step 2 details: スタッフ所属履歴 (at0029, Excel, 本登録) for this year -> 管理ツール\所属履歴\所属履歴_YYYY.xlsx
#      (2026-10-09: moved out of 名簿\, which Domain Users can read; the file holds non-doctor staff codes,
#      login IDs and department history. 管理ツール\ is admin-only. The old 名簿\ copy is deleted after a good write),
#      and from 11/1 to 12/31 also next year's -> 所属履歴_(YYYY+1).xlsx (even if
#      empty). Overwritten every day, no date in the name; copied under a
#      temporary name first and then swapped in, so the app never reads a half
#      file. Also copied to 00_DrJOY\Dr.JOY管理ファイル\スタッフ所属履歴_YYYY年.xlsx
#      (Dr.JOY's own name, overwritten daily, 2026-09-28).
#      A failure here only adds a line to the alert; steps 1 and 3 are not
#      affected; roster_refresh.py then runs on the previous day's file
#      (管理者の決定 2026-09-27, requested by the 当直勤務表アプリ session).
#   4. Normal: log only. Any problem: one mail to user003@example.ac.jp.
#
#   pwsh -File .\Run-TouchokuRosterDaily.ps1                         normal run
#   pwsh -File .\Run-TouchokuRosterDaily.ps1 -NoMail                 print the mail instead of sending
#   pwsh -File .\Run-TouchokuRosterDaily.ps1 -DepartmentHistoryOnly  step 2 (and the report) only
#
# Exit codes: 0 ok / 1 something needs attention (a mail was sent)
#Requires -Version 7

param(
    [switch]$NoMail,
    [switch]$DepartmentHistoryOnly,
    [datetime]$Today = (Get-Date),
    [string]$To = 'user003@example.ac.jp',
    [string]$LogDir = (Join-Path $PSScriptRoot '..\logs'),
    [string]$TouchokuDir = 'C:\tools\touchoku',
    [string]$RosterDir = '\\fileserver\share$\00_DrJOY\当直勤務表\名簿',
    [string]$HistoryDir = '\\fileserver\share$\00_DrJOY\当直勤務表\管理ツール\所属履歴',
    [string]$ShareDir = '\\fileserver\share$\00_DrJOY\Dr.JOY管理ファイル',
    [string]$Downloads = (Join-Path $env:USERPROFILE 'Downloads')
)

# Child processes (node, python, pwsh) write UTF-8, but a scheduled task's
# console starts on the OEM code page (cp932). Without this, their output and
# the paths parsed from it ("OK: <path>") come back garbled (2026-09-28).
try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
$OutputEncoding = [Text.UTF8Encoding]::new($false)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
New-Item -ItemType Directory -Force $LogDir | Out-Null
$LogDir = (Resolve-Path $LogDir).Path
$logPath = Join-Path $LogDir ("touchoku-roster_{0:yyyyMMdd}.log" -f (Get-Date))

function Write-Log([string]$Message) {
    $line = '{0:yyyy-MM-dd HH:mm:ss} | {1}' -f (Get-Date), $Message
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line -Encoding utf8BOM
}

# Newest staffs*.csv with at least 500 lines, the same rule as the consumers.
function Get-NewestStaffs {
    $cp932 = [Text.Encoding]::GetEncoding(932)
    foreach ($f in Get-ChildItem -LiteralPath $Downloads -Filter 'staffs*.csv' -File | Sort-Object LastWriteTime -Descending) {
        $n = [IO.File]::ReadAllLines($f.FullName, $cp932).Count
        if ($n -ge 500) { return [pscustomobject]@{ File = $f; Lines = $n } }
    }
    return $null
}

# Copy to <Dir>\<Name> under a temporary name first, then swap it in, so a
# reader never sees a half-written file. Returns the final path.
function Copy-Swap([string]$Source, [string]$Dir, [string]$Name) {
    $target = Join-Path $Dir $Name
    # single extension (a double one like x.xlsx.tmp is flagged by the EDR as ransomware-like; 2026-10-09)
    $tmp = Join-Path $Dir ("~{0}_{1}_{2}.tmp" -f [IO.Path]::GetFileNameWithoutExtension($Name), [IO.Path]::GetExtension($Name).TrimStart('.'), $PID)
    Copy-Item -LiteralPath $Source -Destination $tmp -Force
    Move-Item -LiteralPath $tmp -Destination $target -Force
    return $target
}

function Update-DepartmentHistory {
    # スタッフ所属履歴 -> 管理ツール\所属履歴\所属履歴_YYYY.xlsx (step 2)
    $years = @($Today.Year)
    if ($Today.Month -ge 11) { $years += $Today.Year + 1 }
    $work = Join-Path $env:TEMP "DrJoyAutomation-depthist-$PID"
    foreach ($y in $years) {
        try {
            New-Item -ItemType Directory -Force $work | Out-Null
            $out = & pwsh -NoProfile -File (Join-Path $repo 'Export-DrJoyAt0029.ps1') -Type DepartmentHistory -Year $y -Headless -OutDir $work 2>&1 | ForEach-Object { "$_" }
            $code = $LASTEXITCODE
            $out | ForEach-Object { Write-Log "  depthist ${y}: $_" }
            $ok = $out | Where-Object { $_ -match 'OK: (.+?) \(' } | Select-Object -Last 1
            if ($code -ne 0 -or -not ($ok -match 'OK: (.+?) \(')) {
                throw "export failed (exit $code): " + (($out | Where-Object { $_ -match 'ERROR:|NG:|Not logged in' } | Select-Object -Last 2) -join ' / ')
            }
            $src = $Matches[1]
            $rows = 0
            try {
                # Row count for the log (openpyxl; informational only).
                $rows = & (Get-Command python).Source -X utf8 -c "import openpyxl,sys,warnings;warnings.filterwarnings('ignore');ws=openpyxl.load_workbook(sys.argv[1],read_only=True).worksheets[0];ws.reset_dimensions();print(sum(1 for _ in ws.iter_rows()))" $src 2>$null
            } catch { }
            New-Item -ItemType Directory -Force $HistoryDir | Out-Null
            $target = Copy-Swap $src $HistoryDir "所属履歴_$y.xlsx"
            $old = Join-Path $RosterDir "所属履歴_$y.xlsx"         # the former, widely readable place (2026-10-09)
            if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Force -ErrorAction SilentlyContinue; Write-Log "  removed old $old" }
            $summary.Add("所属履歴: $target（$rows 行、見出し込み）")
            # Also to the shared Dr.JOY management folder under Dr.JOY's own name (2026-09-28).
            try {
                $shared = Copy-Swap $src $ShareDir (Split-Path $src -Leaf)
                $summary.Add("共有フォルダ: $shared")
            } catch {
                $problems.Add("スタッフ所属履歴 $y 年を共有フォルダ（$ShareDir）に置けませんでした（前日の分がそのまま残っています）: $($_.Exception.Message)")
            }
            Remove-Item -LiteralPath $src -Force -ErrorAction SilentlyContinue
        } catch {
            $script:historyFailed = $true
            $msg = $_.Exception.Message
            if ($msg -match 'Not logged in') { $msg = 'Dr.JOY のセッション切れ' }
            $problems.Add("スタッフ所属履歴 $y 年を 管理ツール\所属履歴\所属履歴_$y.xlsx に置けませんでした（前日の分がそのまま残っています。職員一覧の更新には影響しません）: $msg")
        }
    }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

$problems = [System.Collections.Generic.List[string]]::new()
$summary = [System.Collections.Generic.List[string]]::new()
$exportCode = 0
$refreshCode = 0
$historyFailed = $false
Write-Log "start (PC=$env:COMPUTERNAME$(if ($DepartmentHistoryOnly) { ', department history only' }))"

if (-not $DepartmentHistoryOnly) {
    # --- 1. Dr.JOY export ---------------------------------------------------
    try {
        $previous = Get-NewestStaffs
        $exportOut = & pwsh -NoProfile -File (Join-Path $repo 'Export-DrJoyStaff.ps1') -Headless 2>&1 | ForEach-Object { "$_" }
        $exportCode = $LASTEXITCODE
    } catch {
        $exportOut = @("ERROR: $($_.Exception.Message)")
        $exportCode = -1
    }
    $exportOut | ForEach-Object { Write-Log "  export: $_" }
    if ($exportCode -eq 0) {
        $current = Get-NewestStaffs
        $diff = if ($previous -and $current) { $current.Lines - $previous.Lines } else { $null }
        $diffText = if ($null -eq $diff) { '' } else { '（前回 {0} 行から {1:+0;-0;±0}）' -f $previous.Lines, $diff }
        $summary.Add("エクスポート: $($current.File.Name) $($current.Lines) 行$diffText")
        # Also to the shared Dr.JOY management folder as staffs.csv, overwritten daily (2026-09-28).
        try {
            $shared = Copy-Swap $current.File.FullName $ShareDir 'staffs.csv'
            $summary.Add("共有フォルダ: $shared")
        } catch {
            $problems.Add("スタッフ一覧を共有フォルダ（$ShareDir）に staffs.csv として置けませんでした（前日の分がそのまま残っています。名簿の更新には影響しません）: $($_.Exception.Message)")
        }
    } else {
        $detail = ($exportOut | Where-Object { $_ -match 'ERROR:|NG:|Not logged in|timed out' } | Select-Object -Last 2) -join ' / '
        if ($exportOut -match 'Not logged in') {
            $problems.Add("Dr.JOY のセッションが切れていて、スタッフ一覧をエクスポートできませんでした。見える端末で pwsh -File C:\tools\touchoku\drjoy\Export-DrJoyStaff.ps1 を実行し、開いた Chrome でログインしてください（ログイン状態を保存(3ヶ月間) にチェック）。ログインするとそのままエクスポートまで進みます。")
        } elseif ($exportCode -eq 3) {
            $problems.Add("エクスポートした CSV が検査で不合格でした（rejected-staffs-*.csv として保存。下流には拾われません）。at0006 の絞り込みを確かめてください。$detail")
        } else {
            $problems.Add("スタッフ一覧のエクスポートに失敗しました（終了コード $exportCode）。$detail")
        }
        $summary.Add('エクスポート: 失敗。名簿は Downloads にある直近の staffs CSV で作りました。')
    }

    # --- 2. スタッフ所属履歴 (before roster_refresh, which reads it) ------------
    Update-DepartmentHistory

    # --- 3. roster_refresh.py -----------------------------------------------
    $python = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $python) { $python = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe' }
    try {
        Push-Location $TouchokuDir
        try {
            $env:PYTHONIOENCODING = 'utf-8'
            $refreshOut = & $python -X utf8 'tools\roster_refresh.py' 2>&1 | ForEach-Object { "$_" }
            $refreshCode = $LASTEXITCODE
        } finally {
            Pop-Location
        }
    } catch {
        $refreshOut = @("ERROR: $($_.Exception.Message)")
        $refreshCode = -1
    }
    $refreshOut | ForEach-Object { Write-Log "  roster_refresh: $_" }
    if ($refreshCode -eq 0) {
        $summary.Add((($refreshOut | Where-Object { $_ -match '^(名簿|作り直し|注意|活きている)' }) -join ' / '))
    } else {
        $problems.Add("roster_refresh.py が失敗しました（終了コード $refreshCode）。本番共有 \\fileserver\share$\00_DrJOY\当直勤務表 に届かないときは職場ネットワークを確かめてください。出力: " + (($refreshOut | Select-Object -Last 3) -join ' / '))
    }
} else {
    Update-DepartmentHistory
}

$summary | ForEach-Object { Write-Log "summary: $_" }

# --- 4. Report --------------------------------------------------------------
if ($problems.Count -eq 0) {
    Write-Log 'done: ok'
    exit 0
}

$what = if ($exportCode -ne 0) { 'Dr.JOY エクスポート失敗' } elseif ($refreshCode -ne 0) { '名簿の配置に失敗' } else { '所属履歴の取得に失敗' }
$subject = "【当直名簿更新】要対応: $what ※自動送信"
$body = @(
    '毎日3:30の当直アプリ名簿更新（タスク \Dtx\DrJoyAutomation-TouchokuRoster、職場PC）で問題がありました。'
    ''
    $problems | ForEach-Object { "- $_" }
    ''
    '結果:'
    $summary | ForEach-Object { "- $_" }
    ''
    "ログ: $logPath"
    '手で走らせ直すときは: pwsh -File C:\tools\touchoku\drjoy\tasks\Run-TouchokuRosterDaily.ps1（所属履歴だけなら -DepartmentHistoryOnly）'
    ''
    '問い合わせ先: 当直勤務表の担当'
) -join "`r`n"

if ($NoMail) {
    Write-Log "mail (not sent, -NoMail): $subject"
    Write-Host $body
} else {
    try {
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $To
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.DeleteAfterSubmit = $true   # 20261020 管理者: do not keep these in Sent Items
        $mail.Send()
        Write-Log "mail sent to ${To}: $subject"
    } catch {
        Write-Log "mail FAILED: $($_.Exception.Message)"
    }
}
Write-Log 'done: needs attention'
exit 1
