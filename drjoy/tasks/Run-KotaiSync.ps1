# Run-KotaiSync.ps1  20261103
# Daily 15:30 当直交代 -> Dr.JOY 勤務予定 sync.
# Moved from 17:30 on 2026-10-09: 当直勤務表アプリ's
# daily sync runs at 15:00 and takes about 2 minutes. Steps (the browser parts
# are done by the scripts in drjoy/):
#
#   0. drjoy_kotai_sync.py check          exit 3 = nothing new, or only on-call swaps (recorded) -> done (most days)
#   2. before: 勤務予定 export per month (at0029) + 申請一覧 (at0022 API) -> RUN\前
#   3. drjoy_kotai_sync.py plan --mail    IMPORT=<xlsx> (empty = nothing to import)
#   4. import IMPORT through at0069       (drjoy-at0069-import.js)
#   5. after: the same exports again      -> RUN\後
#   6. drjoy_kotai_sync.py verify --mail  exit 2 = requests decreased (it mails the alert)
#   7. drjoy_kotai_sync.py commit
#
# Safety gate: until tasks\kotai-sync.armed exists, the run stops before step 4
# and mails the plan (the at0069 import has not been watched live yet). Resume
# such a run by hand with -ImportRun <RUN folder>.
#
#   pwsh -File .\Run-KotaiSync.ps1
#   pwsh -File .\Run-KotaiSync.ps1 -ImportRun C:\Users\user003\Desktop\DrJOY交代自動登録\run_YYYYMMDD_HHMM
#   pwsh -File .\Run-KotaiSync.ps1 -CheckOnly        (step 0 only)
#
# Exit codes: 0 done (or nothing to do) / 1 needs attention (a mail was sent)
#Requires -Version 7

param(
    [string]$ImportRun,
    [switch]$CheckOnly,
    [string]$MailTo = 'user003@example.ac.jp',
    [string]$AllowanceDir = 'C:\tools\touchoku\drjoy_import'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'TaskCommon.ps1')
$repo = Split-Path $PSScriptRoot -Parent
$logPath = Initialize-TaskLog 'kotai-sync'
$armed = Test-Path (Join-Path $PSScriptRoot 'kotai-sync.armed')
$py = Get-PythonPath
Write-TaskLog "start (PC=$env:COMPUTERNAME, armed=$armed$(if ($ImportRun) { ", resume $ImportRun" }))"

function Stop-WithAlert([string]$Subject, [string[]]$Lines) {
    $body = (@('当直交代の Dr.JOY 自動登録（タスク \Dtx\DrJoyAutomation-KotaiSync、職場PC、毎日15:30）') + '' + $Lines + '' +
        "ログ: $logPath") -join "`r`n"
    Send-AlertMail -To $MailTo -Subject "【当直交代の自動登録】$Subject" -Body $body -Contact '手当計算'
    Write-TaskLog 'done: needs attention'
    exit 1
}

function Invoke-Kotai([string[]]$KotaiArgs) {
    Invoke-Logged 'kotai' { & $py -X utf8 tools\drjoy_kotai_sync.py @KotaiArgs } -WorkingDirectory $AllowanceDir
}

function Get-Tagged([string[]]$Lines, [string]$Tag) {
    $l = $Lines | Where-Object { $_ -match "^\s*(\[[^\]]*\]\s*)?$Tag=" } | Select-Object -Last 1
    if ($l -match "$Tag=(.*)$") { return $Matches[1].Trim() }
    return $null
}

function Test-LoggedOut([string[]]$Lines) { return [bool]($Lines -match 'Not logged in') }

# Exports into $Dir: 勤務予定 per month, then 申請一覧 for all months.
# Returns @{ Plans; Reqs } or stops with an alert.
function Get-Materials([string]$Dir, [string[]]$Months, [string]$Label) {
    New-Item -ItemType Directory -Force $Dir | Out-Null
    $plans = @()
    foreach ($m in $Months) {
        $r = Invoke-Logged "wp $m" { pwsh -NoProfile -File (Join-Path $repo 'Export-DrJoyAt0029.ps1') -Type WorkingPlan -Month $m -Headless -OutDir $Dir }
        $p = Get-OkPath $r.Lines
        if ($r.Code -ne 0 -or -not $p) {
            if (Test-LoggedOut $r.Lines) { Stop-WithAlert 'Dr.JOY ログイン切れ' @('Dr.JOY のセッションが切れていて止まりました（記録はしていないので、ログイン後の次の回にやり直します）。', '見える端末で pwsh -File C:\tools\touchoku\drjoy\Export-DrJoyStaff.ps1 を実行し、開いた Chrome でログインしてください。') }
            Stop-WithAlert "$Label の勤務予定エクスポートに失敗" @("$m の勤務予定を取れませんでした（終了コード $($r.Code)）。記録はしていないので、次の回にやり直します。", (($r.Lines | Where-Object { $_ -match 'ERROR:|NG:' }) -join ' / '))
        }
        $plans += $p
    }
    $reqArgs = @('drjoy-request-list.js') + ($Months | ForEach-Object { '--month', $_ }) + @('--out-dir', $Dir, '--headless')
    $r = Invoke-Logged 'requests' { & node @reqArgs } -WorkingDirectory $repo
    $reqs = @($r.Lines | Where-Object { $_ -match 'OK: (.+?) \(' } | ForEach-Object { if ($_ -match 'OK: (.+?) \(') { $Matches[1] } })
    if ($r.Code -ne 0 -or $reqs.Count -ne $Months.Count) {
        if (Test-LoggedOut $r.Lines) { Stop-WithAlert 'Dr.JOY ログイン切れ' @('Dr.JOY のセッションが切れていて止まりました。見える端末で pwsh -File C:\tools\touchoku\drjoy\Export-DrJoyStaff.ps1 を実行してログインしてください。') }
        Stop-WithAlert "$Label の申請一覧の取得に失敗" @("申請一覧を取れませんでした（終了コード $($r.Code)）。記録はしていないので、次の回にやり直します。", (($r.Lines | Where-Object { $_ -match 'ERROR:|注意:' }) -join ' / '))
    }
    return [pscustomobject]@{ Plans = $plans; Reqs = $reqs }
}

# ------------------------------------------------------------------ 0 + 2 + 3
if ($ImportRun) {
    $run = $ImportRun
    $pending = Get-Content -LiteralPath (Join-Path $run 'pending.json') -Raw -Encoding utf8 | ConvertFrom-Json
    $planJson = Get-Content -LiteralPath (Join-Path $run 'plan.json') -Raw -Encoding utf8 | ConvertFrom-Json
    $months = @($pending.plan_months)
    $import = $planJson.import
    Write-TaskLog "resume: months $($months -join ', ') / IMPORT=$import"
} else {
    $check = Invoke-Kotai @('check')
    if ($check.Code -eq 3) { Write-TaskLog 'done: 入れるものなし（新しい交代記録なし、またはオンコール同士の交代だけ）'; exit 0 }
    if ($check.Code -ne 0) { Stop-WithAlert 'check が止まりました' @("drjoy_kotai_sync.py check が終了コード $($check.Code) で止まりました。", ($check.Lines | Select-Object -Last 5)) }
    $run = Get-Tagged $check.Lines 'RUN'
    $monthLine = $check.Lines | Where-Object { $_ -match '勤務予定エクスポートを取る月:' } | Select-Object -Last 1
    # 20261103 keep only YYYY-MM: an empty list ("月: ") used to slip through as " " and break the export
    $months = if ($monthLine -match '勤務予定エクスポートを取る月:(.*)$') { @($Matches[1] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d{4}-\d{2}$' }) } else { @() }
    if (-not $run -or $months.Count -eq 0) { Stop-WithAlert 'check の出力が読めません' @('RUN= か「勤務予定エクスポートを取る月」が出力にありません。', ($check.Lines | Select-Object -Last 8)) }
    if ($CheckOnly) { Write-TaskLog "check only: RUN=$run months=$($months -join ', ')"; exit 0 }

    $before = Get-Materials (Join-Path $run '前') $months '取り込み前'
    $planArgs = @('plan', '--run', $run) + ($before.Plans | ForEach-Object { '--plan', $_ }) + ($before.Reqs | ForEach-Object { '--req', $_ }) + '--mail'
    $plan = Invoke-Kotai $planArgs
    switch ($plan.Code) {
        0 { }
        5 { Write-TaskLog 'plan: 二重割当（drjoy_kotai_sync.py が警報メール済み）。取り込まずに止めます。'; Write-TaskLog 'done: needs attention'; exit 1 }
        4 { Stop-WithAlert '取り込む行が60行を超えました' @("取り込まずに止めました。実行フォルダ: $run", ($plan.Lines | Where-Object { $_ -match '取り込む行|上限' })) }
        default { Stop-WithAlert 'plan が止まりました' @("drjoy_kotai_sync.py plan が終了コード $($plan.Code) で止まりました。記録はしていません。実行フォルダ: $run", ($plan.Lines | Select-Object -Last 6)) }
    }
    $import = Get-Tagged $plan.Lines 'IMPORT'
    if (-not $import) {
        Write-TaskLog '取り込むものはありません（一致していた・全部見送り）。記録します。'
        $c = Invoke-Kotai @('commit', '--run', $run)
        if ($c.Code -ne 0) { Stop-WithAlert 'commit が止まりました' @("実行フォルダ: $run", ($c.Lines | Select-Object -Last 4)) }
        Write-TaskLog 'done: ok (nothing to import)'
        exit 0
    }
    if (-not $armed) {
        $rows = $plan.Lines | Where-Object { $_ -match '取り込む行|^\s*(\[[^\]]*\]\s*)?\s+\d{4}-\d{2}-\d{2} |見送り' }
        Stop-WithAlert '取り込み待ち（自動取り込みはまだ無効）' (@(
            '取り込むファイルができましたが、at0069 への自動取り込みはまだ一度も立ち会って試していないので、ここで止めました（記録はしていません）。'
            "取り込むファイル: $import"
            '中身:') + ($rows | ForEach-Object { "  $_" }) + @(
            ''
            '続けるには、職場PCの表示されるターミナルで次を実行してください（取り込み→後の材料→突き合わせ→記録まで進みます）:'
            "  pwsh -File C:\tools\touchoku\drjoy\tasks\Run-KotaiSync.ps1 -ImportRun `"$run`""
            '問題なく終わったら、C:\tools\touchoku\drjoy\tasks\kotai-sync.armed を置くと次からは全部自動になります。'))
    }
}

# ------------------------------------------------------------------ 4 import
$imp = Invoke-Logged 'import' { & node drjoy-at0069-import.js --file $import --headless } -WorkingDirectory $repo
$errors = Get-Tagged $imp.Lines 'ERRORS'
switch ($imp.Code) {
    0 { }
    5 { Write-TaskLog "取り込みで失敗した行があります（エラー報告: $(if ($errors) { $errors } else { '自動では取れず' })）" }
    default {
        if (Test-LoggedOut $imp.Lines) { Stop-WithAlert 'Dr.JOY ログイン切れ（取り込み前）' @("取り込む前に止まりました（記録はしていません）。ログインしてから -ImportRun `"$run`" で続けてください。") }
        Stop-WithAlert '取り込みが確かめられません' @(
            "at0069 への取り込みが終わったか確かめられませんでした（終了コード $($imp.Code)）。記録はしていません。"
            'at0069 の取り込み履歴を見て、取り込まれていなければ次で続けてください（取り込まれていたら、翌日の回が差分なしで記録します）:'
            "  pwsh -File C:\tools\touchoku\drjoy\tasks\Run-KotaiSync.ps1 -ImportRun `"$run`""
            (($imp.Lines | Where-Object { $_ -match 'ERROR:|確認' }) -join ' / '))
    }
}

# ------------------------------------------------------------------ 5 + 6
$after = Get-Materials (Join-Path $run '後') $months '取り込み後'
$verifyArgs = @('verify', '--run', $run) + ($after.Plans | ForEach-Object { '--plan-after', $_ }) + ($after.Reqs | ForEach-Object { '--req-after', $_ })
if ($errors) { $verifyArgs += @('--errors', $errors) }
$verifyArgs += '--mail'
$verify = Invoke-Kotai $verifyArgs
$notes = @()
if ($verify.Code -eq 2) { $notes += '申請が減った人×日があります（drjoy_kotai_sync.py が【警報】メール済み）。' }
elseif ($verify.Code -ne 0) { Stop-WithAlert 'verify が止まりました' @("drjoy_kotai_sync.py verify が終了コード $($verify.Code) で止まりました。取り込みは済んでいます。記録はしていません。実行フォルダ: $run", ($verify.Lines | Select-Object -Last 6)) }
$sumLine = $verify.Lines | Where-Object { $_ -match '入れた形と違う\s*(\d+)\s*/\s*拒否\s*(\d+)' } | Select-Object -Last 1
if ($sumLine -match '入れた形と違う\s*(\d+)\s*/\s*拒否\s*(\d+)' -and ([int]$Matches[1] + [int]$Matches[2]) -gt 0) {
    $notes += "入れた形と違う $($Matches[1]) 件・拒否 $($Matches[2]) 件（見送り一覧.csv に追記済み）。"
}
if ($imp.Code -eq 5 -and -not $errors) { $notes += 'at0069 で失敗した行がありましたが、エラー報告を自動で取れませんでした（at0069 の該当行から確かめてください）。' }

# ------------------------------------------------------------------ 7 commit (also after alerts)
$c = Invoke-Kotai @('commit', '--run', $run)
if ($c.Code -ne 0) { Stop-WithAlert 'commit が止まりました' (@("取り込みと突き合わせは済んでいます。実行フォルダ: $run") + ($c.Lines | Select-Object -Last 4)) }

$summary = ($plan.Lines + $verify.Lines | Where-Object { $_ -match '取り込む行|申請が減った人×日' }) -join ' / '
Write-TaskLog "summary: $summary"
if ($notes.Count -gt 0 -and -not ($verify.Code -eq 2 -and $notes.Count -eq 1)) {
    Send-AlertMail -To $MailTo -Subject '【当直交代の自動登録】確認してください' -Body ((@('当直交代の Dr.JOY 自動登録で、確認が要る結果がありました。', '') + ($notes | ForEach-Object { "- $_" }) + '' + "実行フォルダ: $run" + "ログ: $logPath") -join "`r`n") -Contact '手当計算'
}
Write-TaskLog "done: $(if ($notes.Count) { 'needs attention' } else { 'ok' })"
exit $(if ($notes.Count) { 1 } else { 0 })
