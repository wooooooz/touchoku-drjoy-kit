# TaskCommon.ps1  20261020
# Shared helpers for the scheduled jobs in this folder. Dot-source it:
#   . (Join-Path $PSScriptRoot 'TaskCommon.ps1')
#   Initialize-TaskLog 'gaibu-beacon'
#Requires -Version 7

# Child processes (node, python, pwsh) write UTF-8, but a scheduled task's
# console starts on the OEM code page (cp932). Without this, their output and
# the paths parsed from it ("OK: <path>") come back garbled (2026-09-28).
try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }
$OutputEncoding = [Text.UTF8Encoding]::new($false)
$script:TaskLogPath = $null

function Initialize-TaskLog([string]$Prefix, [string]$LogDir = (Join-Path $PSScriptRoot '..\logs')) {
    New-Item -ItemType Directory -Force $LogDir | Out-Null
    $script:TaskLogPath = Join-Path (Resolve-Path $LogDir).Path ("{0}_{1:yyyyMMdd}.log" -f $Prefix, (Get-Date))
    return $script:TaskLogPath
}

function Write-TaskLog([string]$Message) {
    $line = '{0:yyyy-MM-dd HH:mm:ss} | {1}' -f (Get-Date), $Message
    Write-Host $line
    if ($script:TaskLogPath) { Add-Content -LiteralPath $script:TaskLogPath -Value $line -Encoding utf8BOM }
}

# Run a native command, log every output line with a label, and return
# @{ Code; Lines }. Exceptions (missing exe or folder) come back as Code -1.
function Invoke-Logged([string]$Label, [scriptblock]$Command, [string]$WorkingDirectory) {
    $lines = @()
    $code = -1
    try {
        if ($WorkingDirectory) { Push-Location $WorkingDirectory }
        try {
            $env:PYTHONIOENCODING = 'utf-8'
            $lines = @(& $Command 2>&1 | ForEach-Object { "$_" })
            $code = $LASTEXITCODE
        } finally {
            if ($WorkingDirectory) { Pop-Location }
        }
    } catch {
        $lines = @("ERROR: $($_.Exception.Message)")
        $code = -1
    }
    $lines | ForEach-Object { Write-TaskLog "  ${Label}: $_" }
    return [pscustomobject]@{ Code = $code; Lines = $lines }
}

function Get-PythonPath {
    $p = (Get-Command python -ErrorAction SilentlyContinue).Source
    if (-not $p) { $p = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe' }
    return $p
}

# "OK: <path> (...)" from the Dr.JOY export scripts -> <path>
function Get-OkPath([string[]]$Lines) {
    $ok = $Lines | Where-Object { $_ -match 'OK: (.+?) \(' } | Select-Object -Last 1
    if ($ok -match 'OK: (.+?) \(') { return $Matches[1] }
    return $null
}

# One alert mail through Outlook COM.
# 20261020 (管理者): mail to 管理者 ends its subject with "※自動送信" like every other
# automatic mail; -Contact names who to ask,
# at the end of the body; the sent copy is not kept in Sent Items.
function Send-AlertMail([string]$To, [string]$Subject, [string]$Body, [switch]$NoMail, [string]$Contact) {
    $tag = '※自動送信'
    if ($To -match 'user003@|admin@' -and $Subject -notmatch [regex]::Escape($tag)) { $Subject = "$Subject $tag" }
    if ($Contact) { $Body = "$Body`r`n`r`n問い合わせ先: $Contact の担当" }
    if ($NoMail) {
        Write-TaskLog "mail (not sent, -NoMail): $Subject"
        Write-Host $Body
        return
    }
    try {
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $To
        $mail.Subject = $Subject
        $mail.Body = $Body
        $mail.DeleteAfterSubmit = $true
        $mail.Send()
        Write-TaskLog "mail sent to ${To}: $Subject"
    } catch {
        Write-TaskLog "mail FAILED: $($_.Exception.Message)"
    }
}
