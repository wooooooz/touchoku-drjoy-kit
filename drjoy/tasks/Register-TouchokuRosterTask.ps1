# Register-TouchokuRosterTask.ps1  20261008
# Register (or remove) the scheduled task that runs Run-TouchokuRosterDaily.ps1
# every day at 3:30 as the current user, only while that user is logged on
# (the headless Chrome and Outlook need the user's session). No elevation.
#
#   pwsh -File .\Register-TouchokuRosterTask.ps1
#   pwsh -File .\Register-TouchokuRosterTask.ps1 -Unregister
#Requires -Version 7

param(
    [string]$TaskName = 'DrJoyAutomation-TouchokuRoster',
    [string]$At = '03:30',
    [string]$TaskPath = '\Dtx\',
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false
    Write-Host "Removed: $TaskName"
    return
}

$script = Join-Path $PSScriptRoot 'Run-TouchokuRosterDaily.ps1'
# Prefer the app execution alias: the Store install path has the version in it
# (...\Microsoft.PowerShell_7.6.2.0_x64__...\pwsh.exe) and disappears when PowerShell updates.
$pwsh = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe'
if (-not (Test-Path -LiteralPath $pwsh)) { $pwsh = (Get-Command pwsh).Source }
$user = "$env:USERDOMAIN\$env:USERNAME"

# conhost --headless: no console window flashes on screen (exit code is then always 0).
$action = New-ScheduledTaskAction -Execute (Join-Path $env:SystemRoot 'System32\conhost.exe') `
    -Argument "--headless `"$pwsh`" -NoProfile -ExecutionPolicy Bypass -File `"$script`"" `
    -WorkingDirectory (Split-Path $PSScriptRoot -Parent)
$trigger = New-ScheduledTaskTrigger -Daily -At $At
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
    -Description '当直アプリ名簿更新: Dr.JOY スタッフ一覧をエクスポート → roster_refresh.py → 問題があれば user003 へメール。C:\tools\touchoku\drjoy\tasks\Run-TouchokuRosterDaily.ps1' `
    -Force | Out-Null

$t = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath
$i = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath $TaskPath
Write-Host ("Registered: {0} | {1} | at {2} | {3} | next {4}" -f $t.TaskName, $t.State, $At, $t.Principal.LogonType, $i.NextRunTime)
