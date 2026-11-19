# Register-DrJoyTask.ps1  20261008
# Register (or remove) a scheduled task that runs one of the job scripts in
# this folder as the current user, only while that user is logged on (the
# headless Chrome and Outlook need the user's session). No elevation.
#
#   pwsh -File .\Register-DrJoyTask.ps1 -TaskName DrJoyAutomation-GaibuBeacon -Script Run-GaibuBeaconWatch.ps1 -At 08:45
#   pwsh -File .\Register-DrJoyTask.ps1 -TaskName DrJoyAutomation-PhoneWatch -Script Run-PhoneTableWatch.ps1 -At 09:00 -DaysOfWeek Monday
#   pwsh -File .\Register-DrJoyTask.ps1 -TaskName DrJoyAutomation-GaibuBeacon -Unregister
#Requires -Version 7

param(
    [Parameter(Mandatory)][string]$TaskName,
    [string]$Script,
    [string]$At,
    [System.DayOfWeek[]]$DaysOfWeek,
    [int]$TimeLimitMinutes = 30,
    [string]$Description = '',
    [string]$TaskPath = '\Dtx\',
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false
    Write-Host "Removed: $TaskName"
    return
}
if (-not $Script -or -not $At) { throw '-Script and -At are required.' }

$scriptPath = Join-Path $PSScriptRoot $Script
if (-not (Test-Path $scriptPath)) { throw "Script not found: $scriptPath" }
# Prefer the app execution alias: the Store install path has the version in it
# (...\Microsoft.PowerShell_7.6.2.0_x64__...\pwsh.exe) and disappears when PowerShell updates.
$pwsh = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe'
if (-not (Test-Path -LiteralPath $pwsh)) { $pwsh = (Get-Command pwsh).Source }
$user = "$env:USERDOMAIN\$env:USERNAME"

# conhost --headless: no console window flashes on screen (exit code is then always 0).
$action = New-ScheduledTaskAction -Execute (Join-Path $env:SystemRoot 'System32\conhost.exe') `
    -Argument "--headless `"$pwsh`" -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" `
    -WorkingDirectory (Split-Path $PSScriptRoot -Parent)
$trigger = if ($DaysOfWeek) {
    New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $At
} else {
    New-ScheduledTaskTrigger -Daily -At $At
}
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes $TimeLimitMinutes) -MultipleInstances IgnoreNew

if (-not $Description) { $Description = "DrJoyAutomation: $scriptPath" }
Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
    -Description $Description -Force | Out-Null

$t = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath
$i = Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath $TaskPath
$when = if ($DaysOfWeek) { "$($DaysOfWeek -join ',') $At" } else { "daily $At" }
Write-Host ("Registered: {0} | {1} | {2} | {3} | next {4}" -f $t.TaskName, $t.State, $when, $t.Principal.LogonType, $i.NextRunTime)
