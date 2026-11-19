# Export-DrJoyAt0029.ps1  20261017（WorkRecord を追加）
# Create and download an export on Dr.JOY's data-download screen (at0029)
# by running drjoy-at0029-export.js (Playwright, no screenshots).
# The file is saved to Downloads under its Dr.JOY name.
#
#   pwsh -File .\Export-DrJoyAt0029.ps1 -Type WorkingPlan -Month 2026-10 -Headless
#   pwsh -File .\Export-DrJoyAt0029.ps1 -Type Beacon -From 2026-10-01 -To 2026-10-07 -Headless
#   pwsh -File .\Export-DrJoyAt0029.ps1 -Type Payroll -Month 2026-09 -Headless     (給与連携項目一覧・医師・1日〜月末)
#   pwsh -File .\Export-DrJoyAt0029.ps1 -Type DepartmentHistory -Year 2026 -Headless   (スタッフ所属履歴・Excel・本登録)
#   pwsh -File .\Export-DrJoyAt0029.ps1 -Type WorkRecord -Month 2026-09 -Headless    (勤務実績・医師・本登録・Shift_JIS・HH:MM)
#
# The saved path is on the last line: "OK: <path> (...)".
#Requires -Version 7

param(
    [Parameter(Mandatory)][ValidateSet('WorkingPlan', 'Beacon', 'Payroll', 'DepartmentHistory', 'WorkRecord')][string]$Type,
    [string]$Month,
    [string]$From,
    [string]$To,
    [string]$Year,
    [switch]$Headless,
    [string]$OutDir = (Join-Path $env:USERPROFILE 'Downloads')
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Error 'Node.js is not installed (node.exe not found on PATH).'
    exit 1
}
if (-not (Test-Path (Join-Path $PSScriptRoot 'node_modules\playwright-core'))) {
    Write-Host 'Installing playwright-core (first run only)...'
    npm install --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

$nodeArgs = @('drjoy-at0029-export.js', '--out-dir', $OutDir)
switch ($Type) {
    'WorkingPlan' { $nodeArgs += @('--type', 'working-plan', '--month', $Month) }
    'WorkRecord'  { $nodeArgs += @('--type', 'work-record', '--month', $Month) }
    'Beacon'      { $nodeArgs += @('--type', 'beacon', '--from', $From, '--to', $To) }
    'Payroll'     { $nodeArgs += @('--type', 'payroll', '--month', $Month) }
    'DepartmentHistory' { $nodeArgs += @('--type', 'department-history'); if ($Year) { $nodeArgs += @('--year', $Year) } }
}
if ($Headless) { $nodeArgs += '--headless' }

node @nodeArgs
$code = $LASTEXITCODE

switch ($code) {
    0 { }
    2 { Write-Host 'Login timed out. Run again without -Headless and log in within 5 minutes.' -ForegroundColor Yellow }
    3 { Write-Host 'The file failed validation and was saved as rejected-at0029-*. Check the output above.' -ForegroundColor Yellow }
    default { Write-Host "Failed (exit code $code)." -ForegroundColor Red }
}
exit $code
