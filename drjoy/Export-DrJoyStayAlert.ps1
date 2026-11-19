# Export-DrJoyStayAlert.ps1  20260927
# Export the Dr.JOY 滞在アラート list (at0052) for one month to Downloads as
# 滞在アラート_YYYY_M.csv (Shift_JIS) by running drjoy-stay-alert-export.js
# (Playwright, no screenshots). The screen keeps only last, this and
# next month.
#
#   pwsh -File .\Export-DrJoyStayAlert.ps1 -Month 2026-09 -Headless
#
# The saved path is on the last line: "OK: <path> (...)".
#Requires -Version 7

param(
    [Parameter(Mandatory)][string]$Month,
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

$nodeArgs = @('drjoy-stay-alert-export.js', '--month', $Month, '--out-dir', $OutDir)
if ($Headless) { $nodeArgs += '--headless' }

node @nodeArgs
$code = $LASTEXITCODE

switch ($code) {
    0 { }
    2 { Write-Host 'Login timed out. Run again without -Headless and log in within 5 minutes.' -ForegroundColor Yellow }
    3 { Write-Host 'The CSV failed validation and was saved as rejected-stay-alert-*.csv. Check the output above.' -ForegroundColor Yellow }
    default { Write-Host "Failed (exit code $code)." -ForegroundColor Red }
}
exit $code
