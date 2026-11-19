# Export-DrJoyStaff.ps1  20260923
# Export the Dr.JOY staff list (at0006) to Downloads as "staffs (N).csv"
# by running drjoy-staff-export.js (Playwright, no screenshots).
# First run: a Chrome window opens; log in to Dr.JOY there once.
#
#   pwsh -File .\Export-DrJoyStaff.ps1              export
#   pwsh -File .\Export-DrJoyStaff.ps1 -LoginOnly   log in and save the session only
#   pwsh -File .\Export-DrJoyStaff.ps1 -Inspect     print the at0006 form controls
#Requires -Version 7

param(
    [switch]$LoginOnly,
    [switch]$Inspect,
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

$nodeArgs = @('drjoy-staff-export.js', '--out-dir', $OutDir)
if ($LoginOnly) { $nodeArgs += '--login-only' }
if ($Inspect)   { $nodeArgs += '--inspect' }
if ($Headless)  { $nodeArgs += '--headless' }

node @nodeArgs
$code = $LASTEXITCODE

switch ($code) {
    0 { }
    2 { Write-Host 'Login timed out. Run again and log in within 5 minutes.' -ForegroundColor Yellow }
    3 { Write-Host 'The CSV failed validation and was NOT saved as staffs (N).csv. Check the at0006 filters.' -ForegroundColor Yellow }
    default { Write-Host "Failed (exit code $code)." -ForegroundColor Red }
}
exit $code
