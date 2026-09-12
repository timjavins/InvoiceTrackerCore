# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force

Reset-AssertCounters

$vba = New-VbaHost -SourceFiles @(Join-Path $repo 'FiscalCalendar.vb')
try {
    # Smoke: the harness can call into injected VBA at all.
    Assert-Equal -Expected 'ok' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalCalendarSelfCheck') `
                 -Because 'harness can invoke an injected VBA function'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
}

Write-AssertSummary
