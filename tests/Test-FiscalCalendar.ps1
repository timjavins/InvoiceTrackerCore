# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force

Reset-AssertCounters

$fixture = Import-FiscalFixture -Path (Join-Path $repo 'fiscal-calendar-fixture.csv')

$vba = New-VbaHost -SourceFiles @(Join-Path $repo 'FiscalCalendar.vb')
try {
    Assert-Equal -Expected 'ok' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalCalendarSelfCheck') `
                 -Because 'harness can invoke an injected VBA function'

    foreach ($row in $fixture) {
        $fy = $row.fiscal_year

        Assert-Equal -Expected $row.year_end `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearEnd' -Arguments @($fy)) `
                     -Because "FY$fy year end"

        Assert-Equal -Expected $row.week1_start `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearStart' -Arguments @($fy)) `
                     -Because "FY$fy week 1 start"

        # Every year end is a Saturday, every year start a Sunday. Asserted directly rather
        # than trusted, because a rule that is off by one day still matches on some years.
        $end   = [datetime](Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearEnd'   -Arguments @($fy))
        $start = [datetime](Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearStart' -Arguments @($fy))
        Assert-Equal -Expected 'Saturday' -Actual $end.DayOfWeek   -Because "FY$fy ends on a Saturday"
        Assert-Equal -Expected 'Sunday'   -Actual $start.DayOfWeek -Because "FY$fy starts on a Sunday"
    }

    # The anchor helper, at both tie-break directions. 31 Jan 2026 IS a Saturday (no move);
    # 31 Jan 2027 is a Sunday (move back 1); 31 Jan 2007 is a Wednesday (move forward 3).
    Assert-Equal -Expected '2026-01-31' `
                 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'SaturdayClosestTo' -Arguments @([datetime]'2026-01-31')) `
                 -Because 'anchor already on a Saturday does not move'
    Assert-Equal -Expected '2027-01-30' `
                 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'SaturdayClosestTo' -Arguments @([datetime]'2027-01-31')) `
                 -Because 'anchor on a Sunday moves back one day'
    Assert-Equal -Expected '2007-02-03' `
                 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'SaturdayClosestTo' -Arguments @([datetime]'2007-01-31')) `
                 -Because 'anchor on a Wednesday moves forward three days'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
}

Write-AssertSummary
