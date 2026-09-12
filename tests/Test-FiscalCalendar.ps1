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

        Assert-Equal -Expected $row.weeks `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearWeeks' -Arguments @($fy)) `
                     -Because "FY$fy week count"

        # Every year end is a Saturday, every year start a Sunday. Asserted directly rather
        # than trusted, because a rule that is off by one day still matches on some years.
        $end   = [datetime](Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearEnd'   -Arguments @($fy))
        $start = [datetime](Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearStart' -Arguments @($fy))
        Assert-Equal -Expected 'Saturday' -Actual $end.DayOfWeek   -Because "FY$fy ends on a Saturday"
        Assert-Equal -Expected 'Sunday'   -Actual $start.DayOfWeek -Because "FY$fy starts on a Sunday"
    }

    $expected53 = @(2006, 2012, 2017, 2023, 2028, 2034, 2040, 2045)
    $actual53 = @($fixture | Where-Object {
        (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearWeeks' -Arguments @($_.fiscal_year)) -eq 53
    } | ForEach-Object { $_.fiscal_year })
    Assert-Equal -Expected ($expected53 -join ',') -Actual ($actual53 -join ',') `
                 -Because '53-week years are exactly the eight known ones'

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

    # FY2026 is a 52-week year: 4,5,4 four times over.
    $normal = @(4,5,4,4,5,4,4,5,4,4,5,4)
    for ($m = 1; $m -le 12; $m++) {
        Assert-Equal -Expected $normal[$m - 1] `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthWeeks' -Arguments @(2026, $m)) `
                     -Because "FY2026 fiscal month $m week count"
    }

    # FY2023 is a 53-week year: the extra week lands on January, the twelfth fiscal month.
    $long = @(4,5,4,4,5,4,4,5,4,4,5,5)
    for ($m = 1; $m -le 12; $m++) {
        Assert-Equal -Expected $long[$m - 1] `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthWeeks' -Arguments @(2023, $m)) `
                     -Because "FY2023 (53-week) fiscal month $m week count"
    }

    # Month weeks must always sum to the year's week count, for every year in the fixture.
    foreach ($row in $fixture) {
        $sum = 0
        for ($m = 1; $m -le 12; $m++) {
            $sum += [int](Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthWeeks' -Arguments @($row.fiscal_year, $m))
        }
        Assert-Equal -Expected $row.weeks -Actual $sum -Because "FY$($row.fiscal_year) month weeks sum to the year"
    }

    # Known month boundaries, read from the published FY2026 calendar.
    Assert-Equal -Expected '2026-02-01' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 1))  -Because 'FY2026 February starts'
    Assert-Equal -Expected '2026-03-01' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 2))  -Because 'FY2026 March starts'
    Assert-Equal -Expected '2026-04-05' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 3))  -Because 'FY2026 April starts'
    Assert-Equal -Expected '2026-08-02' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 7))  -Because 'FY2026 August starts (fiscal month 7, the index that is easy to confuse with September)'
    Assert-Equal -Expected '2026-08-30' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 8))  -Because 'FY2026 September starts'
    Assert-Equal -Expected '2027-01-03' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 12)) -Because 'FY2026 January starts'

    # Week starts: first, a mid-year one, and the last week of a 52-week year.
    Assert-Equal -Expected '2026-02-01' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2026, 1))  -Because 'FY2026 week 1 start'
    Assert-Equal -Expected '2026-09-27' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2026, 35)) -Because 'FY2026 week 35 start'
    Assert-Equal -Expected '2027-01-24' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2026, 52)) -Because 'FY2026 week 52 start'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
}

Write-AssertSummary
