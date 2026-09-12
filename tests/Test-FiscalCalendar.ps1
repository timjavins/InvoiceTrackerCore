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

    # Round-trip: every fixture year's boundaries resolve back to that year, and the day
    # either side belongs to the neighbouring year. Boundary-off-by-one is the likely bug here.
    foreach ($row in $fixture) {
        $fy = $row.fiscal_year
        Assert-Equal -Expected $fy -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearOf' -Arguments @($row.week1_start)) -Because "FY$fy starts inside FY$fy"
        Assert-Equal -Expected $fy -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearOf' -Arguments @($row.year_end))    -Because "FY$fy ends inside FY$fy"
        Assert-Equal -Expected ($fy - 1) -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearOf' -Arguments @($row.week1_start.AddDays(-1))) -Because "day before FY$fy belongs to FY$($fy-1)"
        Assert-Equal -Expected ($fy + 1) -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearOf' -Arguments @($row.year_end.AddDays(1)))     -Because "day after FY$fy belongs to FY$($fy+1)"

        Assert-Equal -Expected 1 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekOf' -Arguments @($row.week1_start)) -Because "FY$fy first day is week 1"
        Assert-Equal -Expected $row.weeks -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekOf' -Arguments @($row.year_end)) -Because "FY$fy last day is week $($row.weeks)"
    }

    # Known FY2026 points, including the fiscal September that starts in calendar August.
    Assert-Equal -Expected 2026 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearOf'  -Arguments @([datetime]'2026-09-27')) -Because '2026-09-27 is in FY2026'
    Assert-Equal -Expected 35   -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekOf'  -Arguments @([datetime]'2026-09-27')) -Because '2026-09-27 is fiscal week 35'
    Assert-Equal -Expected 8    -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthOf' -Arguments @([datetime]'2026-09-27')) -Because '2026-09-27 is in fiscal September (month 8)'
    Assert-Equal -Expected 8    -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthOf' -Arguments @([datetime]'2026-08-30')) -Because '2026-08-30 is already fiscal September'
    Assert-Equal -Expected 7    -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthOf' -Arguments @([datetime]'2026-08-29')) -Because '2026-08-29 is still fiscal August'
    Assert-Equal -Expected 1    -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthOf' -Arguments @([datetime]'2026-02-01')) -Because '2026-02-01 is fiscal February (month 1)'
    Assert-Equal -Expected 12   -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthOf' -Arguments @([datetime]'2027-01-30')) -Because '2027-01-30 is fiscal January (month 12)'

    # A mid-week date resolves to its week's Sunday. This is what the guards tracker stores.
    Assert-Equal -Expected '2026-09-27' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStartOf' -Arguments @([datetime]'2026-10-01')) -Because 'a Thursday resolves back to its week Sunday'
    Assert-Equal -Expected '2026-09-27' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStartOf' -Arguments @([datetime]'2026-09-27')) -Because 'a Sunday resolves to itself'

    # The 35 consecutive week starts the guards tracker holds must all resolve to weeks 1-35 of FY2026.
    for ($i = 0; $i -lt 35; $i++) {
        $d = ([datetime]'2026-02-01').AddDays($i * 7)
        Assert-Equal -Expected ($i + 1) -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekOf' -Arguments @($d)) `
                     -Because "guards tracker week $($i + 1) ($($d.ToString('yyyy-MM-dd'))) is FY2026 week $($i + 1)"
    }

    # A date carrying a time component must resolve exactly as its midnight equivalent does.
    # These three are ordinary behavioural checks, not proof of the Int-vs-CLng fix: none of
    # them actually discriminates. 2026-09-27 is itself the Sunday that starts week 35, so
    # rounding 14:00 up to the 28th stays inside the same week. 2027-01-30 23:59 never reaches
    # the CLng/Int conversion at all -- Year() and the date comparisons in FiscalYearOf work
    # from the date's truncated part regardless of time-of-day, so this is exact either way.
    # 2026-10-01 09:30 has a fraction under 0.5, so it rounds down to the same day truncation
    # would give. See the Saturday-afternoon case below for the assertion that actually
    # distinguishes Int from CLng.
    Assert-Equal -Expected 35 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekOf' -Arguments @([datetime]'2026-09-27 14:00')) -Because 'an afternoon time does not shift the fiscal week (mid-week; see the Saturday case for the discriminating test)'
    Assert-Equal -Expected 2026 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearOf' -Arguments @([datetime]'2027-01-30 23:59')) -Because 'a late time on the last day of the fiscal year still resolves to that fiscal year (plain boundary check; Year() and the comparisons are exact regardless of time-of-day)'
    Assert-Equal -Expected '2026-09-27' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStartOf' -Arguments @([datetime]'2026-10-01 09:30')) -Because 'a mid-week morning resolves to its week Sunday'

    # Saturday is the last day of a fiscal week, so rounding up crosses into the next one.
    # This is the case that actually distinguishes Int from CLng: Int keeps week 35, CLng gives 36.
    #
    # Only the FiscalWeekOf assertion below can detect a rounding regression. FiscalWeekStartOf
    # normalises its own input with Int() before calling FiscalWeekOf, so on that path FiscalWeekOf
    # only ever receives a whole date and CLng would be exact. That per-function normalisation is
    # deliberate defence in depth -- a caller passing Now() must get the right answer regardless of
    # what a callee does -- and the cost is that it masks a callee's rounding bug. Hence the direct
    # test. Verified: reverting FiscalWeekOf to CLng fails the first assertion and not the second.
    Assert-Equal -Expected 35 -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekOf' -Arguments @([datetime]'2026-10-03 14:00')) -Because 'a Saturday afternoon stays in its own fiscal week (Int truncates; CLng would round into the next week)'
    Assert-Equal -Expected '2026-09-27' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStartOf' -Arguments @([datetime]'2026-10-03 14:00')) -Because 'a Saturday afternoon resolves to its own week Sunday, not the next one'

    # Week 53 in a long year is otherwise untested. Guard clauses that would raise (invalid
    # month/week numbers) are verified by code review, not by this suite -- see tests/README.md.
    Assert-Equal -Expected '2024-01-28' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2023, 53)) -Because 'FY2023 is a 53-week year, so week 53 is valid and is its last week'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
}

Write-AssertSummary
