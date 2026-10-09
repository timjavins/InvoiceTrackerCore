# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-AccrualCutoff.ps1
#
# The fiscal-month arithmetic behind the accruals report's cut-off picker (AccrualCutoff.vb), run
# over FiscalCalendar.vb in a throwaway standard module. The prompts, the picker and the report
# itself live in GenerateAccrualsReport.vb and are not covered here: they need a worksheet, a tenant
# config and a person at a dialog.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force

Reset-AssertCounters

function Show-Day($value) { ([datetime]$value).ToString('yyyy-MM-dd ddd') }

$vba = New-VbaHost -SourceFiles @((Join-Path $repo 'FiscalCalendar.vb'), (Join-Path $repo 'AccrualCutoff.vb'))
try {
    function Call-Vba($name, $arguments) {
        Invoke-VbaFunction -VbaHost $vba -Name $name -Arguments $arguments -TimeoutSeconds 30
    }

    # Month ends.
    Assert-Equal -Expected '2026-10-03 Sat' -Actual (Show-Day (Call-Vba 'AccrualMonthEnd' @(2026, 8))) `
                 -Because 'FY2026 September is a 5-week month ending 2026-10-03'
    Assert-Equal -Expected '2027-01-30 Sat' -Actual (Show-Day (Call-Vba 'AccrualMonthEnd' @(2026, 12))) `
                 -Because 'FY2026 January ends the fiscal year'
    Assert-Equal -Expected (Show-Day (Call-Vba 'FiscalYearEnd' @(2023))) `
                 -Actual (Show-Day (Call-Vba 'AccrualMonthEnd' @(2023, 12))) `
                 -Because 'FY2023 has 53 weeks, so its January absorbs the extra week and still ends the year'

    # Every month ends on a Saturday, the day before the next month starts. Sampled rather than swept
    # over all 45 years: each check is two COM calls and the full sweep took minutes. The sample holds
    # every 53-week year the module's own tests cover, which is where month boundaries move.
    $bad = 0
    foreach ($fy in 2005, 2006, 2012, 2017, 2023, 2026, 2028, 2034, 2040, 2045, 2049) {
        foreach ($m in 1..11) {
            $end  = [datetime](Call-Vba 'AccrualMonthEnd' @($fy, $m))
            $next = [datetime](Call-Vba 'FiscalMonthStart' @($fy, ($m + 1)))
            if ($end.DayOfWeek -ne 'Saturday' -or $end.AddDays(1) -ne $next) { $bad++ }
        }
    }
    Assert-Equal -Expected 0 -Actual $bad -Because 'month ends are Saturdays that abut the next month start, in the sampled years'

    # Last closed day: the newest month that has finished, not the one asOf falls in.
    Assert-Equal -Expected '2026-10-03 Sat' -Actual (Show-Day (Call-Vba 'AccrualLastClosedDay' @([datetime]'2026-10-09'))) `
                 -Because 'on 2026-10-09 September has closed'
    Assert-Equal -Expected '2026-10-03 Sat' -Actual (Show-Day (Call-Vba 'AccrualLastClosedDay' @([datetime]'2026-10-04'))) `
                 -Because 'the first day of fiscal October: September has just closed'
    Assert-Equal -Expected '2026-08-29 Sat' -Actual (Show-Day (Call-Vba 'AccrualLastClosedDay' @([datetime]'2026-10-03'))) `
                 -Because 'on September''s last day September is still open, so August is the newest closed month'
    Assert-Equal -Expected '2027-01-30 Sat' -Actual (Show-Day (Call-Vba 'AccrualLastClosedDay' @([datetime]'2027-02-03'))) `
                 -Because 'early in FY2027 the newest closed month is the previous year''s January'
    Assert-Equal -Expected '2026-01-31 Sat' -Actual (Show-Day (Call-Vba 'AccrualLastClosedDay' @([datetime]'2026-02-10'))) `
                 -Because 'early in FY2026 the newest closed month is FY2025''s January'

    # Parsing what the user clicked or typed, with eight closed months (as on 2026-10-09).
    $answers = @(
        @('8', 8), @(' 1 ', 1), @('September', 8), @('september', 8), @('february', 1),
        @('9', 0), @('0', 0), @('-1', 0), @('1.5', 0), @('Oct', 0), @('October', 0), @('', 0), @('abc', 0)
    )
    foreach ($case in $answers) {
        Assert-Equal -Expected $case[1] -Actual ([int](Call-Vba 'AccrualMonthFromAnswer' @($case[0], 8))) `
                     -Because "answer '$($case[0])' with 8 closed months"
    }

    Assert-Equal -Expected 'February' -Actual (Call-Vba 'AccrualFiscalMonthName' @(1))  -Because 'month 1 is February'
    Assert-Equal -Expected 'January'  -Actual (Call-Vba 'AccrualFiscalMonthName' @(12)) -Because 'month 12 is January'
    Assert-Equal -Expected ''         -Actual (Call-Vba 'AccrualFiscalMonthName' @(13)) -Because 'out of range yields empty'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
}

Write-AssertSummary
