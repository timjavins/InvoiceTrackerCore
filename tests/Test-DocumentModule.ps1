# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-DocumentModule.ps1
#
# Proves FiscalCalendar.vb survives being pasted into ThisWorkbook, a document (class)
# module -- the only place production code actually lives. Test-FiscalCalendar.ps1 injects
# into a standard module, which allows things a document module forbids (Public Const,
# public fixed-size arrays, fixed-length strings, Declare). A module can pass every
# assertion there and still fail to compile here. See tests/README.md.
#
# Injects Header.vb (from the Securitas variant, the canary used for the stacker checks in
# the fiscal-calendar plan's Task 6) ahead of FiscalCalendar.vb, the same order
# Stack-VBFiles.ps1 pins them in the assembled stack. Header.vb is comment-only in that
# variant, so this is not a stand-in for the full assembled stack -- it is the minimum
# needed to reproduce the stacker's module order.
#
# Does NOT inject a tenant's assembled megastack: Securitas's stack declares variables typed
# as its UserForms, bound at compile time and absent from a bare workbook.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force

Reset-AssertCounters

$headerPath  = Join-Path $repo '..\SecuritasAutomation\Header.vb'
$fiscalPath  = Join-Path $repo 'FiscalCalendar.vb'
$sourceFiles = @($headerPath, $fiscalPath)

# Validate every source path BEFORE any COM object exists, same reasoning as
# VbaHarness.psm1's New-VbaHost: nothing can leak a hidden process on a typo'd path if
# nothing was created yet.
foreach ($file in $sourceFiles) {
    if (-not (Test-Path -LiteralPath $file)) { throw "Source file not found: $file" }
}

$excel = $null
$wb = $null

try {
    # New-Object -ComObject creates a SEPARATE instance. Never use GetActiveObject here:
    # that would attach to the user's Excel and run test code beside live workbooks.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    try {
        $project = $wb.VBProject
    } catch {
        throw "Cannot reach the VBA project. Enable Excel > File > Options > Trust Center > " +
              "Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', " +
              "then re-run. Excel said: $($_.Exception.Message)"
    }

    # ThisWorkbook is a document (class) module that exists on every workbook by default --
    # unlike VbaHarness.psm1's standard module, it is looked up, never Added. This is the
    # injection path that differs from the standard-module harness and is the whole point
    # of this script.
    $thisWorkbook = $project.VBComponents('ThisWorkbook')

    foreach ($file in $sourceFiles) {
        $source = Get-Content -LiteralPath $file -Raw -Encoding UTF8
        $thisWorkbook.CodeModule.AddFromString($source)
    }
} catch {
    if ($excel) { try { $excel.Quit() } catch { } }
    foreach ($obj in $wb, $excel) {
        if ($obj) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { } }
    }
    throw
}

# Shaped like VbaHarness.psm1's host so Remove-VbaHost can be reused as-is for cleanup.
$vba = @{ Excel = $excel; Workbook = $wb }

try {
    # Called as a COM method on the workbook object -- e.g. $wb.FiscalYearWeeks(2026) --
    # NOT Application.Run. Public procedures in a document module are members of that COM
    # object, and calling them this way returns their value. Application.Run executes them
    # but DISCARDS a Function's return value for a document-module procedure, which would
    # make every assertion below impossible.
    #
    # Each call also forces lazy compilation of the module up to that point, so working
    # through all 12 public functions is itself the compile check this script exists for:
    # a class-module violation (Public Const, a public fixed-size array, a fixed-length
    # string, Declare) would surface here as a COM exception, not as a silent pass.

    Assert-Equal -Expected 'ok' -Actual ($wb.FiscalCalendarSelfCheck()) `
                 -Because 'ThisWorkbook.FiscalCalendarSelfCheck compiles and runs as a document-module member'

    Assert-Equal -Expected '2026-01-31' -Actual ($wb.SaturdayClosestTo([datetime]'2026-01-31')) `
                 -Because 'ThisWorkbook.SaturdayClosestTo(2026-01-31)'

    Assert-Equal -Expected '2027-01-30' -Actual ($wb.FiscalYearEnd(2026)) `
                 -Because 'ThisWorkbook.FiscalYearEnd(2026)'

    Assert-Equal -Expected '2026-02-01' -Actual ($wb.FiscalYearStart(2026)) `
                 -Because 'ThisWorkbook.FiscalYearStart(2026)'

    Assert-Equal -Expected 52 -Actual ($wb.FiscalYearWeeks(2026)) `
                 -Because 'ThisWorkbook.FiscalYearWeeks(2026)'

    Assert-Equal -Expected 53 -Actual ($wb.FiscalYearWeeks(2023)) `
                 -Because 'ThisWorkbook.FiscalYearWeeks(2023)'

    Assert-Equal -Expected 5 -Actual ($wb.FiscalMonthWeeks(2026, 8)) `
                 -Because 'ThisWorkbook.FiscalMonthWeeks(2026, 8)'

    Assert-Equal -Expected 5 -Actual ($wb.FiscalMonthWeeks(2023, 12)) `
                 -Because 'ThisWorkbook.FiscalMonthWeeks(2023, 12)'

    Assert-Equal -Expected '2026-08-30' -Actual ($wb.FiscalMonthStart(2026, 8)) `
                 -Because 'ThisWorkbook.FiscalMonthStart(2026, 8)'

    Assert-Equal -Expected '2026-08-02' -Actual ($wb.FiscalMonthStart(2026, 7)) `
                 -Because 'ThisWorkbook.FiscalMonthStart(2026, 7)'

    Assert-Equal -Expected '2026-09-27' -Actual ($wb.FiscalWeekStart(2026, 35)) `
                 -Because 'ThisWorkbook.FiscalWeekStart(2026, 35)'

    Assert-Equal -Expected 2026 -Actual ($wb.FiscalYearOf([datetime]'2026-09-27')) `
                 -Because 'ThisWorkbook.FiscalYearOf(2026-09-27)'

    Assert-Equal -Expected 35 -Actual ($wb.FiscalWeekOf([datetime]'2026-09-27')) `
                 -Because 'ThisWorkbook.FiscalWeekOf(2026-09-27) -- fiscal September starting in calendar August'

    Assert-Equal -Expected 8 -Actual ($wb.FiscalMonthOf([datetime]'2026-08-30')) `
                 -Because 'ThisWorkbook.FiscalMonthOf(2026-08-30) -- fiscal September starting in calendar August'

    Assert-Equal -Expected '2026-09-27' -Actual ($wb.FiscalWeekStartOf([datetime]'2026-10-01')) `
                 -Because 'ThisWorkbook.FiscalWeekStartOf(2026-10-01)'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
}

Write-AssertSummary
