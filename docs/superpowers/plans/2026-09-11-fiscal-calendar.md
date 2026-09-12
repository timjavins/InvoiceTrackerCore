# Fiscal Calendar (core) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `FiscalCalendar.vb` to `InvoiceTrackerCore` — a module that computes Nordstrom's 4-5-4 retail fiscal calendar from one anchor rule, with a PowerShell/COM test harness that asserts it against 46 years of published ground truth.

**Architecture:** One VBA module of pure functions, no worksheet or workbook dependencies, no module-level state. Correctness comes from a single anchor rule (the Saturday closest to 31 January) plus a repeating `4,5,4` month cycle. Tests run outside Excel's UI: a PowerShell script spins up its own hidden Excel instance, injects the module into a throwaway workbook as a standard module, calls each function over COM, and compares results to `fiscal-calendar-fixture.csv`.

**Tech Stack:** VBA (Excel), PowerShell 5.1, Excel COM automation, CSV fixture.

**Spec:** `../../../NordGuardsTracker/docs/superpowers/specs/2026-09-11-po-key-and-core-adoption-design.md` — see the Roadmap section, "`AddNewWeeks`" and "Compute the calendar; use the file as a test fixture, not a dependency". This plan implements effort **B1** from that spec's delivery split.

## Global Constraints

- **No `Option Explicit` outside `Header.vb`.** Every module is concatenated into one file; a module-level option landing mid-file fails to compile.
- **No `Public Const`, public fixed-size arrays, fixed-length strings, or `Declare`.** The stack is pasted into `ThisWorkbook`, a class module, which forbids these as public members. `Private Const` and procedure-local `Const` are fine.
- **Prefer no module-level declarations at all in this module.** `Stack-VBFiles.ps1` hoists them and reports having done so; a module with none is quieter to review.
- **Do not depend on the active sheet, on `ThisWorkbook`, or on any worksheet.** This module is pure computation and must stay callable from a bare standard module in an empty workbook — the test harness relies on exactly that.
- **Filenames are the module system.** `Stack-VBFiles.ps1` discovers `*.vb` **recursively** with no `tests/` exclusion, and collides on **leaf filename, case-insensitively**. Never add a `.vb` file under any folder in this repo that shares a name with an existing one, and never put VBA test code in this repo — it would ship into every tenant's production stack. PowerShell (`.ps1`) files are not swept and are safe.
- **Ordering is `Header.vb`, then `Refresh.vb`, then alphabetical by filename.** Do not rely on any other module having been defined "before" this one; VBA does not care, but reviewers do.
- **Fiscal week 1 of any year starts on a Sunday; every fiscal year ends on a Saturday.** Verified across all 46 fixture years.
- **The fixture is a test oracle, never a runtime input.** `FiscalCalendar.vb` must not read any file, sheet, or network resource.

---

### Task 1: PowerShell/COM VBA test harness

Every later task depends on this, and it is the one piece a reviewer could reject on its own merits (COM lifecycle, cleanup, trust failure messaging) without judging any calendar logic.

**Files:**
- Create: `tests/VbaHarness.psm1`
- Create: `tests/Test-FiscalCalendar.ps1`
- Create: `tests/README.md`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `New-VbaHost -SourceFiles <string[]>` → `[hashtable]` with keys `Excel` (Application), `Workbook`, `Run` (a `[scriptblock]`-free helper is not used; call `$host.Excel.Run(...)` directly), `TempPath`.
  - `Invoke-VbaFunction -Host <hashtable> -Name <string> -Args <object[]>` → the function's return value.
  - `Remove-VbaHost -Host <hashtable>` → `$null`; closes the workbook without saving, quits Excel, releases COM, deletes the temp file.
  - `Import-FiscalFixture -Path <string>` → array of objects with `fiscal_year` (int), `week1_start` (DateTime), `year_end` (DateTime), `weeks` (int).
  - `Assert-Equal -Expected <object> -Actual <object> -Because <string>` → increments script-scope pass/fail counters; writes one line per failure.

- [ ] **Step 1: Write the harness module**

Create `tests/VbaHarness.psm1`:

```powershell
# Runs VBA from InvoiceTrackerCore against a throwaway workbook in its own hidden
# Excel instance, so tests never touch a workbook the user has open.
#
# Requires Excel's "Trust access to the VBA project object model" (VBOM). Without it
# $wb.VBProject throws, and the message Excel gives is unhelpful, so we catch and explain.

Set-StrictMode -Version Latest

function New-VbaHost {
    param([Parameter(Mandatory)][string[]] $SourceFiles)

    # New-Object -ComObject creates a SEPARATE instance. Never use GetActiveObject here:
    # that would attach to the user's Excel and run test code beside live workbooks.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    try {
        $project = $wb.VBProject
    } catch {
        $excel.Quit()
        throw "Cannot reach the VBA project. Enable Excel > File > Options > Trust Center > " +
              "Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', " +
              "then re-run. Excel said: $($_.Exception.Message)"
    }

    # 1 = vbext_ct_StdModule. A standard module (not ThisWorkbook) so Application.Run
    # resolves unqualified names -- a class module would require 'ThisWorkbook.Proc'.
    $module = $project.VBComponents.Add(1)

    foreach ($file in $SourceFiles) {
        if (-not (Test-Path -LiteralPath $file)) { throw "Source file not found: $file" }
        $source = Get-Content -LiteralPath $file -Raw -Encoding UTF8
        $module.CodeModule.AddFromString($source)
    }

    $temp = Join-Path $env:TEMP ("vbatest_{0}.xlsm" -f [guid]::NewGuid().ToString('N'))

    @{ Excel = $excel; Workbook = $wb; TempPath = $temp }
}

function Invoke-VbaFunction {
    param(
        [Parameter(Mandatory)][hashtable] $VbaHost,
        [Parameter(Mandatory)][string] $Name,
        [object[]] $Arguments = @()
    )
    switch ($Arguments.Count) {
        0 { $VbaHost.Excel.Run($Name) }
        1 { $VbaHost.Excel.Run($Name, $Arguments[0]) }
        2 { $VbaHost.Excel.Run($Name, $Arguments[0], $Arguments[1]) }
        default { throw "Invoke-VbaFunction supports up to 2 arguments; got $($Arguments.Count)." }
    }
}

function Remove-VbaHost {
    param([Parameter(Mandatory)][hashtable] $VbaHost)

    try { $VbaHost.Workbook.Close($false) } catch { }
    try { $VbaHost.Excel.Quit() } catch { }
    foreach ($key in 'Workbook', 'Excel') {
        if ($VbaHost[$key]) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($VbaHost[$key]) } catch { }
        }
    }
    if ($VbaHost.TempPath -and (Test-Path -LiteralPath $VbaHost.TempPath)) {
        Remove-Item -LiteralPath $VbaHost.TempPath -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    $null
}

function Import-FiscalFixture {
    param([Parameter(Mandatory)][string] $Path)

    # Import-Csv cannot skip comment lines, so strip them before parsing.
    Get-Content -LiteralPath $Path |
        Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } |
        ConvertFrom-Csv |
        ForEach-Object {
            [pscustomobject]@{
                fiscal_year = [int]    $_.fiscal_year
                week1_start = [datetime]::ParseExact($_.week1_start, 'yyyy-MM-dd', $null)
                year_end    = [datetime]::ParseExact($_.year_end,    'yyyy-MM-dd', $null)
                weeks       = [int]    $_.weeks
            }
        }
}

function Reset-AssertCounters {
    $script:Passed = 0
    $script:Failed = 0
    $script:Failures = New-Object System.Collections.Generic.List[string]
}

function Assert-Equal {
    param($Expected, $Actual, [Parameter(Mandatory)][string] $Because)

    # Normalise dates to yyyy-MM-dd so a COM DateTime with a zero time component
    # compares equal to a fixture date.
    if ($Expected -is [datetime]) { $Expected = $Expected.ToString('yyyy-MM-dd') }
    if ($Actual   -is [datetime]) { $Actual   = $Actual.ToString('yyyy-MM-dd') }

    if ("$Expected" -eq "$Actual") {
        $script:Passed++
    } else {
        $script:Failed++
        $script:Failures.Add("FAIL $Because -- expected '$Expected', got '$Actual'")
    }
}

function Write-AssertSummary {
    foreach ($f in $script:Failures) { Write-Host $f -ForegroundColor Red }
    Write-Host ""
    Write-Host ("passed: {0}  failed: {1}" -f $script:Passed, $script:Failed)
    if ($script:Failed -gt 0) { exit 1 }
    Write-Host "ALL PASS" -ForegroundColor Green
    exit 0
}

Export-ModuleMember -Function New-VbaHost, Invoke-VbaFunction, Remove-VbaHost,
                              Import-FiscalFixture, Reset-AssertCounters, Assert-Equal,
                              Write-AssertSummary
```

- [ ] **Step 2: Write the smoke test that proves the harness works**

Create `tests/Test-FiscalCalendar.ps1`. At this stage it only proves injection and invocation work — calendar assertions arrive in Task 2.

```powershell
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
```

- [ ] **Step 3: Run the test to verify it fails**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: throws `Source file not found: ...\FiscalCalendar.vb`. That is the correct first failure — the module does not exist yet.

- [ ] **Step 4: Create the module with only the self-check**

Create `FiscalCalendar.vb` in the repo root:

```vb
' Nordstrom 4-5-4 retail fiscal calendar, computed from one anchor rule.
'
' Pure computation: no worksheet, no workbook, no file, no network. That is deliberate --
' the calendar used to be looked up in a Finance-owned SharePoint workbook, which meant every
' consumer inherited that file's availability and layout. The rule below reproduces all 46
' published years exactly, so nothing needs to read it at run time.
'
' The rule:
'   1. Fiscal year N ends on the Saturday closest to 31 January of year N+1. Year N+1 starts
'      the Sunday after. Weeks start Sunday; years end Saturday.
'   2. Each quarter is 4,5,4 weeks -- 52 weeks in a normal year.
'   3. A year runs 53 weeks exactly when the gap between consecutive year-end anchors is 371
'      days rather than 364. The extra week lands on the final fiscal month, January.
'
' Rules 1 and 3 absorb leap-year drift at the anchor, which is why nothing here reasons about
' February directly.
'
' Verified against tests/../fiscal-calendar-fixture.csv (FY2005-FY2050) by
' tests/Test-FiscalCalendar.ps1. See ADR and the design doc referenced in that plan.
'
' Fiscal month indexes are 1-12 where 1 = February and 12 = January, because that is the order
' the fiscal year runs in. Callers wanting a calendar month name must map it themselves.

' Proves the module loaded and is callable. Used by the test harness smoke check.
Public Function FiscalCalendarSelfCheck() As String
    FiscalCalendarSelfCheck = "ok"
End Function
```

- [ ] **Step 5: Run the test to verify it passes**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: `passed: 1  failed: 0` then `ALL PASS`.

If it instead throws about the VBA project, VBOM trust is off — enable it as the error text describes and re-run. Do not proceed until this passes; every later task uses this harness.

- [ ] **Step 6: Document how to run the tests**

Create `tests/README.md`:

```markdown
# Tests

PowerShell, driving Excel over COM. There is no VBA test framework here and there cannot be a
convenient one: `Stack-VBFiles.ps1` sweeps `*.vb` recursively with no `tests/` exclusion, so any
VBA test module in this repo would be concatenated into every tenant's production stack. `.ps1`
files are not swept, so the tests live here instead.

Run them:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Each script exits 0 on all-pass, 1 on any failure, so they are CI-usable as-is.

## Requirements

- Excel installed.
- Excel's **Trust access to the VBA project object model** enabled (Trust Center > Macro
  Settings). The harness injects source into a throwaway workbook; without VBOM it cannot.

## What the harness does not do

It never attaches to a running Excel instance and never opens a workbook you have open. Each run
creates its own hidden `Excel.Application`, adds an empty workbook, injects the modules under test
as a **standard** module, calls functions with `Application.Run`, then closes without saving and
deletes its temp file.

The standard module is a deliberate simplification with a real cost. `Application.Run` cannot
resolve an unqualified name in a document module, and even when qualified as `ThisWorkbook.Proc` it
executes the procedure but **discards a `Function`'s return value** — so assertions would be
impossible there. A standard module sidesteps both.

The cost is that a standard module allows things a document module forbids: `Public Const`, public
fixed-size arrays, fixed-length strings, `Declare`. **A module can pass every test here and still
fail to compile in the real stack.** That gap is closed separately, by compiling the assembled stack
in a scratch workbook and asserting against it through the Excel MCP, whose `run_macro` does return
values from document-module functions. Both layers are needed; neither substitutes for the other.
```

- [ ] **Step 7: Commit**

```bash
git add tests/VbaHarness.psm1 tests/Test-FiscalCalendar.ps1 tests/README.md FiscalCalendar.vb
git commit -m "test: add a PowerShell/COM harness for running VBA modules

VBA has no test framework here and cannot easily have an in-repo one: the stacker
sweeps *.vb recursively with no tests/ exclusion, so VBA test modules would ship
into every tenant's production stack. PowerShell files are not swept.

The harness creates its own hidden Excel instance rather than attaching to the
user's, injects modules as a standard module so Application.Run resolves
unqualified names, and cleans up its temp workbook."
```

---

### Task 2: Year boundaries — `FiscalYearEnd` and `FiscalYearStart`

**Files:**
- Modify: `FiscalCalendar.vb`
- Modify: `tests/Test-FiscalCalendar.ps1`

**Interfaces:**
- Consumes: the Task 1 harness (`New-VbaHost`, `Invoke-VbaFunction`, `Import-FiscalFixture`, `Assert-Equal`).
- Produces:
  - `FiscalYearEnd(ByVal fiscalYear As Long) As Date`
  - `FiscalYearStart(ByVal fiscalYear As Long) As Date`
  - `SaturdayClosestTo(ByVal anchor As Date) As Date` (Public — it is independently useful and independently testable)

- [ ] **Step 1: Write the failing tests**

Replace the smoke assertion block in `tests/Test-FiscalCalendar.ps1` (keep the surrounding scaffolding) with:

```powershell
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
```

- [ ] **Step 2: Run the tests to verify they fail**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: a COM error naming `FiscalYearEnd` as unavailable — the function does not exist yet.

- [ ] **Step 3: Implement the year boundary functions**

Append to `FiscalCalendar.vb`:

```vb
' The Saturday nearest a given date. Ties cannot occur: a date is at most 3 days from one
' Saturday and at least 4 from the other, so "forward if 3 or fewer, else back" is total.
Public Function SaturdayClosestTo(ByVal anchor As Date) As Date
    Dim dow As Long
    dow = Weekday(anchor, vbSunday)      ' 1 = Sunday ... 7 = Saturday

    Dim forwardDays As Long
    forwardDays = 7 - dow                ' 0 when anchor is already Saturday

    If forwardDays <= 3 Then
        SaturdayClosestTo = anchor + forwardDays
    Else
        SaturdayClosestTo = anchor - dow ' the previous Saturday
    End If
End Function

' Last day of the fiscal year: the Saturday closest to 31 January of the FOLLOWING calendar year.
Public Function FiscalYearEnd(ByVal fiscalYear As Long) As Date
    FiscalYearEnd = SaturdayClosestTo(DateSerial(fiscalYear + 1, 1, 31))
End Function

' First day of the fiscal year -- the Sunday after the previous year ended. Derived from the
' previous year's anchor rather than from this year's, so the two can never disagree about
' where the boundary is.
Public Function FiscalYearStart(ByVal fiscalYear As Long) As Date
    FiscalYearStart = FiscalYearEnd(fiscalYear - 1) + 1
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: `passed: 231  failed: 0` — 46 years × 4 assertions, plus 1 smoke, plus 3 anchor cases, then `ALL PASS`. If the count differs but failures are 0, recount rather than assuming; a silently skipped fixture row is a real bug in the test.

- [ ] **Step 5: Commit**

```bash
git add FiscalCalendar.vb tests/Test-FiscalCalendar.ps1
git commit -m "feat(FiscalCalendar): compute fiscal year boundaries from the anchor rule

Year N ends on the Saturday closest to 31 January of year N+1; year N starts the
Sunday after N-1 ends. Deriving the start from the previous year's anchor rather
than its own means the two cannot disagree about the boundary.

Verified against all 46 published years, asserting the weekday of every boundary
as well as the date -- a rule that is off by one day still matches on some years."
```

---

### Task 3: Year length — `FiscalYearWeeks`

**Files:**
- Modify: `FiscalCalendar.vb`
- Modify: `tests/Test-FiscalCalendar.ps1`

**Interfaces:**
- Consumes: `FiscalYearEnd` (Task 2).
- Produces: `FiscalYearWeeks(ByVal fiscalYear As Long) As Long` — returns 52 or 53.

- [ ] **Step 1: Write the failing tests**

Insert inside the `foreach ($row in $fixture)` loop in `tests/Test-FiscalCalendar.ps1`:

```powershell
        Assert-Equal -Expected $row.weeks `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearWeeks' -Arguments @($fy)) `
                     -Because "FY$fy week count"
```

And after the loop, assert the 53-week set explicitly. This is belt-and-braces over the loop: it fails loudly if the fixture is ever edited to drop a 53-week year, which would otherwise weaken the suite silently.

```powershell
    $expected53 = @(2006, 2012, 2017, 2023, 2028, 2034, 2040, 2045)
    $actual53 = @($fixture | Where-Object {
        (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalYearWeeks' -Arguments @($_.fiscal_year)) -eq 53
    } | ForEach-Object { $_.fiscal_year })
    Assert-Equal -Expected ($expected53 -join ',') -Actual ($actual53 -join ',') `
                 -Because '53-week years are exactly the eight known ones'
```

- [ ] **Step 2: Run the tests to verify they fail**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: a COM error naming `FiscalYearWeeks` as unavailable.

- [ ] **Step 3: Implement `FiscalYearWeeks`**

Append to `FiscalCalendar.vb`:

```vb
' 52 or 53. A year is long exactly when its anchor-to-anchor span is 371 days rather than 364 --
' which is how the leap-year drift the anchor absorbs becomes visible as a whole extra week.
'
' Deliberately NOT a periodicity rule. The long years are 2006, 2012, 2017, 2023, 2028, 2034,
' 2040, 2045 -- gaps of 6,5,6,5,6,6,5. That is not a 5-6 alternation (2028 -> 2034 -> 2040 is two
' consecutive sixes), so anything counting years since the last long one is wrong twice this century.
Public Function FiscalYearWeeks(ByVal fiscalYear As Long) As Long
    Dim spanDays As Long
    spanDays = CLng(FiscalYearEnd(fiscalYear) - FiscalYearEnd(fiscalYear - 1))
    FiscalYearWeeks = spanDays \ 7
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: `passed: 278  failed: 0` (231 from Task 2, plus 46 week counts, plus the 53-week set assertion), then `ALL PASS`.

- [ ] **Step 5: Commit**

```bash
git add FiscalCalendar.vb tests/Test-FiscalCalendar.ps1
git commit -m "feat(FiscalCalendar): derive 52- vs 53-week years from the anchor span

A year is long exactly when its anchor-to-anchor span is 371 days rather than 364.
Asserts the eight long years as a set as well as per-year, so editing the fixture
cannot silently weaken the suite.

Comments record why this is not a periodicity rule: the gaps run 6,5,6,5,6,6,5, so
counting years since the last long one breaks twice this century."
```

---

### Task 4: Month lengths and month starts

**Files:**
- Modify: `FiscalCalendar.vb`
- Modify: `tests/Test-FiscalCalendar.ps1`

**Interfaces:**
- Consumes: `FiscalYearStart`, `FiscalYearWeeks` (Tasks 2-3).
- Produces:
  - `FiscalMonthWeeks(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Long` — `fiscalMonth` 1-12, 1 = February.
  - `FiscalMonthStart(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Date`
  - `FiscalWeekStart(ByVal fiscalYear As Long, ByVal weekNumber As Long) As Date`

- [ ] **Step 1: Write the failing tests**

Append inside the `try` block of `tests/Test-FiscalCalendar.ps1`, after the fixture loop:

```powershell
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
    Assert-Equal -Expected '2026-08-30' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 7))  -Because 'FY2026 September starts'
    Assert-Equal -Expected '2027-01-03' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalMonthStart' -Arguments @(2026, 12)) -Because 'FY2026 January starts'

    # Week starts: first, a mid-year one, and the last week of a 52-week year.
    Assert-Equal -Expected '2026-02-01' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2026, 1))  -Because 'FY2026 week 1 start'
    Assert-Equal -Expected '2026-09-27' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2026, 35)) -Because 'FY2026 week 35 start'
    Assert-Equal -Expected '2027-01-24' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'FiscalWeekStart' -Arguments @(2026, 52)) -Because 'FY2026 week 52 start'
```

Note on the September figure: FY2026's fiscal September starts **2026-08-30**, which is exactly the sort of month boundary that makes this module worth having.

- [ ] **Step 2: Run the tests to verify they fail**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: a COM error naming `FiscalMonthWeeks` as unavailable.

- [ ] **Step 3: Implement the month and week functions**

Append to `FiscalCalendar.vb`:

```vb
' Weeks in a fiscal month. fiscalMonth is 1-12 where 1 = February, 12 = January.
'
' The cycle is 4,5,4 per quarter, so the 5-week months are the second of each quarter --
' fiscal months 2, 5, 8, 11, i.e. those where fiscalMonth Mod 3 = 2. In a 53-week year the
' extra week lands on January (month 12), making it 5 and its quarter 14 weeks.
Public Function FiscalMonthWeeks(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Long
    If fiscalMonth < 1 Or fiscalMonth > 12 Then
        Err.Raise 5, "FiscalMonthWeeks", _
            "fiscalMonth must be 1-12 (1 = February, 12 = January); got " & fiscalMonth & "."
    End If

    Dim weeks As Long
    If fiscalMonth Mod 3 = 2 Then
        weeks = 5
    Else
        weeks = 4
    End If

    If fiscalMonth = 12 Then
        If FiscalYearWeeks(fiscalYear) = 53 Then weeks = weeks + 1
    End If

    FiscalMonthWeeks = weeks
End Function

' First day (a Sunday) of a fiscal month.
Public Function FiscalMonthStart(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Date
    If fiscalMonth < 1 Or fiscalMonth > 12 Then
        Err.Raise 5, "FiscalMonthStart", _
            "fiscalMonth must be 1-12 (1 = February, 12 = January); got " & fiscalMonth & "."
    End If

    Dim weeksBefore As Long
    Dim i As Long
    For i = 1 To fiscalMonth - 1
        weeksBefore = weeksBefore + FiscalMonthWeeks(fiscalYear, i)
    Next i

    FiscalMonthStart = FiscalYearStart(fiscalYear) + (weeksBefore * 7)
End Function

' First day (a Sunday) of a fiscal week. weekNumber is 1-based and must fall inside the year,
' which is 52 or 53 weeks long -- asking for week 53 of a 52-week year is a caller bug, not a
' value to guess at.
Public Function FiscalWeekStart(ByVal fiscalYear As Long, ByVal weekNumber As Long) As Date
    Dim total As Long
    total = FiscalYearWeeks(fiscalYear)

    If weekNumber < 1 Or weekNumber > total Then
        Err.Raise 5, "FiscalWeekStart", _
            "weekNumber must be 1-" & total & " for FY" & fiscalYear & "; got " & weekNumber & "."
    End If

    FiscalWeekStart = FiscalYearStart(fiscalYear) + ((weekNumber - 1) * 7)
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: `passed: 356  failed: 0` (278 from Task 3, plus 12 + 12 month counts, plus 46 sum checks, plus 5 month starts, plus 3 week starts), then `ALL PASS`.

- [ ] **Step 5: Commit**

```bash
git add FiscalCalendar.vb tests/Test-FiscalCalendar.ps1
git commit -m "feat(FiscalCalendar): month lengths, month starts and week starts

The 5-week months are the second of each quarter (fiscalMonth Mod 3 = 2); in a
53-week year January takes the extra week, making its quarter 14. Fiscal month
indexes run 1-12 with 1 = February, matching the order the year runs in.

Asserts that month weeks sum to the year's week count for all 46 fixture years,
which catches an off-by-one in either function that per-month checks would miss.
Out-of-range month or week numbers raise rather than return a guessed date."
```

---

### Task 5: Reverse lookups — date to fiscal year, week and month

**Files:**
- Modify: `FiscalCalendar.vb`
- Modify: `tests/Test-FiscalCalendar.ps1`

**Interfaces:**
- Consumes: `FiscalYearStart`, `FiscalYearEnd`, `FiscalYearWeeks`, `FiscalMonthWeeks` (Tasks 2-4).
- Produces:
  - `FiscalYearOf(ByVal d As Date) As Long`
  - `FiscalWeekOf(ByVal d As Date) As Long`
  - `FiscalMonthOf(ByVal d As Date) As Long` — 1-12, 1 = February.
  - `FiscalWeekStartOf(ByVal d As Date) As Date` — the Sunday beginning the fiscal week containing `d`.

- [ ] **Step 1: Write the failing tests**

Append inside the `try` block of `tests/Test-FiscalCalendar.ps1`:

```powershell
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
```

- [ ] **Step 2: Run the tests to verify they fail**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: a COM error naming `FiscalYearOf` as unavailable.

- [ ] **Step 3: Implement the reverse lookups**

Append to `FiscalCalendar.vb`:

```vb
' Which fiscal year a date falls in.
'
' Seeded from the calendar year and then corrected, rather than searched. A date in January or
' early February can belong to the previous fiscal year, and a late-January date can already
' belong to the next one, so both directions need checking -- one-sided correction is the bug
' waiting to happen here.
Public Function FiscalYearOf(ByVal d As Date) As Long
    Dim fy As Long
    fy = Year(d)

    If d < FiscalYearStart(fy) Then
        fy = fy - 1
    ElseIf d > FiscalYearEnd(fy) Then
        fy = fy + 1
    End If

    FiscalYearOf = fy
End Function

' 1-based fiscal week number of a date, within its own fiscal year.
Public Function FiscalWeekOf(ByVal d As Date) As Long
    Dim fy As Long
    fy = FiscalYearOf(d)
    FiscalWeekOf = ((CLng(d) - CLng(FiscalYearStart(fy))) \ 7) + 1
End Function

' The Sunday that begins the fiscal week containing a date. This is the value the guard
' tracker stores in its FISCAL WEEK column, and the value the PO key encodes.
Public Function FiscalWeekStartOf(ByVal d As Date) As Date
    Dim fy As Long
    fy = FiscalYearOf(d)
    FiscalWeekStartOf = FiscalWeekStart(fy, FiscalWeekOf(d))
End Function

' Fiscal month of a date: 1-12 where 1 = February, 12 = January.
'
' Walks the month lengths rather than dividing, because month lengths are not uniform -- and in
' a 53-week year the last month is longer still.
Public Function FiscalMonthOf(ByVal d As Date) As Long
    Dim fy As Long
    fy = FiscalYearOf(d)

    Dim week As Long
    week = FiscalWeekOf(d)

    Dim m As Long
    Dim cumulative As Long
    For m = 1 To 12
        cumulative = cumulative + FiscalMonthWeeks(fy, m)
        If week <= cumulative Then
            FiscalMonthOf = m
            Exit Function
        End If
    Next m

    ' Unreachable while month weeks sum to the year's week count, which the tests assert for
    ' every published year. Raising beats returning 0 and letting a caller index an array with it.
    Err.Raise 5, "FiscalMonthOf", _
        "Could not place " & Format$(d, "yyyy-mm-dd") & " in FY" & fy & " (week " & week & ")."
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FiscalCalendar.ps1
```

Expected: `passed: 668  failed: 0` (356 from Task 4, plus 46 × 6 round-trip assertions, plus 9 known FY2026 points, plus 35 guards-tracker weeks), then `ALL PASS`.

Recount if the total differs with zero failures — a skipped loop is a silent hole.

- [ ] **Step 5: Commit**

```bash
git add FiscalCalendar.vb tests/Test-FiscalCalendar.ps1
git commit -m "feat(FiscalCalendar): resolve a date to fiscal year, week and month

FiscalYearOf seeds from the calendar year then corrects in BOTH directions: a
January date can belong to the previous fiscal year and a late-January date to the
next, so one-sided correction is wrong. Asserted at every fixture boundary and the
day either side of it.

FiscalMonthOf walks month lengths rather than dividing, since months are 4 or 5
weeks and January is longer again in a 53-week year.

Also asserts the 35 fiscal week values the guards tracker holds resolve to FY2026
weeks 1-35, which is the case this module was written for."
```

---

### Task 6: Prove it composes into a real stack

The module is worthless if adding it breaks the assembler. This is a separate task because it can fail for reasons entirely unrelated to calendar logic — a filename collision, a hoisting surprise, a duplicate procedure name against an existing tenant module.

**Files:**
- Modify: `docs/superpowers/plans/2026-09-11-fiscal-calendar.md` (tick boxes only)
- No source changes expected. If one is needed, it is a finding — record it in the commit message.

**Interfaces:**
- Consumes: `FiscalCalendar.vb` complete (Tasks 2-5).
- Produces: confirmation that `FiscalCalendar.vb` stacks cleanly into an existing variant.

- [ ] **Step 1: Confirm no filename collision exists**

The shadow rule keys on leaf filename, case-insensitively, across the whole recursive tree. A variant file also named `FiscalCalendar.vb` would silently shadow core's and win.

```powershell
Get-ChildItem -Path 'C:\Users\p4bn\Documents\SecuritasAutomation','C:\Users\p4bn\Documents\JCI-invoice-tracker','C:\Users\p4bn\Documents\InvoiceTrackerCore' -Recurse -File -Filter '*.vb' |
    Where-Object { $_.Name -ieq 'FiscalCalendar.vb' } |
    Select-Object FullName
```

Expected: exactly one row, `InvoiceTrackerCore\FiscalCalendar.vb`. More than one means a collision — stop and rename before continuing.

- [ ] **Step 2: Run the stacker for Securitas**

Securitas is the right canary: its workbook already runs the assembled stack, so a problem shows up against real tenant code rather than a toy.

`Stack-VBFiles.ps1` ends in a bare `Read-Host` to hold the window open when double-clicked, so an automated caller **must** redirect stdin or it hangs waiting for Enter.

```powershell
cd C:\Users\p4bn\Documents\SecuritasAutomation
cmd /c "powershell -NoProfile -ExecutionPolicy Bypass -File .\Stack-VBFiles.ps1 -CorePath ..\InvoiceTrackerCore < NUL"
```

Expected: completes without throwing, and its output mentions `FiscalCalendar.vb` among the stacked modules. Exit code 0.

- [ ] **Step 3: Confirm the module landed in the output**

```powershell
Select-String -Path 'C:\Users\p4bn\Documents\SecuritasAutomation\Securitas-Invoice-Tracker_MegaStack.vb' `
              -Pattern 'from FiscalCalendar\.vb', 'Public Function FiscalYearEnd' |
    Select-Object LineNumber, Line
```

Expected: both patterns found. The `from FiscalCalendar.vb` marker is the assembler's own provenance comment.

- [ ] **Step 4: Confirm nothing was hoisted and no name collides**

Check the stacker's own report from Step 2's output:

- `FiscalCalendar.vb` must **not** appear in any "hoisted declarations from" list. It has no module-level declarations by design; if it is listed, something was added that should not have been.
- No duplicate-filename throw occurred (Step 2 would have failed).
- The stacker does **not** check for duplicate *procedure* names — that only surfaces on compile in the VBE. So grep for the exported names appearing more than once:

```powershell
$stack = 'C:\Users\p4bn\Documents\SecuritasAutomation\Securitas-Invoice-Tracker_MegaStack.vb'
'FiscalYearEnd','FiscalYearStart','FiscalYearWeeks','FiscalMonthWeeks','FiscalMonthStart',
'FiscalWeekStart','FiscalYearOf','FiscalWeekOf','FiscalMonthOf','FiscalWeekStartOf',
'SaturdayClosestTo','FiscalCalendarSelfCheck' | ForEach-Object {
    $n = ([regex]::Matches((Get-Content $stack -Raw), "(?m)^(Public\s+)?Function\s+$_\b")).Count
    [pscustomobject]@{ Name = $_; Definitions = $n }
} | Format-Table -AutoSize
```

Expected: `Definitions = 1` for every name. Any 2 is a collision with existing tenant code and must be resolved by renaming the new function (core is the newcomer here, so core yields).

- [ ] **Step 5: Assert against the assembled stack, in a document module**

Everything up to here tested `FiscalCalendar.vb` injected into a **standard** module. Production is `ThisWorkbook`, a **document** module, which forbids `Public Const`, public fixed-size arrays, fixed-length strings and `Declare`. This module was written to avoid all of those — but "was written to" is not "was verified to".

Paste the assembled stack into `ThisWorkbook` of a **scratch copy** of a tracker workbook (never a live one), Debug > Compile, then call into it over the Excel MCP. `run_macro` returns the value of a `Function` in a document module, reporting `invoked_via: "com_method"`. Requires Excel's Trust Center VBA access (on, on this machine).

| Call | Args | Expected |
| --- | --- | --- |
| `ThisWorkbook.FiscalCalendarSelfCheck` | — | `ok` |
| `ThisWorkbook.FiscalYearWeeks` | `[2026]` | `52` |
| `ThisWorkbook.FiscalYearWeeks` | `[2023]` | `53` |
| `ThisWorkbook.FiscalMonthWeeks` | `[2026, 7]` | `4` |
| `ThisWorkbook.FiscalMonthWeeks` | `[2023, 12]` | `5` |
| `ThisWorkbook.FiscalWeekOf` | `["2026-09-27"]` | `35` |
| `ThisWorkbook.FiscalMonthOf` | `["2026-08-30"]` | `8` |

The last two are the ones worth the trouble: fiscal September starting in calendar August is the case the whole module exists for, and it now gets checked in the environment the code actually runs in.

If a date argument fails to marshal, pass the Excel serial number instead — the date-parsing path is already covered by the PowerShell suite, so this table only needs to prove the module works *here*.

- [ ] **Step 6: Restore the stack file**

Step 2 overwrote a tracked build artifact. Leave the repo as it was found — this plan's deliverable is core's module, not a regenerated Securitas artifact.

```powershell
cd C:\Users\p4bn\Documents\SecuritasAutomation
git checkout -- Securitas-Invoice-Tracker_MegaStack.vb
git status --short
```

Expected: clean, or at least no modification to the megastack.

- [ ] **Step 7: Commit the verification note**

No source changed, so commit the ticked plan and record what was proven.

```bash
git add docs/superpowers/plans/2026-09-11-fiscal-calendar.md
git commit -m "docs: confirm FiscalCalendar stacks and runs inside a document module

Stacked against SecuritasAutomation: no filename collision, module present in the
output with its provenance marker, no declarations hoisted, and every exported
procedure name defined exactly once. The stacker does not check for duplicate
procedure names, so that last check was done by hand.

Then asserted against the assembled stack pasted into ThisWorkbook of a scratch
workbook, since the unit tests inject into a standard module and a document module
forbids things a standard module allows. Fiscal September starting in calendar
August is now checked in the environment the code actually runs in.

Noted for future automation: Stack-VBFiles.ps1 ends in a bare Read-Host, so a
scripted caller must redirect stdin (< NUL) or it hangs."
```

---

## Notes for whoever executes this

**The fixture is already committed** at `fiscal-calendar-fixture.csv` on branch `fiscal-calendar` (commit `6d935ac`). Work on that branch. Do not regenerate the fixture from the SharePoint workbook — it has been extracted and verified once, and re-extracting it introduces a chance of changing the oracle to match a bug.

**Assertion counts in the "expected" lines are cumulative** and assume no test is removed. If you add assertions, the counts shift; the number matters less than `failed: 0` plus a total that moves in the direction you expect.

**`AddNewWeeks` is not in this plan.** It is the first real consumer of this module and belongs to the guards tracker, deliberately deferred. This plan ships the calendar alone, which is why it has no regression surface: nothing calls it yet.
