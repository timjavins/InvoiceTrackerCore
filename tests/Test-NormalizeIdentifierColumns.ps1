# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-NormalizeIdentifierColumns.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

# NormalizeIdentifierColumns.vb (both the Sub and the Function) is one module, and VBA compiles
# a module as a whole once anything in it is entered -- so the Sub's callees (TenantSheetName,
# TenantIdentifierColumns, TenantColLetter, UnprotectSheet, ProtectSheet) must be resolvable even
# though the tests below only ever call NormalizeIdentifier / NormalizeIdentifierColumnsOn
# directly. Same reasoning as LookupReqs needing UnprotectSheet/ProtectSheet stubs -- see
# tests/README.md.
$driver = Join-Path $env:TEMP 'NicDriver.vb'
@'
Public Function TenantSheetName(ByVal role As String) As String
    TenantSheetName = "Sheet1"
End Function

Public Function TenantIdentifierColumns() As Variant
    TenantIdentifierColumns = Array()
End Function

Public Function TenantColLetter(ByVal concept As String) As String
    TenantColLetter = "A"
End Function

Public Sub UnprotectSheet()
End Sub

Public Sub ProtectSheet()
End Sub

Public Function T_Normalize(ByVal v As Variant) As String
    T_Normalize = DescribeVariant(NormalizeIdentifier(v))
End Function

' Empty cannot cross the COM boundary reliably as an argument (a passed $null marshals as
' Null, not Empty), so this calls with an uninitialized Variant instead, which VBA gives Empty.
Public Function T_NormalizeEmpty() As String
    Dim v As Variant
    T_NormalizeEmpty = DescribeVariant(NormalizeIdentifier(v))
End Function

Private Function DescribeVariant(ByVal r As Variant) As String
    If IsEmpty(r) Then
        DescribeVariant = "(empty)"
    ElseIf IsNull(r) Then
        DescribeVariant = "(null)"
    Else
        DescribeVariant = CStr(r)
    End If
End Function

Public Function T_Sweep(ByVal colsJoined As String) As Long
    T_Sweep = NormalizeIdentifierColumnsOn(ThisWorkbook.Worksheets(1), Split(colsJoined, "|"))
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

$src = Join-Path $repo 'NormalizeIdentifierColumns.vb'

# --- Pure NormalizeIdentifier cases -------------------------------------------------------
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $r = { param($n, $a) Invoke-VbaFunction -VbaHost $vba -Name $n -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal '914751' (& $r 'T_Normalize' @([double]914751)) -Because 'a whole number becomes plain digit text'
    Assert-Equal '8059053607' (& $r 'T_Normalize' @([double]8059053607)) -Because 'a large whole number never renders as scientific notation'
    Assert-Equal '810492' (& $r 'T_Normalize' @(' 810492 ')) -Because 'a padded numeric-looking string is trimmed'
    Assert-Equal 'WARRANTY' (& $r 'T_Normalize' @('WARRANTY')) -Because 'a non-numeric marker is left alone'
    Assert-Equal '(empty)' (Invoke-VbaFunction -VbaHost $vba -Name 'T_NormalizeEmpty' -TimeoutSeconds 30) -Because 'Empty passes through unchanged'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Sheet sweep: ordinary (unfiltered) fixture -------------------------------------------
# A numeric-looking string ("1166552") needs NumberFormat = "@" set *before* the value write,
# or Excel re-numericizes it on assignment, exactly the trap guarantee #5 exists to avoid --
# that pre-step is fixture setup, not a mask. A non-numeric string (a WARRANTY-style marker)
# is never re-numericized regardless of format, so it is deliberately left with whatever
# format the cell already has (General, for a fresh fixture cell) instead of forcing "@" --
# forcing "@" on every string cell up front would make a spurious sweep-caused re-set to "@"
# indistinguishable from the fixture's own setup, which is exactly the gap that let a real
# regression (a blanket NumberFormat write over cells that never needed one) through review
# undetected. Blank cells are left with their default format for the same reason.
function Set-Cell($ws, [int]$row, [int]$col, $value, [switch]$AsFormula) {
    $cell = $ws.Cells($row, $col)
    if ($AsFormula) {
        $cell.Formula = $value
    } elseif ($value -is [string]) {
        if ($value -match '^\s*-?\d+(\.\d+)?\s*$') {
            $cell.NumberFormat = '@'
        }
        $cell.Value2 = $value
    } elseif ($null -eq $value) {
        $cell.ClearContents() | Out-Null
    } else {
        $cell.Value2 = [double]$value
    }
}

$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)

    # Column A: numeric -> formula -> WARRANTY -> blank -> already-normalized text
    Set-Cell $ws 1 1 'ID'
    Set-Cell $ws 2 1 914751          # numeric: must become text "914751"
    Set-Cell $ws 3 1 '=1+1' -AsFormula   # formula: must never be touched
    Set-Cell $ws 4 1 'WARRANTY'       # marker text: already normalized, must be untouched
    Set-Cell $ws 5 1 $null            # blank: must stay blank
    Set-Cell $ws 6 1 '1166552'        # already text and already trimmed: not rewritten

    # A deliberately wrong pre-existing format on the already-normalized text cell -- one no
    # real sweep output would ever produce -- proves the sweep issues no format write at all
    # for a cell whose value doesn't change, not merely that the value survives. Set-Cell's own
    # "@" pre-step (needed to store '1166552' as text at all) already ran above; this overrides
    # it afterward, which does not turn the stored string back into a number (verified directly:
    # a NumberFormat change alone never coerces existing cell content).
    $ws.Cells(6,1).NumberFormat = '0.00'
    # WARRANTY (row 4) and the blank (row 5) were never given "@" by Set-Cell in the first
    # place, so they are still whatever a fresh cell defaults to -- also a value the sweep
    # itself would never legitimately write.
    $warrantyFormatBefore = $ws.Cells(4,1).NumberFormat
    $blankFormatBefore = $ws.Cells(5,1).NumberFormat

    # Column C: same shapes, to prove multiple columns in one call are all swept
    Set-Cell $ws 1 3 'ID2'
    Set-Cell $ws 2 3 555555
    Set-Cell $ws 3 3 '=2+2' -AsFormula
    Set-Cell $ws 4 3 $null
    Set-Cell $ws 5 3 'WARRANTY'
    Set-Cell $ws 6 3 '9999'

    $count = Invoke-VbaFunction -VbaHost $vba -Name 'T_Sweep' -Arguments @('A|C') -TimeoutSeconds 30

    Assert-Equal 2 $count -Because 'exactly the two numeric cells (A2, C2) were changed'

    Assert-Equal 'String' ($ws.Cells(2,1).Value2.GetType().Name) -Because 'A2 becomes a text-typed value'
    Assert-Equal '914751' ($ws.Cells(2,1).Value2) -Because 'A2 keeps its digits exactly'
    Assert-Equal '@' ($ws.Cells(2,1).NumberFormat) -Because 'A2 gets @ format before the write'

    Assert-Equal $true ($ws.Cells(3,1).HasFormula) -Because 'a formula cell is never touched'
    Assert-Equal 2 ($ws.Cells(3,1).Value2) -Because 'the formula still calculates'

    Assert-Equal 'WARRANTY' ($ws.Cells(4,1).Value2) -Because 'a marker string is left alone'
    Assert-Equal '' ([string]$ws.Cells(5,1).Value2) -Because 'a blank cell stays blank'
    Assert-Equal '1166552' ($ws.Cells(6,1).Value2) -Because 'an already-normalized text cell is not rewritten'

    # Proving no write happened at all, not just that the value round-tripped unchanged: each
    # of these cells' deliberately-wrong pre-existing format must still be exactly what it was
    # planted as above -- a real sweep write (NicWriteRuns/NicWritePerCell) always sets "@"
    # first, so any of these being "@" now would mean the cell was written despite not needing
    # a change.
    Assert-Equal '0.00' ($ws.Cells(6,1).NumberFormat) -Because 'an already-normalized text cell gets no format write either'
    Assert-Equal $warrantyFormatBefore ($ws.Cells(4,1).NumberFormat) -Because 'a marker string gets no format write'
    Assert-Equal $blankFormatBefore ($ws.Cells(5,1).NumberFormat) -Because 'a blank cell gets no format write'

    Assert-Equal 'String' ($ws.Cells(2,3).Value2.GetType().Name) -Because 'C2 becomes text too -- multiple columns are swept'
    Assert-Equal '555555' ($ws.Cells(2,3).Value2) -Because 'C2 keeps its digits exactly'
    Assert-Equal $true ($ws.Cells(3,3).HasFormula) -Because 'the second column formula cell is never touched either'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Sheet sweep: formula cell in the middle of a literal column -------------------------
# Guarantee #1 (never write a formula cell) must hold even when literal cells needing changes
# surround the formula on both sides. This is also the only case that exercises NicWriteRuns's
# per-run write path, since every currently-configured tenant column has no formula cells at
# all -- the common path is the whole-range write exercised by every other case in this file.
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)

    Set-Cell $ws 1 1 'ID'
    Set-Cell $ws 2 1 111          # numeric, before the formula: must normalize
    Set-Cell $ws 3 1 222          # numeric, before the formula: must normalize
    Set-Cell $ws 4 1 '333'        # already-clean text, before the formula: not rewritten
    Set-Cell $ws 5 1 '=5+5' -AsFormula   # formula, mid-column: must never be touched
    Set-Cell $ws 6 1 444          # numeric, after the formula: must normalize
    Set-Cell $ws 7 1 ' 555 '      # padded string, after the formula: must normalize
    Set-Cell $ws 8 1 'WARRANTY'   # marker text, after the formula: left alone
    Set-Cell $ws 9 1 $null        # blank, after the formula: left alone

    $count = Invoke-VbaFunction -VbaHost $vba -Name 'T_Sweep' -Arguments @('A') -TimeoutSeconds 30

    Assert-Equal 4 $count -Because 'exactly the four cells needing a change were counted, formula excluded'

    Assert-Equal $true ($ws.Cells(5,1).HasFormula) -Because 'the mid-column formula cell is still a formula'
    Assert-Equal '=5+5' ($ws.Cells(5,1).Formula) -Because 'the formula text is unchanged'
    Assert-Equal 10 ($ws.Cells(5,1).Value2) -Because 'the formula still calculates'

    Assert-Equal 'String' ($ws.Cells(2,1).Value2.GetType().Name) -Because 'a numeric cell before the formula becomes text'
    Assert-Equal '111' ($ws.Cells(2,1).Value2) -Because 'A2 keeps its digits exactly'
    Assert-Equal '222' ($ws.Cells(3,1).Value2) -Because 'A3 keeps its digits exactly'
    Assert-Equal '333' ($ws.Cells(4,1).Value2) -Because 'an already-clean text cell before the formula is not rewritten'

    Assert-Equal 'String' ($ws.Cells(6,1).Value2.GetType().Name) -Because 'a numeric cell after the formula becomes text too'
    Assert-Equal '444' ($ws.Cells(6,1).Value2) -Because 'A6 keeps its digits exactly'
    Assert-Equal '555' ($ws.Cells(7,1).Value2) -Because 'a padded string after the formula is trimmed'
    Assert-Equal 'WARRANTY' ($ws.Cells(8,1).Value2) -Because 'a marker string after the formula is left alone'
    Assert-Equal '' ([string]$ws.Cells(9,1).Value2) -Because 'a blank cell after the formula stays blank'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Sheet sweep: larger scale, mixed shapes, and a timing regression guard ---------------
# A generous wall-clock ceiling here is a regression guard against reintroducing a per-cell
# (O(n) COM round trips) sweep by accident -- not a tight performance benchmark.
#
# This fixture has zero formula cells, which is the specific shape that once took a blanket
# NumberFormat="@" + Value=wholeArray write over the *entire* range -- including every
# already-clean-text and blank row that never needed a write at all. Planting a deliberately
# wrong pre-existing format on one of each and asserting it survives the sweep unchanged is
# what catches that regression; checking only the post-sweep *value* does not, since the value
# round-trips to the same thing either way.
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-Cell $ws 1 1 'ID'

    $rowCount = 300
    $expectedChanged = 0
    for ($i = 0; $i -lt $rowCount; $i++) {
        $row = $i + 2
        switch ($i % 3) {
            0 { Set-Cell $ws $row 1 (1000000 + $i); $expectedChanged++ }
            1 { Set-Cell $ws $row 1 ([string]$i) }
            default { Set-Cell $ws $row 1 $null }
        }
    }

    # Row 3 (i=1) is already-clean text ("1"); Set-Cell had to set "@" to store it as text at
    # all, so override it afterward to a format no legitimate sweep output would produce.
    # Row 4 (i=2) is blank and was never given "@" in the first place -- record its natural
    # default so the after-sweep check isn't just hardcoding "General".
    $ws.Cells(3,1).NumberFormat = '0.00'
    $blankFormatBefore = $ws.Cells(4,1).NumberFormat

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $count = Invoke-VbaFunction -VbaHost $vba -Name 'T_Sweep' -Arguments @('A') -TimeoutSeconds 30
    $sw.Stop()

    Assert-Equal $expectedChanged $count -Because 'every numeric row across a few hundred rows was counted, and only those'
    Assert-Equal $true ($sw.Elapsed.TotalSeconds -lt 10) -Because 'the bulk sweep completes well under a generous ceiling, guarding against an accidental per-cell reintroduction'

    Assert-Equal 'String' ($ws.Cells(2,1).Value2.GetType().Name) -Because 'the first numeric row becomes text'
    Assert-Equal '1000000' ($ws.Cells(2,1).Value2) -Because 'the first numeric row keeps its digits exactly'
    Assert-Equal '1' ($ws.Cells(3,1).Value2) -Because 'an already-clean text row is untouched in value'
    Assert-Equal '' ([string]$ws.Cells(4,1).Value2) -Because 'a blank row stays blank'

    Assert-Equal '0.00' ($ws.Cells(3,1).NumberFormat) -Because 'an already-clean text row gets no format write in the no-formula bulk path either'
    Assert-Equal $blankFormatBefore ($ws.Cells(4,1).NumberFormat) -Because 'a blank row gets no format write in the no-formula bulk path either'

    $tailIndex = $rowCount - 1
    while (($tailIndex % 3) -ne 0) { $tailIndex-- }
    $tailRow = $tailIndex + 2
    Assert-Equal 'String' ($ws.Cells($tailRow,1).Value2.GetType().Name) -Because 'the last numeric row in the range also becomes text'
    Assert-Equal ([string](1000000 + $tailIndex)) ($ws.Cells($tailRow,1).Value2) -Because 'the last numeric row keeps its digits exactly'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Sheet sweep: filter-proofing ---------------------------------------------------------
# End(xlUp) and Find both walk visible cells only under an AutoFilter, so either would miss a
# filtered-out trailing row. NormalizeIdentifierColumnsOn instead reads UsedRange's full extent,
# which is unaffected by hidden rows. This fixture hides the sheet's last data row and proves it
# still gets normalized.
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)

    Set-Cell $ws 1 1 'ID'
    Set-Cell $ws 1 2 'TAG'
    Set-Cell $ws 2 1 111111
    Set-Cell $ws 2 2 'show'
    Set-Cell $ws 3 1 222222          # last data row -- will be hidden by the filter
    Set-Cell $ws 3 2 'hide'

    $ws.Range($ws.Cells(1,1), $ws.Cells(3,2)).AutoFilter(2, 'show') | Out-Null
    Assert-Equal $true $ws.FilterMode -Because 'the fixture actually has an active filter'
    Assert-Equal $true $ws.Rows(3).Hidden -Because 'the last data row starts out hidden by the filter'

    $count = Invoke-VbaFunction -VbaHost $vba -Name 'T_Sweep' -Arguments @('A') -TimeoutSeconds 30

    Assert-Equal 2 $count -Because 'both rows were normalized, including the hidden one'
    Assert-Equal 'String' ($ws.Cells(3,1).Value2.GetType().Name) -Because 'the hidden row is normalized to text'
    Assert-Equal '222222' ($ws.Cells(3,1).Value2) -Because 'the hidden row keeps its digits exactly'
    Assert-Equal '@' ($ws.Cells(3,1).NumberFormat) -Because 'the hidden row gets @ format too'
    Assert-Equal $true $ws.Rows(3).Hidden -Because 'the sweep never touches the filter itself'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Document-module compile check ---------------------------------------------------------
# Proves NormalizeIdentifierColumns.vb -- both the Function and the Sub, together, since VBA
# compiles a module as a whole -- survives being pasted into ThisWorkbook, a document (class)
# module, the same kind of module Stack-VBFiles.ps1 pastes it into in every tenant's real
# assembled stack. See tests/README.md.
$vba = New-VbaHost -SourceFiles @($src, $driver) -DocumentModule
try {
    Assert-Equal 0 ($vba.Workbook.T_Sweep('A')) `
                 -Because 'NormalizeIdentifierColumns.vb compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
