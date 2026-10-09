# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-LastRealDataRow.ps1
#
# Covers LastRealDataRow.vb, added after a live "blank cell" rejection in AddNewBills traced
# to a vendor's "Grand Total" footer row landing in the STORE # column itself -- the mirror
# image of the shape the 2026-09-29 fix (BILL CODE blank, STORE # carrying the label) already
# handled. That fix anchored on a single column; this one judges the row's shape instead, so
# it does not need to know which column a future vendor's export will use.
#
# Invoke-VbaFunction supports at most 2 arguments (Application.Run's own ceiling), so every
# driver function below takes exactly one or two, packing extra values into a '|'/'~'-delimited
# string where needed.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$sources = @(
    (Join-Path $repo 'LastRealDataRow.vb'),
    (Join-Path $repo 'CoerceValues.vb')
)

function Set-Cell($ws, [int]$row, [int]$col, $value) {
    $cell = $ws.Cells($row, $col)
    if ($null -eq $value) {
        $cell.ClearContents() | Out-Null
    } elseif ($value -is [string]) {
        $cell.Value2 = $value
    } else {
        $cell.Value2 = [double]$value
    }
}

# --- RowValuesLookLikeSummary: pure signal combinations -------------------------------------
$pureDriver = Join-Path $env:TEMP 'LrdrPureDriver.vb'
@'
Public Function T_RowLooksLikeSummary(ByVal packed As String, ByVal maxAmount As Double) As Boolean
    Dim parts As Variant
    parts = Split(packed, "|")
    T_RowLooksLikeSummary = RowValuesLookLikeSummary(parts, maxAmount)
End Function
'@ | Set-Content -LiteralPath $pureDriver -Encoding UTF8

$vba = New-VbaHost -SourceFiles ($sources + $pureDriver)
try {
    $r = { param($packed, $max) Invoke-VbaFunction -VbaHost $vba -Name 'T_RowLooksLikeSummary' -Arguments @($packed, [double]$max) -TimeoutSeconds 30 }

    # Exact live shape: label + mostly blank + the block's largest amount -> 3 signals.
    Assert-Equal $true (& $r 'Grand Total|||||||||||||40842.95|1444.15|42287.10' 42287.10) `
        -Because 'label + sparse + largest amount together is a summary row (the live fixture''s exact shape)'

    # Label alone, dense row, modest amount -> 1 signal, not a summary row.
    Assert-Equal $false (& $r 'Total|B|C|D|E|F|G|H|I|J|K|L|M|10.00|1.00|11.00' 500) `
        -Because 'a label by itself, on an otherwise fully-populated row, is not enough on its own'

    # Sparse row, no label, no notable amount -> 1 signal only.
    Assert-Equal $false (& $r '||||||||||||||||' 500) `
        -Because 'sparseness alone (no label, no amount) is only one signal'

    # Dense real row that happens to hold the batch's largest single amount, no label -> 1 signal.
    Assert-Equal $false (& $r '234|NORDSTROM INC|4000 WORTH AVE|COLUMBUS|OH|43219|7/22/2026|7/12/2026|08032026|LP REPAIR|notes|9/3/2026|6200042863|330.00|26.40|356.40' 356.40) `
        -Because 'a real, fully-populated row is not a summary row just for holding the largest single amount'

    # "total" as a substring inside free text, not an exact label match -> no label signal;
    # combined with a dense row and a non-maximal amount, stays a real row.
    Assert-Equal $false (& $r '234|NORDSTROM INC|4000 WORTH AVE|COLUMBUS|OH|43219|7/22/2026|7/12/2026|08032026|LP REPAIR|Replaced totaled panel|9/3/2026|6200042863|330.00|26.40|356.40' 999999) `
        -Because 'the substring "total" inside a free-text note must not trigger the exact-label signal'

    # Label + sparse, modest amount (not the block max) -> 2 signals, still a summary row.
    Assert-Equal $true (& $r 'Subtotal||||||||||||||10.00|' 999999) `
        -Because 'label + sparse is two signals even when the amount is not the block max'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }
Remove-Item $pureDriver -Force -ErrorAction SilentlyContinue

# --- Worksheet entry point ------------------------------------------------------------------
$wsDriver = Join-Path $env:TEMP 'LrdrWsDriver.vb'
@'
Public Function T_Worksheet(ByVal packed As String) As Long
    Dim parts As Variant
    parts = Split(packed, "|")
    If UBound(parts) >= 3 Then
        Dim names As Variant, mc() As Variant, i As Long
        names = Split(parts(3), ",")
        ReDim mc(0 To UBound(names))
        For i = 0 To UBound(names)
            mc(i) = CLng(names(i))
        Next i
        T_Worksheet = LastRealDataRow(ThisWorkbook.Worksheets(1), CLng(parts(0)), CLng(parts(1)), 1, CLng(parts(2)), mc)
    Else
        T_Worksheet = LastRealDataRow(ThisWorkbook.Worksheets(1), CLng(parts(0)), CLng(parts(1)), 1, CLng(parts(2)))
    End If
End Function
'@ | Set-Content -LiteralPath $wsDriver -Encoding UTF8

# Live fixture shape: "Grand Total" in the first (anchor) column.
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-Cell $ws 1 1 'STORE #'; Set-Cell $ws 1 2 'BILL CODE'; Set-Cell $ws 1 3 'SUBTOTAL'
    Set-Cell $ws 2 1 234; Set-Cell $ws 2 2 6200042863; Set-Cell $ws 2 3 330.00
    Set-Cell $ws 3 1 600; Set-Cell $ws 3 2 6200041701; Set-Cell $ws 3 3 416.28
    Set-Cell $ws 4 1 'Grand Total'; Set-Cell $ws 4 2 $null; Set-Cell $ws 4 3 746.28

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('4|1|3') -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'the Grand Total row (label in STORE #, the anchor column) is excluded'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# Original bug shape: label in a DIFFERENT column, anchor column blank.
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-Cell $ws 1 1 'STORE #'; Set-Cell $ws 1 2 'BILL CODE'; Set-Cell $ws 1 3 'SUBTOTAL'
    Set-Cell $ws 2 1 234; Set-Cell $ws 2 2 6200042863; Set-Cell $ws 2 3 330.00
    Set-Cell $ws 3 1 600; Set-Cell $ws 3 2 6200041701; Set-Cell $ws 3 3 416.28
    Set-Cell $ws 4 1 $null; Set-Cell $ws 4 2 'Grand Total'; Set-Cell $ws 4 3 746.28

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('4|1|3') -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'the original bug shape (label in BILL CODE, STORE # blank) is still handled -- not a column-specific fix'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# Multiple trailing rows: a blank row, then a subtotal row, then the real data.
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-Cell $ws 1 1 'STORE #'; Set-Cell $ws 1 2 'BILL CODE'; Set-Cell $ws 1 3 'SUBTOTAL'
    Set-Cell $ws 2 1 234; Set-Cell $ws 2 2 6200042863; Set-Cell $ws 2 3 330.00
    Set-Cell $ws 3 1 600; Set-Cell $ws 3 2 6200041701; Set-Cell $ws 3 3 416.28
    Set-Cell $ws 4 1 'Subtotal'; Set-Cell $ws 4 2 $null; Set-Cell $ws 4 3 746.28
    Set-Cell $ws 5 1 $null; Set-Cell $ws 5 2 $null; Set-Cell $ws 5 3 $null

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('5|1|3') -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'a trailing blank row and a trailing subtotal row are both walked past'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# No footer at all: the raw guess is already correct.
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-Cell $ws 1 1 'STORE #'; Set-Cell $ws 1 2 'BILL CODE'; Set-Cell $ws 1 3 'SUBTOTAL'
    Set-Cell $ws 2 1 234; Set-Cell $ws 2 2 6200042863; Set-Cell $ws 2 3 330.00
    Set-Cell $ws 3 1 600; Set-Cell $ws 3 2 6200041701; Set-Cell $ws 3 3 416.28

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('3|1|3') -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'with no footer present, the raw last row is returned unchanged'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# False-positive guard: a real last row with the batch's largest amount, fully dense, no
# label -- must NOT be excluded.
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-Cell $ws 1 1 'STORE #'; Set-Cell $ws 1 2 'BILL CODE'; Set-Cell $ws 1 3 'SUBTOTAL'
    Set-Cell $ws 2 1 234; Set-Cell $ws 2 2 6200042863; Set-Cell $ws 2 3 330.00
    Set-Cell $ws 3 1 600; Set-Cell $ws 3 2 6200041701; Set-Cell $ws 3 3 9999.99  # priciest single repair in the batch

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('3|1|3') -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'a real, fully-populated last row is not excluded just for holding the largest single amount'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# Live shape (2026-10-08, Securitas July RMR re-send): the footer carries NO label and is blank
# everywhere except the three money columns, so the label signal never fires and "sparse" is the
# only signal left -- unless the largest-amount signal works. It could not: the max-amount scan
# read EVERY column, and a text invoice # like "0906964326" (906,964,326) and date serials
# (46,240) dwarf the footer's 27,511.33, so the footer was never "the largest amount" and the
# row was treated as real (STORE # blank -> "blank cell" rejection).
function Set-LiveFooterFixture($ws) {
    # Columns: 1 STORE #, 2 INV DATE, 3 INVOICE # (text), 4 SUBTOTAL, 5 SALES TAX, 6 TOTAL
    $ws.Columns(3).NumberFormat = '@'
    Set-Cell $ws 1 1 'STORE #'; Set-Cell $ws 1 2 'INV DATE'; Set-Cell $ws 1 3 'INVOICE #'
    Set-Cell $ws 1 4 'SUBTOTAL'; Set-Cell $ws 1 5 'SALES TAX'; Set-Cell $ws 1 6 'TOTAL'
    Set-Cell $ws 2 1 '0001'; Set-Cell $ws 2 2 46240; Set-Cell $ws 2 3 '0906964326'
    Set-Cell $ws 2 4 93.61; Set-Cell $ws 2 5 9.89; Set-Cell $ws 2 6 103.50
    Set-Cell $ws 3 1 '7771'; Set-Cell $ws 3 2 46240; Set-Cell $ws 3 3 '0906964326'
    Set-Cell $ws 3 4 52.00; Set-Cell $ws 3 5 0; Set-Cell $ws 3 6 52.00
    # Row 4: label-less footer, money columns only. Rows 5-6: blank tail UsedRange over-includes.
    Set-Cell $ws 4 4 145.61; Set-Cell $ws 4 5 9.89; Set-Cell $ws 4 6 155.50
}

$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    Set-LiveFooterFixture $vba.Workbook.Worksheets(1)
    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('6|1|6|4,5,6') -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'a label-less footer holding only the money columns is excluded when the caller names its money columns'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# Money columns scoped: a big identifier or date elsewhere in the row must not make the footer
# look small. Without the money-column list, behavior is unchanged (old all-column scan).
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    Set-LiveFooterFixture $vba.Workbook.Worksheets(1)
    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('6|1|6') -TimeoutSeconds 30
    Assert-Equal 4 $result -Because 'without a money-column list the legacy all-column scan is kept (documents why callers should pass one)'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# False-positive guard with money columns named: a dense real last row that holds the batch's
# largest money amount is still a real row (only one signal).
$vba = New-VbaHost -SourceFiles ($sources + $wsDriver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    Set-LiveFooterFixture $ws
    Set-Cell $ws 4 1 '9000'; Set-Cell $ws 4 2 46240; Set-Cell $ws 4 3 '0906964326'
    Set-Cell $ws 4 4 5000; Set-Cell $ws 4 5 0; Set-Cell $ws 4 6 5000
    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Worksheet' -Arguments @('6|1|6|4,5,6') -TimeoutSeconds 30
    Assert-Equal 4 $result -Because 'a fully populated last row is not excluded just for holding the largest money amount'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }
Remove-Item $wsDriver -Force -ErrorAction SilentlyContinue

# --- Array entry point: parity with the worksheet path on the live fixture's shape ----------
$arrDriver = Join-Path $env:TEMP 'LrdrArrDriver.vb'
@'
Public Function T_Array(ByVal packed As String) As Long
    Dim fields As Variant
    fields = Split(packed, "~")
    Dim packedRows As String, headerRow As Long, lastCol As Long
    packedRows = fields(0)
    headerRow = CLng(fields(1))
    lastCol = CLng(fields(2))

    Dim rowsSplit As Variant
    rowsSplit = Split(packedRows, ";")

    Dim data() As Variant
    ReDim data(1 To UBound(rowsSplit) + 1, 1 To lastCol)

    Dim r As Long, c As Long, cells As Variant
    For r = 0 To UBound(rowsSplit)
        cells = Split(rowsSplit(r), "|")
        For c = 0 To UBound(cells)
            data(r + 1, c + 1) = cells(c)
        Next c
    Next r

    If UBound(fields) >= 3 Then
        Dim names As Variant, mc() As Variant, i As Long
        names = Split(fields(3), ",")
        ReDim mc(0 To UBound(names))
        For i = 0 To UBound(names)
            mc(i) = CLng(names(i))
        Next i
        T_Array = LastRealDataRowInArray(data, UBound(data, 1), headerRow, 1, lastCol, mc)
    Else
        T_Array = LastRealDataRowInArray(data, UBound(data, 1), headerRow, 1, lastCol)
    End If
End Function
'@ | Set-Content -LiteralPath $arrDriver -Encoding UTF8

$vba = New-VbaHost -SourceFiles ($sources + $arrDriver)
try {
    # Row 1 = header, rows 2-3 = data, row 4 = Grand Total in the first column.
    $packed = 'STORE #|BILL CODE|SUBTOTAL;234|6200042863|330.00;600|6200041701|416.28;Grand Total||746.28~1~3'
    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Array' -Arguments @($packed) -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'the array entry point (JCI''s shape) excludes the same Grand Total row the worksheet entry point does'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles ($sources + $arrDriver)
try {
    # Live footer shape: label-less, money columns only; big text invoice # and date serial.
    $packed = 'STORE #|INV DATE|INVOICE #|SUBTOTAL|SALES TAX|TOTAL;0001|46240|0906964326|93.61|9.89|103.50;7771|46240|0906964326|52.00|0|52.00;||||9.89|155.50~1~6~4,5,6'
    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Array' -Arguments @($packed) -TimeoutSeconds 30
    Assert-Equal 3 $result -Because 'the array entry point also excludes a label-less money-only footer when money columns are named'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }
Remove-Item $arrDriver -Force -ErrorAction SilentlyContinue

Write-AssertSummary
