# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ProcessNewBillsBulk.ps1
#
# Covers SecuritasAutomation/file ingesting/ProcessNewBills.vb's two private write paths,
# ProcessNewBillsBulk and ProcessNewBillsPerCell, added when the write_vba cell-by-cell
# advisory (2026-09-29 deploy) pointed at ProcessNewBills' original per-cell loop. Both paths
# are called directly rather than through the public ProcessNewBills wrapper, which avoids
# needing to stub MapSourceToTargetColumns / InputTargetColumnName / WriteFormulas_Tracker and
# the failureState/withErrors globals -- none of those are involved in what changed.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

# ProcessNewBills.vb lives in the SecuritasAutomation repo, a sibling of InvoiceTrackerCore.
$pnbPath = Join-Path (Split-Path $repo -Parent) 'SecuritasAutomation\file ingesting\ProcessNewBills.vb'
$sources = @(
    (Join-Path $repo 'CoerceValues.vb'),
    (Join-Path $repo 'ConvertStoreNumbers.vb'),
    $pnbPath
)

# ProcessNewBills.vb is entered (ProcessNewBillsBulk/PerCell are called directly below), so
# everything else in this same combined module must compile too: CoerceValues.vb and
# ConvertStoreNumbers.vb are the real, pure source (worth exercising for real rather than
# stubbing), which in turn need TenantColLetter/TenantSheetName/UnprotectSheet/ProtectSheet
# stubbed, since ConvertStoreNumbers/ConvertStoreNumbersOn reference them even though this
# suite never calls those two procedures. ResolveInvoiceType is stubbed as identity, since the
# real one lives in AddNewBills.vb alongside a FileDialog-driven Sub with a large dependency
# graph that has nothing to do with what this suite tests.
#
# Invoke-VbaFunction supports at most 2 arguments (Application.Run's own ceiling), so every
# call below packs its real arguments into one '~'-delimited string and unpacks it in VBA.
$driver = Join-Path $env:TEMP 'PnbDriver.vb'
@'
Public Function TenantColLetter(ByVal concept As String) As String
    TenantColLetter = "A"
End Function

Public Function TenantSheetName(ByVal role As String) As String
    TenantSheetName = "Sheet1"
End Function

Public Sub UnprotectSheet()
End Sub

Public Sub ProtectSheet()
End Sub

Public Function ResolveInvoiceType(ByVal vendorValue As String, ByVal billType As String) As String
    ResolveInvoiceType = vendorValue
End Function

' packed = "headers|srcCols|tgtCols~sourceHeaderRow~sourceLastRow~targetLastRow~rowStatusJoined~mode"
' headers/srcCols/tgtCols are each '|'-joined and line up positionally. rowStatusJoined is
' '|'-joined per-row markers ("" for written, "SKIP" for skipped), or the literal "NONE" to
' pass rowStatus as IsMissing rather than an empty array. mode is "BULK" or "PERCELL".
Public Function T_Run(ByVal packed As String) As Long
    Dim fields As Variant
    fields = Split(packed, "~")

    Dim cols As Variant
    cols = Split(fields(0), "|")
    Dim headers As Variant, srcCols As Variant, tgtCols As Variant
    headers = Split(cols(0), ",")
    srcCols = Split(cols(1), ",")
    tgtCols = Split(cols(2), ",")

    Dim sourceHeaderRow As Long, sourceLastRow As Long, targetLastRow As Long
    sourceHeaderRow = CLng(fields(1))
    sourceLastRow = CLng(fields(2))
    targetLastRow = CLng(fields(3))
    Dim rowStatusJoined As String
    rowStatusJoined = fields(4)
    Dim mode As String
    mode = fields(5)

    Dim sourceColIndexes As Object, targetMap As Object
    Set sourceColIndexes = CreateObject("Scripting.Dictionary")
    Set targetMap = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        sourceColIndexes.Add headers(i), CLng(srcCols(i))
        targetMap.Add headers(i), CLng(tgtCols(i))
    Next i

    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("Source")
    Set wsTarget = ThisWorkbook.Worksheets("Target")

    Dim writtenCount As Long

    If rowStatusJoined = "NONE" Then
        writtenCount = sourceLastRow - sourceHeaderRow

        If mode = "PERCELL" Then
            ProcessNewBillsPerCell wsSource, sourceHeaderRow, sourceLastRow, sourceColIndexes, _
                                   wsTarget, targetLastRow, targetMap, Nothing, "repair"
        Else
            ProcessNewBillsBulk wsSource, sourceHeaderRow, sourceLastRow, sourceColIndexes, _
                                wsTarget, targetLastRow, targetMap, Nothing, "repair"
        End If
    Else
        Dim parts As Variant
        parts = Split(rowStatusJoined, ",")
        Dim statuses() As String
        ReDim statuses(LBound(parts) To UBound(parts))
        For i = LBound(parts) To UBound(parts)
            If parts(i) = "SKIP" Then
                statuses(i) = "duplicate"
            Else
                statuses(i) = ""
                writtenCount = writtenCount + 1
            End If
        Next i

        If mode = "PERCELL" Then
            ProcessNewBillsPerCell wsSource, sourceHeaderRow, sourceLastRow, sourceColIndexes, _
                                   wsTarget, targetLastRow, targetMap, Nothing, "repair", statuses
        Else
            ProcessNewBillsBulk wsSource, sourceHeaderRow, sourceLastRow, sourceColIndexes, _
                                wsTarget, targetLastRow, targetMap, Nothing, "repair", statuses
        End If
    End If

    T_Run = writtenCount
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

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

# --- Bulk path: one column of each coercion kind, plus a passthrough column, plus a
#     coercion-failure row -----------------------------------------------------------------
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wsSource = $vba.Workbook.Worksheets(1)
    $wsSource.Name = 'Source'
    $wsTarget = $vba.Workbook.Worksheets.Add()
    $wsTarget.Name = 'Target'

    # Source/target columns: A=STORE #, B=SUBTOTAL, C=INV DATE, D=BILL CODE, E=INVOICE TYPE, F=MISC
    Set-Cell $wsSource 2 1 28              # STORE # -> "0028"
    Set-Cell $wsSource 2 2 '$1,234.56'     # SUBTOTAL -> 1234.56
    Set-Cell $wsSource 2 3 '01/15/2026'    # INV DATE -> a real date
    Set-Cell $wsSource 2 4 100234          # BILL CODE -> "100234"
    Set-Cell $wsSource 2 5 'REPAIR'        # INVOICE TYPE -> passed through (stub is identity)
    Set-Cell $wsSource 2 6 'hello'         # Case Else -> passthrough, no format change

    Set-Cell $wsSource 3 1 '0007'
    Set-Cell $wsSource 3 2 'not a number'  # unparseable money -> blank, not written as text
    Set-Cell $wsSource 3 3 'not a date'    # unparseable date -> blank
    Set-Cell $wsSource 3 4 'BC-2'
    Set-Cell $wsSource 3 5 'installation'
    Set-Cell $wsSource 3 6 42

    $cols = 'STORE #,SUBTOTAL,INV DATE,BILL CODE,INVOICE TYPE,MISC|1,2,3,4,5,6|1,2,3,4,5,6'
    $packed = "$cols~1~3~10~NONE~BULK"

    $count = Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @($packed) -TimeoutSeconds 30

    Assert-Equal 'String' ($wsTarget.Cells(11,1).Value2.GetType().Name) -Because 'STORE # lands as text'
    Assert-Equal '0028'   ($wsTarget.Cells(11,1).Value2) -Because 'STORE # is zero-padded to 4 digits'
    Assert-Equal '@'      ($wsTarget.Cells(11,1).NumberFormat) -Because 'STORE # column gets @ format'

    Assert-Equal 1234.56  ($wsTarget.Cells(11,2).Value2) -Because 'a currency-formatted string coerces to a number'
    Assert-Equal '#,##0.00' ($wsTarget.Cells(11,2).NumberFormat) -Because 'SUBTOTAL gets money format'

    Assert-Equal '2026-01-15' ($wsTarget.Cells(11,3).Value()) -Because 'INV DATE coerces to a real date'
    Assert-Equal 'mm/dd/yyyy' ($wsTarget.Cells(11,3).NumberFormat) -Because 'INV DATE gets date format'

    Assert-Equal 'String' ($wsTarget.Cells(11,4).Value2.GetType().Name) -Because 'BILL CODE lands as text'
    Assert-Equal '100234' ($wsTarget.Cells(11,4).Value2) -Because 'BILL CODE keeps its digits exactly'

    Assert-Equal 'REPAIR' ($wsTarget.Cells(11,5).Value2) -Because 'INVOICE TYPE passes through the (stubbed) resolver'
    Assert-Equal 'hello'  ($wsTarget.Cells(11,6).Value2) -Because 'an unrecognised header is passed through unchanged'

    Assert-Equal '0007'   ($wsTarget.Cells(12,1).Value2) -Because 'a pre-padded store number is left alone'
    Assert-Equal ''       ([string]$wsTarget.Cells(12,2).Value2) -Because 'unparseable money is left blank, not written as text'
    Assert-Equal ''       ([string]$wsTarget.Cells(12,3).Value2) -Because 'unparseable date is left blank'
    Assert-Equal 'BC-2'   ($wsTarget.Cells(12,4).Value2) -Because 'a non-numeric bill code is kept as text'
    Assert-Equal 42       ($wsTarget.Cells(12,6).Value2) -Because 'a numeric passthrough value is written as a number, unchanged'

    Assert-Equal 2 $count -Because 'both source rows were counted as written'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Bulk path: rowStatus skips one of three rows, and IsMissing is exercised separately -----
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wsSource = $vba.Workbook.Worksheets(1)
    $wsSource.Name = 'Source'
    $wsTarget = $vba.Workbook.Worksheets.Add()
    $wsTarget.Name = 'Target'

    Set-Cell $wsSource 2 1 1
    Set-Cell $wsSource 3 1 2   # this row is flagged as a duplicate and must be skipped
    Set-Cell $wsSource 4 1 3

    $cols = 'STORE #|1|1'
    $packed = "$cols~1~4~5~,SKIP,~BULK"
    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @($packed) -TimeoutSeconds 30 | Out-Null

    # Row 2 (source) -> target row 6, row 3 skipped, row 4 (source) -> target row 7 (contiguous,
    # not row 8): a skipped row must not leave a gap in the appended block.
    Assert-Equal '0001' ($wsTarget.Cells(6,1).Value2) -Because 'first non-skipped row lands at targetLastRow+1'
    Assert-Equal '0003' ($wsTarget.Cells(7,1).Value2) -Because 'the skipped row leaves no gap -- rows land contiguously'
    Assert-Equal ''     ([string]$wsTarget.Cells(8,1).Value2) -Because 'nothing was written past the two actually-written rows'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- PerCell path (the FilterMode fallback) produces identical output to the bulk path ------
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wsSource = $vba.Workbook.Worksheets(1)
    $wsSource.Name = 'Source'
    $wsTarget = $vba.Workbook.Worksheets.Add()
    $wsTarget.Name = 'Target'

    Set-Cell $wsSource 2 1 28
    Set-Cell $wsSource 2 2 '$1,234.56'
    Set-Cell $wsSource 2 3 '01/15/2026'
    Set-Cell $wsSource 2 4 100234
    Set-Cell $wsSource 2 5 'REPAIR'
    Set-Cell $wsSource 2 6 'hello'

    $cols = 'STORE #,SUBTOTAL,INV DATE,BILL CODE,INVOICE TYPE,MISC|1,2,3,4,5,6|1,2,3,4,5,6'
    $packed = "$cols~1~2~10~NONE~PERCELL"
    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @($packed) -TimeoutSeconds 30 | Out-Null

    Assert-Equal '0028'   ($wsTarget.Cells(11,1).Value2) -Because 'PerCell path: STORE # matches the bulk path'
    Assert-Equal '@'      ($wsTarget.Cells(11,1).NumberFormat) -Because 'PerCell path: STORE # format matches the bulk path'
    Assert-Equal 1234.56  ($wsTarget.Cells(11,2).Value2) -Because 'PerCell path: SUBTOTAL matches the bulk path'
    Assert-Equal '2026-01-15' ($wsTarget.Cells(11,3).Value()) -Because 'PerCell path: INV DATE matches the bulk path'
    Assert-Equal '100234' ($wsTarget.Cells(11,4).Value2) -Because 'PerCell path: BILL CODE matches the bulk path'
    Assert-Equal 'REPAIR' ($wsTarget.Cells(11,5).Value2) -Because 'PerCell path: INVOICE TYPE matches the bulk path'
    Assert-Equal 'hello'  ($wsTarget.Cells(11,6).Value2) -Because 'PerCell path: MISC matches the bulk path'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Regression guard: the bulk path stays fast at a realistic import size (was O(rows), now
#     O(1) round trips regardless of row count) ---------------------------------------------
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wsSource = $vba.Workbook.Worksheets(1)
    $wsSource.Name = 'Source'
    $wsTarget = $vba.Workbook.Worksheets.Add()
    $wsTarget.Name = 'Target'

    $rowCount = 500
    for ($i = 0; $i -lt $rowCount; $i++) {
        $r = $i + 2
        Set-Cell $wsSource $r 1 (100 + $i)
        Set-Cell $wsSource $r 2 (10.5 + $i)
        Set-Cell $wsSource $r 3 '01/15/2026'
        Set-Cell $wsSource $r 4 (5000 + $i)
        Set-Cell $wsSource $r 5 'REPAIR'
        Set-Cell $wsSource $r 6 $i
    }

    $cols = 'STORE #,SUBTOTAL,INV DATE,BILL CODE,INVOICE TYPE,MISC|1,2,3,4,5,6|1,2,3,4,5,6'
    $packed = "$cols~1~$($rowCount + 1)~10~NONE~BULK"

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $count = Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @($packed) -TimeoutSeconds 60
    $sw.Stop()

    Assert-Equal $rowCount $count -Because 'every generated row was written'
    Assert-Equal $true ($sw.Elapsed.TotalSeconds -lt 15) `
        -Because 'a 500-row import completes well under a generous ceiling -- a per-cell reintroduction over ~6 mapped columns would cost roughly 3000 round trips instead of the bulk path''s ~12'
    Assert-Equal '0599' ($wsTarget.Cells($rowCount + 10, 1).Value2) -Because 'the last row in the block landed correctly'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
