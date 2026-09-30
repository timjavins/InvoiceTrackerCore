# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-PopulateInvoiceTypeForNewRows.ps1
#
# Covers SecuritasAutomation/file ingesting/AddNewBills.vb's PopulateInvoiceTypeForNewRows,
# rewritten alongside ProcessNewBills.vb when the write_vba cell-by-cell advisory (2026-09-29
# deploy) was investigated. The Sub is extracted verbatim (by regex, from the real file) rather
# than injecting the whole AddNewBills.vb: that file's `Sub AddNewBills()` is a large
# FileDialog-driven flow with ~20 further callees (HashFile, ClassifyBillCodes, SyncAllYears,
# ...) that have nothing to do with this Sub and would all need stubbing just to make the
# combined module compile. Extracting only the Sub under test keeps this suite honest about
# what it covers -- the real production text -- without that unrelated stubbing cost.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$addNewBillsPath = Join-Path (Split-Path $repo -Parent) 'SecuritasAutomation\file ingesting\AddNewBills.vb'
$fullSource = Get-Content -LiteralPath $addNewBillsPath -Raw

$match = [regex]::Match($fullSource, '(?ms)^Sub PopulateInvoiceTypeForNewRows.*?^End Sub')
if (-not $match.Success) {
    throw "Could not extract PopulateInvoiceTypeForNewRows from $addNewBillsPath -- has its signature changed?"
}
$extracted = Join-Path $env:TEMP 'PopulateInvoiceType.vb'
Set-Content -LiteralPath $extracted -Value $match.Value -Encoding UTF8

$driver = Join-Path $env:TEMP 'PitDriver.vb'
@'
' Real protect/unprotect nesting is exercised by other suites (ProtectSheet.vb has none of its
' own here, but LookupReqs/NormalizeIdentifierColumns tests do). This Sub's own behaviour does
' not depend on nesting depth, so no-op stubs are enough to let it compile and run.
Public Sub UnprotectSheetOn(ByVal ws As Worksheet)
End Sub

Public Sub ProtectSheetOn(ByVal ws As Worksheet)
End Sub

' The real ResolveInvoiceType lives elsewhere in AddNewBills.vb and is exercised by its own
' logic in Test-ProcessNewBillsBulk.ps1 (via the bulk/per-cell paths' INVOICE TYPE case). This
' Sub only cares that whatever ResolveInvoiceType returns gets written into every blank cell,
' so a fixed stub value is enough to prove that.
Public Function ResolveInvoiceType(ByVal vendorValue As String, ByVal billType As String) As String
    ResolveInvoiceType = "REPAIR"
End Function

' Invoke-VbaFunction supports at most 2 arguments, so billType/detectedInvoiceType are fixed
' here rather than threaded through -- ResolveInvoiceType above ignores them anyway.
Public Function T_Run(ByVal firstDataRow As Long, ByVal lastDataRow As Long) As Long
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets("Target")

    Dim headerDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    headerDict.Add "INVOICE TYPE", 5

    PopulateInvoiceTypeForNewRows wsTarget, firstDataRow, lastDataRow, headerDict, "repair", "REPAIR"
    T_Run = 0
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

# --- Unfiltered: bulk array path fills only the blank cells --------------------------------
$vba = New-VbaHost -SourceFiles @($extracted, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    $ws.Name = 'Target'

    Set-Cell $ws 11 5 $null           # blank -> gets the default
    Set-Cell $ws 12 5 'MONITORING'    # already set (source file carried its own) -> untouched
    Set-Cell $ws 13 5 $null           # blank -> gets the default

    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @(11, 13) -TimeoutSeconds 30 | Out-Null

    Assert-Equal 'REPAIR' ($ws.Cells(11,5).Value2) -Because 'a blank INVOICE TYPE cell gets the resolved default'
    Assert-Equal 'MONITORING' ($ws.Cells(12,5).Value2) -Because 'a non-blank cell already set from the source file is left alone'
    Assert-Equal 'REPAIR' ($ws.Cells(13,5).Value2) -Because 'a second blank cell in the same block also gets the default'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Every cell already populated: no write at all (guard against a pointless write-back) --
$vba = New-VbaHost -SourceFiles @($extracted, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    $ws.Name = 'Target'

    Set-Cell $ws 11 5 'REPAIR'
    Set-Cell $ws 12 5 'INSTALLATION'
    $ws.Cells(11,5).NumberFormat = '0.00'   # a format no real write of this Sub would produce
    $formatBefore = $ws.Cells(11,5).NumberFormat

    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @(11, 12) -TimeoutSeconds 30 | Out-Null

    Assert-Equal 'REPAIR' ($ws.Cells(11,5).Value2) -Because 'an already-populated cell keeps its own value'
    Assert-Equal 'INSTALLATION' ($ws.Cells(12,5).Value2) -Because 'a second already-populated cell also keeps its own value'
    Assert-Equal $formatBefore ($ws.Cells(11,5).NumberFormat) -Because 'no write happened at all when nothing was blank -- not even a format touch'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- firstDataRow > lastDataRow: no-op, matching the original guard clause -----------------
$vba = New-VbaHost -SourceFiles @($extracted, $driver)
try {
    $ws = $vba.Workbook.Worksheets(1)
    $ws.Name = 'Target'
    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @(11, 10) -TimeoutSeconds 30 | Out-Null
    Assert-Equal '' ([string]$ws.Cells(11,5).Value2) -Because 'firstDataRow > lastDataRow is a no-op, same as the original'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $extracted -Force -ErrorAction SilentlyContinue
Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
