# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ProcessedBatchLog.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$stub = Join-Path $env:TEMP 'StubPblTenant.vb'
@'
Public Function TenantProcessedBatchSheet() As String
    TenantProcessedBatchSheet = "Processed Batches"
End Function

Public Function T_IsProcessed(ByVal h As String) As Boolean
    T_IsProcessed = IsAlreadyProcessed(ThisWorkbook, h)
End Function

Public Function T_Record(ByVal h As String, ByVal f As String) As String
    RecordProcessedBatch ThisWorkbook, h, f, "TestTenant", 3, 1
    T_Record = "ok"
End Function

Public Function T_Summary(ByVal h As String) As String
    T_Summary = ProcessedBatchSummary(ThisWorkbook, h)
End Function

Public Function T_SheetVisibility() As Long
    T_SheetVisibility = ThisWorkbook.Worksheets("Processed Batches").Visible
End Function

Public Function T_LogRows() As Long
    With ThisWorkbook.Worksheets("Processed Batches")
        T_LogRows = .Cells(.Rows.Count, 1).End(xlUp).Row - 1
    End With
End Function

Public Function T_Header(ByVal c As Long) As String
    T_Header = CStr(ThisWorkbook.Worksheets("Processed Batches").Cells(1, c).Value)
End Function
'@ | Set-Content -LiteralPath $stub -Encoding UTF8

$src = Join-Path $repo 'ProcessedBatchLog.vb'
$h1 = 'a' * 64
$h2 = 'b' * 64

$vba = New-VbaHost -SourceFiles @($stub, $src)
try {
    $run = { param($n, $a) Invoke-VbaFunction -VbaHost $vba -Name $n -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal -Expected $false -Actual (& $run 'T_IsProcessed' @($h1)) -Because 'nothing is processed before the log sheet exists'
    Assert-Equal -Expected ''     -Actual (& $run 'T_Summary' @($h1))     -Because 'no summary for an unlogged hash'
    Assert-Equal -Expected 'ok'   -Actual (& $run 'T_Record' @($h1, 'bills-sept.xlsx')) -Because 'first record creates the sheet'
    Assert-Equal -Expected $true  -Actual (& $run 'T_IsProcessed' @($h1)) -Because 'a recorded hash is processed'
    Assert-Equal -Expected $true  -Actual (& $run 'T_IsProcessed' @($h1.ToUpper())) -Because 'hash match ignores case'
    Assert-Equal -Expected $false -Actual (& $run 'T_IsProcessed' @($h2)) -Because 'a different hash is not processed'
    Assert-Equal -Expected $false -Actual (& $run 'T_IsProcessed' @(''))  -Because 'an empty hash is never processed'
    Assert-Equal -Expected 0      -Actual (& $run 'T_SheetVisibility' @()) -Because 'the log sheet is hidden (xlSheetHidden = 0)'
    Assert-Equal -Expected 'HASH' -Actual (& $run 'T_Header' @(1)) -Because 'header A1'
    Assert-Equal -Expected 'ROWS SKIPPED' -Actual (& $run 'T_Header' @(6)) -Because 'header F1'

    $summary = [string](& $run 'T_Summary' @($h1))
    Assert-Equal -Expected $true -Actual ($summary -like '*bills-sept.xlsx*') -Because "summary names the file ($summary)"
    Assert-Equal -Expected $true -Actual ($summary -like '*3 added, 1 skipped*') -Because "summary gives the counts ($summary)"

    Assert-Equal -Expected 'ok' -Actual (& $run 'T_Record' @($h2, 'bills-oct.xlsx')) -Because 'second record reuses the sheet'
    Assert-Equal -Expected 2    -Actual (& $run 'T_LogRows' @()) -Because 'two batches, two rows, one sheet'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles @($stub, $src) -DocumentModule
try {
    Assert-Equal -Expected $false -Actual ($vba.Workbook.T_IsProcessed($h1)) `
                 -Because 'ProcessedBatchLog compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $stub -Force -ErrorAction SilentlyContinue
Write-AssertSummary
