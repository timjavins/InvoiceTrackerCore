# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-BillCodeScreen.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

# Drivers take "|"-joined strings so no array has to cross the COM boundary, and map ""
# to "ok" so expected values stay readable.
$driver = Join-Path $env:TEMP 'BcsDriver.vb'
@'
Private Function BcsTestRun(ByVal existingJoined As String, ByVal incomingJoined As String, ByVal allow As Boolean) As String
    Dim existing As Variant, incoming As Variant, result As Variant, i As Long, parts() As String
    existing = Split(existingJoined, "|")
    incoming = Split(incomingJoined, "|")
    result = ClassifyBillCodes(existing, incoming, allow)
    If UBound(result) < LBound(result) Then BcsTestRun = "(empty)": Exit Function
    ReDim parts(LBound(result) To UBound(result))
    For i = LBound(result) To UBound(result)
        parts(i) = IIf(result(i) = "", "ok", result(i))
    Next i
    BcsTestRun = Join(parts, "|")
End Function

Public Function T_Repair(ByVal e As String, ByVal n As String) As String
    T_Repair = BcsTestRun(e, n, False)
End Function

Public Function T_Monitoring(ByVal e As String, ByVal n As String) As String
    T_Monitoring = BcsTestRun(e, n, True)
End Function

Public Function T_Column(ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim v As Variant, i As Long, s As String
    v = BillCodeColumnValues(ThisWorkbook.Worksheets(1), 1, firstRow, lastRow)
    s = CStr(UBound(v) - LBound(v) + 1) & ":"
    For i = LBound(v) To UBound(v)
        s = s & CStr(v(i)) & ";"
    Next i
    T_Column = s
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

$src = Join-Path $repo 'BillCodeScreen.vb'
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $r = { param($n, $a) Invoke-VbaFunction -VbaHost $vba -Name $n -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal 'ok|ok' (& $r 'T_Repair' @('A|B', 'C|D')) -Because 'repair: new codes pass'
    Assert-Equal 'already-processed|ok' (& $r 'T_Repair' @('A|B', 'B|C')) -Because 'repair: a code already in the tracker is skipped'
    Assert-Equal 'repeat-in-file|repeat-in-file|ok' (& $r 'T_Repair' @('', 'C|C|D')) -Because 'repair: every row of a repeated code is skipped'
    Assert-Equal 'already-processed|already-processed' (& $r 'T_Repair' @('C', 'C|C')) -Because 'repair: already-processed wins over repeat-in-file'
    Assert-Equal 'ok|ok|ok' (& $r 'T_Monitoring' @('', 'C|C|D')) -Because 'monitoring: a shared code across store rows passes'
    Assert-Equal 'already-processed|already-processed|ok' (& $r 'T_Monitoring' @('C', 'C|C|D')) -Because 'monitoring: a whole already-processed group is skipped'
    Assert-Equal 'already-processed' (& $r 'T_Repair' @('  abc ', 'ABC')) -Because 'match is trimmed and case-insensitive'
    Assert-Equal 'ok|ok' (& $r 'T_Repair' @('', '|X')) -Because 'a blank incoming code is never flagged'
    Assert-Equal 'ok|ok' (& $r 'T_Repair' @('', '|')) -Because 'two blanks are not a repeat'

    $ws = $vba.Workbook.Worksheets(1)
    $ws.Range('A2').Value2 = [double]6005463095
    $ws.Range('A3').NumberFormat = '@'
    $ws.Range('A3').Value2 = '6005463095'
    Assert-Equal '2:6005463095;6005463095;' (& $r 'T_Column' @(2, 3)) -Because 'numeric and text codes read back as the same string'
    Assert-Equal '1:6005463095;' (& $r 'T_Column' @(2, 2)) -Because 'a one-cell range still returns an array'
    Assert-Equal '0:' (& $r 'T_Column' @(3, 2)) -Because 'an empty range returns an empty array'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles @($src, $driver) -DocumentModule
try {
    Assert-Equal 'ok|ok|ok' ($vba.Workbook.T_Monitoring('', 'C|C|D')) -Because 'BillCodeScreen compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
