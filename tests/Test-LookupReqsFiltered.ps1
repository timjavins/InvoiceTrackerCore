# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-LookupReqsFiltered.ps1
#
# LookupReqs must give the same result whether or not the tracker has an active AutoFilter.
# Regression for the 2026-10-07 incident on the guards tracker: with a filter on, the old bulk
# write-back (Range.Value = array) gave every VISIBLE cell the array's first element, so blank rows
# that matched nothing and even existing REQ #s were overwritten with Invoices!R2's value. Hidden
# rows must also be filled, and a hidden trailing row must not truncate the scan.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$stub = Join-Path $env:TEMP 'StubLookupFiltered.vb'
@'
Public Function TenantSheetName(ByVal role As String) As String
    Select Case LCase$(Trim$(role))
        Case "tracker":    TenantSheetName = "Invoices"
        Case "coupa-reqs": TenantSheetName = "Coupa Reqs"
        Case Else: Err.Raise 5, "TenantSheetName", "Unknown role '" & role & "'."
    End Select
End Function
Public Function TenantColLetter(ByVal concept As String) As String
    Select Case LCase$(Trim$(concept))
        Case "req-join-key":       TenantColLetter = "M"
        Case "requisition-number": TenantColLetter = "R"
        Case Else: Err.Raise 5, "TenantColLetter", "Unknown '" & concept & "'."
    End Select
End Function
' Compile-only: LookupReqs False, False never reaches UnprotectSheet/ProtectSheet, but VBA resolves
' their callees when it compiles LookupReqs. Same reasoning as Test-MarkerPreservation in guards.
Public Function TenantSheetPassword() As String
    TenantSheetPassword = vbNullString
End Function
Public Function SetupFilteredSheets() As String
    Dim wsT As Worksheet, wsR As Worksheet
    Set wsT = ThisWorkbook.Sheets(1)
    wsT.Name = "Invoices"
    Set wsR = ThisWorkbook.Sheets.Add
    wsR.Name = "Coupa Reqs"

    wsT.Range("A1").Value = "STORE": wsT.Range("M1").Value = "PO KEY": wsT.Range("R1").Value = "REQ #"
    ' Row 2: hidden, existing REQ #. It is also the array's first element, which is what the old
    ' write-back smeared over every visible cell.
    wsT.Range("A2").Value = "S001": wsT.Range("R2").Value = 777
    ' Row 3: visible, fillable.
    wsT.Range("A3").Value = "S847": wsT.Range("M3").Value = "KEY-3"
    ' Row 4: visible, a key Coupa does not know. Must stay blank.
    wsT.Range("A4").Value = "S847": wsT.Range("M4").Value = "KEY-4"
    ' Row 5: visible, existing numeric REQ #, no key.
    wsT.Range("A5").Value = "S847": wsT.Range("R5").Value = 1108423
    ' Row 6: visible, legacy marker on a row whose key Coupa DOES know. Must survive.
    wsT.Range("A6").Value = "S847": wsT.Range("M6").Value = "KEY-6": wsT.Range("R6").Value = "NOT ME"
    ' Row 7: hidden, fillable.
    wsT.Range("A7").Value = "S001": wsT.Range("M7").Value = "KEY-7"
    ' Row 8: hidden, fillable.
    wsT.Range("A8").Value = "S001": wsT.Range("M8").Value = "KEY-8"
    ' Row 9: hidden AND the last row, fillable, and Coupa holds its REQ # as a true NUMBER. A
    ' visible-only last-row scan would stop at row 6; a plain Value write would store a Double.
    wsT.Range("A9").Value = "S001": wsT.Range("M9").Value = "KEY-9"
    wsT.Range("A1:R9").AutoFilter 1, "S847"

    wsR.Range("A1").Value = "Req #": wsR.Range("C1").Value = "Supplier Part Number"
    wsR.Range("A2").Value = "1000003": wsR.Range("C2").Value = "KEY-3"
    wsR.Range("A3").Value = "1000006": wsR.Range("C3").Value = "KEY-6"
    wsR.Range("A4").Value = "1000007": wsR.Range("C4").Value = "KEY-7"
    wsR.Range("A5").Value = "1000008": wsR.Range("C5").Value = "KEY-8"
    wsR.Range("A6").Value = 1000009: wsR.Range("C6").Value = "KEY-9"
    SetupFilteredSheets = "FilterMode=" & CStr(wsT.FilterMode)
End Function
Public Function ReadCell(ByVal ref As String) As String
    ReadCell = CStr(ThisWorkbook.Sheets("Invoices").Range(ref).Value)
End Function
Public Function ReadCellType(ByVal ref As String) As String
    ReadCellType = TypeName(ThisWorkbook.Sheets("Invoices").Range(ref).Value)
End Function
Public Function ReadFilterMode() As String
    ReadFilterMode = CStr(ThisWorkbook.Sheets("Invoices").FilterMode)
End Function
Public Sub RunLookup()
    LookupReqs False, False
End Sub
'@ | Set-Content -LiteralPath $stub -Encoding UTF8

$sources = @(
    $stub,
    (Join-Path $repo 'LookupReqs.vb'),
    (Join-Path $repo 'CoerceValues.vb'),
    (Join-Path $repo 'GetHeaderColumnIndexes.vb'),
    (Join-Path $repo 'UnprotectSheet.vb'),
    (Join-Path $repo 'ProtectSheet.vb')
)
$vba = $null
try {
    $vba = New-VbaHost -SourceFiles $sources
    $setup = Invoke-VbaFunction -VbaHost $vba -Name 'SetupFilteredSheets' -TimeoutSeconds 60
    Assert-Equal -Expected 'FilterMode=True' -Actual $setup -Because 'the fixture really has an active filter'
    Invoke-VbaFunction -VbaHost $vba -Name 'RunLookup' -TimeoutSeconds 60 | Out-Null

    $read = { param($ref) Invoke-VbaFunction -VbaHost $vba -Name 'ReadCell' -Arguments @($ref) -TimeoutSeconds 60 }
    Assert-Equal -Expected '777'     -Actual (& $read 'R2') -Because 'a hidden row with an existing REQ # is untouched'
    Assert-Equal -Expected '1000003' -Actual (& $read 'R3') -Because 'a visible blank row takes its own matching REQ #'
    Assert-Equal -Expected ''        -Actual (& $read 'R4') -Because 'a visible blank row that matches nothing stays blank (not R2''s value)'
    Assert-Equal -Expected '1108423' -Actual (& $read 'R5') -Because 'an existing REQ # on a visible row is untouched'
    Assert-Equal -Expected 'NOT ME'  -Actual (& $read 'R6') -Because 'a marker survives even though Coupa knows its key'
    Assert-Equal -Expected '1000007' -Actual (& $read 'R7') -Because 'a hidden blank row is filled too'
    Assert-Equal -Expected '1000008' -Actual (& $read 'R8') -Because 'a hidden blank row is filled too (second one)'
    Assert-Equal -Expected '1000009' -Actual (& $read 'R9') -Because 'a hidden trailing row is still scanned and filled'
    $type = { param($ref) Invoke-VbaFunction -VbaHost $vba -Name 'ReadCellType' -Arguments @($ref) -TimeoutSeconds 60 }
    Assert-Equal -Expected 'String' -Actual (& $type 'R3') -Because 'a REQ # written from text stays text'
    Assert-Equal -Expected 'String' -Actual (& $type 'R9') -Because 'a REQ # Coupa holds as a NUMBER is written as text, not a Double'
    Assert-Equal -Expected 'True' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'ReadFilterMode' -TimeoutSeconds 60) `
                 -Because 'the operator''s filter is still on afterwards'
} finally { if ($vba) { Remove-VbaHost -VbaHost $vba | Out-Null }; Remove-Item $stub -Force -ErrorAction SilentlyContinue }
Write-AssertSummary
