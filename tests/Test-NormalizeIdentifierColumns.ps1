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
function Set-Cell($ws, [int]$row, [int]$col, $value, [switch]$AsFormula) {
    $cell = $ws.Cells($row, $col)
    if ($AsFormula) {
        $cell.Formula = $value
    } elseif ($value -is [string]) {
        $cell.NumberFormat = '@'
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

    Assert-Equal 'String' ($ws.Cells(2,3).Value2.GetType().Name) -Because 'C2 becomes text too -- multiple columns are swept'
    Assert-Equal '555555' ($ws.Cells(2,3).Value2) -Because 'C2 keeps its digits exactly'
    Assert-Equal $true ($ws.Cells(3,3).HasFormula) -Because 'the second column formula cell is never touched either'
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
