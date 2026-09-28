# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MirrorInvoiceBlock.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$driver = Join-Path $env:TEMP 'MibDriver.vb'
@'
Public Function T_Mirror() As Long
    T_Mirror = MirrorInvoiceBlock(ThisWorkbook.Worksheets("Src"), ThisWorkbook.Worksheets("Dst"), _
                                  "2026:Invoices", Array("2026:Invoices", "2026Model:Invoices"))
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

$xlUp = -4162

function New-Sheets($wb) {
    $dst = $wb.Worksheets.Add(); $dst.Name = 'Dst'
    $src = $wb.Worksheets.Add(); $src.Name = 'Src'
    @{ Src = $src; Dst = $dst }
}

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        $cell = $ws.Cells($row, $c + 1)
        if ($v -is [string] -and $v.StartsWith('=')) { $cell.Formula = $v }
        elseif ($v -is [string]) { $cell.NumberFormat = '@'; $cell.Value2 = $v }
        # PowerShell 5.1 caches the COM property-put marshalling for Value2 by call site: once
        # this line has set it from a string, an unconverted boxed Int32 on a later loop
        # iteration throws "Unable to cast object of type 'System.Int32' to type 'System.String'".
        # Casting to [double] first side-steps that cache instead of changing test intent.
        else { $cell.Value2 = [double]$v }
    }
}

function Get-LastRow($ws, [int]$col) { $ws.Cells($ws.Rows.Count, $col).End($xlUp).Row }

function Build-Fixture($s) {
    Set-Row $s.Src 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES')
    # A monitoring bill: three store rows sharing one code.
    Set-Row $s.Src 2 @('0003', 'M1', 10, 'store a')
    Set-Row $s.Src 3 @('0071', 'M1', 20, 'store b')
    Set-Row $s.Src 4 @('0227', 'M1', 30, 'store c')
    Set-Row $s.Src 5 @('0414', 'R9', '=2*50', 'repair')

    Set-Row $s.Dst 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES', 'Source File')
    Set-Row $s.Dst 2 @('0175', 'L1', 1, 'legacy one', '2025:Invoices')
    Set-Row $s.Dst 3 @('0999', 'M1', 60, 'collapsed by migration', '2026Model:Invoices')
    Set-Row $s.Dst 4 @('0176', 'L2', 2, 'PAID', 'RootNoYear:Invoices')
    Set-Row $s.Dst 5 @('0998', 'OLD', 5, 'stale mirror row', '2026:Invoices')
    Set-Row $s.Dst 6 @('0177', 'L3', 3, 'legacy three', '2025:Invoices')
}

$src = Join-Path $repo 'MirrorInvoiceBlock.vb'
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $s = New-Sheets $vba.Workbook
    Build-Fixture $s
    $d = $s.Dst
    $mirror = { Invoke-VbaFunction -VbaHost $vba -Name 'T_Mirror' -TimeoutSeconds 60 }

    Assert-Equal 4 (& $mirror) -Because 'four source rows mirrored'
    Assert-Equal 8 (Get-LastRow $d 5) -Because '3 legacy + 4 mirrored rows, header in row 1'
    Assert-Equal 'L1|L2|L3' (($d.Cells(2,2).Value2, $d.Cells(3,2).Value2, $d.Cells(4,2).Value2) -join '|') `
                 -Because 'legacy rows keep their order, owned rows (both tags) are gone'
    Assert-Equal 'PAID' ($d.Cells(3,4).Value2) -Because 'a hand-set PAID on a legacy row survives'
    Assert-Equal 'M1|M1|M1|R9' (($d.Cells(5,2).Value2, $d.Cells(6,2).Value2, $d.Cells(7,2).Value2, $d.Cells(8,2).Value2) -join '|') `
                 -Because 'all three monitoring store rows arrive, not one'
    Assert-Equal '0003' ($d.Cells(5,1).Value2) -Because 'a text store number keeps its leading zeros'
    Assert-Equal '0175' ($d.Cells(2,1).Value2) -Because 'a legacy text store number is not rewritten to a number'
    Assert-Equal 100 ($d.Cells(8,3).Value2) -Because 'a formula arrives as its value'
    Assert-Equal $false ($d.Cells(8,3).HasFormula) -Because 'no formula is copied'
    Assert-Equal '2026:Invoices' ($d.Cells(8,5).Value2) -Because 'mirrored rows get the write tag'

    Assert-Equal 4 (& $mirror) -Because 'a re-run mirrors the same four rows'
    Assert-Equal 8 (Get-LastRow $d 5) -Because 'a re-run does not duplicate anything'

    $s.Src.Rows(5).Delete() | Out-Null
    Assert-Equal 3 (& $mirror) -Because 'a smaller source mirrors fewer rows'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'the dropped source row is gone from the mirror'

    $d.Cells(1,3).Value2 = 'AMOUNT'
    Assert-Equal -1 (& $mirror) -Because 'a header mismatch refuses'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'a refused mirror writes nothing'
    $d.Cells(1,3).Value2 = ' TOTAL '

    $d.Cells(1,6).Value2 = 'EXTRA'
    Assert-Equal -1 (& $mirror) -Because 'a column after Source File refuses'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'a refused mirror writes nothing'
    $d.Cells(1,6).ClearContents() | Out-Null

    $s.Src.Cells(1,2).Value2 = 'CODE'
    $d.Cells(1,2).Value2 = 'CODE'
    Assert-Equal -1 (& $mirror) -Because 'a source with no BILL CODE header refuses'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'a refused mirror writes nothing'
    $s.Src.Cells(1,2).Value2 = ' BILL CODE '
    $d.Cells(1,2).Value2 = ' BILL CODE '

    $s.Src.Range('A2:D10').ClearContents() | Out-Null
    Assert-Equal -1 (& $mirror) -Because 'an empty source refuses rather than wiping the owned block'
    Assert-Equal 7 (Get-LastRow $d 5) -Because 'the owned block survives an empty source'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles @($src, $driver) -DocumentModule
try {
    New-Sheets $vba.Workbook | Out-Null
    Assert-Equal -1 ($vba.Workbook.T_Mirror()) -Because 'MirrorInvoiceBlock compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
