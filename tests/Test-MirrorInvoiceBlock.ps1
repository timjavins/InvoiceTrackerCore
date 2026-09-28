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

# --- C1: filter-proofing -----------------------------------------------------------------
# End(xlUp) walks like Ctrl+Up and returns the last VISIBLE row on a filtered sheet. The first
# fix tried was Find(SearchDirection:=xlPrevious) instead, on the assumption that Find ignores
# hidden rows. Measured directly against a real AutoFilter (see the header comment and the
# report for the raw evidence), Find turned out to skip filter-hidden rows exactly like
# End(xlUp) -- e.g. case (a) below, run against the Find-only version, put the append on top
# of a hidden legacy row instead of after it. So MirrorInvoiceBlock instead refuses outright
# (-1, nothing written) whenever either sheet has FilterMode = True. These two cases prove the
# refusal: a fixture where the OLD End(xlUp)/Find approach would truncate or corrupt the
# mirror instead now returns -1 having touched nothing.

# (a) dst filtered: an owned row interleaved (hidden) and a trailing legacy row (hidden) both
# sit below the last row visible under the filter -- exactly the shape that used to make the
# append clobber the hidden legacy row. The mirror must refuse instead.
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $s = New-Sheets $vba.Workbook
    $d = $s.Dst

    Set-Row $s.Src 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES')
    Set-Row $s.Src 2 @('0003', 'M1', 10, 'store a')
    Set-Row $s.Src 3 @('0071', 'M1', 20, 'store b')

    Set-Row $d 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES', 'Source File')
    Set-Row $d 2 @('0175', 'L1', 1, 'legacy one', '2025:Invoices')
    Set-Row $d 3 @('0999', 'M1', 60, 'collapsed by migration', '2026:Invoices')   # owned, interleaved, hidden
    Set-Row $d 4 @('0176', 'L2', 2, 'PAID', '2025:Invoices')                     # last row VISIBLE under filter
    Set-Row $d 5 @('0177', 'L3', 3, 'PAID', 'RootNoYear:Invoices')               # hidden, AFTER the last visible row

    # Show only rows tagged 2025:Invoices; this hides row 3 (owned) and row 5 (trailing legacy).
    $d.Range($d.Cells(1,1), $d.Cells(5,5)).AutoFilter(5, '2025:Invoices') | Out-Null

    Assert-Equal $true $d.FilterMode -Because 'the fixture actually has an active filter'
    Assert-Equal $true $d.Rows(5).Hidden -Because 'the trailing legacy row starts out hidden by the filter'

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Mirror' -TimeoutSeconds 60
    Assert-Equal -1 $result -Because 'a filtered dst is refused rather than mirrored under a guess'

    Assert-Equal $true $d.FilterMode -Because 'the refusal never clears or changes the filter'
    Assert-Equal '=2025:Invoices' ($d.AutoFilter.Filters.Item(5).Criteria1) -Because 'the filter criteria is untouched'

    Assert-Equal 'L1' ($d.Cells(2,2).Value2) -Because 'row 2 is untouched'
    Assert-Equal 'M1' ($d.Cells(3,2).Value2) -Because 'the owned row is untouched -- nothing was deleted'
    Assert-Equal 'L2' ($d.Cells(4,2).Value2) -Because 'row 4 is untouched -- rows did not shift'
    Assert-Equal 'PAID' ($d.Cells(4,4).Value2) -Because 'its PAID note is untouched'
    Assert-Equal 'L3' ($d.Cells(5,2).Value2) -Because 'the hidden trailing legacy row is untouched'
    Assert-Equal 'PAID' ($d.Cells(5,4).Value2) -Because 'the hidden legacy row keeps its own hand-set PAID'
    Assert-Equal $true $d.Rows(5).Hidden -Because 'still hidden by the filter -- nothing about it changed'
    Assert-Equal '' ([string]$d.Cells(6,5).Value2) -Because 'nothing was appended at all'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# (b) src filtered: its last data row is hidden -- exactly the shape that used to make the
# mirror silently drop that row. The mirror must refuse instead of copying a truncated source.
$vba = New-VbaHost -SourceFiles @($src, $driver)
try {
    $s = New-Sheets $vba.Workbook
    $d = $s.Dst

    Set-Row $s.Src 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES')
    Set-Row $s.Src 2 @('0003', 'M1', 10, 'store a')
    Set-Row $s.Src 3 @('0071', 'M1', 20, 'store b')   # will be hidden -- last data row

    Set-Row $d 1 @('STORE #', ' BILL CODE ', ' TOTAL ', 'NOTES', 'Source File')
    Set-Row $d 2 @('0175', 'L1', 1, 'legacy one', 'Other:Invoices')

    # Show only 'store a'; this hides row 3, the source's last data row.
    $s.Src.Range($s.Src.Cells(1,1), $s.Src.Cells(3,4)).AutoFilter(4, 'store a') | Out-Null
    Assert-Equal $true $s.Src.FilterMode -Because 'the fixture actually has an active filter'
    Assert-Equal $true $s.Src.Rows(3).Hidden -Because 'the last source row starts out hidden'

    $result = Invoke-VbaFunction -VbaHost $vba -Name 'T_Mirror' -TimeoutSeconds 60
    Assert-Equal -1 $result -Because 'a filtered source is refused rather than mirrored truncated'

    Assert-Equal $true $s.Src.FilterMode -Because 'the refusal never clears or changes the source filter either'
    Assert-Equal 'L1' ($d.Cells(2,2).Value2) -Because 'the archive is completely untouched'
    Assert-Equal '' ([string]$d.Cells(3,2).Value2) -Because 'nothing was appended'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
