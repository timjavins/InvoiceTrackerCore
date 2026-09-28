# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-MoveProcessedFile.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$dir = Join-Path $env:TEMP ('mpf-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$first = Join-Path $dir 'bills.csv'
Set-Content -LiteralPath $first -Value 'one'

$srcFiles = @(
    (Join-Path $repo 'GetBaseFileName.vb'),
    (Join-Path $repo 'GetFileExt.vb'),
    (Join-Path $repo 'MoveProcessedFile.vb')
)
$vba = New-VbaHost -SourceFiles $srcFiles
try {
    $run = { param($a) Invoke-VbaFunction -VbaHost $vba -Name 'MoveProcessedFile' -Arguments $a -TimeoutSeconds 30 }

    Assert-Equal -Expected '' -Actual (& $run @($first, '')) -Because 'an empty subfolder name disables the move'
    Assert-Equal -Expected $true -Actual (Test-Path -LiteralPath $first) -Because 'the file is untouched when the move is disabled'

    $expected = Join-Path $dir 'Processed\bills.csv'
    Assert-Equal -Expected $expected -Actual (& $run @($first, 'Processed')) -Because 'moves into a new sibling subfolder'
    Assert-Equal -Expected $false -Actual (Test-Path -LiteralPath $first) -Because 'the original is gone after the move'
    Assert-Equal -Expected $true  -Actual (Test-Path -LiteralPath $expected) -Because 'the file is at the returned path'

    Set-Content -LiteralPath $first -Value 'two'
    $second = [string](& $run @($first, 'Processed'))
    Assert-Equal -Expected $true -Actual ($second -match '\\Processed\\bills \(\d{8}-\d{6}\)\.csv$') `
                 -Because "a name collision gets a timestamp suffix ($second)"
    Assert-Equal -Expected 'one' -Actual ((Get-Content -LiteralPath $expected -Raw).Trim()) `
                 -Because 'the earlier processed file is not overwritten'

    Assert-Equal -Expected '' -Actual (& $run @((Join-Path $dir 'missing.csv'), 'Processed')) -Because 'a missing file returns ""'
    Assert-Equal -Expected '' -Actual (& $run @('https://example.com/x.csv', 'Processed')) -Because 'a URL returns ""'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

$vba = New-VbaHost -SourceFiles $srcFiles -DocumentModule
try {
    Assert-Equal -Expected '' -Actual ($vba.Workbook.MoveProcessedFile((Join-Path $dir 'missing.csv'), 'Processed')) `
                 -Because 'MoveProcessedFile compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
Write-AssertSummary
