# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-HashFile.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$dir = Join-Path $env:TEMP ('hashfile-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir | Out-Null
$abc = Join-Path $dir 'abc.txt'
[IO.File]::WriteAllText($abc, 'abc')                        # no BOM, no newline
$spaced = Join-Path $dir 'bill file (2) & co.txt'
[IO.File]::WriteAllText($spaced, 'abc')
$missing = Join-Path $dir 'nope.txt'

# Known SHA-256 of the three bytes "abc" (FIPS 180-2 test vector).
$abcHash = 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'

$src = Join-Path $repo 'HashFile.vb'
$vba = New-VbaHost -SourceFiles @($src)
try {
    Assert-Equal -Expected $abcHash -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @($abc) -TimeoutSeconds 30) `
                 -Because 'hash of "abc" matches the FIPS test vector'
    Assert-Equal -Expected $abcHash -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @($spaced) -TimeoutSeconds 30) `
                 -Because 'a path with spaces, parentheses and & still hashes'
    Assert-Equal -Expected '' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @($missing) -TimeoutSeconds 30) `
                 -Because 'a missing file returns "" rather than raising'
    Assert-Equal -Expected '' -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'HashFile' -Arguments @('https://example.com/x.xlsx') -TimeoutSeconds 30) `
                 -Because 'a URL returns "" rather than raising'
    $leftovers = @(Get-ChildItem -Path $env:TEMP -Filter 'hashfile_*.txt' -ErrorAction SilentlyContinue |
                   Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) })
    Assert-Equal -Expected 0 -Actual $leftovers.Count -Because 'HashFile deletes its temp output file'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# Same module as a document-module member: proves it compiles where production puts it.
$vba = New-VbaHost -SourceFiles @($src) -DocumentModule
try {
    Assert-Equal -Expected $abcHash -Actual ($vba.Workbook.HashFile($abc)) `
                 -Because 'HashFile compiles and runs inside ThisWorkbook'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
Write-AssertSummary
