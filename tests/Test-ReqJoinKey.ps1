# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ReqJoinKey.ps1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

# Two stub tenants, written to temp files so the harness can inject them like real modules.
$declaring = Join-Path $env:TEMP 'StubTenantDeclaring.vb'
@'
Public Function TenantColLetter(ByVal concept As String) As String
    Select Case LCase$(Trim$(concept))
        Case "req-join-key":             TenantColLetter = "M"
        Case "submitted-invoice-number": TenantColLetter = "J"
        Case Else: Err.Raise 5, "TenantColLetter", "Unknown concept '" & concept & "'."
    End Select
End Function
'@ | Set-Content -LiteralPath $declaring -Encoding UTF8

$legacy = Join-Path $env:TEMP 'StubTenantLegacy.vb'
@'
Public Function TenantColLetter(ByVal concept As String) As String
    Select Case LCase$(Trim$(concept))
        Case "submitted-invoice-number": TenantColLetter = "J"
        Case Else: Err.Raise 5, "TenantColLetter", "Unknown concept '" & concept & "'."
    End Select
End Function
'@ | Set-Content -LiteralPath $legacy -Encoding UTF8

# Only the resolver is under test, so inject it alone rather than all of LookupReqs -- the rest of
# that module needs worksheets we deliberately do not have here.
$resolver = Join-Path $env:TEMP 'ResolverUnderTest.vb'
$lookupSrc = Get-Content -LiteralPath (Join-Path $repo 'LookupReqs.vb') -Raw -Encoding UTF8
$match = [regex]::Match($lookupSrc, '(?ms)^' + [regex]::Escape("' Which tracker column holds the value") + '.*?^End Function\s*$')
if (-not $match.Success) { throw "Could not find ReqJoinKeyColumn in LookupReqs.vb. Has it been written yet?" }
$match.Value | Set-Content -LiteralPath $resolver -Encoding UTF8

foreach ($case in @(
    @{ Stub = $declaring; Expected = 'M'; Because = 'a tenant declaring req-join-key gets its column' },
    @{ Stub = $legacy;    Expected = 'J'; Because = 'a tenant without req-join-key falls back to submitted-invoice-number' }
)) {
    $vba = New-VbaHost -SourceFiles @($case.Stub, $resolver)
    try {
        Assert-Equal -Expected $case.Expected `
                     -Actual (Invoke-VbaFunction -VbaHost $vba -Name 'ReqJoinKeyColumn') `
                     -Because $case.Because
    } finally { Remove-VbaHost -VbaHost $vba | Out-Null }
}

Remove-Item $declaring, $legacy, $resolver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
