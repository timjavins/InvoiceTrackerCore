# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-FlatFileHelpers.ps1
#
# The credit-chart decision in FlatFileHelpers.vb: which business unit is credit, which chart of
# accounts a row posts to, and whether a supplier site agrees with that chart. FlatFileHelpers.vb is
# compiled alone, over a made-up tenant config defined below -- the real BU and chart names are
# configuration, so nothing here depends on (or contains) a real tenant's values.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force

Reset-AssertCounters

# A made-up tenant: credit BU 900, charts "Standard Chart" (normal) and "Credit Chart" (credit). The stub
# also covers the cases a PowerShell argument cannot express as a Variant -- Empty, Null and an Excel error.
$stub = Join-Path $env:TEMP 'FlatFileHelpersStub.vb'
@'
Public Function TenantCreditBusinessUnit() As String
    TenantCreditBusinessUnit = "900"
End Function
Public Function TenantChartOfAccountName(Optional ByVal isCredit As Boolean = False) As String
    If isCredit Then TenantChartOfAccountName = "Credit Chart" Else TenantChartOfAccountName = "Standard Chart"
End Function
Public Function CreditOf(ByVal kind As String) As String
    Dim v As Variant
    Select Case kind
        Case "empty":   v = Empty
        Case "null":    v = Null
        Case "error":   v = CVErr(2042)
        Case "double":  v = 900#
        Case "long":    v = 900&
        Case "text":    v = " 900 "
        Case "other":   v = 200
        Case "words":   v = "credit"
        Case Else:      v = kind
    End Select
    CreditOf = CStr(IsCreditBusinessUnit(v)) & "|" & ChartOfAccountNameForBU(v)
End Function
Public Function IdentityOf(ByVal spec As String) As String
    Dim p As Variant, n As Variant
    p = Split(spec, ";")
    If IsNumeric(p(2)) Then n = CDbl(p(2)) Else n = p(2)
    IdentityOf = LineIdentityProblem(p(0), CStr(p(1)), n)
End Function
'@ | Set-Content -LiteralPath $stub -Encoding UTF8

$vba = New-VbaHost -SourceFiles @((Join-Path $repo 'FlatFileHelpers.vb'), $stub)
try {
    function Call-Vba($name, $arguments) {
        Invoke-VbaFunction -VbaHost $vba -Name $name -Arguments $arguments -TimeoutSeconds 30
    }
    # One string in, because the harness passes at most two arguments: "bu;supplier name;supplier number".
    function Identity($a) { Call-Vba 'IdentityOf' @("$($a[0]);$($a[1]);$($a[2])") }

    # IsCreditBusinessUnit and ChartOfAccountNameForBU always agree, so each case asserts both.
    Assert-Equal -Expected 'True|Credit Chart'     -Actual (Call-Vba 'CreditOf' @('900'))    -Because 'text 900 is the credit BU'
    Assert-Equal -Expected 'True|Credit Chart'     -Actual (Call-Vba 'CreditOf' @('text'))   -Because 'surrounding spaces are trimmed'
    Assert-Equal -Expected 'True|Credit Chart'     -Actual (Call-Vba 'CreditOf' @('double')) -Because 'a Double 900.0 is credit'
    Assert-Equal -Expected 'True|Credit Chart'     -Actual (Call-Vba 'CreditOf' @('long'))   -Because 'a Long 900 is credit'
    Assert-Equal -Expected 'True|Credit Chart'     -Actual (Call-Vba 'CreditOf' @('0900'))   -Because 'a zero-padded 0900 is the same number'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @('other'))  -Because 'any other BU posts to the standard chart'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @('9000'))   -Because '9000 is not 900'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @('words'))  -Because 'text that is not a number is not credit'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @('empty'))  -Because 'a blank BU is not credit'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @('null'))   -Because 'Null is not credit'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @('error'))  -Because 'an Excel error value is not credit'
    Assert-Equal -Expected 'False|Standard Chart'  -Actual (Call-Vba 'CreditOf' @(''))       -Because 'an empty string is not credit'

    # SupplierSiteMatchesChart: False only on a contradiction. A name that ends in neither chart's name
    # (a placeholder such as "ODP -", or a blank) gives no verdict.
    $std    = 'ACME GUARD CO (ACME GUARD_1 MAIN ST_SPRINGFIELD_Standard Chart)'
    $credit = 'ACME GUARD CO (ACME GUARD_1 MAIN ST_SPRINGFIELD_Credit Chart)'
    $plainStd    = 'ACME ALARM CORP_PO BOX 100_SHELBYVILLE_Standard Chart'
    $plainCredit = 'ACME ALARM CORP_DEPT 2_CAPITAL CITY_Credit Chart'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @($credit, 'Credit Chart'))   -Because 'a credit site on the credit chart agrees'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @($std, 'Standard Chart'))    -Because 'a standard site on the standard chart agrees'
    Assert-Equal -Expected 'False' -Actual (Call-Vba 'SupplierSiteMatchesChart' @($credit, 'Standard Chart')) -Because 'a credit site on the standard chart is the mismatch Coupa rejects'
    Assert-Equal -Expected 'False' -Actual (Call-Vba 'SupplierSiteMatchesChart' @($std, 'Credit Chart'))      -Because 'a standard site on the credit chart is the reverse mismatch'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @($plainCredit, 'Credit Chart'))   -Because 'an unwrapped credit site on the credit chart agrees'
    Assert-Equal -Expected 'False' -Actual (Call-Vba 'SupplierSiteMatchesChart' @($plainCredit, 'Standard Chart')) -Because 'an unwrapped credit site on the standard chart is a mismatch too'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @($plainStd, 'Standard Chart'))    -Because 'an unwrapped standard site on the standard chart agrees'
    Assert-Equal -Expected 'False' -Actual (Call-Vba 'SupplierSiteMatchesChart' @($plainStd, 'Credit Chart'))      -Because 'an unwrapped standard site on the credit chart is a mismatch too'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @('SITE_credit chart ', 'Credit Chart')) -Because 'case and trailing spaces are ignored'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @('ODP -', 'Credit Chart'))     -Because 'a name with no recognised suffix gives no verdict'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @('', 'Credit Chart'))          -Because 'a blank supplier name is reported elsewhere, not here'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @('X_Credit Chart Union', 'Standard Chart')) -Because 'the suffix must be the whole site name, after an underscore'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @($credit, 'Some Other Chart')) -Because 'a chart the tenant did not configure gives no verdict'
    Assert-Equal -Expected 'True'  -Actual (Call-Vba 'SupplierSiteMatchesChart' @($credit, ''))                 -Because 'a blank chart gives no verdict'

    # LineIdentityProblem: "" when a line's business unit, supplier and chart agree, else a reason a person
    # can act on. A supplier with no credit supplier number is the second way a credit line is rejected, and
    # a tenant declares it by returning 0 and "" for its credit identity.
    Assert-Equal -Expected '' -Actual (Identity @('900', $credit, 540))        -Because 'a credit row with a credit site and a number is fine'
    Assert-Equal -Expected '' -Actual (Identity @('900', $plainCredit, '5473')) -Because 'a text supplier number is fine too'
    Assert-Equal -Expected '' -Actual (Identity @('200', $plainStd, 2362))      -Because 'a normal row with a standard site is fine'
    Assert-Equal -Expected '' -Actual (Identity @('200', '', ''))               -Because 'a normal row with no supplier yet is not this check''s business'
    Assert-Equal -Expected '' -Actual (Identity @('200', 'ODP -', 'N/A'))       -Because 'a normal row with a placeholder supplier is not this check''s business'
    foreach ($none in @(0, '', 'N/A', 'n/a', ' ', '0')) {
        $got = Identity @('900', $credit, $none)
        Assert-Equal -Expected $true -Actual ($got -like '*has no Credit Chart supplier number*') `
                     -Because "a credit row whose supplier number is '$none' has no credit supplier (got: $got)"
    }
    $got = Identity @('900', '', 5473)
    Assert-Equal -Expected $true -Actual ($got -like '*has no Credit Chart supplier number*') -Because 'a credit row with no supplier name has no credit supplier'
    $got = Identity @('900', $std, 540)
    Assert-Equal -Expected $true -Actual ($got -like '*other chart of accounts*' -and $got -like '*Credit Chart*') `
                 -Because "a credit row on a standard site is the mismatch, named with its chart (got: $got)"
    $got = Identity @('200', $credit, 540)
    Assert-Equal -Expected $true -Actual ($got -like '*other chart of accounts*' -and $got -like '*Standard Chart*') `
                 -Because "a normal row on a credit site is the reverse mismatch (got: $got)"
    Assert-Equal -Expected '' -Actual (Identity @('900', 'ODP -', 99)) `
                 -Because 'a credit row with a real number and an unrecognised site name gives no verdict'
} finally {
    Remove-VbaHost -VbaHost $vba | Out-Null
    Remove-Item $stub -Force -ErrorAction SilentlyContinue
}
Write-AssertSummary
