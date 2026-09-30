# Run: powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-ExtractInvApprovalDate.ps1
#
# Covers InvoiceTrackerCore/ExtractInvApprovalDate.vb, rewritten from a per-row status
# read + per-row approval-date write into a bulk read/compute/write, when the write_vba
# cell-by-cell advisory (2026-09-29 deploy) was investigated. Unlike ProcessNewBills/
# PopulateInvoiceTypeForNewRows (both append-only), this Sub sweeps the WHOLE existing
# tracker range on every Refresh, so it is the one candidate in this effort where the
# AutoFilter-misalignment hazard (an array write spanning a filter-hidden row can land
# values on the wrong VISIBLE rows) is a live risk, not just a defensive pattern kept for
# consistency -- see the filter-proofing case below.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $PSScriptRoot 'VbaHarness.psm1') -Force
Reset-AssertCounters

$sources = @(
    (Join-Path $repo 'ExtractInvApprovalDate.vb'),
    (Join-Path $repo 'ColumnValues.vb')
)

# ExtractInvApprovalDate.vb is entered, so its callees must resolve: TenantSheetName,
# TenantColLetter (tracker layout), PauseThinking/RestoreThinking/UnprotectSheet/ProtectSheet
# (orchestration side-effects this Sub's own logic does not depend on). All are stubbed as
# no-ops/fixed constants, same convention as the other suites in this directory.
$driver = Join-Path $env:TEMP 'EiadDriver.vb'
@'
Public Function TenantSheetName(ByVal role As String) As String
    Select Case role
        Case "tracker": TenantSheetName = "Tracker"
        Case "coupa-invs": TenantSheetName = "CoupaInvs"
    End Select
End Function

Public Function TenantColLetter(ByVal concept As String) As String
    Select Case concept
        Case "store-number": TenantColLetter = "A"
        Case "invoice-status": TenantColLetter = "B"
        Case "purchase-order-number": TenantColLetter = "C"
        Case "invoice-approval-date": TenantColLetter = "D"
    End Select
End Function

Public Sub PauseThinking()
End Sub

Public Sub RestoreThinking()
End Sub

Public Sub UnprotectSheet()
End Sub

Public Sub ProtectSheet()
End Sub

Public Function T_Run(ByVal manageProtectionInt As Long) As Long
    ExtractInvApprovalDate False, CBool(manageProtectionInt)
    T_Run = 0
End Function
'@ | Set-Content -LiteralPath $driver -Encoding UTF8

function Set-Cell($ws, [int]$row, [int]$col, $value) {
    $cell = $ws.Cells($row, $col)
    if ($null -eq $value) { $cell.ClearContents() | Out-Null } else { $cell.Value2 = $value }
}

# History-log text shaped like a real Coupa export, with both the 2-digit and 4-digit year
# forms this Sub is specifically documented to accept.
$historyFourDigit = '07/29/2026 - Pay Invoice Status: [pending_document_approval] to [ready_to_pay] by someone'
$historyTwoDigit   = '03/02/26 - Pay Invoice Status: [pending_document_approval] to [ready_to_pay] by someone'

# --- Unfiltered: bulk read/compute/write, approved + not-approved + no-match rows ----------
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wb = $vba.Workbook
    $tracker = $wb.Worksheets(1); $tracker.Name = 'Tracker'
    $coupa = $wb.Worksheets.Add(); $coupa.Name = 'CoupaInvs'

    # Tracker: A=store (last-row anchor), B=status, C=PO, D=approval date (written)
    Set-Cell $tracker 2 1 '0001'; Set-Cell $tracker 2 2 'Approved'; Set-Cell $tracker 2 3 'PO-1'
    Set-Cell $tracker 3 1 '0002'; Set-Cell $tracker 3 2 'Pending Approval'; Set-Cell $tracker 3 3 'PO-2'
    Set-Cell $tracker 4 1 '0003'; Set-Cell $tracker 4 2 'Approved'; Set-Cell $tracker 4 3 'PO-3'  # PO not in CoupaInvs
    Set-Cell $tracker 4 4 'stale date from a previous run'                                        # must be cleared: not approved-with-match
    Set-Cell $tracker 5 1 '0004'; Set-Cell $tracker 5 2 'Approved'; Set-Cell $tracker 5 3 'PO-4'   # 2-digit-year history

    # CoupaInvs: A=PO, M=history text
    Set-Cell $coupa 1 1 'PO-1'; Set-Cell $coupa 1 13 $historyFourDigit
    Set-Cell $coupa 2 1 'PO-4'; Set-Cell $coupa 2 13 $historyTwoDigit

    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @(0) -TimeoutSeconds 30 | Out-Null

    # Excel's .Value setter (used here, matching the original per-cell code) auto-converts a
    # date-shaped string into a real date on write, so read back via .Value() -- not .Value2,
    # which would hand back the raw OLE date serial -- and compare as a date.
    Assert-Equal ([datetime]'2026-07-29') ($tracker.Cells(2,4).Value()) -Because 'an approved row with a matching 4-digit-year history gets its approval date'
    Assert-Equal ''           ([string]$tracker.Cells(3,4).Value2) -Because 'a non-approved row is left blank'
    Assert-Equal ''           ([string]$tracker.Cells(4,4).Value2) -Because 'an approved row whose PO has no CoupaInvs match is cleared, not left stale'
    Assert-Equal ([datetime]'2026-03-02') ($tracker.Cells(5,4).Value()) -Because 'a 2-digit year in the history log is still accepted'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- No data rows: must not raise (ColumnValues would hand back Empty; UBound on that is a
#     type mismatch without the lastRowInv < 2 guard added alongside the bulk rewrite) -------
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wb = $vba.Workbook
    $tracker = $wb.Worksheets(1); $tracker.Name = 'Tracker'
    $coupa = $wb.Worksheets.Add(); $coupa.Name = 'CoupaInvs'
    Set-Cell $tracker 1 1 'STORE #'   # header only, no data rows

    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @(0) -TimeoutSeconds 30 | Out-Null
    Assert-Equal '' ([string]$tracker.Cells(2,4).Value2) -Because 'an empty tracker is a no-op, not a crash'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

# --- Filter-proofing: a hidden row must not get its value misaligned onto a visible one -----
# This is the hazard MirrorInvoiceBlock refuses outright and NormalizeIdentifierColumnsOn
# falls back for, and the one candidate in this effort where it is a genuine risk rather than
# defensive boilerplate: this Sub sweeps pre-existing rows, not an append-only block.
$vba = New-VbaHost -SourceFiles ($sources + $driver)
try {
    $wb = $vba.Workbook
    $tracker = $wb.Worksheets(1); $tracker.Name = 'Tracker'
    $coupa = $wb.Worksheets.Add(); $coupa.Name = 'CoupaInvs'

    # Column E is a dedicated filter tag, separate from the data columns this Sub actually
    # reads/writes (A/B/C/D), so the filter's own criteria column cannot be confused with the
    # approval-date logic under test. Row 4 (visible) sits AFTER the hidden row 3 deliberately:
    # End(xlUp) -- which lastRowInv is computed with, unchanged by this rewrite -- walks like
    # Ctrl+Up and skips filter-hidden rows same as it always did, so if the true last row were
    # the hidden one, lastRowInv would undercount and this fixture would exercise that
    # pre-existing, out-of-scope limitation instead of the misalignment hazard this case exists
    # to catch.
    Set-Cell $tracker 1 1 'STORE #'; Set-Cell $tracker 1 5 'TAG'
    Set-Cell $tracker 2 1 '0001'; Set-Cell $tracker 2 2 'Approved'; Set-Cell $tracker 2 3 'PO-1'; Set-Cell $tracker 2 5 'show'
    Set-Cell $tracker 3 1 '0002'; Set-Cell $tracker 3 2 'Approved'; Set-Cell $tracker 3 3 'PO-HIDE'; Set-Cell $tracker 3 5 'hide'
    Set-Cell $tracker 4 1 '0003'; Set-Cell $tracker 4 2 'Approved'; Set-Cell $tracker 4 3 'PO-1'; Set-Cell $tracker 4 5 'show'

    Set-Cell $coupa 1 1 'PO-HIDE'; Set-Cell $coupa 1 13 $historyFourDigit

    # Row 2 and row 4 (PO-1) have no CoupaInvs match and must end up blank; row 3 (PO-HIDE,
    # hidden by the filter) has a real match. A misaligned bulk write would swap these outcomes.
    $tracker.Range($tracker.Cells(1,1), $tracker.Cells(4,5)).AutoFilter(5, 'show') | Out-Null
    Assert-Equal $true $tracker.FilterMode -Because 'the fixture actually has an active filter'
    Assert-Equal $true $tracker.Rows(3).Hidden -Because 'row 3 starts out hidden by the filter'
    Assert-Equal $false $tracker.Rows(4).Hidden -Because 'row 4, the true last row, stays visible'

    Invoke-VbaFunction -VbaHost $vba -Name 'T_Run' -Arguments @(0) -TimeoutSeconds 30 | Out-Null

    Assert-Equal '' ([string]$tracker.Cells(2,4).Value2) -Because 'the visible row (PO-1, no CoupaInvs match) is not corrupted by the hidden row''s value'
    Assert-Equal ([datetime]'2026-07-29') ($tracker.Cells(3,4).Value()) -Because 'the hidden row still gets its own correct approval date, not skipped and not misaligned'
    Assert-Equal '' ([string]$tracker.Cells(4,4).Value2) -Because 'the trailing visible row (also PO-1, no match) is unaffected too'
    Assert-Equal $true $tracker.Rows(3).Hidden -Because 'the sweep never touches the filter itself'
} finally { Remove-VbaHost -VbaHost $vba | Out-Null }

Remove-Item $driver -Force -ErrorAction SilentlyContinue
Write-AssertSummary
