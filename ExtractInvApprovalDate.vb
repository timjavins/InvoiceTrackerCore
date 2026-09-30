' Fills in each approved invoice's approval date by reading the Coupa invoice history log.
'
' For every tracker row whose invoice status is "Approved", the matching Coupa invoice is
' found by PO number and its History text searched for the transition
' "<date> - Pay Invoice Status: [pending_document_approval] to [ready_to_pay] by".
' Rows that are not approved have the date cleared.
'
' The year in that log is FOUR digits in real exports (verified 2026-07-30, e.g.
' "07/29/2026"). Securitas allowed \d{2,4} and worked; JCI allowed only \d{2} and therefore
' matched nothing -- checked against JCI's own export, its pattern found 0 of the 173 rows
' Securitas's pattern found. Core accepts 2- or 4-digit years.
'
' Sheets and columns come from TenantConfig.

' manageProtection defaults to True: JCI's version unprotected and reprotected the tracker
' itself, and its callers rely on that. Securitas's did not, so its callers pass False.
Public Sub ExtractInvApprovalDate(Optional ByVal announce As Boolean = True, _
                                  Optional ByVal manageProtection As Boolean = True)
    Dim wsInv As Worksheet
    Dim wsCoupaInvs As Worksheet

    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets(TenantSheetName("tracker"))
    Set wsCoupaInvs = ThisWorkbook.Sheets(TenantSheetName("coupa-invs"))
    On Error GoTo 0

    If wsInv Is Nothing Or wsCoupaInvs Is Nothing Then
        MsgBox "Required sheets are missing: '" & TenantSheetName("tracker") & "' and '" & _
               TenantSheetName("coupa-invs") & "'.", vbCritical
        Exit Sub
    End If

    Dim colStatus As String, colPO As String, colApproval As String
    colStatus = TenantColLetter("invoice-status")
    colPO = TenantColLetter("purchase-order-number")
    colApproval = TenantColLetter("invoice-approval-date")

    ' Column layout of the Coupa Invs sheet, as pinned by the Process*DataFile step.
    Const COUPA_PO_COL As String = "A"
    Const COUPA_HISTORY_COL As String = "M"

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False
    regex.pattern = "\s*(\d{2}/\d{2}/\d{2,4})\s*-\s*Pay\s*Invoice\s*Status:\s*" & _
                    "\[pending_document_approval\]\s*to\s*\[ready_to_pay\]\s*by"

    Dim lastRowInv As Long
    lastRowInv = wsInv.Cells(wsInv.Rows.Count, TenantColLetter("store-number")).End(xlUp).Row

    PauseThinking

    If manageProtection Then
        wsInv.Activate
        UnprotectSheet
    End If

    ' Nothing below the header row -- ColumnValues would otherwise hand back Empty (lastRow <
    ' firstRow), and UBound on that raises a type mismatch rather than the no-op the original
    ' per-row loop (For i = 2 To lastRowInv, zero iterations) fell into.
    If lastRowInv < 2 Then
        If manageProtection Then ProtectSheet
        RestoreThinking
        If announce Then MsgBox "0 approval date(s) extracted from invoice history logs.", vbInformation
        Exit Sub
    End If

    ' Read status and PO once, compute every approval date in memory, then write the whole
    ' column back in a single operation -- rather than one status read, one PO read and one
    ' approval-date write per row, which is what a few hundred tracker rows had been costing
    ' on every Refresh.
    Dim extracted As Long
    Dim statusArr As Variant, poArr As Variant
    statusArr = ColumnValues(wsInv, colStatus, 2, lastRowInv)
    poArr = ColumnValues(wsInv, colPO, 2, lastRowInv)

    Dim resultArr() As Variant
    ReDim resultArr(1 To UBound(statusArr, 1), 1 To 1)

    Dim i As Long
    For i = 1 To UBound(statusArr, 1)
        If SafeText(statusArr(i, 1)) = "Approved" Then
            Dim approvalDate As String
            approvalDate = FindApprovalDateForPO(wsCoupaInvs, SafeText(poArr(i, 1)), _
                                                 COUPA_PO_COL, COUPA_HISTORY_COL, regex)

            If Len(approvalDate) > 0 Then
                resultArr(i, 1) = approvalDate
                extracted = extracted + 1
            Else
                resultArr(i, 1) = vbNullString
            End If
        Else
            ' Not approved: clear any date left from a previous run.
            resultArr(i, 1) = vbNullString
        End If
    Next i

    ' lastRowInv..2 is the tracker's whole existing row range, which a user-applied AutoFilter
    ' can hide rows within -- unlike the append-only write loops elsewhere in this codebase, this
    ' is exactly the hazard MirrorInvoiceBlock refuses outright and NormalizeIdentifierColumnsOn
    ' falls back for: an array write spanning a filter-hidden row can misalign values onto the
    ' wrong VISIBLE rows, not just skip the hidden one. Fall back to the per-cell path here too.
    If wsInv.FilterMode Then
        For i = 1 To UBound(statusArr, 1)
            wsInv.Cells(i + 1, colApproval).Value = resultArr(i, 1)
        Next i
    Else
        wsInv.Range(wsInv.Cells(2, colApproval), wsInv.Cells(lastRowInv, colApproval)).Value = resultArr
    End If

    If manageProtection Then ProtectSheet

    RestoreThinking

    If announce Then
        MsgBox extracted & " approval date(s) extracted from invoice history logs.", vbInformation
    End If
End Sub

' Searches every Coupa invoice row matching a PO number, returning the first approval date
' found in its history, formatted MM/dd/yyyy. Returns "" when none matches.
'
' A PO can carry several invoices, so all matches are walked rather than just the first.
Private Function FindApprovalDateForPO(ByVal wsCoupaInvs As Worksheet, _
                                       ByVal poNumber As String, _
                                       ByVal poCol As String, _
                                       ByVal historyCol As String, _
                                       ByVal regex As Object) As String
    If Len(Trim$(poNumber)) = 0 Then Exit Function

    Dim foundCell As Range
    Set foundCell = wsCoupaInvs.Range(poCol & ":" & poCol) _
                               .Find(What:=poNumber, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then Exit Function

    Dim firstAddr As String
    firstAddr = foundCell.Address

    Do
        Dim historyText As String
        historyText = CStr(wsCoupaInvs.Cells(foundCell.Row, historyCol).Value)

        If regex.Test(historyText) Then
            Dim matches As Object
            Set matches = regex.Execute(historyText)

            On Error Resume Next
            Dim dt As Date
            dt = DateValue(matches(0).SubMatches(0))
            If Err.Number = 0 Then
                FindApprovalDateForPO = VBA.Format$(dt, "MM/dd/yyyy")
            End If
            On Error GoTo 0

            If Len(FindApprovalDateForPO) > 0 Then Exit Function
        End If

        Set foundCell = wsCoupaInvs.Range(poCol & ":" & poCol).FindNext(foundCell)
    Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddr
End Function
