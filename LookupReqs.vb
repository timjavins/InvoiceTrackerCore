' Backfills the tracker's requisition number for rows that do not have one yet.
'
' This is the inbound half of the REQ# round-trip (see ADR-0001): MakeFlatFile stamps the
' submitted invoice # into Coupa's "Supplier Part Number" when raising a requisition, Coupa
' assigns the REQ #, and this reads it back on the next export. It therefore only resolves
' requisitions that originated from a flat file this system generated.
'
' Rows whose requisition column already holds a number are left alone; so are markers like
' "WARRANTY" or "DUPLICATE", which are not requisition numbers but must not be overwritten
' either -- only genuinely empty cells are filled.
'
' The join column on the Coupa Reqs sheet is resolved BY HEADER NAME. Securitas hardcoded
' column H, which only worked by luck: ProcessRequisitionDataFile pins Req #, Status,
' PO Number, Current Approver, Time with Current Approver and Copied From to fixed columns
' and appends everything else in export order, so "Supplier Part Number" has no guaranteed
' position.

' Which tracker column holds the value that joins a row to its Coupa requisition.
'
' Securitas and JCI join on the submitted invoice # -- a number the supplier gave them. A tracker
' that raises requisitions BEFORE any supplier document exists has no such number and mints its own
' key instead, which is not an invoice number of either kind (see ADR-0001 on why conflating the
' two identities is a bug waiting to happen).
'
' So: prefer an explicit req-join-key, fall back to submitted-invoice-number. The fallback keeps
' every existing tenant working without touching its TenantConfig.
Public Function ReqJoinKeyColumn() As String
    Dim col As String

    ' On Error Resume Next below is there to absorb one specific case: a tenant that has not
    ' declared "req-join-key" in its TenantConfig, which is the expected, common state for a
    ' tenant that hasn't opted in yet. But it necessarily absorbs more than that -- it will just
    ' as silently swallow a genuine bug inside that tenant's TenantColLetter. There is no
    ' discriminator available to narrow on: every tenant's TenantColLetter raises the same
    ' Err.Raise 5 for an unknown concept (see TenantConfig.vb in each variant repo), and error 5
    ' from a real defect is not distinguishable from error 5 for "concept not declared". Adding
    ' one (e.g. matching on Err.Source or Err.Description text) would be guessing at an implicit
    ' contract those functions never promised to keep, for no proven gain, and it would risk the
    ' two live workbooks (Securitas, JCI) this fallback exists to protect. So this breadth is
    ' accepted deliberately, not missed: a broken req-join-key case in a tenant's config degrades
    ' quietly to the submitted-invoice-number fallback below instead of failing loudly.
    On Error Resume Next
    col = TenantColLetter("req-join-key")
    On Error GoTo 0
    Err.Clear   ' GoTo 0 disables the handler but leaves Err.Number populated from the attempt
                ' above; clear it explicitly so a future insertion between here and the return
                ' cannot mistake that stale error for one of its own.

    If Len(col) = 0 Then col = TenantColLetter("submitted-invoice-number")

    ReqJoinKeyColumn = col
End Function

' manageProtection defaults to True so the sub works from a caller that left the sheet
' protected (UpdateCoupaData does). The two sibling Refresh paths differ, for real reasons:
' JCI's Refresh brackets its whole sequence in its own UnprotectSheet/ProtectSheet window
' (Refresh.vb:21,56) and passes False so this sub does not redundantly re-toggle inside it.
' Securitas's Refresh does not unprotect around the sequence at all -- its sheets are
' genuinely protected, and it defines a TenantSheetPassword -- so it passes True and relies
' on this sub to manage protection itself.
Public Sub LookupReqs(Optional ByVal announce As Boolean = True, _
                      Optional ByVal manageProtection As Boolean = True)
    Dim wsTracker As Worksheet
    Dim wsCoupaReqs As Worksheet

    On Error Resume Next
    Set wsTracker = ThisWorkbook.Sheets(TenantSheetName("tracker"))
    Set wsCoupaReqs = ThisWorkbook.Sheets(TenantSheetName("coupa-reqs"))
    On Error GoTo 0

    If wsTracker Is Nothing Or wsCoupaReqs Is Nothing Then
        MsgBox "Required sheets are missing: '" & TenantSheetName("tracker") & "' and '" & _
               TenantSheetName("coupa-reqs") & "'.", vbCritical
        Exit Sub
    End If

    ' Where the submitted invoice # lives on each sheet.
    Dim colTrackerInvoice As String, colTrackerReq As String
    colTrackerInvoice = ReqJoinKeyColumn()
    colTrackerReq = TenantColLetter("requisition-number")

    Dim reqsHeaders As Object
    Set reqsHeaders = GetHeaderColumnIndexes(wsCoupaReqs, 1, _
                          Array("Req #", "Supplier Part Number"))

    If reqsHeaders Is Nothing Then GoTo MissingHeaders
    If Not reqsHeaders.Exists("Req #") Then GoTo MissingHeaders
    If Not reqsHeaders.Exists("Supplier Part Number") Then GoTo MissingHeaders

    Dim colReqNum As Long, colPartNum As Long
    colReqNum = reqsHeaders("Req #")
    colPartNum = reqsHeaders("Supplier Part Number")

    Dim lastRow As Long, lastRowReqs As Long
    lastRow = wsTracker.Cells(wsTracker.Rows.Count, colTrackerInvoice).End(xlUp).Row
    lastRowReqs = wsCoupaReqs.Cells(wsCoupaReqs.Rows.Count, colReqNum).End(xlUp).Row

    If lastRow < 2 Or lastRowReqs < 2 Then Exit Sub

    If manageProtection Then
        wsTracker.Activate
        UnprotectSheet
    End If

    ' Read both sides into memory; write the tracker column back in one operation.
    Dim coupaReqNums As Variant, coupaPartNums As Variant
    coupaReqNums = wsCoupaReqs.Range(wsCoupaReqs.Cells(2, colReqNum), _
                                     wsCoupaReqs.Cells(lastRowReqs, colReqNum)).Value
    coupaPartNums = wsCoupaReqs.Range(wsCoupaReqs.Cells(2, colPartNum), _
                                      wsCoupaReqs.Cells(lastRowReqs, colPartNum)).Value

    Dim trackerReqNums As Variant, trackerInvoices As Variant
    trackerReqNums = wsTracker.Range(colTrackerReq & "2:" & colTrackerReq & lastRow).Value
    trackerInvoices = wsTracker.Range(colTrackerInvoice & "2:" & colTrackerInvoice & lastRow).Value

    ' Map submitted invoice # -> REQ #, keeping the first requisition seen for each.
    Dim reqLookup As Object
    Set reqLookup = CreateObject("Scripting.Dictionary")
    reqLookup.CompareMode = 1   ' vbTextCompare

    Dim i As Long
    Dim partNum As String
    For i = 1 To UBound(coupaPartNums, 1)
        partNum = Trim$(CStr(coupaPartNums(i, 1)))
        If Len(partNum) > 0 Then
            If Not reqLookup.Exists(partNum) Then
                reqLookup.Add partNum, coupaReqNums(i, 1)
            End If
        End If
    Next i

    Dim foundCount As Long
    Dim invoiceNum As String
    For i = 1 To UBound(trackerReqNums, 1)
        If IsBlankRequisition(trackerReqNums(i, 1)) Then
            invoiceNum = Trim$(CStr(trackerInvoices(i, 1)))
            If Len(invoiceNum) > 0 Then
                If reqLookup.Exists(invoiceNum) Then
                    trackerReqNums(i, 1) = reqLookup(invoiceNum)
                    foundCount = foundCount + 1
                End If
            End If
        End If
    Next i

    wsTracker.Range(colTrackerReq & "2:" & colTrackerReq & lastRow).Value = trackerReqNums

    If manageProtection Then ProtectSheet

    If announce And foundCount > 0 Then
        MsgBox foundCount & " requisition number(s) found for invoices that had none listed.", _
               vbInformation
    End If

    Exit Sub

MissingHeaders:
    MsgBox "Could not find 'Req #' and 'Supplier Part Number' on '" & _
           TenantSheetName("coupa-reqs") & "'. Re-import the requisition export.", vbCritical
End Sub

' Whether a requisition cell is genuinely empty and so safe to fill.
'
' Deliberately conservative: markers such as WARRANTY or DUPLICATE live in this column and
' are not requisition numbers, but overwriting them would destroy information. Securitas's
' version filled any non-numeric value, which would have clobbered them.
Private Function IsBlankRequisition(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsNull(value) Then
        IsBlankRequisition = True
        Exit Function
    End If
    IsBlankRequisition = (Len(Trim$(CStr(value))) = 0)
End Function
