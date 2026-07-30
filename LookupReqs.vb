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

Public Sub LookupReqs(Optional ByVal announce As Boolean = True)
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
    colTrackerInvoice = TenantColLetter("submitted-invoice-number")
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
