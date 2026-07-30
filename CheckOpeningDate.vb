' Flags invoices that may be covered by a new-store warranty: work billed within the
' warranty window of a store's opening. Matching rows are highlighted, marked WARRANTY in
' the requisition column when it is still blank, and annotated in notes.
'
' Union of both variants. Securitas gated the check to repair work (a Service Order number
' in transaction details, or "REPAIR" in the invoice type) -- JCI had no such gate and
' checked every row. JCI had the error handling, sheet validation, and a summary of store
' numbers missing from the BU List, which Securitas lacked. Core does both: the repair gate
' applies when the tenant declares the columns it needs, and the safety net is always on.
'
' Window length, sheet names, and columns all come from TenantConfig.

Private Const WARRANTY_NOTE As String = "May be a warranty repair. The invoice date was "

Public Sub CheckOpeningDate(Optional ByVal announce As Boolean = True, _
                            Optional ByVal manageProtection As Boolean = True)
    Dim wsInvoices As Worksheet
    Dim wsBUList As Worksheet
    Dim missingStores As Object
    Dim protectionOn As Boolean

    Set missingStores = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set wsInvoices = ThisWorkbook.Sheets(TenantSheetName("tracker"))
    Set wsBUList = ThisWorkbook.Sheets(TenantSheetName("bu-list"))
    On Error GoTo 0

    If wsInvoices Is Nothing Or wsBUList Is Nothing Then
        MsgBox "Required sheets are missing: '" & TenantSheetName("tracker") & "' and '" & _
               TenantSheetName("bu-list") & "'.", vbCritical
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    If manageProtection Then
        wsInvoices.Activate
        UnprotectSheet
        protectionOn = True
    End If

    ' BU List columns are resolved by header name, not position.
    Dim buHeaderDict As Object
    Set buHeaderDict = GetHeaderColumnIndexes(wsBUList, 1, Array("Store", "StoreOpenDate"))

    If buHeaderDict Is Nothing Then GoTo MissingHeaders
    If Not buHeaderDict.Exists("Store") Then GoTo MissingHeaders
    If Not buHeaderDict.Exists("StoreOpenDate") Then GoTo MissingHeaders

    Dim colBUStore As Long, colBUOpen As Long
    colBUStore = buHeaderDict("Store")
    colBUOpen = buHeaderDict("StoreOpenDate")

    ' Tracker columns come from the tenant's layout.
    Dim colStore As String, colDate As String, colReq As String, colNotes As String
    colStore = TenantColLetter("store-number")
    colDate = TenantColLetter("request-date")
    colReq = TenantColLetter("requisition-number")
    colNotes = TenantColLetter("notes")

    Dim windowDays As Long
    windowDays = TenantWarrantyWeeks() * 7

    Dim lastRowInvoices As Long, lastRowBUList As Long
    lastRowInvoices = wsInvoices.Cells(wsInvoices.Rows.Count, colStore).End(xlUp).Row
    lastRowBUList = wsBUList.Cells(wsBUList.Rows.Count, colBUStore).End(xlUp).Row

    Dim flagged As Long
    Dim i As Long
    For i = 2 To lastRowInvoices
        If IsRepairRow(wsInvoices, i) Then

            If Not IsEmpty(wsInvoices.Cells(i, colStore).Value) And _
               Not IsEmpty(wsInvoices.Cells(i, colDate).Value) Then

                Dim storeNumber As String
                storeNumber = CStr(wsInvoices.Cells(i, colStore).Value)

                Dim foundCell As Range
                Set foundCell = wsBUList.Range(wsBUList.Cells(2, colBUStore), _
                                               wsBUList.Cells(lastRowBUList, colBUStore)) _
                                       .Find(What:=storeNumber, LookIn:=xlValues, _
                                             LookAt:=xlWhole, MatchCase:=False)

                If foundCell Is Nothing Then
                    ' Report each unknown store once rather than per row.
                    If Not missingStores.Exists(storeNumber) Then
                        missingStores.Add storeNumber, True
                    End If
                Else
                    Dim requestDate As Date, storeOpenDate As Date
                    requestDate = wsInvoices.Cells(i, colDate).Value
                    storeOpenDate = wsBUList.Cells(foundCell.Row, colBUOpen).Value

                    If requestDate < storeOpenDate + windowDays Then
                        Dim days As Long
                        days = DateDiff("d", storeOpenDate, requestDate)

                        wsInvoices.Range(wsInvoices.Cells(i, colStore), _
                                         wsInvoices.Cells(i, colNotes)).Interior.Color = RGB(255, 192, 203)

                        If Len(CStr(wsInvoices.Cells(i, colReq).Value)) = 0 Then
                            wsInvoices.Cells(i, colReq).Value = "WARRANTY"
                        End If

                        If InStr(1, CStr(wsInvoices.Cells(i, colNotes).Value), WARRANTY_NOTE, vbTextCompare) = 0 Then
                            wsInvoices.Cells(i, colNotes).Value = WARRANTY_NOTE & days & _
                                " days from store opening. " & wsInvoices.Cells(i, colNotes).Value
                        End If

                        flagged = flagged + 1
                    End If
                End If
            End If
        End If
    Next i

    If manageProtection Then
        ProtectSheet
        protectionOn = False
    End If

    If missingStores.Count > 0 Then
        MsgBox "These store numbers were not found in '" & TenantSheetName("bu-list") & "':" & _
               vbCrLf & Join(missingStores.Keys, vbCrLf), vbExclamation, "Missing Store Numbers"
    End If

    If announce Then
        MsgBox flagged & " row(s) highlighted for warranty review. Request dates were " & _
               "compared against store opening dates.", vbInformation
    End If

    Exit Sub

MissingHeaders:
    If protectionOn Then ProtectSheet
    MsgBox "Required headers were not found on '" & TenantSheetName("bu-list") & _
           "'. Expected: Store and StoreOpenDate.", vbCritical
    Exit Sub

ErrorHandler:
    If protectionOn Then ProtectSheet
    MsgBox "CheckOpeningDate failed: " & Err.Description, vbCritical
End Sub

' Whether a row represents repair work, and so is a warranty candidate.
'
' Securitas gates on a Service Order number in transaction details or "REPAIR" in the
' invoice type. A tenant without those columns has no way to distinguish repair rows, so
' every row is a candidate -- which is what JCI did.
Private Function IsRepairRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    Dim colDetails As String, colType As String

    On Error Resume Next
    colDetails = TenantColLetter("transaction-details")
    colType = TenantColLetter("invoice-type")
    On Error GoTo 0

    If Len(colDetails) = 0 And Len(colType) = 0 Then
        IsRepairRow = True
        Exit Function
    End If

    If Len(colType) > 0 Then
        If InStr(1, CStr(ws.Cells(rowIndex, colType).Value), "REPAIR", vbTextCompare) > 0 Then
            IsRepairRow = True
            Exit Function
        End If
    End If

    If Len(colDetails) > 0 Then
        IsRepairRow = HasServiceOrderNumber(CStr(ws.Cells(rowIndex, colDetails).Value))
    End If
End Function

' True when the text starts with a Service Order reference, e.g. "SO#12345", "SO: 900".
Private Function HasServiceOrderNumber(ByVal text As String) As Boolean
    Static regex As Object

    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = "^\s*SO[#:\s]*\d"
        regex.IgnoreCase = True
        regex.Global = False
    End If

    HasServiceOrderNumber = regex.Test(text)
End Function
