' Replaces empty cells with a single space in the Coupa columns the tracker's XLOOKUP
' formulas read.
'
' XLOOKUP cannot tell "not found" from "found a blank cell" — both come back empty — so a
' blank in a lookup column silently matches and returns nothing useful. Writing a space makes
' the cell non-blank without changing what a human sees.
'
' Columns are resolved BY HEADER NAME. Both variants hardcoded F / I / J / K, and Securitas's
' comment misdescribed them: against the real exports those are PO Number on Coupa Reqs, and
' Payment Date / Payment Due Date (Oracle) / Payment Num's on Coupa Invs -- not the
' "Supplier Invoice #, Supplier Invoice Date, Payment #" the comment claimed. Only pinned
' columns have fixed positions, so a letter is not a reliable way to identify these.
'
' Grafted from both variants: JCI's table-driven loop over sheet/column pairs, with
' Securitas's blank test, which also catches empty strings and not just IsEmpty.

Public Sub ReplaceNullValues(Optional ByVal announce As Boolean = False)
    Dim wsCoupaReqs As Worksheet
    Dim wsCoupaInvs As Worksheet

    On Error Resume Next
    Set wsCoupaReqs = ThisWorkbook.Sheets(TenantSheetName("coupa-reqs"))
    Set wsCoupaInvs = ThisWorkbook.Sheets(TenantSheetName("coupa-invs"))
    On Error GoTo 0

    If wsCoupaReqs Is Nothing Or wsCoupaInvs Is Nothing Then
        MsgBox "Required sheets are missing: '" & TenantSheetName("coupa-reqs") & "' and '" & _
               TenantSheetName("coupa-invs") & "'.", vbCritical
        Exit Sub
    End If

    PauseThinking

    Dim coerced As Long
    coerced = CoerceBlanksOn(wsCoupaReqs, Array("PO Number"))
    coerced = coerced + CoerceBlanksOn(wsCoupaInvs, _
                  Array("Payment Date", "Payment Due Date (Oracle)", "Payment Num's"))

    RestoreThinking

    If announce Then
        MsgBox coerced & " blank cell(s) replaced so lookups resolve correctly.", vbInformation
    End If
End Sub

' Replaces blanks with a space in the named columns of one sheet. Returns how many cells
' changed. Columns that are not present are skipped, since a saved Coupa view may omit them.
Private Function CoerceBlanksOn(ByVal ws As Worksheet, ByVal headerNames As Variant) As Long
    ' These sheets can be protected, and writing a protected cell fails.
    UnprotectSheetOn ws

    Dim headers As Object
    Set headers = GetHeaderColumnIndexes(ws, 1, headerNames)
    If headers Is Nothing Then
        ProtectSheetOn ws   ' unprotected above; do not leave it open
        Exit Function
    End If

    Dim i As Long
    Dim headerName As String
    For i = LBound(headerNames) To UBound(headerNames)
        headerName = CStr(headerNames(i))

        If Not headers.Exists(headerName) Then
            Debug.Print "ReplaceNullValues: '" & headerName & "' not found on " & ws.name & _
                        " -- skipped."
        Else
            CoerceBlanksOn = CoerceBlanksInColumn(ws, headers(headerName)) + CoerceBlanksOn
        End If
    Next i

    ProtectSheetOn ws
End Function

' Reads one column, replaces blanks in memory, writes it back in a single operation.
Private Function CoerceBlanksInColumn(ByVal ws As Worksheet, ByVal colIndex As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    If lastRow < 1 Then Exit Function

    Dim target As Range
    Set target = ws.Range(ws.Cells(1, colIndex), ws.Cells(lastRow, colIndex))

    Dim columnData As Variant
    columnData = target.Value2

    ' A single-cell range comes back as a bare value rather than a 2D array.
    If Not IsArray(columnData) Then
        Dim single_ As Variant
        ReDim single_(1 To 1, 1 To 1)
        single_(1, 1) = columnData
        columnData = single_
    End If

    Dim changed As Long
    Dim r As Long
    For r = 1 To UBound(columnData, 1)
        ' Catch both a truly empty cell and one holding an empty string.
        If IsEmpty(columnData(r, 1)) Or Len(CStr(columnData(r, 1) & "")) = 0 Then
            columnData(r, 1) = " "
            changed = changed + 1
        End If
    Next r

    If changed > 0 Then target.Value2 = columnData

    CoerceBlanksInColumn = changed
End Function
