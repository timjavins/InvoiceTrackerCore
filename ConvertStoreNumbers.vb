' Normalizes the tracker's store number column to 4-digit zero-padded text
' (e.g. 1 -> "0001"), which is the Nordstrom store number format. See CONTEXT.md.
'
' The sheet and column come from TenantConfig, so this works against either tracker
' layout without knowing it.

' Normalizes store numbers on the tenant's tracker sheet.
'
' Silent by default -- it is called from orchestration (Refresh, AddNewBills) where a
' modal dialog per step is noise. Pass announce:=True for interactive use.
'
' manageProtection defaults to True because some callers protect the sheet before
' calling and rely on each step unprotecting itself. Callers that already unprotect
' around a whole sequence should pass False to avoid redundant toggling.
Public Sub ConvertStoreNumbers(Optional ByVal announce As Boolean = False, _
                               Optional ByVal manageProtection As Boolean = True)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TenantSheetName("tracker"))

    If manageProtection Then
        ws.Activate
        UnprotectSheet
    End If

    Dim converted As Long
    converted = ConvertStoreNumbersOn(ws, TenantColLetter("store-number"))

    If manageProtection Then ProtectSheet

    If announce Then
        MsgBox converted & " store number(s) normalized to four-digit text.", vbInformation
    End If
End Sub

' Does the work against an explicit sheet and column, so it is callable without the
' tenant's defaults and testable in isolation. Returns the count of non-empty values
' written.
Public Function ConvertStoreNumbersOn(ByVal ws As Worksheet, ByVal colLetter As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row

    ' Nothing below the header row.
    If lastRow < 2 Then Exit Function

    Dim targetRange As Range
    Set targetRange = ws.Range(colLetter & "2:" & colLetter & lastRow)

    Dim storeNumbers As Variant
    storeNumbers = targetRange.Value

    ' A single cell comes back as a bare value rather than a 2D array.
    If Not IsArray(storeNumbers) Then
        Dim single_ As Variant
        ReDim single_(1 To 1, 1 To 1)
        single_(1, 1) = storeNumbers
        storeNumbers = single_
    End If

    Dim updated() As Variant
    ReDim updated(1 To UBound(storeNumbers, 1), 1 To 1)

    Dim i As Long
    Dim count As Long
    For i = 1 To UBound(storeNumbers, 1)
        updated(i, 1) = NormalizeStoreNumber(storeNumbers(i, 1))
        If Len(updated(i, 1)) > 0 Then count = count + 1
    Next i

    targetRange.NumberFormat = "@"
    targetRange.Value = updated

    ConvertStoreNumbersOn = count
End Function

' Pure conversion: 4-digit zero-padded text, or "" when there is nothing to convert.
Public Function NormalizeStoreNumber(ByVal value As Variant) As String
    If IsEmpty(value) Then Exit Function

    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function

    If IsNumeric(value) Then
        NormalizeStoreNumber = Format$(value, "0000")
    Else
        ' Already-padded or non-numeric identifiers keep their trailing 4 characters.
        NormalizeStoreNumber = Right$("0000" & text, 4)
    End If
End Function
