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
Public Function ConvertStoreNumbersOn(ByVal ws As Worksheet, _
                                      ByVal colLetter As String, _
                                      Optional ByVal resolutions As Object) As Long
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
        updated(i, 1) = NormalizeStoreNumber(storeNumbers(i, 1), resolutions)
        If Len(updated(i, 1)) > 0 Then count = count + 1
    Next i

    targetRange.NumberFormat = "@"
    targetRange.Value = updated

    ConvertStoreNumbersOn = count
End Function

' Pure conversion: 4-digit zero-padded text, or "" when there is nothing to convert.
'
' resolutions, when supplied, maps a raw supplier value that is not a plain number -- "0391-A",
' "2242-B" -- to the store number a human said it means. See ResolveStoreNumbers.vb, which
' collects the answers before an import writes any rows.
Public Function NormalizeStoreNumber(ByVal value As Variant, _
                                     Optional ByVal resolutions As Object) As String
    If IsEmpty(value) Then Exit Function

    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function

    ' Covers every ordinary case, including the supplier's inconsistent zero-padding: 28,
    ' 163, 00403 and 02223 all land on four digits. Format$ pads to a minimum width and never
    ' truncates, so a genuine five-digit number survives intact rather than being mangled.
    If IsNumeric(value) Then
        NormalizeStoreNumber = Format$(value, "0000")
        Exit Function
    End If

    If Not resolutions Is Nothing Then
        If resolutions.Exists(text) Then
            NormalizeStoreNumber = CStr(resolutions(text))
            Exit Function
        End If
    End If

    ' Nobody has said what this means, so hand it back untouched and let it fail a lookup
    ' loudly.
    '
    ' This used to be Right$("0000" & text, 4), which is how "0391-A" became "91-A". That
    ' guess only looked harmless because the surviving "-A" broke a VLOOKUP: one character
    ' shorter and "S0391" would have yielded a clean "0391", posting the bill to a real store
    ' on no evidence. A wrong store that validates is far worse than an #N/A.
    NormalizeStoreNumber = text
End Function

' Distinct store values on a source sheet that NormalizeStoreNumber cannot resolve by itself.
'
' Returns a Dictionary of raw text -> the first row it appeared on, so a caller can show that
' row's context when asking a human what the value means. Keyed on the raw text so a value
' repeated across fifty rows is only asked about once.
Public Function CollectNonConformingStores(ByVal ws As Worksheet, _
                                           ByVal headerRow As Long, _
                                           ByVal lastRow As Long, _
                                           ByVal storeCol As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set CollectNonConformingStores = result

    If ws Is Nothing Then Exit Function
    If storeCol < 1 Then Exit Function

    Dim r As Long
    Dim raw As Variant
    Dim text As String
    For r = headerRow + 1 To lastRow
        raw = ws.Cells(r, storeCol).Value
        If Not IsEmpty(raw) Then
            text = Trim$(CStr(raw))
            If Len(text) > 0 Then
                If Not IsNumeric(raw) Then
                    If Not result.Exists(text) Then result.Add text, r
                End If
            End If
        End If
    Next r
End Function
