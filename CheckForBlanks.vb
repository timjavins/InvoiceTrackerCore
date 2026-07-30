' This function checks for blank cells in the specified range and returns True if any are found.
Function CheckForBlanks(ws As Worksheet, headerRow As Long, lastRow As Long, colIndexes As Object, targetHeaders As Variant) As Boolean
    Dim i As Long, r As Long
    Dim colName As String
    Dim colNum As Long
    Dim hasBlanks As Boolean
    hasBlanks = False
    For i = LBound(targetHeaders) To UBound(targetHeaders)
        colName = targetHeaders(i)
        If colIndexes.Exists(colName) Then
            colNum = colIndexes(colName)
            For r = headerRow + 1 To lastRow
                If VBA.Trim(ws.Cells(r, colNum).Value) = "" Then
                    hasBlanks = True
                    MsgBox "Blank cell found in the " & colName & " column of the " & ws.name & " worksheet. Please fix the data.", vbExclamation
                    Exit For
                End If
            Next r
        Else
            MsgBox "Column """ & colName & """ was not found in the """ & ws.name & """ sheet's header row. Please make sure the """ & ws.name & _
                """ worksheet has the """ & colName & """ data. Please look for typos, missing data, and any other issues.", vbExclamation
            hasBlanks = True
        End If
        If hasBlanks Then Exit For
    Next i
    
    CheckForBlanks = hasBlanks
End Function