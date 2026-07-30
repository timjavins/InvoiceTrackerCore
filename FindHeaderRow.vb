' Finds the row carrying a set of required column headers, scanning the top of a sheet.
' Returns 0 when no row within maxRowsToCheck contains all of them.
'
' Vendor bill files often carry title or metadata rows above the real header, so the header
' is not necessarily row 1.
'
' Counts DISTINCT headers found. The previous version incremented a counter on every match,
' so a row where one required header appeared twice could reach the target count without
' every required header actually being present.
Public Function FindHeaderRow(ByVal ws As Worksheet, _
                              ByVal requiredCols As Variant, _
                              Optional ByVal maxRowsToCheck As Long = 40) As Long
    If ws Is Nothing Then Exit Function

    Dim requiredCount As Long
    requiredCount = UBound(requiredCols) - LBound(requiredCols) + 1
    If requiredCount < 1 Then Exit Function

    Dim found As Object
    Set found = CreateObject("Scripting.Dictionary")

    Dim r As Long, c As Long, i As Long
    Dim lastCol As Long
    Dim cellVal As String

    For r = 1 To maxRowsToCheck
        found.RemoveAll
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column

        For c = 1 To lastCol
            cellVal = VBA.Trim$(VBA.UCase$(CStr(ws.Cells(r, c).Value)))
            If Len(cellVal) > 0 Then
                For i = LBound(requiredCols) To UBound(requiredCols)
                    If cellVal = VBA.Trim$(VBA.UCase$(CStr(requiredCols(i)))) Then
                        If Not found.Exists(cellVal) Then found.Add cellVal, c
                        Exit For
                    End If
                Next i
            End If
        Next c

        If found.Count = requiredCount Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r

    FindHeaderRow = 0
End Function
