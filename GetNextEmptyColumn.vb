' Finds the first column at or after startCol whose header cell (row 1) is empty. Used to
' append extra columns after the ones a caller pinned to fixed positions.
Public Function GetNextEmptyColumn(ByVal wsTarget As Worksheet, ByVal startCol As Long) As Long
    If wsTarget Is Nothing Then Exit Function
    If startCol < 1 Then startCol = 1

    Do While Len(CStr(wsTarget.Cells(1, startCol).Value)) > 0
        startCol = startCol + 1
    Loop

    GetNextEmptyColumn = startCol
End Function
