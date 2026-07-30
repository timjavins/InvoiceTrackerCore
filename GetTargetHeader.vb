Function GetTargetHeader(ws As Worksheet) As Variant
    Dim targetHeaders As Variant
    Dim col As Long
    Dim colDict As Object
    Set colDict = CreateObject("Scripting.Dictionary")

    ' Read the headers from the first row of the target worksheet
    targetHeaders = ws.Rows(1).Value

    ' Loop through the target headers and find their columns
    For col = LBound(targetHeaders, 2) To UBound(targetHeaders, 2)
        If Not IsEmpty(targetHeaders(1, col)) Then
            ' Check if the header is already in the dictionary
            If Not colDict.Exists(targetHeaders(1, col)) Then
                ' Add the header to the dictionary if it's not already there.
                ' We use the header as the key and the column number as the value
                ' to ensure we can retrieve the column number later.
                colDict.Add targetHeaders(1, col), col
            Else
                ' If the header already exists, exit the function
                MsgBox "This script does not support duplicate header names. Please remove the duplicate header " & targetHeaders(1, col) & " in the " & ws.Name & " sheet.", vbExclamation
                Set GetTargetHeader = Nothing
                Exit Function
            End If
        End If
    Next col

    ' Return the found headers as a dictionary
    Set GetTargetHeader = colDict
End Function
