' Returns a file name without its extension, given a full path.
'
' Works on the path string rather than calling Dir(), which both variants did. Dir() hits
' the filesystem and returns an empty string for a file that does not exist, so callers
' silently got "" instead of the name they asked about.
Public Function GetBaseFileName(ByVal filePath As String) As String
    Dim fName As String
    fName = FileNameFromPath(filePath)

    Dim dotPos As Long
    dotPos = InStrRev(fName, ".")

    If dotPos > 0 Then
        GetBaseFileName = VBA.Left$(fName, dotPos - 1)
    Else
        GetBaseFileName = fName
    End If
End Function

' Strips any directory portion from a path. Handles both separators.
Public Function FileNameFromPath(ByVal filePath As String) As String
    Dim result As String
    result = VBA.Trim$(filePath)

    Dim sep As Long
    sep = InStrRev(result, "\")
    If sep = 0 Then sep = InStrRev(result, "/")

    If sep > 0 Then result = VBA.Mid$(result, sep + 1)

    FileNameFromPath = result
End Function
