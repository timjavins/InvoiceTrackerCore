' Returns a file's extension without the leading dot, or "" when there is none.
' Case is preserved; callers compare case-insensitively.
'
' Works on the path string rather than calling Dir() -- see GetBaseFileName.vb for why.
Public Function GetFileExt(ByVal filePath As String) As String
    Dim fName As String
    fName = FileNameFromPath(filePath)

    Dim dotPos As Long
    dotPos = InStrRev(fName, ".")

    If dotPos > 0 Then
        GetFileExt = VBA.Mid$(fName, dotPos + 1)
    End If
End Function
