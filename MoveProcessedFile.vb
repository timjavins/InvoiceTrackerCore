' Moves a processed bill file into a subfolder beside it, and returns the new path, or "" if
' it did not move.
'
' Cosmetic only: nothing decides anything from where a file sits. The content hash in
' ProcessedBatchLog is what stops a re-run. So every failure here is a quiet "" for the
' caller to mention, never an error. An empty subfolderName (TenantProcessedFolder) disables
' the move.
Public Function MoveProcessedFile(ByVal filePath As String, ByVal subfolderName As String) As String
    Dim folder As String
    Dim destFolder As String
    Dim target As String
    Dim ext As String

    MoveProcessedFile = vbNullString
    On Error GoTo Failed

    If Len(Trim$(subfolderName)) = 0 Then Exit Function
    If LCase$(Left$(filePath, 4)) = "http" Then Exit Function
    If Len(Dir$(filePath)) = 0 Then Exit Function

    folder = Left$(filePath, InStrRev(filePath, "\"))
    destFolder = folder & Trim$(subfolderName)
    If Len(Dir$(destFolder, vbDirectory)) = 0 Then MkDir destFolder

    target = destFolder & "\" & FileNameFromPath(filePath)
    If Len(Dir$(target)) > 0 Then
        ext = GetFileExt(filePath)
        target = destFolder & "\" & GetBaseFileName(filePath) & " (" & Format$(Now, "yyyymmdd-hhnnss") & ")" & _
                 IIf(Len(ext) > 0, "." & ext, "")
    End If

    Name filePath As target
    MoveProcessedFile = target
    Exit Function

Failed:
    MoveProcessedFile = vbNullString
End Function
