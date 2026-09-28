' SHA-256 of a file's bytes, as 64 lowercase hex characters, or "" if it cannot be computed.
'
' Identifies a bill file by content rather than by name: a vendor re-sending the same file
' under a new name, or someone copying it back into the inbox, still hashes the same. See the
' batch-tracking spec, section 1.
'
' Shells to certutil, which ships with Windows, rather than calling CryptoAPI: that needs
' Declare, which a class module (ThisWorkbook, where the stack lives) forbids as a public
' member and which would tie this module to PtrSafe signatures. Output goes through a temp
' file because WScript.Shell.Exec always flashes a console window; Run with style 0 does not.
'
' "" is the failure signal, never Err.Raise: callers treat "" as "stop before opening the
' file", and the test harness cannot survive a raised error.
'
' Returns the hash only; it stores nothing. AddNewBills passes it to ProcessedBatchLog, which
' keeps it in column A of the hidden "Processed Batches" sheet in the tracker workbook.
Public Function HashFile(ByVal filePath As String) As String
    Dim outPath As String
    Dim wsh As Object
    Dim fileNum As Integer
    Dim textLine As String
    Dim candidate As String

    HashFile = vbNullString
    On Error GoTo Failed

    If LCase$(Left$(filePath, 4)) = "http" Then Exit Function
    If Len(Dir$(filePath)) = 0 Then Exit Function

    outPath = Environ$("TEMP") & "\hashfile_" & Format$(Now, "yyyymmddhhnnss") & "_" & _
              CStr(Int(Rnd * 1000000)) & ".txt"
    Set wsh = CreateObject("WScript.Shell")
    If wsh.Run("cmd /c certutil -hashfile """ & filePath & """ SHA256 > """ & outPath & """", 0, True) <> 0 Then
        GoTo Cleanup
    End If

    ' certutil prints a label line, the digest, then a status line. Older Windows puts a
    ' space between each byte of the digest, so strip spaces before testing the shape.
    fileNum = FreeFile
    Open outPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, textLine
        candidate = LCase$(Replace(Trim$(textLine), " ", ""))
        If HfIsHexDigest(candidate) Then
            HashFile = candidate
            Exit Do
        End If
    Loop
    Close #fileNum
    fileNum = 0

Cleanup:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    If Len(outPath) > 0 Then If Len(Dir$(outPath)) > 0 Then Kill outPath
    Exit Function

Failed:
    HashFile = vbNullString
    Resume Cleanup
End Function

Private Function HfIsHexDigest(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) <> 64 Then Exit Function
    For i = 1 To 64
        If InStr(1, "0123456789abcdef", Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    HfIsHexDigest = True
End Function
