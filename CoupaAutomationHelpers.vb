' Small helpers shared by the Coupa browser-automation routines. These have no WebDriver
' dependency, so they are safe to share; the Selenium routines themselves stay in each
' variant until they can be exercised.

' Ensures a folder path ends in a separator, so it can be concatenated with a file name.
Public Function EnsureTrailingBackslash(ByVal folderPath As String) As String
    If Len(folderPath) = 0 Then Exit Function

    If Right$(folderPath, 1) = "\" Then
        EnsureTrailingBackslash = folderPath
    Else
        EnsureTrailingBackslash = folderPath & "\"
    End If
End Function

' Records a requisition as skipped for this run, so a retry loop does not reprocess it.
' Keyed adds on a Collection raise when the key already exists; ignoring that is the
' intended behaviour.
Public Sub MarkReqSkipped(ByVal skippedReqs As Collection, ByVal reqId As String)
    On Error Resume Next
    skippedReqs.Add True, reqId
    On Error GoTo 0
End Sub

' Whether a requisition was already marked skipped this run.
Public Function IsReqSkipped(ByVal skippedReqs As Collection, ByVal reqId As String) As Boolean
    On Error Resume Next
    Dim value As Variant
    value = skippedReqs(reqId)
    IsReqSkipped = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Prompts for a folder, returning the raw path or "" if the user cancelled.
' Callers apply EnsureTrailingBackslash themselves, as they did before.
Public Function PickFolder(ByVal promptText As String) As String
    Dim folderDialog As Object
    Set folderDialog = Application.FileDialog(4)   ' msoFileDialogFolderPicker

    With folderDialog
        .Title = promptText
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = vbNullString
        End If
    End With
End Function
