' Re-applies an AutoFilter to a worksheet, normally to the row RemoveAutoFilter reported.
' Defaults to row 1.
'
' Sets failureState when the requested row is out of range, so callers that check it can
' react. JCI previously had this line commented out with a TODO because it had no such
' global; core declares it in Header.vb, so the check is live in both variants now.
Public Sub RestoreAutoFilter(ByVal wsTarget As Worksheet, Optional ByVal filterRow As Long = 1)
    If wsTarget Is Nothing Then Exit Sub

    On Error Resume Next

    If filterRow > 0 And filterRow <= wsTarget.Rows.Count Then
        wsTarget.Rows(filterRow).AutoFilter
    Else
        failureState = True
        Debug.Print "RestoreAutoFilter: Invalid row specified for AutoFilter."
    End If

    On Error GoTo 0
End Sub
