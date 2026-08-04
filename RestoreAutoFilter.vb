' Re-applies an AutoFilter to a worksheet, normally to the row RemoveAutoFilter reported.
' Defaults to row 1.
'
' Unprotects the sheet for the moment it takes to apply the filter, then restores protection.
' Protection blocks this even with AllowFiltering:=True -- that permission lets a user operate
' a filter that already exists, but creating one still needs the sheet unprotected. Every step
' in a refresh now reprotects on its way out, so by the time this runs the sheet is protected
' and the call was failing silently under On Error Resume Next: filters removed, never restored.
'
' Sets failureState when the row is out of range or the filter could not be applied, so callers
' that check it can report rather than assume success.
Public Sub RestoreAutoFilter(ByVal wsTarget As Worksheet, Optional ByVal filterRow As Long = 1)
    If wsTarget Is Nothing Then Exit Sub

    If filterRow < 1 Or filterRow > wsTarget.Rows.Count Then
        failureState = True
        Debug.Print "RestoreAutoFilter: row " & filterRow & " is out of range."
        Exit Sub
    End If

    ' Remember whether it was protected so the sheet is left as it was found.
    Dim wasProtected As Boolean
    wasProtected = wsTarget.ProtectContents

    If wasProtected Then UnprotectSheetOn wsTarget

    On Error Resume Next
    Err.Clear
    wsTarget.Rows(filterRow).AutoFilter
    Dim applyError As String
    If Err.Number <> 0 Then applyError = Err.Description
    Err.Clear
    On Error GoTo 0

    If wasProtected Then ProtectSheetOn wsTarget

    If Len(applyError) > 0 Then
        failureState = True
        Debug.Print "RestoreAutoFilter: could not apply a filter to row " & filterRow & _
                    " on " & wsTarget.name & " -- " & applyError
    ElseIf Not wsTarget.AutoFilterMode Then
        ' No error, but no filter either. Worth reporting rather than assuming it worked.
        failureState = True
        Debug.Print "RestoreAutoFilter: no AutoFilter present on " & wsTarget.name & _
                    " after applying to row " & filterRow & "."
    End If
End Sub
