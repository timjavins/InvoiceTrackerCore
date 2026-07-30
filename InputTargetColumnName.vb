Function InputTargetColumnName(sourceCol As String, targetHeaderDict As Object) As Long
    Dim targetCol As Long
    Dim colName As String
    Dim answer As VbMsgBoxResult

GetInput:
    ' Prompt the user to input the target column name
    colName = InputBox("Where should '" & sourceCol & "' data go?" & vbCrLf & "Enter the name of the target column:", "Input Target Column Name")

    ' Check if the user provided a valid column name
    If colName <> "" Then
        ' Look up the column index in the target header dictionary
        If targetHeaderDict.Exists(colName) Then
            targetCol = targetHeaderDict(colName)
        Else
            answer = MsgBox("The column '" & colName & "' does not exist in the " & ActiveSheet.Name & " sheet of " & ThisWorkbook.Name & "." & vbCrLf & _
                            "Please find the correct column name in the " & ActiveSheet.Name & " sheet and try again." & vbCrLf & _
                            "Would you like to re-enter the column name now?", vbYesNo + vbQuestion, "Re-enter Column Name")
            If answer = vbYes Then
                GoTo GetInput
            Else
                failureState = True ' Set failure state if the user chooses not to re-enter
                Exit Function
            End If
        End If
    Else
        GoTo GetInput
    End If
    InputTargetColumnName = targetCol
End Function