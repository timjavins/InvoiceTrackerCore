' Decides which sheet of an imported workbook holds the bill data.
'
' Tries the tenant's expected sheet name first. If that sheet is absent, asks the user --
' through the variant's sheet picker if it has one, otherwise through a plain prompt listing
' the available sheets.
'
' The picker is a UserForm, which cannot live in core: form code binds to controls on its own
' form object, and the assembler excludes UserForm paths for that reason. A variant that has
' imported the picker provides PickSheetName; one that has not gets the prompt fallback, so
' this works either way.
'
' Sets failureState when the user declines to choose.
Public Function InputSheetName(ByVal wb As Workbook) As String
    If wb Is Nothing Then
        failureState = True
        Exit Function
    End If

    Dim expected As String
    expected = TenantImportSheetName()

    Dim candidate As Worksheet
    Set candidate = FindWorksheetByName(wb, expected)

    If Not candidate Is Nothing Then
        InputSheetName = candidate.Name
        Exit Function
    End If

    ' Expected sheet absent: ask. Prefer the variant's picker.
    Dim chosen As String
    chosen = AskForSheetName(wb, expected)

    If Len(Trim$(chosen)) = 0 Then
        MsgBox "As no sheet was chosen, the operation is cancelled.", vbExclamation
        failureState = True
        Exit Function
    End If

    Set candidate = FindWorksheetByName(wb, chosen)
    If candidate Is Nothing Then
        MsgBox "There is no sheet named '" & chosen & "' in that workbook.", vbExclamation
        failureState = True
        Exit Function
    End If

    InputSheetName = candidate.Name
End Function

' Asks the variant's picker, falling back to a prompt when it declines.
'
' Every variant must define PickSheetName -- VBA binds calls at compile time, so a missing
' one is a compile error, not something an error handler could rescue. A variant with no
' picker form defines it as a one-liner returning "", which lands us on the prompt.
Private Function AskForSheetName(ByVal wb As Workbook, ByVal expected As String) As String
    Dim chosen As String

    On Error Resume Next
    chosen = PickSheetName(wb, expected)
    On Error GoTo 0

    If Len(Trim$(chosen)) > 0 Then
        AskForSheetName = chosen
    Else
        AskForSheetName = PromptForSheetName(wb, expected)
    End If
End Function

' Plain-prompt fallback: lists the sheets and asks for one by name.
Private Function PromptForSheetName(ByVal wb As Workbook, ByVal expected As String) As String
    Dim names As String
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        names = names & vbCrLf & "    " & ws.Name
    Next ws

    PromptForSheetName = InputBox( _
        "No sheet named '" & expected & "' was found in that workbook." & vbCrLf & vbCrLf & _
        "Available sheets:" & names & vbCrLf & vbCrLf & _
        "Type the name of the sheet holding the bill data:", _
        "Choose a Sheet")
End Function

' Case- and whitespace-insensitive sheet lookup.
Private Function FindWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If Len(Trim$(sheetName)) = 0 Then Exit Function

    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If VBA.LCase$(VBA.Trim$(ws.Name)) = VBA.LCase$(VBA.Trim$(sheetName)) Then
            Set FindWorksheetByName = ws
            Exit Function
        End If
    Next ws
End Function
