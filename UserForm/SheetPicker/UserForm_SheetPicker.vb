' These scripts are for the SheetPicker UserForm.
Option Explicit

Public SelectedSheetName As String

Private mWorkbook As Workbook
Private mExpectedSheetName As String

Public Sub InitializeSheets(ByVal wb As Workbook, Optional ByVal expectedSheetName As String = "INVOICE DETAILS")
    Dim ws As Worksheet

    Set mWorkbook = wb
    mExpectedSheetName = expectedSheetName

    Me.Caption = "Target Sheet Name"
    lblMessage.Caption = "Select the source worksheet to import from:"

    lstSheets.Clear
    For Each ws In mWorkbook.Worksheets
        lstSheets.AddItem ws.Name
    Next ws

    SelectExpectedSheet
End Sub

Private Sub UserForm_Initialize()
    SelectedSheetName = ""
    Me.Caption = "Target Sheet Name"
End Sub

Private Sub UserForm_Activate()
    If lstSheets.ListCount > 0 And lstSheets.ListIndex = -1 Then
        SelectExpectedSheet
    End If
End Sub

Private Sub btnOK_Click()
    If lstSheets.ListIndex = -1 Then
        MsgBox "Please select a sheet name from the list.", vbExclamation, "Select Sheet"
        Exit Sub
    End If

    SelectedSheetName = lstSheets.Value
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    SelectedSheetName = ""
    Me.Hide
End Sub

Private Sub lstSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOK_Click
End Sub

Private Sub SelectExpectedSheet()
    Dim i As Long
    Dim expectedName As String

    expectedName = VBA.LCase(VBA.Trim(mExpectedSheetName))
    If expectedName = "" Then Exit Sub

    For i = 0 To lstSheets.ListCount - 1
        If VBA.LCase(VBA.Trim(CStr(lstSheets.List(i)))) = expectedName Then
            lstSheets.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub
