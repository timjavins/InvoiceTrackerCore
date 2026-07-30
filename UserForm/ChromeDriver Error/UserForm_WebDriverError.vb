' These scipts are for the WebDriverError UserForm.
Option Explicit

Private Sub UserForm_Initialize()
    ' Set default values - these will be overridden by calling code
    Me.Caption = "WebDriver Error"
End Sub

Private Sub btnOpenDir_Click()
    ' Open the directory in Windows Explorer
    Dim dirPath As String
    dirPath = txtDirectory.Text
    
    ' Check if directory exists
    If Len(Dir(dirPath, vbDirectory)) > 0 Then
        Shell "explorer.exe """ & dirPath & """", vbNormalFocus
    Else
        MsgBox "Directory does not exist:" & vbCrLf & dirPath, vbExclamation, "Directory Not Found"
    End If
End Sub

Private Sub btnCopyURL_Click()
    ' Copy URL to clipboard using DataObject (more reliable)
    On Error Resume Next
    Dim dataObj As Object
    ' Create a DataObject via its 128-bit MSForms.DataObject CLSID (AKA GUID) number in the Windows registry
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText txtURL.Text
    dataObj.PutInClipboard
    Set dataObj = Nothing
    On Error GoTo 0
    
    ' Visual feedback
    Dim originalCaption As String
    originalCaption = btnCopyURL.Caption
    btnCopyURL.Caption = "Copied!"
    Application.Wait Now + TimeValue("0:00:01")
    btnCopyURL.Caption = originalCaption
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' Allow double-click on textboxes to select all
Private Sub txtDirectory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    txtDirectory.SelStart = 0
    txtDirectory.SelLength = Len(txtDirectory.Text)
End Sub

Private Sub txtURL_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    txtURL.SelStart = 0
    txtURL.SelLength = Len(txtURL.Text)
End Sub