' Finds the approver currently blocking each pending requisition, and records their name
' and email on the tracker.
'
' Coupa host and the tracker columns come from TenantConfig. The WebDriver error dialog is
' shown through ShowWebDriverError, which each variant owns because a UserForm class cannot
' be named from config.
'
' Ported from the Securitas version. JCI's copy had no error dialog at all, so that tracker
' gains one.

' This script is designed for the "2025 SECURITAS Nordstrom Repair_Installation Invoices.xlsm" file in the AP PMO Securitas folder.
' Its purpose is to find any items that are marked as "Pending Approval" in the "Invoices" sheet and then look up the corresponding
' email addresses of those blockers in Coupa. The email addresses are then written to the "Relevant email" column with the first
' names written to the "Relevant name" column.

Sub GetBlockerEmails()
    Dim ws As Worksheet
    Dim wsCoupaInvs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim url As String
    Dim reqNum As String
    Dim userResponse As Integer
    Dim loggedIn As Boolean
    Dim response As String
    Dim emailData As Variant
    Dim timeout As Double
    Dim firstName As String
    Dim loadCounter As Long
    Dim statusArr As Variant
    Dim reqArr As Variant
    Dim emailArr() As Variant
    Dim nameArr() As Variant
    Dim driver As WebDriver
    Dim seleniumInstalled As Boolean

    ' Check if Selenium is installed
    On Error Resume Next
    seleniumInstalled = Not (CreateObject("Selenium.WebDriver") Is Nothing)
    On Error GoTo 0

    If Not seleniumInstalled Then
        MsgBox "Selenium is not installed. Please install Selenium Basic and the appropriate Chrome driver to get the approver email addresses.", vbExclamation
        Exit Sub
    End If

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets(TenantSheetName("tracker"))
    Set wsCoupaInvs = ThisWorkbook.Sheets(TenantSheetName("coupa-invs"))

    ' Find the last row with data in column O
    lastRow = ws.Cells(ws.Rows.Count, TenantColLetter("submitted-invoice-number")).End(xlUp).Row
   
    ' Clear existing values in columns AF and AG before starting
    ws.Range(TenantColLetter("blocker-email") & "2:" & TenantColLetter("blocker-email") & lastRow).ClearContents
    ws.Range(TenantColLetter("blocker-name") & "2:" & TenantColLetter("blocker-name") & lastRow).ClearContents

    Debug.Print "Beginning to parse data in " + ws.Name
    ' Read all relevant columns into arrays at once
    statusArr = ws.Range(TenantColLetter("requisition-status") & "2:" & TenantColLetter("requisition-status") & lastRow).Value ' Status column
    reqArr = ws.Range(TenantColLetter("requisition-number") & "2:" & TenantColLetter("requisition-number") & lastRow).Value    ' REQ # column

    ' Count the work before launching a browser -- starting Chrome and waiting for an Okta
    ' login only to find nothing to look up is wasted effort.
    Dim pendingCount As Long
    For i = 1 To UBound(statusArr, 1)
        If statusArr(i, 1) = "Pending Approval" Then
            If Len(Trim$(CStr(reqArr(i, 1) & ""))) > 0 Then pendingCount = pendingCount + 1
        End If
    Next i

    If pendingCount = 0 Then
        ' Nothing paused or unprotected yet, so there is nothing to undo here.
        MsgBox "No requisitions are pending approval, so there are no approvers to look up.", _
               vbInformation
        Exit Sub
    End If

    ' Initialize WebDriver
    Debug.Print "Initializing web driver"
    Set driver = New WebDriver
    Debug.Print "Starting web driver"

    On Error GoTo DriverError
    driver.Start "chrome", ""
    On Error GoTo 0

    Debug.Print "Maximizing window"
    driver.Window.Maximize

    ' Navigate to the login page
    Debug.Print "Navigating to Coupa."
    driver.Get "https://" & TenantCoupaHost() & "/"

    ' Initialize fresh empty arrays for results
    ReDim emailArr(1 To UBound(statusArr, 1), 1 To 1)
    ReDim nameArr(1 To UBound(statusArr, 1), 1 To 1)

    For i = 1 To UBound(statusArr, 1)
        ' Reset login state for each iteration
        loggedIn = False
        
        ' Initialize default empty values
        emailArr(i, 1) = ""
        nameArr(i, 1) = ""
        
        ' Check if the requisition is pending approval
        If statusArr(i, 1) = "Pending Approval" Then
CheckURL:
            timeout = Timer + 10
            Do While Timer < timeout
                If IsCoupaSignedIn(driver) Then
                    loggedIn = True
                    Exit Do
                End If
                Application.Wait Now + TimeValue("0:00:01")
            Loop
            
            If Not loggedIn Then
                GoTo UserLogin
            End If

            reqNum = reqArr(i, 1)
            url = "https://" & TenantCoupaHost() & "/requisition_headers/" & reqNum
            driver.Get url

            loadCounter = 0
CheckStatus:
            loadCounter = loadCounter + 1
            If loadCounter > 10 Then
                emailData = Array("error in page load", "")
                firstName = ""
                GoTo SetArrayValue
            End If

            response = driver.PageSource
            If InStr(response, "<h3 class=""requisitionTitle__status s-requisitionTitleStatus"">") = 0 Then
                If Not IsCoupaSignedIn(driver) Then
                    GoTo CheckURL
                End If
                Application.Wait Now + TimeValue("0:00:01")
                GoTo CheckStatus
            End If
            
            If InStr(response, "<h3 class=""requisitionTitle__status s-requisitionTitleStatus"">(Pending Approval)</h3>") = 0 Then
                emailArr(i, 1) = ""
                nameArr(i, 1) = "update 'Coupa Reqs' sheet"
                GoTo NextIteration
            End If

            timeout = Timer + 10
            Do While Timer < timeout
                response = driver.PageSource
                emailData = ExtractContact(response)
                If emailData(0) <> "error finding email" Then
                    Exit Do
                End If
                Application.Wait Now + TimeValue("0:00:01")
            Loop

            ' Initialize firstName
            firstName = ""
            If IsArray(emailData) And UBound(emailData) >= 0 Then
                firstName = HandleAllCaps(CStr(emailData(0)))
            End If
            
SetArrayValue:
            If IsArray(emailData) And UBound(emailData) >= 1 Then
                emailArr(i, 1) = emailData(1)
            Else
                emailArr(i, 1) = ""
            End If
            nameArr(i, 1) = firstName
        End If
NextIteration:
    Next i

    ' Write all results back at once (single write operation per column)
    ws.Range(TenantColLetter("blocker-email") & "2:" & TenantColLetter("blocker-email") & lastRow).Value = emailArr
    ws.Range(TenantColLetter("blocker-name") & "2:" & TenantColLetter("blocker-name") & lastRow).Value = nameArr

    driver.Quit
    Exit Sub

DriverError:
    ' The form class lives in the variant, since VBA binds it at compile time.
    ShowWebDriverError Err.Description
    
    Exit Sub

UserLogin:
    userResponse = MsgBox("Please log into the Coupa website. Click 'Yes' once logged in.", vbYesNo + vbQuestion, "Sign In")
    If userResponse = vbYes Then
        GoTo CheckURL
    Else
        MsgBox "Login not confirmed. Exiting script.", vbExclamation
        driver.Quit
        Exit Sub
    End If
End Sub

' Whether the browser is actually on a signed-in Coupa page.
'
' Checking only that the URL starts with the Coupa host is not enough: Coupa's own login and
' SSO hand-off pages live on that host too, so the old test reported "logged in" while the
' user was still staring at a sign-in screen. Require the Coupa host AND the absence of the
' identity provider and any login path.
Private Function IsCoupaSignedIn(ByVal driver As WebDriver) As Boolean
    Dim url As String

    On Error Resume Next
    url = LCase$(driver.Url)
    On Error GoTo 0

    If Len(url) = 0 Then Exit Function
    If InStr(url, LCase$("https://" & TenantCoupaHost())) <> 1 Then Exit Function

    ' Still authenticating.
    If InStr(url, "okta.com") > 0 Then Exit Function
    If InStr(url, "/sessions/new") > 0 Then Exit Function
    If InStr(url, "/login") > 0 Then Exit Function
    If InStr(url, "/user/session") > 0 Then Exit Function
    If InStr(url, "saml") > 0 Then Exit Function

    IsCoupaSignedIn = True
End Function
