' Attaches a PDF to each of this supplier's draft Coupa requisitions created on a given
' date, then submits them for approval.
'
' Supplier keyword, Coupa host and the custom-view filter id come from TenantConfig, so one
' implementation serves every tenant. The WebDriver error dialog is shown through
' ShowWebDriverError, which each variant owns because a UserForm class cannot be named from
' config.
'
' Unified on the Securitas settings: 45s page-load timeout and MM/DD/YYYY target dates. The
' JCI values (20s, MM/DD/YY) were incidental drift, not tenant requirements.

' Coupa Requisition Attachment Automation
' This script navigates to a custom Coupa view filtered for draft requisitions created on a specific date,
' and attaches corresponding PDF files based on supplier part numbers found in requisition details.
'
' PDF files must be named with the supplier part number (e.g., 6200003393.pdf)
'
' Custom Coupa View Filter ID: -4814 (includes Created Date column)
'
' Note: Coupa's activity/custom view paginates results (90 items per page). This
' script reloads the custom view after processing each batch and repeats until
' no draft requisitions matching the `targetDate` remain, ensuring all matches
' are processed beyond the first page.

Sub AttachCoupaDocs()
    Dim driver As WebDriver
    Dim seleniumInstalled As Boolean
    Dim loggedIn As Boolean
    Dim userResponse As Integer
    Dim timeout As Double
    Dim draftReqs As Collection
    Dim skippedReqs As Collection
    Dim req As Variant
    Dim submitBtn As WebElement
    Dim reqId As String
    Dim partNumber As String
    Dim pdfPath As String
    Dim pdfSourceDir As String
    Dim targetDate As String
    Dim selectedFolder As String
    Dim successCount As Long
    Dim failCount As Long
    Dim notFoundCount As Long

    ' Prompt for configuration at runtime
    selectedFolder = PickFolder("Select the directory that contains the requisition PDF files.")
    If selectedFolder = "" Then
        MsgBox "No search directory selected. Exiting script.", vbExclamation
        Exit Sub
    End If

    pdfSourceDir = EnsureTrailingBackslash(selectedFolder)

    targetDate = PromptForTargetDate()
    If targetDate = "" Then
        MsgBox "No valid target date was provided. Exiting script.", vbExclamation
        Exit Sub
    End If

    ' Check if Selenium is installed
    On Error Resume Next
    seleniumInstalled = Not (CreateObject("Selenium.WebDriver") Is Nothing)
    On Error GoTo 0

    If Not seleniumInstalled Then
        MsgBox "Selenium is not installed. Please install Selenium Basic and the appropriate Chrome driver.", vbExclamation
        Exit Sub
    End If

    ' Initialize counters
    successCount = 0
    failCount = 0
    notFoundCount = 0
    Set skippedReqs = New Collection

    ' Initialize WebDriver
    Debug.Print "Initializing web driver"
    Set driver = New WebDriver
    Debug.Print "Starting web driver"

    On Error GoTo DriverError
    driver.Start "chrome", ""
    On Error GoTo 0

    Debug.Print "Maximizing window"
    driver.Window.Maximize

    ' Navigate to custom Coupa view (filter=-4814 includes Created Date column)
    Dim filteredUrl As String
    filteredUrl = "https://" & TenantCoupaHost() & "/user/account?filter=" & TenantCoupaFilterId()

    Debug.Print "Navigating to Coupa custom view"
    driver.Get filteredUrl

    ' Wait for login
    loggedIn = False
    timeout = Timer + 300 ' 5 minute timeout for login
    Do While Timer < timeout
        If InStr(driver.Url, "https://" & TenantCoupaHost()) = 1 And InStr(driver.Url, "okta.com") = 0 Then
            loggedIn = True
            Exit Do
        End If
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    If Not loggedIn Then
        GoTo UserLogin
    End If

    ' Wait for page to load
    Application.Wait Now + TimeValue("0:00:03")
    ' Loop: repeatedly fetch the custom view and process any draft requisitions
    Do
        Debug.Print "Extracting requisitions created on " & targetDate
        Set draftReqs = GetDraftRequisitionsByDate(driver, targetDate, TenantSupplierKeyword(), skippedReqs)

        If draftReqs.Count = 0 Then
            Debug.Print "No more processable draft requisitions found for " & targetDate
            Exit Do
        End If

        Debug.Print "Found " & draftReqs.Count & " draft requisitions to process"

        ' Process each draft requisition found in this batch
        For Each req In draftReqs
            reqId = req
            Debug.Print "Processing requisition: " & reqId

            ' Navigate to edit page
            driver.Get "https://" & TenantCoupaHost() & "/requisition_headers/" & reqId & "/edit"

            ' Wait for page to fully load (check for Submit button in footer)
            timeout = Timer + 45
            Do While Timer < timeout
                On Error Resume Next
                Set submitBtn = driver.FindElementById("submit_for_approval_link")
                If Not submitBtn Is Nothing Then Exit Do
                On Error GoTo 0
                Application.Wait Now + TimeValue("0:00:01")
            Loop

            If submitBtn Is Nothing Then
                Debug.Print "  [SKIP] Page did not load within timeout for " & reqId
                failCount = failCount + 1
                MarkReqSkipped skippedReqs, reqId
                GoTo NextReq
            End If

            ' Extract part number from page
            partNumber = ExtractPartNumber(driver)

            If partNumber = "" Then
                Debug.Print "  [SKIP] Could not extract part number from " & reqId
                failCount = failCount + 1
                MarkReqSkipped skippedReqs, reqId
                GoTo NextReq
            End If

            Debug.Print "  Part number: " & partNumber

            ' Check if PDF exists
            pdfPath = pdfSourceDir & partNumber & ".pdf"
            If Dir(pdfPath) = "" Then
                Debug.Print "  [NOT FOUND] " & partNumber & ".pdf"
                notFoundCount = notFoundCount + 1
                MarkReqSkipped skippedReqs, reqId
                GoTo NextReq
            End If

            ' Attach file and submit
            If AttachFileAndSubmit(driver, pdfPath) Then
                Debug.Print "  [SUCCESS] Attached and submitted " & reqId
                successCount = successCount + 1
            Else
                Debug.Print "  [FAILED] Could not attach/submit " & reqId
                failCount = failCount + 1
                MarkReqSkipped skippedReqs, reqId
            End If

            ' Brief pause between requisitions
            Application.Wait Now + TimeValue("0:00:03")

NextReq:
        Next req

        ' After processing the current batch, reload the custom view and get a fresh list
        driver.Get filteredUrl
        Application.Wait Now + TimeValue("0:00:03")
    Loop

    ' Show summary
    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Submitted: " & successCount & vbCrLf & _
           "Failed: " & failCount & vbCrLf & _
           "Not Found: " & notFoundCount, vbInformation

    driver.Quit
    Exit Sub

DriverError:
    Debug.Print "Driver error: " & Err.Description
    ' The form class lives in the variant, since VBA binds it at compile time.
    ShowWebDriverError Err.Description
    
    Exit Sub

UserLogin:
    userResponse = MsgBox("Please log into the Coupa website. Click 'Yes' once logged in.", vbYesNo + vbQuestion, "Sign In")
    If userResponse = vbYes Then
        loggedIn = False
        timeout = Timer + 60
        Do While Timer < timeout
            If InStr(driver.Url, "https://" & TenantCoupaHost()) = 1 And InStr(driver.Url, "okta.com") = 0 Then
                loggedIn = True
                Exit Do
            End If
            Application.Wait Now + TimeValue("0:00:01")
        Loop

        If loggedIn Then
            driver.Get "https://" & TenantCoupaHost() & "/user/account?filter=" & TenantCoupaFilterId()
            Application.Wait Now + TimeValue("0:00:03")
            Set draftReqs = GetDraftRequisitionsByDate(driver, targetDate, TenantSupplierKeyword(), skippedReqs)
            If draftReqs.Count > 0 Then
                Resume Next
            End If
        End If
    Else
        MsgBox "Login not confirmed. Exiting script.", vbExclamation
        driver.Quit
        Exit Sub
    End If
End Sub

' Prompt the user for a target date and validate MM/DD/YYYY format.
Function PromptForTargetDate() As String
    Dim response As String

    Do
        response = InputBox("On what day were the requisition drafts created?" & vbCrLf & vbCrLf & _
                            "Enter the date in MM/DD/YYYY format.", _
                            "Target Date")

        If response = "" Then
            PromptForTargetDate = ""
            Exit Function
        End If

        response = Trim(response)

        If IsValidTargetDate(response) Then
            PromptForTargetDate = Format(CDate(response), "mm/dd/yyyy")
            Exit Function
        End If

        MsgBox "Please enter a valid date in MM/DD/YYYY format.", vbExclamation
    Loop
End Function

' Validate that the response is a real date in MM/DD/YYYY format.
Function IsValidTargetDate(value As String) As Boolean
    Dim parts() As String
    Dim monthPart As Integer
    Dim dayPart As Integer
    Dim yearPart As Integer
    Dim normalizedDate As String

    parts = Split(value, "/")
    If UBound(parts) <> 2 Then
        IsValidTargetDate = False
        Exit Function
    End If

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        IsValidTargetDate = False
        Exit Function
    End If

    If Len(parts(2)) <> 4 Then
        IsValidTargetDate = False
        Exit Function
    End If

    monthPart = CInt(parts(0))
    dayPart = CInt(parts(1))
    yearPart = CInt(parts(2))

    On Error GoTo InvalidDate
    normalizedDate = Format(DateSerial(yearPart, monthPart, dayPart), "mm/dd/yyyy")
    IsValidTargetDate = (normalizedDate = value)
    Exit Function

InvalidDate:
    IsValidTargetDate = False
End Function

' Extract draft requisitions from custom view filtered by Created Date
Function GetDraftRequisitionsByDate(driver As WebDriver, targetDate As String, requiredKeyword As String, skippedReqs As Collection) As Collection
    Dim draftReqs As Collection
    Dim tableRows As WebElements
    Dim row As WebElement
    Dim cells As WebElements
    Dim reqLink As WebElement
    Dim createdDate As String
    Dim reqId As String

    Set draftReqs = New Collection

    On Error Resume Next

    ' Find all table rows
    Set tableRows = driver.FindElementsByTag("tr")

    For Each row In tableRows
        ' Get all cells in this row
        Set cells = row.FindElementsByTag("td")

        ' Check if row has enough cells (Created Date is 2nd column, index 1)
        If cells.Count >= 2 Then
            createdDate = cells(2).Text ' VBA WebElements collection is 1-based, Created Date is 2nd column

            ' Only process rows with matching date, keyword, and not already skipped
            If createdDate = targetDate And InStr(1, row.Text, requiredKeyword, vbTextCompare) > 0 Then
                ' Look for requisition link in this row
                Set reqLink = row.FindElementByXPath(".//a[contains(@href, '/requisition_headers/')]")

                If Not reqLink Is Nothing Then
                    reqId = reqLink.Text

                    ' Only add if reqId looks valid
                    If Len(reqId) > 0 And IsNumeric(reqId) And Not IsReqSkipped(skippedReqs, reqId) Then
                        ' Add to collection if not already present
                        On Error Resume Next
                        draftReqs.Add reqId, reqId ' Key = reqId to prevent duplicates
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next row

    On Error GoTo 0
    Set GetDraftRequisitionsByDate = draftReqs
End Function

' Extract part number from requisition edit page
' Tries multiple sources: title, description, supplier part number
Function ExtractPartNumber(driver As WebDriver) As String
    Dim response As String
    Dim partNumber As String
    Dim regex As Object
    Dim matches As Object

    ' Pattern 1: Requisition title "REQ for 6200003393"
    partNumber = ExtractPartFromTitle(driver)
    If partNumber <> "" Then
        ExtractPartNumber = partNumber
        Exit Function
    End If

    ' Pattern 2: Look for 10-digit numbers in page source
    response = driver.PageSource
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "\b(\d{9,10})\b"
    End With

    Set matches = regex.Execute(response)
    If matches.Count > 0 Then
        partNumber = matches(0).SubMatches(0)

        ' Validate it's a 10-digit number
        If (Len(partNumber) = 9 Or Len(partNumber) = 10) And IsNumeric(partNumber) Then
            ExtractPartNumber = partNumber
            Exit Function
        End If
    End If

    ExtractPartNumber = ""
End Function

' Extract part number from requisition title (h1 tag)
Function ExtractPartFromTitle(driver As WebDriver) As String
    Dim titleElem As WebElement
    Dim titleText As String
    Dim regex As Object
    Dim matches As Object

    On Error Resume Next
    Set titleElem = driver.FindElementByTag("h1")

    If Not titleElem Is Nothing Then
        titleText = titleElem.Text

        ' Pattern: "REQ for 6200003393"
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Global = False
            .pattern = "REQ for (\d{9,10})"
        End With

        Set matches = regex.Execute(titleText)
        If matches.Count > 0 Then
            ExtractPartFromTitle = matches(0).SubMatches(0)
            Exit Function
        End If
    End If

    On Error GoTo 0
    ExtractPartFromTitle = ""
End Function

' Attach PDF file and submit requisition
Function AttachFileAndSubmit(driver As WebDriver, pdfPath As String) As Boolean
    Dim attachLink As WebElement
    Dim fileInput As WebElement
    Dim checkbox As WebElement
    Dim addBtn As WebElement
    Dim submitBtn As WebElement
    Dim timeout As Double
    Dim success As Boolean

    On Error GoTo AttachError

    ' Click the "File" link in the requisition attachments section.
    ' The page has two attachment UIs: one for the requisition (div.attachments.attachified)
    ' and one for comments (div#comment_form). We must target only the requisition one.
    Set attachLink = driver.FindElementByCss(".attachments.attachified a.file-attachment")

    If attachLink Is Nothing Then
        Debug.Print "  [ERROR] Requisition file attachment link not found"
        AttachFileAndSubmit = False
        Exit Function
    End If

    driver.ExecuteScript "arguments[0].scrollIntoView({block: 'center'});", attachLink
    Application.Wait Now + TimeValue("0:00:01")
    driver.ExecuteScript "arguments[0].click();", attachLink
    Application.Wait Now + TimeValue("0:00:02")

    ' Check "Send to Supplier" checkbox
    On Error Resume Next
    Set checkbox = driver.FindElementById("requisition_header_attachments_attributes_attachment_intent")
    On Error GoTo AttachError
    If Not checkbox Is Nothing Then
        If Not checkbox.IsSelected Then
            checkbox.Click
            Application.Wait Now + TimeValue("0:00:01")
        End If
    End If

    ' SendKeys triggers FineUploader's upload. FineUploader immediately replaces the input
    ' DOM node after picking up the file, so the dispatchEvent call will hit a stale
    ' element reference — that's expected and safe to ignore.
    Set fileInput = driver.FindElementByCss("#requisition_header_attachments_attributes_attachment input[type='file']")
    fileInput.SendKeys pdfPath
    On Error Resume Next
    driver.ExecuteScript "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", fileInput
    On Error GoTo AttachError

    ' Wait for FineUploader to auto-close the attachment panel — happens when upload completes.
    ' Returns a number (1/0) rather than boolean to avoid Selenium Basic type coercion issues.
    timeout = Timer + 30
    success = False
    Do While Timer < timeout
        Dim panelState As Long
        panelState = CLng(driver.ExecuteScript("var p = document.querySelector('.attachments.attachified .fields'); return (!p || p.classList.contains('flash_hidden')) ? 1 : 0;"))
        If panelState = 1 Then
            success = True
            Exit Do
        End If
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    If Not success Then
        Debug.Print "  [ERROR] File upload did not complete within timeout"
        AttachFileAndSubmit = False
        Exit Function
    End If

    Set submitBtn = driver.FindElementById("submit_for_approval_link")
    driver.ExecuteScript "arguments[0].click();", submitBtn

    ' Wait for submission (URL should change from /edit)
    timeout = Timer + 15
    success = False
    Do While Timer < timeout
        If InStr(driver.Url, "/edit") = 0 Then
            success = True
            Exit Do
        End If
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    AttachFileAndSubmit = success
    Exit Function

AttachError:
    Debug.Print "Error in AttachFileAndSubmit: " & Err.Description
    AttachFileAndSubmit = False
End Function
