' Pulls the pending approver's first name and email out of a Coupa requisition page's HTML.
'
' Returns Array(firstName, emailAddress). Team approvers (Hardware, Tech) have no
' individual mailbox, so the full team name is returned as the name with an empty email.
' Returns Array("error finding email", "") when the tooltip is not present.
'
' The email domain comes from TenantConfig rather than being hardcoded.
Public Function ExtractContact(ByVal responseText As String) As Variant
    Dim cleanedResponse As String

    ' Collapse line breaks so the tooltip markup can be matched as one run of text.
    cleanedResponse = Replace(responseText, vbCrLf, " ")
    cleanedResponse = Replace(cleanedResponse, vbLf, " ")

    Dim domain As String
    domain = TenantApproverEmailDomain()

    Dim pattern As String
    pattern = "<div class=""ApprovalTooltip__container"">.*?Pending Approval.*?" & _
              "<div class=""ApprovalTooltip__details"">.*?" & _
              "<div class=""ApprovalTooltip__title s-tt-title"">(\w+)\s[^<]*</div>.*?" & _
              "<div class=""ApprovalTooltip__mention s-tt-mention"">\(\s*(@[\w\.-]+" & _
              Replace(domain, ".", "\.") & ")\s*\)</div>.*?</div>.*?</div>"

    Dim matches As Object
    Set matches = ExecuteHtmlPattern(cleanedResponse, pattern)

    If matches.Count = 0 Then
        ExtractContact = Array("error finding email", "")
        Exit Function
    End If

    Dim firstName As String
    Dim emailAddress As String
    firstName = matches(0).SubMatches(0)
    emailAddress = matches(0).SubMatches(1)

    If IsTeamApprover(firstName) Then
        ' A team has no individual mailbox: return the whole team name, no email.
        Dim teamPattern As String
        teamPattern = "<div class=""ApprovalTooltip__title s-tt-title"">(" & firstName & _
                      "[^<]+)</div>"

        Dim teamMatches As Object
        Set teamMatches = ExecuteHtmlPattern(cleanedResponse, teamPattern)

        If teamMatches.Count > 0 Then
            firstName = teamMatches(0).SubMatches(0)
            emailAddress = vbNullString
        End If
    Else
        ' The tooltip renders the address as "@local.partdomain"; move the @ into place.
        emailAddress = NormalizeMentionToEmail(emailAddress, domain)
    End If

    ExtractContact = Array(firstName, emailAddress)
End Function

' Approver names that denote a team rather than a person.
Private Function IsTeamApprover(ByVal firstName As String) As Boolean
    Select Case firstName
        Case "Hardware", "Tech"
            IsTeamApprover = True
    End Select
End Function

' Turns a tooltip mention like "@jane.smithnordstrom.com" into "jane.smith@nordstrom.com".
Private Function NormalizeMentionToEmail(ByVal mention As String, ByVal domain As String) As String
    Dim localPart As String
    localPart = mention

    ' Drop the leading @ the mention markup carries.
    If Left$(localPart, 1) = "@" Then localPart = Mid$(localPart, 2)

    Dim domainPos As Long
    domainPos = InStr(1, localPart, domain, vbTextCompare)

    If domainPos > 1 Then
        NormalizeMentionToEmail = Left$(localPart, domainPos - 1) & "@" & Mid$(localPart, domainPos)
    Else
        NormalizeMentionToEmail = localPart
    End If
End Function

Private Function ExecuteHtmlPattern(ByVal text As String, ByVal pattern As String) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .MultiLine = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    Set ExecuteHtmlPattern = regex.Execute(text)
End Function
