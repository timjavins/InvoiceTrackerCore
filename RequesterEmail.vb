' Who a requisition is raised for.
'
' The two trackers used to disagree about this. Securitas looked the store up in 'Site RPs' and took
' its "Email" column; JCI prompted once and put that single address on every requisition it
' generated, so every JCI requisition was raised in the operator's name regardless of which site the
' work happened at. This module is the one answer for both -- see ADR-0003, and the note on scope at
' the bottom of FlatFileHelpers.vb for what stays per-variant.
'
' THE CASCADE
'
' 'Site RPs' names four people per store, in escalation order:
'
'     Email          the site's responsible party
'     Tier 1 email   their escalation contact
'     Tier 2 email
'     Tier 3 email
'
' The first of the four that is both present and an active Coupa user wins. Skipping past an
' inactive address is the point of the check: Coupa rejects an upload row whose requester is not a
' live user. Escalating up the tiers rather than straight to the operator keeps the requisition with
' someone who actually owns the site -- a departed responsible party hands off to their manager, not
' to whoever happened to run the export.
'
' Only when all four fail does the caller fall back to the prompted default requester.
'
' Columns are resolved by header name, so a reordered or re-exported 'Site RPs' still reads. The
' store and email headers are required; the tier headers are optional, and a sheet without them
' simply has a shorter cascade.

' Resolves the requester for a store from 'Site RPs'.
'
' Returns False only when the store is ABSENT from the sheet, or the sheet cannot be read at all.
' That is a data gap the caller should report and skip rather than paper over: raising a requisition
' for a store nobody has claimed puts it on the wrong person's approval queue, and the fix is one
' row in a shared list.
'
' Returns True with requesterEmail = "" when the store IS listed but none of its four contacts is an
' active Coupa user. That is the case the prompted default requester exists for.
'
' resolvedTier names which of the four answered, and reason carries text for the run's error
' summary. reason is set on failure, AND on success when the answer did not come from the site's own
' responsible party -- an escalation is worth seeing even though it exported fine.
Public Function TryResolveSiteRequester(ByVal wsSiteRPs As Worksheet, _
                                        ByVal storeKey As String, _
                                        ByVal altStoreKey As String, _
                                        ByRef requesterEmail As String, _
                                        ByRef resolvedTier As String, _
                                        ByRef reason As String) As Boolean
    requesterEmail = vbNullString
    resolvedTier = vbNullString
    reason = vbNullString

    If wsSiteRPs Is Nothing Then
        reason = "No site responsible-party sheet was supplied."
        Exit Function
    End If

    Dim tierHeaders As Variant
    Dim tierLabels As Variant
    tierHeaders = TierEmailHeaders()
    tierLabels = TierLabelNames()

    ' The store and responsible-party headers are the contract. Failing loudly here beats guessing
    ' at column positions: a wrong guess raises requisitions for the wrong people, which is far
    ' worse than an export that stops and says which headers it wanted.
    Dim headerRow As Long
    headerRow = FindHeaderRow(wsSiteRPs, Array(StoreHeaderName(), tierHeaders(0)))
    If headerRow = 0 Then
        reason = "Sheet '" & wsSiteRPs.Name & "' has no row carrying both a '" & _
                 StoreHeaderName() & "' and an '" & tierHeaders(0) & "' header."
        Exit Function
    End If

    Dim wanted As Variant
    wanted = Array(StoreHeaderName(), tierHeaders(0), tierHeaders(1), tierHeaders(2), tierHeaders(3))

    Dim cols As Object
    Set cols = GetHeaderColumnIndexes(wsSiteRPs, headerRow, wanted)

    Dim storeRow As Long
    storeRow = FindStoreRow(wsSiteRPs, cols(StoreHeaderName()), headerRow, storeKey, altStoreKey)
    If storeRow = 0 Then
        reason = "Store " & storeKey & " is not listed on '" & wsSiteRPs.Name & "'."
        Exit Function
    End If

    ' The store exists, so from here on the caller has something to work with either way.
    TryResolveSiteRequester = True

    Dim skipped As String
    Dim i As Long
    Dim candidate As String

    For i = LBound(tierHeaders) To UBound(tierHeaders)
        If cols.Exists(tierHeaders(i)) Then
            candidate = Trim$(CStr(wsSiteRPs.Cells(storeRow, cols(tierHeaders(i))).Value))
            If Len(candidate) > 0 Then
                If IsActiveCoupaUser(candidate) Then
                    requesterEmail = candidate
                    resolvedTier = CStr(tierLabels(i))
                    If Len(skipped) > 0 Then
                        ' Phrased without a verb agreeing on number, because the skipped list
                        ' can hold one address or several and this text goes straight into the
                        ' run's error summary.
                        reason = "Store " & storeKey & ": escalated to " & resolvedTier & " (" & _
                                 candidate & ") -- not active in Coupa: " & skipped & "."
                    End If
                    Exit Function
                End If
                skipped = AppendToList(skipped, candidate)
            End If
        End If
    Next i

    If Len(skipped) = 0 Then
        reason = "Store " & storeKey & " has no requester email on '" & wsSiteRPs.Name & "'."
    Else
        reason = "Store " & storeKey & " has no active Coupa user among its contacts (" & _
                 skipped & ")."
    End If
End Function

' The fallback requester, asked for once per run and remembered in the caller's state.
'
' Moved here from Securitas's MakeFlatFile so both trackers ask the same question and apply the same
' checks. Two things changed on the way: the domain comes from TenantApproverEmailDomain instead of a
' literal "@nordstrom.com" and its hardcoded length, and the address is now checked against the
' active-Coupa-user roster. The old prompt could only advise "be sure to use an email that is
' registered in Coupa"; it had no way to check, so a typo or a departed colleague was discovered by
' Coupa rejecting every row in the file.
'
' State stays with the caller rather than in this module so it resets naturally per run. Returns ""
' when the user cancels, which callers treat as abandon-the-export.
Public Function GetOrPromptDefaultEmail(ByRef defaultEmail As String, _
                                        ByRef defaultEmailNeeded As Boolean) As String
    If Not defaultEmailNeeded Then
        GetOrPromptDefaultEmail = defaultEmail
        Exit Function
    End If

    Dim domain As String
    domain = "@" & LCase$(Trim$(TenantApproverEmailDomain()))

    ' The generators run with screen updating suspended, so a modal prompt would not paint.
    Dim wasPaused As Boolean
    wasPaused = SuspendThinkingForDialog()

    ' Carries the last attempt back into the box, so a rejected address can be corrected rather
    ' than retyped.
    Dim prefill As String
    prefill = defaultEmail

    Dim entered As String
    Do
        entered = InputBox( _
            "Please enter the default requester email address. This is used for stores with no " & _
            "active requester of their own, and should probably be your own address.", _
            "Default Email Address", prefill)

        ' StrPtr distinguishes Cancel from OK on an empty box; both return "".
        If StrPtr(entered) = 0 Then
            ResumeThinkingAfterDialog wasPaused
            Exit Function
        End If

        entered = Replace(entered, " ", "")
        prefill = entered

        If Len(entered) = 0 Then
            MsgBox "A default email address is required to proceed.", vbExclamation, "Error"

        ElseIf Len(entered) <= Len(domain) Or LCase$(Right$(entered, Len(domain))) <> domain Then
            MsgBox "Please enter a " & Mid$(domain, 2) & " email address.", _
                   vbExclamation, "Try Again"

        ElseIf Not IsActiveCoupaUser(entered) Then
            MsgBox "'" & entered & "' is not an active Coupa user, so every row using it would " & _
                   "be rejected on upload." & vbCrLf & vbCrLf & _
                   "Check that the Coupa users sheet is up to date, then try again.", _
                   vbExclamation, "Not an Active Coupa User"

        ElseIf MsgBox("Please confirm: use '" & entered & "' as the default requester email " & _
                      "address?", vbYesNo, "Confirm Default Email Address") = vbYes Then
            defaultEmail = entered
            defaultEmailNeeded = False
            GetOrPromptDefaultEmail = entered
            ResumeThinkingAfterDialog wasPaused
            Exit Function
        End If
    Loop
End Function

' The row a store sits on, or 0 when it is not listed.
'
' altStoreKey is tried second because some 'Site RPs' rows are keyed by the store's GL code rather
' than its store number. Pass "" when the caller has no second key.
Private Function FindStoreRow(ByVal ws As Worksheet, _
                              ByVal storeCol As Long, _
                              ByVal headerRow As Long, _
                              ByVal storeKey As String, _
                              ByVal altStoreKey As String) As Long
    If storeCol < 1 Then Exit Function

    Dim hit As Range
    Set hit = ws.Columns(storeCol).Find(What:=storeKey, LookIn:=xlValues, LookAt:=xlWhole)

    If hit Is Nothing And Len(altStoreKey) > 0 Then
        Set hit = ws.Columns(storeCol).Find(What:=altStoreKey, LookIn:=xlValues, LookAt:=xlWhole)
    End If

    If hit Is Nothing Then Exit Function
    If hit.Row <= headerRow Then Exit Function

    FindStoreRow = hit.Row
End Function

' Header naming the store column on 'Site RPs'.
Private Function StoreHeaderName() As String
    StoreHeaderName = "Store"
End Function

' The email columns to try, in escalation order. Index 0 is the site's own responsible party.
Private Function TierEmailHeaders() As Variant
    TierEmailHeaders = Array("Email", "Tier 1 email", "Tier 2 email", "Tier 3 email")
End Function

' How each tier is described in a warning, parallel to TierEmailHeaders.
'
' NAMED "...Names" DELIBERATELY -- do not shorten it to TierLabels. VBA identifiers are
' case-insensitive, so a procedure called TierLabels is the SAME identifier as the caller's
' local `tierLabels`. The Dim shadows the function for the whole procedure, and
' `tierLabels = TierLabels()` then compiles as a subscript-less array access on the
' uninitialised local, failing at run time with "Subscript out of range" -- not at compile
' time, because the expression is syntactically valid. That bug reached a live workbook.
'
' The "...Names" suffix matches EmailHeaderNames / StatusHeaderNames in CoupaUsers.vb and
' keeps these constant-providers clear of the local names that read them.
Private Function TierLabelNames() As Variant
    TierLabelNames = Array("the responsible party", "tier 1", "tier 2", "tier 3")
End Function

' Appends an item to a comma-separated list, for warning text, skipping one already present.
'
' The de-duplication is not cosmetic tidiness: 'Site RPs' frequently names the SAME person as
' both the site's responsible party and its tier 1 contact, so when that person is inactive
' the cascade skips their address twice and the warning read "not active in Coupa:
' michael.kelly@nordstrom.com, michael.kelly@nordstrom.com". Observed on stores 0515 and 0788.
Private Function AppendToList(ByVal existing As String, ByVal item As String) As String
    AppendToList = existing

    If Len(item) = 0 Then Exit Function

    If Len(existing) = 0 Then
        AppendToList = item
        Exit Function
    End If

    ' Compare whole entries, so a shorter address cannot match inside a longer one.
    Dim entry As Variant
    For Each entry In Split(existing, ", ")
        If StrComp(CStr(entry), item, vbTextCompare) = 0 Then Exit Function
    Next entry

    AppendToList = existing & ", " & item
End Function
