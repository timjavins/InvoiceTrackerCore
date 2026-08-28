' The active-Coupa-user roster, and the question every requester email has to pass before it
' reaches a flat file.
'
' Coupa rejects an upload row whose "Requested By (Email)" is not a live user. A rejected row is
' invisible until someone reconciles the tracker against Coupa days later, by which time the
' invoice is closer to being late. This module turns that into a warning at export time.
'
' WHERE THE ROSTER COMES FROM
'
' A sheet named by TenantSheetName("coupa-users"), fed the same way 'BU List' is: a Power Query
' against the SharePoint list holding the shared active-users report. This module does not care how
' the sheet is filled -- Power Query, the legacy SharePoint list link, or a paste all read the same.
' See the "Coupa Users roster" section of README.md for the query.
'
' Columns are resolved by HEADER NAME rather than position, because a Power Query refresh reorders
' and renames columns freely, and because the report is a shared resource neither tenant controls.
' Several spellings are accepted for each of the two columns that matter -- see EmailHeaderNames
' and StatusHeaderNames.
'
' A roster with no recognizable status column is treated as already filtered to active users, which
' is what an "active users report" normally is. Only an explicitly inactive value excludes an
' address.
'
' UNAVAILABLE IS NOT INACTIVE
'
' When the sheet is missing, empty, or has no recognizable email column, verification is
' UNAVAILABLE and every address passes. A workbook that has not been wired up to the SharePoint
' list yet must still be able to export; it simply exports without the extra check, and
' CoupaUsersUnavailableReason says why.
'
' The alternative -- treating unavailable as inactive -- would fall every single store back to the
' prompted default requester and quietly undo the responsible-party lookup, which is the exact
' failure this module exists to prevent. Fail open, and say so.
'
' STALE IS NOT UNAVAILABLE
'
' A roster whose publishing pipeline stopped still reads perfectly -- it just describes who was
' active last week. Nothing about the sheet says so, so the pipeline stamps every row with the day
' it ran and CoupaUsersStalenessWarning reports a roster older than CoupaUsersStalenessDays. A stale
' roster is still used, because last week's list beats no list at all; the run just says how old it
' is. And an undated roster is UNKNOWN rather than stale, so a hand-pasted one does not cry wolf.

' Lazily built map of normalized email -> True. Nothing means "not available", which is why the
' loaded flag is tracked separately: an unavailable roster must not be rebuilt on every lookup.
Private activeCoupaUsers As Object
Private coupaUsersLoaded As Boolean
Private coupaUsersReason As String

' Captured from the roster's GeneratedOn column, so a pipeline that silently stopped is
' visible instead of serving last-good data forever.
Private coupaUsersGeneratedOn As Date
Private coupaUsersGeneratedOnKnown As Boolean

' Forces the roster to be re-read. Call at the start of a run so refreshing the SharePoint query
' mid-session takes effect without reopening the workbook.
Public Sub ResetCoupaUsersCache()
    Set activeCoupaUsers = Nothing
    coupaUsersLoaded = False
    coupaUsersReason = vbNullString
    coupaUsersGeneratedOn = 0
    coupaUsersGeneratedOnKnown = False
End Sub

' Whether the roster could be read at all. False means IsActiveCoupaUser passes everything.
Public Function CoupaUsersAvailable() As Boolean
    EnsureCoupaUsersLoaded
    CoupaUsersAvailable = Not (activeCoupaUsers Is Nothing)
End Function

' Why the roster is unavailable, or "" when it is available. Worth putting in a run's error
' summary: an export that silently skipped the check should say so once.
Public Function CoupaUsersUnavailableReason() As String
    EnsureCoupaUsersLoaded
    CoupaUsersUnavailableReason = coupaUsersReason
End Function

' How many active users the roster holds. 0 when unavailable.
Public Function CoupaUsersCount() As Long
    EnsureCoupaUsersLoaded
    If Not activeCoupaUsers Is Nothing Then CoupaUsersCount = activeCoupaUsers.Count
End Function

' A warning when the roster looks stale, or "" when it is fresh, undated, or unavailable.
'
' An absent or unparseable stamp is treated as UNKNOWN rather than stale, consistent with
' this module's fail-open posture: a roster hand-pasted without the column should not
' generate a warning on every run.
Public Function CoupaUsersStalenessWarning() As String
    EnsureCoupaUsersLoaded
    If activeCoupaUsers Is Nothing Then Exit Function
    If Not coupaUsersGeneratedOnKnown Then Exit Function

    Dim ageDays As Long
    ageDays = DateDiff("d", coupaUsersGeneratedOn, Date)
    If ageDays <= CoupaUsersStalenessDays() Then Exit Function

    CoupaUsersStalenessWarning = _
        "The Coupa users roster was generated " & Format$(coupaUsersGeneratedOn, "yyyy-mm-dd") & _
        ", " & ageDays & " days ago. Requester checks may be using stale data -- refresh " & _
        "the '" & TenantSheetName("coupa-users") & "' query, and check the roster pipeline ran."
End Function

' Whether an address belongs to an active Coupa user.
'
' A blank address is never a user, even when the roster is unavailable -- blank is a data gap the
' caller has to handle, not something to wave through.
Public Function IsActiveCoupaUser(ByVal email As Variant) As Boolean
    Dim key As String
    key = NormalizeEmail(email)
    If Len(key) = 0 Then Exit Function

    EnsureCoupaUsersLoaded
    If activeCoupaUsers Is Nothing Then
        IsActiveCoupaUser = True
        Exit Function
    End If

    IsActiveCoupaUser = activeCoupaUsers.Exists(key)
End Function

' An email reduced to its comparison key. Lowercased and trimmed; tolerates errors and nulls.
'
' Only ever used as a dictionary key. Addresses written to a flat file keep the casing they were
' entered with, because that is what the person recognizes when reading the export back.
Public Function NormalizeEmail(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    NormalizeEmail = LCase$(Trim$(CStr(value)))
End Function

Private Sub EnsureCoupaUsersLoaded()
    If coupaUsersLoaded Then Exit Sub

    ' Set before any early exit, so a failed load is not retried on every lookup.
    coupaUsersLoaded = True
    Set activeCoupaUsers = Nothing
    coupaUsersReason = vbNullString

    Dim sheetName As String
    On Error Resume Next
    sheetName = TenantSheetName("coupa-users")
    On Error GoTo 0
    If Len(sheetName) = 0 Then
        coupaUsersReason = "TenantConfig declares no 'coupa-users' sheet, so requester " & _
                           "addresses were not verified against Coupa."
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        coupaUsersReason = "Sheet '" & sheetName & "' was not found, so requester addresses " & _
                           "were not verified against Coupa."
        Exit Sub
    End If

    Dim headerRow As Long
    Dim emailCol As Long
    headerRow = FindRosterHeaderRow(ws, emailCol)
    If headerRow = 0 Then
        coupaUsersReason = "No email column was found on '" & sheetName & "' (looked for " & _
                           Join(EmailHeaderNames(), ", ") & "), so requester addresses were " & _
                           "not verified against Coupa."
        Exit Sub
    End If

    ' Absent is fine: a report that already lists only active users has nothing to filter on.
    Dim statusCol As Long
    statusCol = FindRosterColumn(ws, headerRow, StatusHeaderNames())

    ' Same value on every row, so the first data row is enough.
    Dim generatedCol As Long
    generatedCol = FindRosterColumn(ws, headerRow, GeneratedOnHeaderNames())
    If generatedCol > 0 Then
        coupaUsersGeneratedOnKnown = _
            TryParseIsoDate(ws.Cells(headerRow + 1, generatedCol).Value, coupaUsersGeneratedOn)
    End If

    Dim roster As Object
    Set roster = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, emailCol).End(xlUp).Row

    If lastRow > headerRow Then
        ' Read the two columns in one hit rather than touching cells individually. The real export
        ' runs to 42,000+ rows, so a per-cell loop would be ~85,000 worksheet reads on every run.
        ' Only the span between the two columns is pulled, not the whole sheet.
        Dim firstCol As Long
        Dim lastColRead As Long
        firstCol = emailCol
        lastColRead = emailCol
        If statusCol > 0 Then
            If statusCol < firstCol Then firstCol = statusCol
            If statusCol > lastColRead Then lastColRead = statusCol
        End If

        Dim block As Variant
        block = ws.Range(ws.Cells(headerRow + 1, firstCol), _
                         ws.Cells(lastRow, lastColRead)).Value2

        Dim emailIdx As Long
        Dim statusIdx As Long
        emailIdx = emailCol - firstCol + 1
        If statusCol > 0 Then statusIdx = statusCol - firstCol + 1

        Dim key As String

        If Not IsArray(block) Then
            ' A single cell does not come back as an array. Only reachable when the roster holds one
            ' user and has no status column.
            key = NormalizeEmail(block)
            If Len(key) > 0 Then roster.Add key, True
        Else
            Dim r As Long
            For r = 1 To UBound(block, 1)
                key = NormalizeEmail(block(r, emailIdx))
                If Len(key) > 0 Then
                    ' Inactive rows are never added, so a key present means at least one ACTIVE
                    ' record carries that address. That matters: 807 addresses in the real export
                    ' appear on both an active and an inactive record, and Coupa resolves such an
                    ' address to the live user. Any-active-wins is therefore the right rule, and
                    ' filtering before the add gets it without comparing rows.
                    If statusIdx = 0 Or Not IsInactiveStatus(block(r, statusIdx)) Then
                        If Not roster.Exists(key) Then roster.Add key, True
                    End If
                End If
            Next r
        End If
    End If

    ' An empty roster is a not-yet-refreshed query, not a company with no users. Treated as
    ' unavailable rather than as "nobody is active", which would fall every store back to the
    ' default requester.
    If roster.Count = 0 Then
        coupaUsersReason = "Sheet '" & sheetName & "' holds no active users and may not have " & _
                           "been refreshed, so requester addresses were not verified against Coupa."
        Exit Sub
    End If

    Set activeCoupaUsers = roster
End Sub

' The row carrying the roster's headers, and the column holding the email, or 0 when neither is
' found within the first few rows.
'
' Not necessarily row 1: Coupa's scheduled reports carry title, account and run-time rows above the
' real header, the same way vendor bill files do (see FindHeaderRow).
Private Function FindRosterHeaderRow(ByVal ws As Worksheet, _
                                     ByRef emailCol As Long, _
                                     Optional ByVal maxRowsToCheck As Long = 20) As Long
    emailCol = 0

    Dim r As Long
    Dim c As Long
    For r = 1 To maxRowsToCheck
        c = FindRosterColumn(ws, r, EmailHeaderNames())
        If c > 0 Then
            emailCol = c
            FindRosterHeaderRow = r
            Exit Function
        End If
    Next r
End Function

' The column matching the first candidate header present on a row, or 0 when none is.
'
' Candidate order is preference order, so a report carrying both "Email" and "Login" resolves to
' "Email" -- GetHeaderColumnIndexes returns a dictionary, which has no useful order of its own.
Private Function FindRosterColumn(ByVal ws As Worksheet, _
                                  ByVal headerRow As Long, _
                                  ByVal candidates As Variant) As Long
    Dim found As Object
    Set found = GetHeaderColumnIndexes(ws, headerRow, candidates)
    If found.Count = 0 Then Exit Function

    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If found.Exists(candidates(i)) Then
            FindRosterColumn = found(candidates(i))
            Exit Function
        End If
    Next i
End Function

' Header names the roster's email column may go by, most preferred first.
'
' "Login" is last because a Coupa login is usually the email address but is not required to be; it
' is a fallback for a report that omits the email column outright.
Private Function EmailHeaderNames() As Variant
    EmailHeaderNames = Array("Email", "Email Address", "E-mail", "Login")
End Function

' Header names the roster's active-or-not column may go by, most preferred first.
Private Function StatusHeaderNames() As Variant
    StatusHeaderNames = Array("Active", "Active?", "Status", "Account Status", "User Status")
End Function

' Header names the roster's generated-on stamp may go by.
Private Function GeneratedOnHeaderNames() As Variant
    GeneratedOnHeaderNames = Array("GeneratedOn", "Generated On", "Generated")
End Function

' Whether a status value says this user cannot be a requester.
'
' Verified against a real "Coupa Users List" export (2026-08-25, 42,399 rows). Its "Active" column
' holds lowercase words, not booleans:
'
'     active     27,669
'     inactive   14,727
'     deleted         3
'
' So "deleted" is here despite Coupa's user state being described as just active/inactive -- the
' report emits a third value, and a deleted user plainly cannot own a requisition. "false"/"no"/"0"
' are kept for a differently-built report that renders the column as a boolean; a real Boolean from
' a live query is handled above as one.
'
' Phrased as "cannot be a requester" rather than "is active" deliberately. A value that is none of
' the above most likely means FindRosterColumn picked the wrong column, and what it picked could be
' anything: a store, a full name, a GL string. Keeping the user is the safe reading of that.
' Reporting it as inactive would empty the roster and fall every store back to the default
' requester, which is the failure this module exists to prevent. Same fail-open reasoning as an
' unavailable roster.
'
' A blank status is kept for the same reason: the export never leaves this column empty, so a blank
' is evidence about the column, not about the user.
Private Function IsInactiveStatus(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function

    ' A boolean column comes through as a real Boolean, not as text.
    If VarType(value) = vbBoolean Then
        IsInactiveStatus = Not CBool(value)
        Exit Function
    End If

    Dim statusText As String
    statusText = LCase$(Trim$(CStr(value)))
    If Len(statusText) = 0 Then Exit Function

    Select Case statusText
        Case "inactive", "deleted", "false", "no", "0"
            IsInactiveStatus = True
    End Select
End Function

' Parses "YYYY-MM-DD" without going through CDate.
'
' CDate honours the machine's locale, so the same text can parse differently on two
' machines -- and silently, since 2026-08-26 is a valid date read either way round. Parsing
' the parts explicitly makes the reading of an ISO stamp locale-independent.
Private Function TryParseIsoDate(ByVal value As Variant, ByRef parsed As Date) As Boolean
    If IsError(value) Or IsNull(value) Then Exit Function

    ' A column typed as a date by Power Query arrives as a real Date already.
    If VarType(value) = vbDate Then
        parsed = CDate(value)
        TryParseIsoDate = True
        Exit Function
    End If

    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) < 10 Then Exit Function

    Dim parts As Variant
    parts = Split(Left$(text, 10), "-")
    If UBound(parts) <> 2 Then Exit Function

    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then
        Exit Function
    End If

    On Error Resume Next
    parsed = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
    TryParseIsoDate = (Err.Number = 0)
    On Error GoTo 0
End Function

' Days after which the roster counts as stale. A core rule, not a tenant fact: it describes
' how often the shared pipeline is expected to publish.
Private Function CoupaUsersStalenessDays() As Long
    CoupaUsersStalenessDays = 3
End Function
