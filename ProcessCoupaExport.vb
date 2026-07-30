' Replaces a Coupa export sheet with freshly loaded data.
'
' One routine serves all three export types. Each type pins a known set of headers to fixed
' columns -- because the tracker's formulas reference those positions -- and appends whatever
' else the export contained after them. Both the pin layout and the sheet name come from
' TenantConfig, so the two trackers can pin differently (they do for requisitions).
'
' Replaces the six near-identical Process{Requisition,Order,Invoice}DataFile modules that
' each variant carried. They differed only in pin lists, debug strings, and whether an import
' timestamp was written.

Public Sub ProcessCoupaExport(ByVal objectType As String, _
                              ByVal sourceData As Variant, _
                              ByVal headersDict As Object, _
                              ByVal fileExt As String)
    If IsEmpty(sourceData) Then
        MsgBox "No data was loaded from the export.", vbExclamation
        Exit Sub
    End If

    Dim sheetRole As String
    sheetRole = SheetRoleForObjectType(objectType)
    If Len(sheetRole) = 0 Then
        MsgBox "Unknown Coupa object type '" & objectType & "'.", vbExclamation
        Exit Sub
    End If

    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(TenantSheetName(sheetRole))
    On Error GoTo 0

    If wsTarget Is Nothing Then
        MsgBox "Sheet '" & TenantSheetName(sheetRole) & "' was not found.", vbCritical
        Exit Sub
    End If

    ' Split the tenant's pin list into parallel header/column arrays.
    Dim pins As Variant
    pins = TenantCoupaPins(objectType)

    Dim pinNames() As String
    Dim pinCols() As Long
    Dim pinCount As Long
    pinCount = ParsePins(pins, pinNames, pinCols)

    If pinCount = 0 Then
        MsgBox "No column pins are configured for '" & objectType & "'.", vbExclamation
        Exit Sub
    End If

    ' Every pinned header must be present, or the tracker's formulas would read blanks.
    Dim missing As String
    Dim i As Long
    For i = 1 To pinCount
        If Not headersDict.Exists(pinNames(i)) Then
            missing = missing & vbCrLf & "    " & pinNames(i)
        End If
    Next i

    If Len(missing) > 0 Then
        MsgBox "The export is missing required column(s):" & missing & vbCrLf & vbCrLf & _
               "Check the saved Coupa view includes them.", vbExclamation
        Exit Sub
    End If

    PauseThinking

    wsTarget.Cells.Clear

    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = 1   ' vbTextCompare

    ' Pinned columns first, at their fixed positions.
    For i = 1 To pinCount
        CopyData sourceData, fileExt, pinNames(i), headersDict(pinNames(i)), pinCols(i), wsTarget
        If Not processed.Exists(pinNames(i)) Then processed.Add pinNames(i), True
    Next i

    ' Then everything else, appended after the last pinned column.
    Dim maxPin As Long
    For i = 1 To pinCount
        If pinCols(i) > maxPin Then maxPin = pinCols(i)
    Next i

    Dim nextCol As Long
    nextCol = GetNextEmptyColumn(wsTarget, maxPin + 1)

    Dim key As Variant
    For Each key In headersDict.Keys
        If Not processed.Exists(CStr(key)) Then
            CopyData sourceData, fileExt, CStr(key), headersDict(key), nextCol, wsTarget
            nextCol = GetNextEmptyColumn(wsTarget, nextCol + 1)
        End If
    Next key

    StampImportTime objectType

    RestoreThinking
End Sub

' Records when this export type was last imported, so staleness is visible on the sheet.
' Written as a value rather than a formula so it does not recalculate.
Private Sub StampImportTime(ByVal objectType As String)
    Dim cellAddr As String
    cellAddr = TenantImportTimestampCell(objectType)
    If Len(cellAddr) = 0 Then Exit Sub

    Dim wsHelper As Worksheet
    On Error Resume Next
    Set wsHelper = ThisWorkbook.Sheets(TenantSheetName("helper"))
    On Error GoTo 0
    If wsHelper Is Nothing Then Exit Sub

    On Error Resume Next
    wsHelper.Range(cellAddr).Value = Now
    On Error GoTo 0
End Sub

' Turns Array("Req #=1", "Status=4") into parallel name/column arrays. Returns the count.
Private Function ParsePins(ByVal pins As Variant, _
                           ByRef pinNames() As String, _
                           ByRef pinCols() As Long) As Long
    On Error GoTo Empty_
    Dim total As Long
    total = UBound(pins) - LBound(pins) + 1
    If total < 1 Then GoTo Empty_

    ReDim pinNames(1 To total)
    ReDim pinCols(1 To total)

    Dim i As Long, n As Long
    Dim parts() As String
    For i = LBound(pins) To UBound(pins)
        parts = Split(CStr(pins(i)), "=")
        If UBound(parts) = 1 Then
            n = n + 1
            pinNames(n) = Trim$(parts(0))
            pinCols(n) = CLng(Trim$(parts(1)))
        End If
    Next i

    ParsePins = n
    Exit Function

Empty_:
    ParsePins = 0
End Function

Private Function SheetRoleForObjectType(ByVal objectType As String) As String
    Select Case LCase$(Trim$(objectType))
        Case "requisitions": SheetRoleForObjectType = "coupa-reqs"
        Case "orders":       SheetRoleForObjectType = "coupa-pos"
        Case "invoices":     SheetRoleForObjectType = "coupa-invs"
    End Select
End Function
