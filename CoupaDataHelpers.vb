' Loads a Coupa export (CSV or XLSX) into a uniform 1-based 2D array, so downstream code
' handles both formats the same way. Row 1 is the header row.
'
' Grafted from both variants: JCI returned the data as a value (no global state) but split
' CSV on every comma; Securitas had a quote-aware parser (itself buggy -- see
' ParseCSVLine.vb) but stashed the loaded file in module-level globals that callers had to
' remember to clean up. Core keeps the value-returning shape and parses quotes correctly.
'
' Returns Empty when nothing could be loaded.

' Loads an export by extension. Unknown extensions return Empty.
Public Function LoadCoupaSourceData(ByVal selectedFile As String, ByVal fileExt As String) As Variant
    Select Case LCase$(Trim$(fileExt))
        Case "csv"
            LoadCoupaSourceData = LoadCoupaCsvSourceData(selectedFile)
        Case "xlsx", "xlsm", "xls"
            LoadCoupaSourceData = LoadCoupaXlsxSourceData(selectedFile)
        Case Else
            Debug.Print "LoadCoupaSourceData: unsupported extension '" & fileExt & "'."
            LoadCoupaSourceData = Empty
    End Select
End Function

Private Function LoadCoupaCsvSourceData(ByVal selectedFile As String) As Variant
    Dim fileNum As Integer
    Dim lineData As String
    Dim lines() As String
    Dim lineCount As Long

    On Error GoTo Failed

    fileNum = FreeFile()
    Open selectedFile For Input As #fileNum

    ' Read once; parse once. The previous JCI version parsed every line twice.
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineData
        lineCount = lineCount + 1
        ReDim Preserve lines(1 To lineCount)
        lines(lineCount) = lineData
    Loop

    Close #fileNum
    fileNum = 0

    If lineCount = 0 Then
        LoadCoupaCsvSourceData = Empty
        Exit Function
    End If

    ' The BOM only ever appears at the very start of the file.
    lines(1) = StripUtf8Bom(lines(1))

    Dim parsed() As Variant
    ReDim parsed(1 To lineCount)

    Dim maxCols As Long
    Dim r As Long
    For r = 1 To lineCount
        parsed(r) = ParseCSVLine(lines(r))
        If UBound(parsed(r)) + 1 > maxCols Then maxCols = UBound(parsed(r)) + 1
    Next r

    If maxCols = 0 Then
        LoadCoupaCsvSourceData = Empty
        Exit Function
    End If

    Dim sourceData() As Variant
    ReDim sourceData(1 To lineCount, 1 To maxCols)

    Dim c As Long
    Dim fields As Variant
    For r = 1 To lineCount
        fields = parsed(r)
        For c = 0 To UBound(fields)
            sourceData(r, c + 1) = fields(c)
        Next c
    Next r

    LoadCoupaCsvSourceData = sourceData
    Exit Function

Failed:
    Debug.Print "LoadCoupaCsvSourceData: " & Err.Description
    If fileNum <> 0 Then Close #fileNum
    LoadCoupaCsvSourceData = Empty
End Function

Private Function LoadCoupaXlsxSourceData(ByVal selectedFile As String) As Variant
    Dim wbData As Workbook
    Dim sourceData As Variant

    On Error GoTo Failed

    Set wbData = Workbooks.Open(selectedFile, ReadOnly:=True)
    wbData.Windows(1).Visible = False
    sourceData = wbData.Sheets(1).UsedRange.Value2
    wbData.Close SaveChanges:=False
    Set wbData = Nothing

    If IsEmpty(sourceData) Then
        LoadCoupaXlsxSourceData = Empty
        Exit Function
    End If

    ' A single-cell UsedRange comes back as a bare value, not a 2D array.
    If Not IsArray(sourceData) Then
        Dim single_() As Variant
        ReDim single_(1 To 1, 1 To 1)
        single_(1, 1) = sourceData
        LoadCoupaXlsxSourceData = single_
        Exit Function
    End If

    ' UsedRange is already a 1-based 2D array, but it can start below row 1 if the sheet
    ' has leading blank rows, so normalise to a 1-based copy.
    Dim rows_ As Long, cols_ As Long
    rows_ = UBound(sourceData, 1) - LBound(sourceData, 1) + 1
    cols_ = UBound(sourceData, 2) - LBound(sourceData, 2) + 1

    Dim normalized() As Variant
    ReDim normalized(1 To rows_, 1 To cols_)

    Dim r As Long, c As Long
    For r = 1 To rows_
        For c = 1 To cols_
            normalized(r, c) = sourceData(LBound(sourceData, 1) + r - 1, _
                                         LBound(sourceData, 2) + c - 1)
        Next c
    Next r

    ' A leading BOM survives into the first header cell when Excel opens a UTF-8 file.
    If VarType(normalized(1, 1)) = vbString Then
        normalized(1, 1) = StripUtf8Bom(CStr(normalized(1, 1)))
    End If

    LoadCoupaXlsxSourceData = normalized
    Exit Function

Failed:
    Debug.Print "LoadCoupaXlsxSourceData: " & Err.Description
    If Not wbData Is Nothing Then wbData.Close SaveChanges:=False
    LoadCoupaXlsxSourceData = Empty
End Function

' Removes a UTF-8 byte-order mark if present.
Public Function StripUtf8Bom(ByVal text As String) As String
    Dim bom As String
    bom = Chr$(239) & Chr$(187) & Chr$(191)

    If Left$(text, 3) = bom Then
        StripUtf8Bom = Mid$(text, 4)
    Else
        StripUtf8Bom = text
    End If
End Function

' Copies one column out of loaded source data into a target worksheet column, writing
' colName as the header. Bulk-writes the column in a single operation.
Public Sub WriteCoupaColumnFromSourceData(ByVal sourceData As Variant, _
                                          ByVal colName As String, _
                                          ByVal headerIndex As Long, _
                                          ByVal targetCol As Long, _
                                          ByVal wsTarget As Worksheet, _
                                          Optional ByVal handlerName As String = "WriteCoupaColumn")
    If Len(Trim$(colName)) = 0 Or wsTarget Is Nothing Then
        Debug.Print handlerName & ": missing parameter(s). Exiting."
        Exit Sub
    End If

    If IsEmpty(sourceData) Then
        Debug.Print handlerName & ": source data is empty. Exiting."
        Exit Sub
    End If

    Dim sourceRowCount As Long, sourceColCount As Long
    sourceRowCount = UBound(sourceData, 1)
    sourceColCount = UBound(sourceData, 2)

    wsTarget.Cells(1, targetCol).Value = colName

    If sourceRowCount < 2 Or headerIndex < 1 Or headerIndex > sourceColCount Then
        Debug.Print handlerName & ": no data rows or header index " & headerIndex & _
                    " out of range (1.." & sourceColCount & ")."
        Exit Sub
    End If

    Dim outputData() As Variant
    ReDim outputData(1 To sourceRowCount - 1, 1 To 1)

    ' Identifier columns are written as text. These are the join keys the tracker's XLOOKUPs
    ' search, and the tracker holds its side as text -- a numeric Req # or PO # here would not
    ' match, and every dependent formula would return not-found.
    Dim asText As Boolean
    asText = IsIdentifierHeader(colName)

    Dim r As Long
    For r = 2 To sourceRowCount
        If asText Then
            outputData(r - 1, 1) = CoerceIdentifier(sourceData(r, headerIndex))
        Else
            outputData(r - 1, 1) = sourceData(r, headerIndex)
        End If
    Next r

    Dim targetRange As Range
    Set targetRange = wsTarget.Range(wsTarget.Cells(2, targetCol), _
                                     wsTarget.Cells(sourceRowCount, targetCol))

    If asText Then targetRange.NumberFormat = "@"
    targetRange.Value = outputData
End Sub

' Whether a Coupa export column holds an identifier that the tracker joins on, and so must be
' stored as text to match the tracker's side.
'
' Amounts, dates and statuses are deliberately excluded -- those want to stay numeric or
' date-typed so they sum and sort correctly.
Private Function IsIdentifierHeader(ByVal colName As String) As Boolean
    Select Case UCase$(Trim$(colName))
        Case "REQ #", "REQ", "PO NUMBER", "INVOICE #", "SUPPLIER PART NUMBER", _
             "INVOICE ID", "PAYMENT NUM'S", "COPIED FROM"
            IsIdentifierHeader = True
    End Select
End Function
