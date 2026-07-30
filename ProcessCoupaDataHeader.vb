' Builds a {column index -> header name} dictionary from loaded Coupa export data.
'
' Takes the array from LoadCoupaSourceData rather than reopening the file, so CSV and XLSX
' are handled identically.
'
' The former "skip any header containing a digit, except Custom Field 4" rule is gone. It
' did not skip such a header -- it did Exit For, silently discarding that column and every
' column after it. Checked against real exports for both suppliers (2026-07-30): no header
' in any of the six contains a digit, and "Custom Field 4" appears in none of them, so the
' rule never fired and its exception was dead. Keeping it would risk truncating the header
' set the first time someone adds a column like "Custom Field 2" to a saved view.
Public Function ProcessCoupaDataHeader(ByVal sourceData As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set ProcessCoupaDataHeader = dict

    If IsEmpty(sourceData) Then
        Debug.Print "ProcessCoupaDataHeader: source data is empty."
        Exit Function
    End If

    Dim colCount As Long
    colCount = UBound(sourceData, 2)

    Dim i As Long
    Dim headerValue As String
    For i = 1 To colCount
        headerValue = VBA.Trim$(CStr(sourceData(1, i)))

        ' Blank cells past the last real header: stop, rather than indexing empty names.
        If Len(headerValue) = 0 Then Exit For

        dict.Add i, headerValue
    Next i
End Function
