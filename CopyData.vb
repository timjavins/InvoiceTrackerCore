' Copies a named column from loaded source data into a target worksheet column.
'
' Both variants routed this through per-format handlers. The handlers no longer differ in
' behaviour now that loading is unified -- CSV and XLSX both arrive as the same 2D array --
' so the format argument is kept only to preserve the existing call signature and the
' debug trail. HandlerCoupaCSV/HandlerCoupaXLSX remain as thin named entry points.
Public Sub CopyData(ByVal sourceData As Variant, _
                    ByVal fileExt As String, _
                    ByVal colName As String, _
                    ByVal headerIndex As Long, _
                    ByVal targetCol As Long, _
                    ByVal wsTarget As Worksheet)
    Select Case LCase$(Trim$(fileExt))
        Case "csv"
            HandlerCoupaCSV sourceData, colName, headerIndex, targetCol, wsTarget
        Case "xlsx", "xlsm", "xls"
            HandlerCoupaXLSX sourceData, colName, headerIndex, targetCol, wsTarget
        Case Else
            Debug.Print "CopyData: unsupported extension '" & fileExt & "'."
    End Select
End Sub
