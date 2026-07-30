' Writes one column of CSV-sourced data into a target worksheet column.
' Kept as a named entry point so the debug trail says which format the data came from.
Public Sub HandlerCoupaCSV(ByVal sourceData As Variant, _
                           ByVal colName As String, _
                           ByVal headerIndex As Long, _
                           ByVal targetCol As Long, _
                           ByVal wsTarget As Worksheet)
    WriteCoupaColumnFromSourceData sourceData, colName, headerIndex, targetCol, wsTarget, _
                                   "HandlerCoupaCSV"
End Sub
