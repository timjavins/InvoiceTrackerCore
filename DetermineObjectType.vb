' Identifies which kind of Coupa export a file is, from its header names.
'
' Verified against real exports for both suppliers (2026-07-30). All three carry
' PO Number, Status, Supplier, and Total, so those cannot discriminate. The identifier
' columns are what distinguish them:
'
'   Requisition export : "Req #"
'   Order (PO) export  : "Req"        <- note: no "#"
'   Invoice export     : "Invoice #"
'
' IMPORTANT: "Req" is a prefix of "Req #", so these must be compared EXACTLY. A
' substring or Like "*Req*" test would match a requisition export as an order.
' The order of the checks below is therefore not load-bearing, but exact comparison is.
'
' Secondary markers exist and are used as fallbacks in case a saved Coupa view omits the
' identifier column: orders also carry "Order Date" and "Uninvoiced Amount"; invoices also
' carry "History", "Invoice ID", and "Invoice Date"; requisitions also carry
' "Supplier Part Number" and "Submitted On".
Public Function DetermineObjectType(ByVal headersDict As Object) As String
    DetermineObjectType = "unknown"
    If headersDict Is Nothing Then Exit Function

    ' Index the header names once for exact, case-insensitive lookup.
    Dim present As Object
    Set present = CreateObject("Scripting.Dictionary")
    present.CompareMode = 1   ' vbTextCompare

    Dim key As Variant
    Dim name As String
    For Each key In headersDict.Keys
        name = VBA.Trim$(CStr(headersDict(key)))
        If Len(name) > 0 Then
            If Not present.Exists(name) Then present.Add name, True
        End If
    Next key

    ' Primary: the identifier column unique to each export type.
    If present.Exists("Invoice #") Then
        DetermineObjectType = "invoices"
    ElseIf present.Exists("Req #") Then
        DetermineObjectType = "requisitions"
    ElseIf present.Exists("Req") Then
        DetermineObjectType = "orders"

    ' Fallback: a saved view that dropped the identifier column.
    ElseIf present.Exists("History") Or present.Exists("Invoice ID") _
           Or present.Exists("Invoice Date") Then
        DetermineObjectType = "invoices"
    ElseIf present.Exists("Order Date") Or present.Exists("Uninvoiced Amount") _
           Or present.Exists("Transmission Status") Then
        DetermineObjectType = "orders"
    ElseIf present.Exists("Supplier Part Number") Or present.Exists("Submitted On") Then
        DetermineObjectType = "requisitions"
    End If
End Function
