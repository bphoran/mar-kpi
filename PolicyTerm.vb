' PolicyTerm ------------------------------------------------------------------------------------------------
 
Option Explicit
 
'version 2022.04.25
'author: Brendan Horan
 
Private sPolicy As String
Private sEffective As String
Private sExpiry As String
 
Private sCountry As String
 
Private colDetails As New Collection
Private colPremiums As New Collection
Private colInvoices As New Collection
Private colInterface As New Collection
 
Private colGroups As New Collection
Private colGroupsKeys As New Collection
 
Private colCoverages As New Collection
Private colCoveragesKeys As New Collection
 
Private colGlobals As New Collection
Private colGlobalsKeys As New Collection
 
Public Sub init(newPolicy As String, newEffective As String, newExpiry As String)
   
    sPolicy = newPolicy
    sEffective = newEffective
    sExpiry = newExpiry
   
End Sub
 
Public Sub AddDetailLine(ByRef lineItems As Variant, lineRow As Long, ByRef covLayerTerm As LineItem)
   
    Dim newItem As LineItem
    Set newItem = New LineItem
   
    newItem.dTria = lineItems(lineRow, DETAIL_FIELDS(1))
    newItem.dAmount = lineItems(lineRow, DETAIL_FIELDS(2))
    newItem.sType = "Premium"
    newItem.sLine = lineItems(lineRow, DETAIL_FIELDS(3))
    newItem.sGroup = lineItems(lineRow, DETAIL_FIELDS(4))
    newItem.sCoverage = lineItems(lineRow, DETAIL_FIELDS(5))
    newItem.sGlobal = lineItems(lineRow, DETAIL_FIELDS(6))
    newItem.sLocal = lineItems(lineRow, DETAIL_FIELDS(7))
    newItem.sCountry = lineItems(lineRow, DETAIL_FIELDS(8))
    sCountry = lineItems(lineRow, DETAIL_FIELDS(8))
   
    Set newItem.oCovLayerTerm = covLayerTerm
   
    colDetails.Add newItem
   
    ' Share with Invoice Details
    AddConnection colGroups, colGroupsKeys, newItem.sGroup, newItem.sLocal
    AddConnection colGroups, colGroupsKeys, newItem.sGroup, newItem.sCoverage
    AddConnection colGlobals, colGlobalsKeys, newItem.sGlobal, newItem.sLocal
    AddConnection colGlobals, colGlobalsKeys, newItem.sGlobal, newItem.sCoverage
   
    ' Share with DQ Premiums
    AddConnection colCoverages, colCoveragesKeys, newItem.sCoverage, newItem.sGroup
    AddConnection colCoverages, colCoveragesKeys, newItem.sCoverage, newItem.sLocal
   
End Sub
 
Public Sub AddPremiumLine(ByRef lineItems As Variant, lineRow As Long)
   
    Dim newItem As LineItem
    Set newItem = New LineItem
   
    newItem.dAmount = lineItems(lineRow, PREMIUM_FIELDS(1))
    newItem.sType = lineItems(lineRow, PREMIUM_FIELDS(2))
    newItem.sLine = lineItems(lineRow, PREMIUM_FIELDS(3))
    newItem.sGroup = lineItems(lineRow, PREMIUM_FIELDS(4))
    newItem.sGlobal = lineItems(lineRow, PREMIUM_FIELDS(5))
    newItem.sLocal = lineItems(lineRow, PREMIUM_FIELDS(6))
    newItem.sCountry = lineItems(lineRow, PREMIUM_FIELDS(7))
    sCountry = lineItems(lineRow, PREMIUM_FIELDS(7))
   
    'Store Client Share
    newItem.dTria = lineItems(lineRow, 21)
   
    colPremiums.Add newItem
   
    ' Share with Invoice Details
    AddConnection colGroups, colGroupsKeys, newItem.sGroup, newItem.sLocal
    AddConnection colGlobals, colGlobalsKeys, newItem.sGlobal, newItem.sLocal
   
End Sub
 
Public Sub AddInvoiceLine(ByRef lineItems As Variant, lineRow As Long)
   
    Dim newItem As LineItem
    Set newItem = New LineItem
   
    newItem.dAmount = lineItems(lineRow, INVOICE_FIELDS(1))
    newItem.sType = lineItems(lineRow, INVOICE_FIELDS(2))
    newItem.sLine = lineItems(lineRow, INVOICE_FIELDS(3))
    newItem.sCoverage = lineItems(lineRow, INVOICE_FIELDS(4))
    newItem.sLocal = lineItems(lineRow, INVOICE_FIELDS(5))
    newItem.sDocDate = Format(lineItems(lineRow, INVOICE_FIELDS(6)), "mm/dd/yyyy")
    newItem.sDocName = lineItems(lineRow, INVOICE_FIELDS(7))
    newItem.sCountry = sCountry
   
    colInvoices.Add newItem
   
    'Share with DQ Premiums
    AddConnection colCoverages, colCoveragesKeys, newItem.sCoverage, newItem.sLocal
   
End Sub
 
Public Sub AddInterfaceLine(ByRef lineItems As Variant, lineRow As Long)
   
    Dim newItem As LineItem
    Set newItem = New LineItem
   
    newItem.sDocName = lineItems(lineRow, 1)
    newItem.sType = lineItems(lineRow, 3)
    newItem.sDocDate = lineItems(lineRow, 6)
    newItem.sLine = lineItems(lineRow, 10)
    newItem.sGlobal = lineItems(lineRow, 12)
    newItem.sLocal = lineItems(lineRow, 13)
   
    colInterface.Add newItem
   
End Sub
 
Public Function HasPremiumLines() As Boolean
   
    HasPremiumLines = colDetails.Count > 0 Or colPremiums.Count > 0 Or colInvoices.Count > 0
   
End Function
 
Public Function GetPremiumLines() As Variant
   
    Dim result As Variant
    ReDim result(1 To colDetails.Count + colPremiums.Count + colInvoices.Count, _
                    1 To 20)
   
    'Default values of zero
    Dim dMartTotal As Double
    Dim dPremiumTotal As Double
    Dim dInvoiceTotal As Double
    Dim dPremFeeTotal As Double
    Dim dInvFeeTotal As Double
   
    Dim dPremRebate As Double
   
    Dim lResultRow As Long: lResultRow = 1
    Dim colItem As Variant
    For Each colItem In colDetails
       
        result(lResultRow, 1) = "DQ Detail"
        result(lResultRow, 2) = sEffective
        result(lResultRow, 3) = sExpiry
        result(lResultRow, 4) = sPolicy
        result(lResultRow, 6) = colItem.dTria
        result(lResultRow, 7) = colItem.dAmount
        result(lResultRow, 8) = 0#
        result(lResultRow, 9) = 0#
        result(lResultRow, 10) = 0#
        result(lResultRow, 11) = 0#
        result(lResultRow, 12) = colItem.sType
        result(lResultRow, 13) = colItem.sDocDate
       
        ' dTria contains layer key #, dAmount contains layer participation sum
        ' If participation sum for layer key 1 and above is not 100, add info
        ' to [Invoice #] column of spreadsheet
        If colItem.oCovLayerTerm.dTria > 0 And colItem.oCovLayerTerm.dAmount <> 100 Then
           
            result(lResultRow, 14) = "Key " & colItem.oCovLayerTerm.dTria & _
                ",Part " & colItem.oCovLayerTerm.dAmount & "%"
           
        End If
       
        result(lResultRow, 15) = colItem.sGlobal
        result(lResultRow, 16) = colItem.sLocal
        result(lResultRow, 17) = colItem.sCoverage
        result(lResultRow, 18) = colItem.sGroup
        result(lResultRow, 19) = colItem.sLine
        result(lResultRow, 20) = colItem.sCountry
       
        'Only count if country matches DQ Premium
        If colItem.sCountry = sCountry Then
            dMartTotal = dMartTotal + colItem.dAmount + colItem.dTria
        End If
       
        lResultRow = lResultRow + 1
       
    Next colItem
   
    For Each colItem In colPremiums
       
        result(lResultRow, 1) = "DQ Premium"
        result(lResultRow, 2) = sEffective
        result(lResultRow, 3) = sExpiry
        result(lResultRow, 4) = sPolicy
        result(lResultRow, 6) = 0#
        result(lResultRow, 8) = 0#
        result(lResultRow, 9) = 0#
        result(lResultRow, 10) = 0#
        result(lResultRow, 11) = 0#
        result(lResultRow, 12) = colItem.sType
        result(lResultRow, 13) = colItem.sDocDate
        result(lResultRow, 14) = colItem.sDocName
        result(lResultRow, 15) = colItem.sGlobal
        result(lResultRow, 16) = colItem.sLocal
        result(lResultRow, 17) = colItem.sCoverage
        result(lResultRow, 18) = colItem.sGroup
        result(lResultRow, 19) = colItem.sLine
        result(lResultRow, 20) = colItem.sCountry
       
        'Placing Client Share into DQ Detail premium column
        result(lResultRow, 7) = colItem.dTria
        dPremRebate = dPremRebate + colItem.dTria
       
        colItem.sType = UCase(colItem.sType)
       
        If IsFee(colItem.sType) Then
            result(lResultRow, 10) = colItem.dAmount
            dPremFeeTotal = dPremFeeTotal + colItem.dAmount
        Else
            result(lResultRow, 8) = colItem.dAmount
            dPremiumTotal = dPremiumTotal + colItem.dAmount
        End If
       
        Dim item As Variant
       
        If LINQ.ContainsKey(colCoveragesKeys, colItem.sLocal) Then
            Set item = colCoverages.item(colItem.sLocal)
            result(lResultRow, 17) = item.sType
        End If
       
        If LINQ.ContainsKey(colCoveragesKeys, colItem.sGroup) Then
            Set item = colCoverages.item(colItem.sGroup)
            result(lResultRow, 17) = item.sType
        End If
       
        lResultRow = lResultRow + 1
       
    Next colItem
   
    For Each colItem In colInvoices
       
        result(lResultRow, 1) = "Invoice Detail"
        result(lResultRow, 2) = sEffective
        result(lResultRow, 3) = sExpiry
        result(lResultRow, 4) = sPolicy
        result(lResultRow, 6) = colItem.dTria
        result(lResultRow, 7) = 0#
        result(lResultRow, 8) = 0#
        result(lResultRow, 9) = 0#
        result(lResultRow, 10) = 0#
        result(lResultRow, 11) = 0#
        result(lResultRow, 12) = colItem.sType
        result(lResultRow, 13) = colItem.sDocDate
        result(lResultRow, 14) = colItem.sDocName
        result(lResultRow, 15) = colItem.sGlobal
        result(lResultRow, 16) = colItem.sLocal
        result(lResultRow, 17) = colItem.sCoverage
        result(lResultRow, 18) = colItem.sGroup
        result(lResultRow, 19) = colItem.sLine
        result(lResultRow, 20) = colItem.sCountry
       
        colItem.sType = UCase(colItem.sType)
       
        If IsFee(colItem.sType) Then
            result(lResultRow, 11) = colItem.dAmount
            dInvFeeTotal = dInvFeeTotal + colItem.dAmount
        Else
            result(lResultRow, 9) = colItem.dAmount
            dInvoiceTotal = dInvoiceTotal + colItem.dAmount
        End If
       
        If LINQ.ContainsKey(colGlobalsKeys, colItem.sCoverage) Then
            Set item = colGlobals.item(colItem.sCoverage)
            result(lResultRow, 15) = item.sType
        End If
       
        If LINQ.ContainsKey(colGlobalsKeys, colItem.sLocal) Then
            Set item = colGlobals.item(colItem.sLocal)
            result(lResultRow, 15) = item.sType
        End If
       
        If LINQ.ContainsKey(colGroupsKeys, colItem.sLocal) Then
            Set item = colGroups.item(colItem.sLocal)
            result(lResultRow, 18) = item.sType
        End If
       
        If LINQ.ContainsKey(colGroupsKeys, colItem.sCoverage) Then
            Set item = colGroups.item(colItem.sCoverage)
            result(lResultRow, 18) = item.sType
        End If
       
        lResultRow = lResultRow + 1
       
    Next colItem
   
    GetPremiumLines = result
   
    If Not LINQ.DIAGNOSE_PREMIUM Then
        GetPremiumLines = result
        Exit Function
    End If
   
    Dim sIssue As String: sIssue = "No Issue"
   
    If dPremFeeTotal > 1 Or dInvFeeTotal > 1 Then
        sIssue = "Taxes/Fees"
    End If
   
    If Abs(dMartTotal - dInvoiceTotal) > LINQ.TOL Then
        sIssue = "DQ Detail vs Invoice"
       
        If Abs(dMartTotal + dPremRebate - dPremiumTotal) < LINQ.TOL Then
            sIssue = "Commission Rebate"
        End If
       
    End If
   
    If Abs(dPremiumTotal - dInvoiceTotal) > LINQ.TOL Then
        sIssue = "DQ Premium vs Invoice"
        
        If dPremiumTotal > 1 Then
           
            Dim dRatio As Double
            dRatio = dInvoiceTotal / dPremiumTotal
           
            'Cancelled/Reissued invoice premiums will appear as multiples of the DQ Premiums
            If Abs(dRatio - Int(dRatio)) < 0.01 Then
                sIssue = "DQ Premium vs Invoice (" & CStr(Int(dRatio)) & "x)"
            End If
           
        End If
       
    End If
   
    If dPremiumTotal < LINQ.TOL Or dInvoiceTotal < LINQ.TOL Then
       sIssue = "Not in DQ Premium/Invoice"
    End If
   
    If dMartTotal < LINQ.TOL Then
        sIssue = "Not in DQ Detail"
    End If
   
    For lResultRow = 1 To UBound(result, 1)
        If result(lResultRow, 20) = sCountry Then
            result(lResultRow, 5) = sIssue
        Else
            result(lResultRow, 5) = "Multiple Countries"
        End If
    Next lResultRow
   
    GetPremiumLines = result
   
End Function
 
Public Function HasInterfaceLines() As Boolean
   
    HasInterfaceLines = colDetails.Count > 0 Or colInterface.Count > 0
   
End Function
 
Public Function GetInterfaceLines() As Variant
   
    Dim result As Variant
   
    Dim issueIndex As Integer: issueIndex = 1
    Dim lResultRow As Long: lResultRow = 1
   
    If Not LINQ.DIAGNOSE_INTERFACE Then
        issueIndex = 6
    End If
   
    Dim colItem As LineItem
   
    If colInterface.Count < 1 Then
       
        ReDim result(1 To 1, 1 To 11)
        result(1, 1) = "DQ Detail"
        result(1, 2) = sEffective
        result(1, 3) = sExpiry
        result(1, 4) = sPolicy
       
        ' No Match - Review Needed
        result(1, 5) = InterfaceIssueArray(issueIndex)
       
        Set colItem = colDetails.item(1)
        result(1, 11) = colItem.sLine
       
        GetInterfaceLines = result
        Exit Function
       
    End If
   
    If colDetails.Count > 0 Then
       
        ReDim result(1 To colInterface.Count + 1, 1 To 11)
        result(1, 1) = "DQ Detail"
        result(1, 2) = sEffective
        result(1, 3) = sExpiry
        result(1, 4) = sPolicy
       
        Set colItem = colDetails.item(1)
        result(1, 11) = colItem.sLine
       
        If issueIndex = 1 Then
            issueIndex = 2
        End If
       
        lResultRow = lResultRow + 1
   
    Else
        ReDim result(1 To colInterface.Count, 1 To 11)
    End If
   
    For Each colItem In colInterface
       
        result(lResultRow, 1) = "Document"
        result(lResultRow, 2) = sEffective
        result(lResultRow, 3) = ""
        result(lResultRow, 4) = sPolicy
        result(lResultRow, 6) = colItem.sLocal
        result(lResultRow, 7) = colItem.sType
        result(lResultRow, 8) = colItem.sGlobal
        result(lResultRow, 9) = colItem.sDocDate
        result(lResultRow, 10) = colItem.sDocName
        result(lResultRow, 11) = colItem.sLine
       
        If UCase(Trim(result(lResultRow, 7))) = "POLICY" And issueIndex = 2 Then
            issueIndex = 3
        End If
       
        If UCase(result(lResultRow, 8)) = "PDF" And issueIndex = 3 Then
            issueIndex = 4
        End If
       
        If UCase(result(lResultRow, 6)) = "YES" And issueIndex = 4 Then
            issueIndex = 5
        End If
       
        lResultRow = lResultRow + 1
       
    Next colItem
   
    For lResultRow = 1 To UBound(result, 1)
        result(lResultRow, 5) = InterfaceIssueArray(issueIndex)
    Next lResultRow
   
    GetInterfaceLines = result
   
End Function
 
Private Function IsFee(sType As String) As Boolean
   
    Select Case UCase(sType)
       
        Case "FEE"
            IsFee = True
           
        Case "LESS COMMISSION"
            IsFee = True
           
        Case "MB RST NON RECOVERABLE"
            IsFee = True
           
        Case "NL RST NON RECOVERABLE"
            IsFee = True
           
        Case "POLICY FEE"
            IsFee = True
           
        Case "PREMIUM TAX"
            IsFee = True
           
        Case "PST - SASK NON RECOVERABLE"
            IsFee = True
           
        Case "PST-ONT NON RECOVERABLE"
            IsFee = True
           
        Case "QST NON RECOVERABLE"
            IsFee = True
           
        Case "STAMPING FEE"
            IsFee = True
       
        Case "SURCHARGE"
            IsFee = True
           
        Case "SURPLUS LINES TAX"
            IsFee = True
           
        Case Else
            IsFee = False
           
    End Select
   
End Function
 
' Allow missing coverages/groups/global carriers to be found by borrowing
' the information from other data sources, for the same PolicyTerm
Private Sub AddConnection(coll As Collection, collKeys As Collection, newItem As String, newKey As String)
   
    Dim LineItem As LineItem
   
    If LINQ.ContainsKey(collKeys, newKey) Then
        Set LineItem = coll.item(newKey)
    Else
        Set LineItem = New LineItem
        coll.Add LineItem, key:=newKey
        collKeys.Add newKey
    End If
   
    ' Using sType to hold info
    LineItem.sType = newItem
   
End Sub
 