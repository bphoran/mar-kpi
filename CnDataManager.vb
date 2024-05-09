' CnDataManager ---------------------------------------------------------------------------------------------
 
Option Explicit
 
'version 2022.04.25
'author: Brendan Horan
 
' CnDataManager
' Maintains PolicyTerm and CovLayerTerms Collections
'
' A PolicyTerm is established for each unique policy number and
' effective term combination
'
' Example
' Policy 1234 effective 5/1/19 to 5/1/20 -> PolicyTerm "1234_2019050120200501"
' Policy 1234 effective 5/1/20 to 5/1/21 -> PolicyTerm "1234_2020050120210501"
'
' A CovLayerTerm is established for each unique coverage, layer key and
' effective term combination
'
' Example
' Coverage Cyber Risk effective 5/1/19 to 5/1/20, layer key 2 ->
'     CovLayerTerm "Cyber Risk_2019050120200501_2"
 
Private colPolicyTerms As New Collection
Private colPolicyTermsKeys As New Collection
 
Private colCovLayerTerms As New Collection
Private colCovLayerTermsKeys As New Collection
Private lCovLayerTermCount As Long
 
Private lLineCount As Long
Private lUniqueDetailCount As Long
 
Private lInterfaceCount As Long
 
Public Sub AddDqDetails(newItems As Variant)
   
    ' Exit if no rows under header
    If UBound(newItems, 1) < 2 Then
        Exit Sub
    End If
    
    Dim headerCol As Integer
    For headerCol = 1 To UBound(newItems, 2)
       
        Select Case newItems(1, headerCol)
           
            Case "TRIA"
                DETAIL_FIELDS(1) = headerCol
               
            Case "Premium"
                DETAIL_FIELDS(2) = headerCol
               
            Case "Product Line"
                DETAIL_FIELDS(3) = headerCol
               
            Case "Product Group"
                DETAIL_FIELDS(4) = headerCol
               
            Case "Coverage"
                DETAIL_FIELDS(5) = headerCol
               
            Case "Global Carrier"
                DETAIL_FIELDS(6) = headerCol
               
            Case "Local Carrier"
                DETAIL_FIELDS(7) = headerCol
               
            Case "Country"
                DETAIL_FIELDS(8) = headerCol
               
            Case "Layer Key"
                DETAIL_FIELDS(9) = headerCol
               
            Case "Policy Number"
                DETAIL_FIELDS(10) = headerCol
               
            Case "Effective"
                DETAIL_FIELDS(11) = headerCol
               
            Case "Expiry"
                DETAIL_FIELDS(12) = headerCol
               
            Case "Participation %"
                DETAIL_FIELDS(13) = headerCol
               
            Case Else
           
        End Select
       
    Next headerCol
   
    Dim lineRow As Long
    For lineRow = 2 To UBound(newItems, 1)
       
        If newItems(lineRow, DETAIL_FIELDS(9)) <> -1 Then
           
            'Remove brackets from policy number
            Dim newPolicy As String
           
            newPolicy = Replace(Format(newItems(lineRow, DETAIL_FIELDS(10)), "@"), "[", "")
            newPolicy = Replace(newPolicy, "]", "")
            newPolicy = Replace(newPolicy, """", "")
           
            Dim newEffective As String, newExpiry As String
            newEffective = Format(newItems(lineRow, DETAIL_FIELDS(11)), "mm/dd/yyyy")
            newExpiry = Format(newItems(lineRow, DETAIL_FIELDS(12)), "mm/dd/yyyy")
           
            Dim policyTermKey As String
            policyTermKey = newPolicy & "_" & Format(newEffective, "yyyymmdd") & _
                            Format(newExpiry, "yyyymmdd")
           
            Dim newPolicyTerm As policyTerm
           
            'Check if policy term is already in collection
            If LINQ.ContainsKey(colPolicyTermsKeys, policyTermKey) Then
                Set newPolicyTerm = colPolicyTerms.item(policyTermKey)
            Else
               
                Set newPolicyTerm = New policyTerm
                newPolicyTerm.init newPolicy, newEffective, newExpiry
                colPolicyTerms.Add newPolicyTerm, key:=policyTermKey
                colPolicyTermsKeys.Add policyTermKey
                lUniqueDetailCount = lUniqueDetailCount + 1
               
            End If
           
            Dim covLayerTermKey As String
            covLayerTermKey = newItems(lineRow, DETAIL_FIELDS(5)) & "_" & _
                            Format(newEffective, "yyyymmdd") & _
                            Format(newExpiry, "yyyymmdd") & "_" & newItems(lineRow, DETAIL_FIELDS(9))
           
            Dim newCovLayerTerm As LineItem
           
            'Check if coverage term is already in collection
            If LINQ.ContainsKey(colCovLayerTermsKeys, covLayerTermKey) Then
                Set newCovLayerTerm = colCovLayerTerms.item(covLayerTermKey)
            Else
               
                Set newCovLayerTerm = New LineItem
                newCovLayerTerm.sLine = newItems(lineRow, DETAIL_FIELDS(3))
                newCovLayerTerm.sGroup = newItems(lineRow, DETAIL_FIELDS(4))
                newCovLayerTerm.sCoverage = newItems(lineRow, DETAIL_FIELDS(5))
               
                'Reusing fields to hold dates
                newCovLayerTerm.sDocDate = Format(newEffective, "yyyymmdd")
                newCovLayerTerm.sDocName = Format(newExpiry, "yyyymmdd")
               
                'Store layer key
                newCovLayerTerm.dTria = newItems(lineRow, DETAIL_FIELDS(9))
               
                colCovLayerTerms.Add newCovLayerTerm, key:=covLayerTermKey
                colCovLayerTermsKeys.Add covLayerTermKey
                lCovLayerTermCount = lCovLayerTermCount + 1
               
            End If
           
            'Maintain layer participation
            newCovLayerTerm.dAmount = newCovLayerTerm.dAmount + newItems(lineRow, DETAIL_FIELDS(13))
           
            newPolicyTerm.AddDetailLine newItems, lineRow, newCovLayerTerm
            lLineCount = lLineCount + 1
           
        End If
       
    Next lineRow
   
End Sub
 
Public Sub AddDqPremiums(newItems As Variant)
   
    ' Exit if no rows under header
    If UBound(newItems, 1) < 2 Then
        Exit Sub
    End If
   
    Dim headerCol As Integer
    For headerCol = 1 To UBound(newItems, 2)
       
        Select Case newItems(1, headerCol)
           
            Case "Line Amount"
                PREMIUM_FIELDS(1) = headerCol
               
            Case "Invoice Line Type"
                PREMIUM_FIELDS(2) = headerCol
               
            Case "Coverage Line"
                PREMIUM_FIELDS(3) = headerCol
               
            Case "Product Group"
                PREMIUM_FIELDS(4) = headerCol
               
            Case "Global_carrier"
                PREMIUM_FIELDS(5) = headerCol
               
            Case "Local Carrier"
                PREMIUM_FIELDS(6) = headerCol
               
            Case "Country"
                PREMIUM_FIELDS(7) = headerCol
               
            Case "Carrier Policy ID"
                PREMIUM_FIELDS(8) = headerCol
               
            Case "Effective"
                PREMIUM_FIELDS(9) = headerCol
               
            Case "Expiry"
                PREMIUM_FIELDS(10) = headerCol
               
            Case Else
           
        End Select
       
    Next headerCol
   
    Dim lineRow As Long
    For lineRow = 2 To UBound(newItems, 1)
       
        Dim newPolicy As String
        newPolicy = newItems(lineRow, PREMIUM_FIELDS(8))
       
        Dim newEffective As String, newExpiry As String
        newEffective = Format(newItems(lineRow, PREMIUM_FIELDS(9)), "mm/dd/yyyy")
        newExpiry = Format(newItems(lineRow, PREMIUM_FIELDS(10)), "mm/dd/yyyy")
       
        Dim policyTermKey As String
        policyTermKey = newPolicy & "_" & Format(newEffective, "yyyymmdd") & _
                        Format(newExpiry, "yyyymmdd")
       
        Dim newPolicyTerm As policyTerm
       
        'Check if policy term is already in collection
        If LINQ.ContainsKey(colPolicyTermsKeys, policyTermKey) Then
            Set newPolicyTerm = colPolicyTerms.item(policyTermKey)
        Else
           
            Set newPolicyTerm = New policyTerm
            newPolicyTerm.init newPolicy, newEffective, newExpiry
            colPolicyTerms.Add newPolicyTerm, key:=policyTermKey
            colPolicyTermsKeys.Add policyTermKey
           
        End If
       
        newPolicyTerm.AddPremiumLine newItems, lineRow
        lLineCount = lLineCount + 1
       
    Next lineRow
   
End Sub
 
Public Sub AddInvoiceDetails(newItems As Variant)
   
    ' Exit if no rows under header
    If UBound(newItems, 1) < 2 Then
        Exit Sub
    End If
   
    Dim headerCol As Integer
    For headerCol = 1 To UBound(newItems, 2)
       
        Select Case newItems(1, headerCol)
           
            Case "Amount"
                INVOICE_FIELDS(1) = headerCol
               
            Case "Line Item"
                INVOICE_FIELDS(2) = headerCol
               
            Case "Product Line"
                INVOICE_FIELDS(3) = headerCol
               
            Case "Coverage"
                INVOICE_FIELDS(4) = headerCol
                
            Case "Insurer"
                INVOICE_FIELDS(5) = headerCol
               
            Case "Billing Date"
                INVOICE_FIELDS(6) = headerCol
               
            Case "Invoice Number"
                INVOICE_FIELDS(7) = headerCol
               
            Case "Policy/ Project No"
                INVOICE_FIELDS(8) = headerCol
               
            Case "Effective Date"
                INVOICE_FIELDS(9) = headerCol
                
            Case "Expiration Date"
                INVOICE_FIELDS(10) = headerCol
               
            Case Else
           
        End Select
       
    Next headerCol
   
    Dim lineRow As Long
    For lineRow = 2 To UBound(newItems, 1)
       
        Dim newPolicy As String
        newPolicy = newItems(lineRow, INVOICE_FIELDS(8))
       
        Dim newEffective As String, newExpiry As String
        newEffective = Format(newItems(lineRow, INVOICE_FIELDS(9)), "mm/dd/yyyy")
        newExpiry = Format(newItems(lineRow, INVOICE_FIELDS(10)), "mm/dd/yyyy")
       
        Dim policyTermKey As String
        policyTermKey = newPolicy & "_" & Format(newEffective, "yyyymmdd") & _
                        Format(newExpiry, "yyyymmdd")
       
        Dim newPolicyTerm As policyTerm
       
        'Check if policy term is already in collection
        If LINQ.ContainsKey(colPolicyTermsKeys, policyTermKey) Then
            Set newPolicyTerm = colPolicyTerms.item(policyTermKey)
        Else
           
            Set newPolicyTerm = New policyTerm
            newPolicyTerm.init newPolicy, newEffective, newExpiry
            colPolicyTerms.Add newPolicyTerm, key:=policyTermKey
            colPolicyTermsKeys.Add policyTermKey
            
        End If
       
        newPolicyTerm.AddInvoiceLine newItems, lineRow
        lLineCount = lLineCount + 1
       
    Next lineRow
   
End Sub
 
Public Sub AddInterfaceDetails(newItems As Variant)
   
    ' Exit if no rows under header
    If UBound(newItems, 1) < 3 Then
        Exit Sub
    End If
   
    Dim lineRow As Long
    For lineRow = 3 To UBound(newItems, 1)
       
        ' NOTE: If the columns of Document Interface change, the column
        ' numbers will need to change, ex. newItems(lineRow, [7])
       
        Dim newPolicy As String
        newPolicy = newItems(lineRow, 7)
       
        Dim newEffective As String
        newEffective = Format(newItems(lineRow, 8), "mm/dd/yyyy")
       
        Dim policyTermKey As String, policyAndTerm As String
        policyAndTerm = newPolicy & "_" & Format(newEffective, "yyyymmdd")
        policyTermKey = LINQ.GetFullKey(colPolicyTermsKeys, policyAndTerm)
       
        Dim newPolicyTerm As policyTerm
       
        'Check if policy term is already in collection
        If policyTermKey <> "" Then
            Set newPolicyTerm = colPolicyTerms.item(policyTermKey)
        Else
           
            Set newPolicyTerm = New policyTerm
            newPolicyTerm.init newPolicy, newEffective, ""
            colPolicyTerms.Add newPolicyTerm, key:=policyAndTerm
            colPolicyTermsKeys.Add policyAndTerm
           
        End If
       
        newPolicyTerm.AddInterfaceLine newItems, lineRow
        lInterfaceCount = lInterfaceCount + 1
       
    Next lineRow
   
End Sub
 
Public Function HasPremiumData() As Boolean
   
    HasPremiumData = lLineCount > 0
   
End Function
 
Public Function GetPremiumData() As Variant
   
    Dim result() As Variant
    ReDim result(1 To lLineCount, 1 To 20)
   
    Dim lLineRow As Long
    Dim lLineCol As Long
    Dim lResultRow As Long: lResultRow = 1
   
    Dim oPolicyTerm As Variant
    For Each oPolicyTerm In colPolicyTerms
       
        If oPolicyTerm.HasPremiumLines() Then
           
            Dim policyTermLines As Variant
            policyTermLines = oPolicyTerm.GetPremiumLines()
           
            For lLineRow = 1 To UBound(policyTermLines, 1)
               
                For lLineCol = 1 To UBound(policyTermLines, 2)
                    result(lResultRow, lLineCol) = policyTermLines(lLineRow, lLineCol)
                Next lLineCol
               
                lResultRow = lResultRow + 1
               
            Next lLineRow
           
        End If
       
    Next oPolicyTerm
   
    GetPremiumData = result
   
End Function
 
Public Function HasInterfaceData() As Boolean
   
    HasInterfaceData = lInterfaceCount > 0
   
End Function
 
Public Function GetInterfaceData() As Variant
   
    Dim result() As Variant
    ReDim result(1 To lUniqueDetailCount + lInterfaceCount, 1 To 11)
   
    Dim lLineRow As Long
    Dim lLineCol As Long
    Dim lResultRow As Long: lResultRow = 1
   
    Dim oPolicyTerm As Variant
    For Each oPolicyTerm In colPolicyTerms
       
        If oPolicyTerm.HasInterfaceLines() Then
           
            Dim policyTermLines As Variant
            policyTermLines = oPolicyTerm.GetInterfaceLines()
           
            For lLineRow = 1 To UBound(policyTermLines, 1)
               
                For lLineCol = 1 To UBound(policyTermLines, 2)
                    result(lResultRow, lLineCol) = policyTermLines(lLineRow, lLineCol)
                Next lLineCol
               
                lResultRow = lResultRow + 1
               
            Next lLineRow
           
        End If
       
    Next oPolicyTerm
   
    GetInterfaceData = result
   
End Function
 