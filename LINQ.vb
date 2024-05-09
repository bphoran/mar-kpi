' LINQ ------------------------------------------------------------------------------------------------------
 
Option Explicit
 
'version 2022.04.25
'author: Brendan Horan
 
' LINQ
' Module that responds to button clicks, reads appropriate spreadsheets
' into 2d arrays, feeds the data into a CnDataManager, then displays
' the manager's results
 
Public Const DIAGNOSE_PREMIUM As Boolean = True
Public Const DIAGNOSE_INTERFACE As Boolean = True
Public Const COLOR_POLICY_LINES As Boolean = True
 
Public Const INFO_SHEET As String = "Info"
Public Const PREMIUM_RESULT_SHEET As String = "Premium Review"
Public Const INTERFACE_RESULT_SHEET As String = "Interface Review"
 
'Public KpiWorkbook As Workbook
'Public CurrentWorkbook As Workbook
Public CurrentKpiSheet As Worksheet
 
' DQ Premium / Invoice difference tolerance
' Example
' MART premium: $1,000 AND Invoice premium: $1,006.89 => Not an issue
Public Const TOL As Double = 10
 
Public Const DETAIL_COLOR As Long = 10213316 'RGB(196,215,155)
Public Const PREMIUM_COLOR As Long = 14336204 'RGB(204,192,218)
Public Const INVOICE_COLOR As Long = 15261367 'RGB(183,222,232)
Public Const DOCUMENT_COLOR As Long = 15849925 'RGB(197,217,241)
 
' These arrays capture column numbers, in case the KPI download columns
' change due to an update
Public DETAIL_FIELDS(1 To 13) As Integer
Public PREMIUM_FIELDS(1 To 10) As Integer
Public INVOICE_FIELDS(1 To 10) As Integer
 
Public InterfaceIssueArray() As String
 
Private dataManager As CnDataManager
 
Sub PremiumReviewButton_Click()
   
    'Check for 3 sheets
    Dim detailFound As Boolean, premiumFound As Boolean, invoiceFound As Boolean
    Dim shSheet As Worksheet
    For Each shSheet In ThisWorkbook.Worksheets
       
        Select Case shSheet.Name
           
            Case "Download to Excel"
                detailFound = True
           
            Case "Client_Premium Data-Excel"
                premiumFound = True
           
            Case "Invoice Detail"
                invoiceFound = True
            
            Case Else 'ignore other sheets
           
        End Select
       
    Next shSheet
   
    Dim Answer As VbMsgBoxResult
    If Not (detailFound And premiumFound And invoiceFound) Then
        Answer = MsgBox("WARNING" & vbNewLine & vbNewLine & _
            "DQ Details, DQ Premiums and Invoice Details Required", vbOKOnly, "Missing Sheets")
        Exit Sub
    End If
   
    ActiveSheet.Buttons(1).Visible = False
    Call ReviewPremiums
   
End Sub
 
Sub InterfaceReviewButton_Click()
   
    'Check for 3 sheets
    Dim detailFound As Boolean, interfaceFound As Boolean
    Dim shSheet As Worksheet
    For Each shSheet In ThisWorkbook.Worksheets
       
        Select Case shSheet.Name
           
            Case "Download to Excel"
                detailFound = True
               
            Case "Sheet1"
                interfaceFound = True
               
            Case Else 'ignore other sheets
       
        End Select
       
    Next shSheet
   
    Dim Answer As VbMsgBoxResult
    If Not (detailFound And interfaceFound) Then
        Answer = MsgBox("WARNING" & vbNewLine & vbNewLine & _
            "DQ Details and Document Interface Download Required", vbOKOnly, "Missing Sheets")
        Exit Sub
    End If
   
    ActiveSheet.Buttons(1).Visible = False
    Call ReviewInterface
   
End Sub
 
Private Sub ReviewPremiums()
   
    Dim book As Workbook
    Set book = ThisWorkbook
    Dim Answer As VbMsgBoxResult
   
    'Check if data already present
    If book.Worksheets(PREMIUM_RESULT_SHEET).Range("A20").Value <> "" Then
       
        Answer = MsgBox("WARNING" & vbNewLine & vbNewLine & "Overwrite existing table?", _
                            vbYesNo + vbQuestion, "Totals")
       
        If Answer = vbNo Then
            Exit Sub
        End If
    End If
   
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
   
    Set dataManager = New CnDataManager
   
    Dim newItems As Variant
    newItems = book.Worksheets("Download to Excel").Range("A1").CurrentRegion.Value
    dataManager.AddDqDetails newItems
   
    newItems = book.Worksheets("Client_Premium Data-Excel").Range("A1").CurrentRegion.Value
    dataManager.AddDqPremiums newItems
   
    newItems = book.Worksheets("Invoice Detail").Range("A1").CurrentRegion.Value
    dataManager.AddInvoiceDetails newItems
   
    'Add Company Name and Number to Info Sheet
    If book.Worksheets(INFO_SHEET).Range("C2").Value = "" Then
       
        book.Worksheets(INFO_SHEET).Range("C2").Value = _
            book.Worksheets("Download to Excel").Range("B2").Value
        
        book.Worksheets(INFO_SHEET).Range("C3").Value = _
            book.Worksheets("Download to Excel").Range("A2").Value
       
    End If
   
    If Not dataManager.HasPremiumData() Then
       
        Debug.Print "No Premium Data."
        Exit Sub
       
    End If
   
    Dim allItems As Variant
    allItems = dataManager.GetPremiumData()
   
    ' Format output
    With book.Worksheets(PREMIUM_RESULT_SHEET)
       
        .Range("A20").CurrentRegion.Columns.Delete
       
        .Range("B:C,M:M").NumberFormat = "m/d/yyyy"
        .Range("D:D,N:N").NumberFormat = "@"
        .Range("F:K,E16:E19").NumberFormat = "#,##0.00"
       
        .Range("A20").Resize(1, 21) = Array("Data", "Effective", "Expiry", _
            "Policy Number", "Issue", "TRIA", "DQ Detail", "DQ Premium", "Invoice Prem", _
            "PremFee", "InvFee", "Line Type", "Billing Date", "Invoice #", "Global Carrier", _
            "Local Carrier", "Coverage", "Group", "Line", "Country", "Active")
       
        'Paste data to worksheet
        .Range("A21").Resize(UBound(allItems, 1), UBound(allItems, 2)).Value = allItems
       
        .Range("A20").CurrentRegion.Sort key1:=.Range("M20"), key2:=.Range("N20"), Header:=xlYes
       
        .Range("A20").CurrentRegion.Sort key1:=.Range("D20"), key2:=.Range("A20"), _
            key3:=.Range("B20"), Header:=xlYes
       
        'Create table
        .ListObjects.Add(xlSrcRange, .Range("A20").CurrentRegion, , xlYes).Name = "tblMerge"
       
        'Add Active formula
        .Range("U21").Formula = "=TODAY()<[@Expiry]"
       
        'Add all borders
        .Range("A20").CurrentRegion.Borders.LineStyle = xlContinuous
       
        .Range("F19").FormulaR1C1 = "=SUBTOTAL(109,tblMerge[TRIA])"
        .Range("G19").FormulaR1C1 = "=SUBTOTAL(109,tblMerge[DQ Detail])"
        .Range("H19").FormulaR1C1 = "=SUBTOTAL(109,tblMerge[DQ Premium])"
        .Range("I19").FormulaR1C1 = "=SUBTOTAL(109,tblMerge[Invoice Prem])"
        .Range("J19").FormulaR1C1 = "=SUBTOTAL(109,tblMerge[PremFee])"
        .Range("K19").FormulaR1C1 = "=SUBTOTAL(109,tblMerge[InvFee])"
       
        .Range("E16:E19").HorizontalAlignment = xlLeft
        .Range("E16").Formula = "=F19+G19"
        .Range("E17").Formula = "=H19+J19"
        .Range("E18").Formula = "=I19+K19"
        .Range("E19").Formula = "=ABS(E16-E18)"
        .Range("H18").Formula = "=ABS(F19+G19-H19)"
        .Range("I18").Formula = "=ABS(F19+G19-I19)"
       
        .Range("D15") = "Totals"
        .Range("G18") = "Prem Diff:"
       
        .Range("D16:D19,G18").HorizontalAlignment = xlRight
        .Range("D16") = "DQ Detail:"
        .Range("D17") = "DQ Premium:"
        .Range("D18") = "Invoices:"
        .Range("D19") = "Difference:"
       
        With .Range("D16:E19,G18:I18,F19:K19")
            .Borders.LineStyle = xlContinuous
            With .Interior
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
            End With
        End With
       
        .Range("E16,F19:G19").Interior.Color = DETAIL_COLOR
        .Range("E17,H19,J19").Interior.Color = PREMIUM_COLOR
        .Range("E18,I19,K19").Interior.Color = INVOICE_COLOR
       
        'Color table lines
        Dim rgRow As Range
        For Each rgRow In .Range("tblMerge").Rows
           
            If rgRow.Cells(1, 1) = "DQ Detail" Then
                rgRow.Interior.Color = DETAIL_COLOR
               
            ElseIf rgRow.Cells(1, 1) = "DQ Premium" Then
                rgRow.Interior.Color = PREMIUM_COLOR
               
            ElseIf rgRow.Cells(1, 1) = "Invoice Detail" Then
                rgRow.Interior.Color = INVOICE_COLOR
               
            End If
           
        Next rgRow
       
        .Range("A:C,E:T").EntireColumn.AutoFit
       
        'Add conditional highlighting
        With .Columns("E:E")
           
            'Green
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=""No Issue"""
            With .FormatConditions(.FormatConditions.Count)
                .Font.Color = -16752384
                .Interior.Color = 13561798
            End With
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=""Taxes/Fees"""
            With .FormatConditions(.FormatConditions.Count)
                .Font.Color = -16752384
                .Interior.Color = 13561798
            End With
           
            'Red
            .FormatConditions.Add Type:=xlExpression, Formula1:= _
                "=OR(ISNUMBER(FIND(""vs"",E1)),ISNUMBER(FIND(""Not"",E1)))"
            With .FormatConditions(.FormatConditions.Count)
                .Font.Color = -16383844
                .Interior.Color = 13551615
            End With
           
        End With
       
        'Add slicers
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Data") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Data", Top:=8, Left:=10, _
            Width:=100, Height:=100
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Active") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Active", Top:=116, Left:=10, _
            Width:=100, Height:=77
       
        With book.SlicerCaches(book.SlicerCaches.Count)
            .Slicers(.Slicers.Count).SlicerCache.SortItems = xlSlicerSortDescending
            .SlicerItems("TRUE").Selected = True
            .SlicerItems("FALSE").Selected = False
        End With
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Issue") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Issue", Top:=8, Left:=120, _
            Width:=180, Height:=188
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Line") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Line", Top:=8, Left:=310, _
            Width:=150, Height:=210
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Group") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Group", Top:=8, Left:=475, _
            Width:=190, Height:=210
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Coverage") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Coverage", Top:=8, Left:=675, _
            Width:=200, Height:=210
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Global Carrier") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Global Carrier", Top:=8, Left:=890, _
            Width:=200, Height:=210
       
        book.SlicerCaches.Add2(.ListObjects("tblMerge"), "Local Carrier") _
            .Slicers.Add book.Worksheets(PREMIUM_RESULT_SHEET), Caption:="Local Carrier", Top:=8, Left:=1100, _
            Width:=220, Height:=210
       
        'Tooltips
        With .Range("E16:E19,F19:K19,H18:I18").Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        End With
       
        .Range("E16").Validation.InputMessage = "DQ Details Premiums + TRIA Total"
        
        .Range("E17").Validation.InputMessage = "DQ Premiums Fees + Premiums Total"
       
        .Range("E18").Validation.InputMessage = "Invoice Details Fees + Premiums Total"
       
        .Range("E19").Validation.InputMessage = "Difference Between DQ Detail and Invoice Detail Totals"
       
        .Range("F19").Validation.InputMessage = "TRIA Total"
       
        .Range("G19").Validation.InputMessage = "DQ Details Premium Total"
       
        .Range("H19").Validation.InputMessage = "DQ Premiums Premium Total"
       
        .Range("I19").Validation.InputMessage = "Invoice Details Premium Total"
       
        .Range("J19").Validation.InputMessage = "DQ Premiums Fee Total"
       
        .Range("K19").Validation.InputMessage = "Invoice Details Fee Total"
       
        .Range("H18").Validation.InputMessage = "Difference Between DQ Premiums and DQ Details Premium Totals"
       
        .Range("I18").Validation.InputMessage = "Difference Between Invoice Details and DQ Details Premium Totals"
       
        .Activate
       
        ' .Rows("21:21").Select
        ' ActiveWindow.FreezePanes = True ' Don't freeze for all users
         .Range("A20").Select
       
    End With
   
    Call DistinguishPolicyNumbers
   
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
   
    Debug.Print "Premium Review Complete."
   
End Sub
 
Private Sub ReviewInterface()
   
    Dim book As Workbook
    Set book = ThisWorkbook
    Dim Answer As VbMsgBoxResult
   
    'Check if data already present
    If book.Worksheets(INTERFACE_RESULT_SHEET).Range("A1").Value <> "" Then
       
        Answer = MsgBox("WARNING" & vbNewLine & vbNewLine & "Overwrite existing table?", _
                            vbYesNo + vbQuestion, "Totals")
       
        If Answer = vbNo Then
            Exit Sub
        End If
    End If
   
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
   
    ReDim InterfaceIssueArray(1 To 6)
   
    InterfaceIssueArray(1) = "No Match - Review Needed"
    InterfaceIssueArray(2) = "No Policy Subtype"
    InterfaceIssueArray(3) = "Not PDF"
    InterfaceIssueArray(4) = "Not Acct Ready"
    InterfaceIssueArray(5) = "No Issue - Verify on LINQ"
    InterfaceIssueArray(6) = ""
   
    Dim newItems As Variant
    If dataManager Is Nothing Then
        Set dataManager = New CnDataManager
        newItems = book.Worksheets("Download to Excel").Range("A1").CurrentRegion.Value
        dataManager.AddDqDetails newItems
    End If
   
    newItems = book.Worksheets("Sheet1").Range("A1").CurrentRegion.Value
    dataManager.AddInterfaceDetails newItems
   
    'Add Company Name and Number to Info Sheet
    If book.Worksheets(INFO_SHEET).Range("C2").Value = "" Then
       
        book.Worksheets(INFO_SHEET).Range("C2").Value = _
            book.Worksheets("Download to Excel").Range("B2").Value
       
        book.Worksheets(INFO_SHEET).Range("C3").Value = _
            book.Worksheets("Download to Excel").Range("A2").Value
       
    End If
   
    If Not dataManager.HasInterfaceData() Then
        
        Debug.Print "No Interface Data."
        'Exit Sub
       
    End If
   
    Dim allItems As Variant
    allItems = dataManager.GetInterfaceData()
   
    ' Format output
    With book.Worksheets(INTERFACE_RESULT_SHEET)
       
        .Range("A10").CurrentRegion.Columns.Delete
       
        .Range("B:C,I:I").NumberFormat = "m/d/yyyy"
        .Range("D:D,J:J").NumberFormat = "@"
       
        .Range("A10").Resize(1, 12) = Array("Data", "Effective", "Expiry", "Policy Number", _
            "Issue", "AR", "Subtype", "Format", "Last Modified", "Document Name", "Line", "Active")
       
        'Paste data to worksheet
        .Range("A11").Resize(UBound(allItems, 1), UBound(allItems, 2)).Value = allItems
       
        .Range("A10").CurrentRegion.Sort key1:=.Range("D10"), key2:=.Range("B10"), _
            Header:=xlYes
       
        'Create table
        .ListObjects.Add(xlSrcRange, .Range("A10").CurrentRegion, , xlYes).Name = "tblInterface"
       
        'Add Active formula
        .Range("L11").Formula = "=OR(TODAY()<[@Expiry],[@Expiry]="""")"
       
        'Add all borders
        .Range("A10").CurrentRegion.Borders.LineStyle = xlContinuous
       
        'Color table lines
        Dim rgRow As Range
        For Each rgRow In .Range("tblInterface").Rows
           
            If rgRow.Cells(1, 1) = "DQ Detail" Then
                rgRow.Interior.Color = DETAIL_COLOR
               
            ElseIf rgRow.Cells(1, 1) = "Document" Then
                rgRow.Interior.Color = INVOICE_COLOR
               
            End If
           
        Next rgRow
       
        .Range("A:C,E:I").EntireColumn.AutoFit
       
        'Add slicers
        book.SlicerCaches.Add2(.ListObjects("tblInterface"), "Data") _
            .Slicers.Add book.Worksheets(INTERFACE_RESULT_SHEET), Caption:="Data", Top:=8, Left:=10, _
            Width:=100, Height:=77
       
        With book.SlicerCaches(book.SlicerCaches.Count)
            .Slicers(.Slicers.Count).SlicerCache.SortItems = xlSlicerSortDescending
        End With
       
        book.SlicerCaches.Add2(.ListObjects("tblInterface"), "Issue") _
            .Slicers.Add book.Worksheets(INTERFACE_RESULT_SHEET), Caption:="Issue", Top:=8, Left:=230, _
            Width:=320, Height:=100
       
        With book.SlicerCaches(book.SlicerCaches.Count)
            .Slicers(.Slicers.Count).NumberOfColumns = 2
        End With
       
        book.SlicerCaches.Add2(.ListObjects("tblInterface"), "Line") _
            .Slicers.Add book.Worksheets(INTERFACE_RESULT_SHEET), Caption:="Line", Top:=8, Left:=560, _
            Width:=200, Height:=120
       
        .Activate
       
        .Rows("11:11").Select
        ActiveWindow.FreezePanes = True
        .Range("A10").Select
       
        book.SlicerCaches.Add2(.ListObjects("tblInterface"), "Active") _
            .Slicers.Add book.Worksheets(INTERFACE_RESULT_SHEET), Caption:="Active", Top:=8, Left:=120, _
            Width:=100, Height:=77
       
        With book.SlicerCaches(book.SlicerCaches.Count)
            .Slicers(.Slicers.Count).SlicerCache.SortItems = xlSlicerSortDescending
            .SlicerItems("TRUE").Selected = True
            .SlicerItems("FALSE").Selected = False
        End With
        
    End With
   
    Call DistinguishPolicyNumbers
   
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
   
    Debug.Print "Interface Review Complete."
   
End Sub
 
Public Sub DistinguishPolicyNumbers()
   
    If CurrentKpiSheet Is Nothing Then
        Exit Sub
    End If
   
    With CurrentKpiSheet
        If Not COLOR_POLICY_LINES Or _
            (.Name <> PREMIUM_RESULT_SHEET And .Name <> INTERFACE_RESULT_SHEET) Or _
            .ListObjects.Count < 1 Then
           
            Exit Sub
        End If
       
        Debug.Print "DistinguishPolicyNumbers: " & .Name
       
    End With
   
    Dim sPrevPolicy As String
   
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
   
    With CurrentKpiSheet.ListObjects(1).DataBodyRange.SpecialCells(xlCellTypeVisible)
       
        Dim rngArea As Range
        For Each rngArea In .Areas
           
            Dim rngRow As Range
            For Each rngRow In rngArea.Rows
               
                If rngRow.Cells(1, 4).Value <> sPrevPolicy Then
                    rngRow.Borders(xlEdgeTop).Weight = xlThick
                Else
                    rngRow.Borders(xlEdgeTop).Weight = xlThin
                End If
               
                sPrevPolicy = rngRow.Cells(1, 4).Value
               
            Next rngRow
           
        Next rngArea
       
    End With
   
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
   
End Sub
 
' Remove special characters, then leading zeroes
Public Function CleanPolicyNumber(rng As Range) As String
   
    Dim result As String
    Dim ch, bytes() As Byte: bytes = rng.Value
    For Each ch In bytes
        If Chr(ch) Like "[A-Z,a-z0-9]" Then
            result = result & Chr(ch)
        End If
    Next ch
   
    bytes = result
    result = ""
   
    Dim bNonZeroFound As Boolean
    For Each ch In bytes
       
        'Every other byte is 0x00
        If ch <> 0 Then
           
            If Chr(ch) <> "0" Then bNonZeroFound = True
            
            If bNonZeroFound Then result = result & Chr(ch)
           
        End If
       
    Next ch
   
    CleanPolicyNumber = result
   
End Function
 
' Check whether item is member of the Collection
Public Function ContainsKey(colKeys As Collection, key As String) As Boolean
   
    ContainsKey = False
   
    Dim colItem As Variant
    For Each colItem In colKeys
       
        If StrComp(UCase(colItem), UCase(key)) = 0 Then
           
            ContainsKey = True
            Exit Function
           
        End If
       
    Next colItem
   
End Function
 
' If item value starts with 'key', return full value. Otherwise return
' empty String
Public Function GetFullKey(colKeys As Collection, key As String) As String
   
    GetFullKey = ""
   
    Dim colItem As Variant
    For Each colItem In colKeys
       
        If Len(colItem) >= Len(key) Then
           
            If StrComp(Left(UCase(colItem), Len(key)), UCase(key)) = 0 Then
                GetFullKey = colItem
                Exit Function
            End If
           
        End If
       
    Next colItem
   
End Function
 