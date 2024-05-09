' KpiWorkbook -----------------------------------------------------------------------------------------------
 
Option Explicit
 
'version 2022.04.25
'author: Brendan Horan
 
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Set LINQ.CurrentKpiSheet = Sh
    Debug.Print "Workbook_SheetActivate: " & Sh.Name
End Sub
 
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    Set LINQ.CurrentKpiSheet = Wn.ActiveSheet
    Debug.Print "Workbook_WindowActivate: " & LINQ.CurrentKpiSheet.Name
End Sub
 
Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    Set LINQ.CurrentKpiSheet = Nothing
    Debug.Print "Workbook_WindowDeactivate: " & Wn.ActiveSheet.Name
End Sub
 
' shPremiumReview -------------------------------------------------------------------------------------------
 
Option Explicit
 
'version 2022.04.25
'author: Brendan Horan
 
Private Sub Worksheet_Calculate()
   
    Debug.Print "Worksheet_Calculate"
    Call LINQ.DistinguishPolicyNumbers
   
End Sub
 