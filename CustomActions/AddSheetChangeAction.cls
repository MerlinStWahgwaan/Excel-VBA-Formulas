VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Application.CalculateFull ' Recalculate all formulas when switching sheets
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.CalculateFull ' Recalculate all formulas when any cell changes
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    Application.CalculateFull ' Recalculate all formulas when leaving a sheet
End Sub
