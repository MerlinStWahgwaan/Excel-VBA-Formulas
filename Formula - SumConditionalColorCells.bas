Attribute VB_Name = "Module2"
Function SumConditionColorCells(CellsRange As Range, ColorRng As Range)
Dim Bambo As Boolean
Dim dbw As String
Dim CFCELL As Range
Dim CF1 As Single
Dim CF2 As Double
Dim CF3 As Long
Application.Volatile
Bambo = False
For CF1 = 1 To CellsRange.FormatConditions.Count
If CellsRange.FormatConditions(CF1).Interior.ColorIndex = ColorRng.Interior.ColorIndex Then
Bambo = True
Exit For
End If
Next CF1
CF2 = 0
CF3 = 0
If Bambo = True Then
For Each CFCELL In CellsRange
dbw = CFCELL.FormatConditions(CF1).Formula1
dbw = Application.ConvertFormula(dbw, xlA1, xlR1C1)
dbw = Application.ConvertFormula(dbw, xlR1C1, xlA1, , ActiveCell.Resize(CellsRange.Rows.Count, CellsRange.Columns.Count).Cells(CF3 + 1))
If Evaluate(dbw) = True Then CF2 = CF2 + CFCELL.Value
CF3 = CF3 + 1
Next CFCELL
Else
SumConditionColorCells = "NO-COLOR"
Exit Function
End If
SumConditionColorCells = CF2
End Function
