Attribute VB_Name = "Module1"
Function SumByColor(SumRange As Range, SumColor As Range)
Application.Volatile
Dim SumColorValue As Integer
Dim TotalSum As Double
SumColorValue = SumColor.Interior.ColorIndex
Set rCell = SumRange
For Each rCell In SumRange
If rCell.Interior.ColorIndex = SumColorValue Then
TotalSum = TotalSum + rCell.Value
End If
Next rCell
SumByColor = TotalSum
End Function
