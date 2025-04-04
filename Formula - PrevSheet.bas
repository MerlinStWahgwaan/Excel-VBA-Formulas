Attribute VB_Name = "Module1"
Function PrevSheet(rng As Range) As Range
    Application.Volatile
    Dim ws As Worksheet
    Set ws = rng.Parent
    If ws.Index > 1 Then
        Set PrevSheet = ws.Parent.Worksheets(ws.Index - 1).Range(rng.Address)
    Else
        Set PrevSheet = ws.Range(rng.Address) ' Fallback to current sheet if it's the first one
    End If
End Function
