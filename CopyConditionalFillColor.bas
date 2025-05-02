Attribute VB_Name = "CopyConditionalFillColor"
Option Explicit

Sub CopyConditionalFillColor()
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim cell As Range
    Dim ws As Worksheet
    Dim form As CopyConditionalFillColorForm
    Dim border As border
    
    ' Show UserForm
    Set form = New CopyConditionalFillColorForm
    form.Show
    
    ' Check if user canceled
    If form.Tag = "Cancel" Then
        Unload form
        Exit Sub
    End If
    
    ' Validate worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(form.ComboBoxSheet.Value)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Selected worksheet '" & form.ComboBoxSheet.Value & "' does not exist.", vbExclamation
        Unload form
        Exit Sub
    End If
    
    ' Validate source range
    On Error Resume Next
    Set sourceRange = ws.Range(form.TextBoxSourceRange.Value)
    On Error GoTo 0
    If sourceRange Is Nothing Then
        MsgBox "Invalid source range: '" & form.TextBoxSourceRange.Value & "'.", vbExclamation
        Unload form
        Exit Sub
    End If
    
    ' Validate target range
    On Error Resume Next
    Set targetRange = ThisWorkbook.ActiveSheet.Range(form.TextBoxTargetRange.Value)
    On Error GoTo 0
    If targetRange Is Nothing Then
        MsgBox "Invalid target range: '" & form.TextBoxTargetRange.Value & "'.", vbExclamation
        Unload form
        Exit Sub
    End If
    
    ' Ensure target range has the same size as source range
    If targetRange.Cells.Count <> sourceRange.Cells.Count Then
        MsgBox "Target range must have the same number of cells as the source range.", vbExclamation
        Unload form
        Exit Sub
    End If
    
    ' Check if at least one property is selected
    If Not (form.CheckBoxFill.Value Or form.CheckBoxFontColor.Value Or _
            form.CheckBoxBold.Value Or form.CheckBoxItalic.Value Or _
            form.CheckBoxBorders.Value Or form.CheckBoxNumberFormat.Value) Then
        MsgBox "Please select at least one property to copy.", vbExclamation
        Unload form
        Exit Sub
    End If
    
    ' Loop through each cell in the source range and copy selected properties
    Application.ScreenUpdating = False ' Improve performance
    For Each cell In sourceRange
        With targetRange.Cells(cell.Row - sourceRange.Row + 1, cell.Column - sourceRange.Column + 1)
            If form.CheckBoxFill.Value Then
                .Interior.Color = cell.DisplayFormat.Interior.Color
            End If
            If form.CheckBoxFontColor.Value Then
                .Font.Color = cell.DisplayFormat.Font.Color
            End If
            If form.CheckBoxBold.Value Then
                .Font.Bold = cell.DisplayFormat.Font.Bold
            End If
            If form.CheckBoxItalic.Value Then
                .Font.Italic = cell.DisplayFormat.Font.Italic
            End If
            If form.CheckBoxBorders.Value Then
                ' Copy border properties for all sides
                Set border = cell.DisplayFormat.Borders(xlLeft)
                .Borders(xlLeft).LineStyle = border.LineStyle
                .Borders(xlLeft).Color = border.Color
                .Borders(xlLeft).Weight = border.Weight
                
                Set border = cell.DisplayFormat.Borders(xlRight)
                .Borders(xlRight).LineStyle = border.LineStyle
                .Borders(xlRight).Color = border.Color
                .Borders(xlRight).Weight = border.Weight
                
                Set border = cell.DisplayFormat.Borders(xlTop)
                .Borders(xlTop).LineStyle = border.LineStyle
                .Borders(xlTop).Color = border.Color
                .Borders(xlTop).Weight = border.Weight
                
                Set border = cell.DisplayFormat.Borders(xlBottom)
                .Borders(xlBottom).LineStyle = border.LineStyle
                .Borders(xlBottom).Color = border.Color
                .Borders(xlBottom).Weight = border.Weight
            End If
            If form.CheckBoxNumberFormat.Value Then
                .NumberFormat = cell.DisplayFormat.NumberFormat
            End If
        End With
    Next cell
    Application.ScreenUpdating = True
    
    MsgBox "Selected properties copied successfully!", vbInformation
    Unload form
End Sub
