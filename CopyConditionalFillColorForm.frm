VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyConditionalFillColorForm 
   Caption         =   "Copy Conditional Formatting Properties"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "CopyConditionalFillColorForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CopyConditionalFillColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ComboBoxSheet.AddItem ws.Name
        ComboBoxTargetSheet.AddItem ws.Name
    Next ws
    ComboBoxSheet.Value = ActiveSheet.Name
    ComboBoxTargetSheet.Value = ActiveSheet.Name
    TextBoxSourceRange.Value = "A1"
    TextBoxTargetRange.Value = "A1"
    CheckBoxFill.Value = True
    CheckBoxFontColor.Value = True
    CheckBoxBold.Value = False
    CheckBoxItalic.Value = False
    CheckBoxBorders.Value = False
    CheckBoxNumberFormat.Value = False
End Sub

Private Sub ButtonOK_Click()
    On Error Resume Next
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Validate worksheet selections
    Set sourceWs = ThisWorkbook.Worksheets(ComboBoxSheet.Value)
    Set targetWs = ThisWorkbook.Worksheets(ComboBoxTargetSheet.Value)
    If sourceWs Is Nothing Or targetWs Is Nothing Then
        MsgBox "Please select valid source and target worksheets.", vbExclamation
        Exit Sub
    End If
    
    ' Validate range inputs
    Set sourceRange = sourceWs.Range(TextBoxSourceRange.Value)
    Set targetRange = targetWs.Range(TextBoxTargetRange.Value)
    If sourceRange Is Nothing Or targetRange Is Nothing Then
        MsgBox "Please enter valid source and target ranges.", vbExclamation
        Exit Sub
    End If
    
    ' Validate range sizes
    If sourceRange.Cells.Count <> targetRange.Cells.Count Then
        MsgBox "Source and target ranges must have the same number of cells.", vbExclamation
        Exit Sub
    End If
    
    ' Validate at least one property is selected
    If Not (CheckBoxFill.Value Or CheckBoxFontColor.Value Or CheckBoxBold.Value Or _
            CheckBoxItalic.Value Or CheckBoxBorders.Value Or CheckBoxNumberFormat.Value) Then
        MsgBox "Please select at least one property to copy.", vbExclamation
        Exit Sub
    End If
    
    Me.Tag = "OK"
    Me.Hide
    On Error GoTo 0
End Sub

Private Sub ButtonCancel_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub ButtonSelectSource_Click()
    On Error Resume Next
    Dim rng As Range
    Set rng = Application.InputBox("Select Source Range:", "Select Range", TextBoxSourceRange.Value, Type:=8)
    If Not rng Is Nothing Then
        If rng.Parent.Name = ComboBoxSheet.Value Then
            TextBoxSourceRange.Value = rng.Address
        Else
            MsgBox "Selected range must be on the source worksheet (" & ComboBoxSheet.Value & ").", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub ButtonSelectTarget_Click()
    On Error Resume Next
    Dim rng As Range
    Set rng = Application.InputBox("Select Target Range:", "Select Range", TextBoxTargetRange.Value, Type:=8)
    If Not rng Is Nothing Then
        If rng.Parent.Name = ComboBoxTargetSheet.Value Then
            TextBoxTargetRange.Value = rng.Address
        Else
            MsgBox "Selected range must be on the target worksheet (" & ComboBoxTargetSheet.Value & ").", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub

