Attribute VB_Name = "ExportFormCntrols"
Option Explicit

Sub ExportUserFormControlsToTextFile()
    Dim uf As Object
    Dim ctrl As Object
    Dim filePath As String
    Dim fileNum As Integer
    Dim controlInfo As String
    
    ' Set the output file path (modify as needed)
    filePath = ThisWorkbook.Path & "\FormControlsList.txt"
    
    ' Create or reference the UserForm
    Set uf = ThisWorkbook.VBProject.VBComponents("CopyConditionalFillColorForm").Designer
    
    ' Open the text file for writing
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write header
    Print #fileNum, "Control List for CopyConditionalFillColorForm"
    Print #fileNum, "Generated on: " & Now
    Print #fileNum, String(50, "-")
    
    ' Loop through all controls in the UserForm
    For Each ctrl In uf.Controls
        controlInfo = ""
        controlInfo = controlInfo & "Control Type: " & TypeName(ctrl) & vbCrLf
        controlInfo = controlInfo & "Name: " & ctrl.Name & vbCrLf
        
        ' Add Caption if applicable (e.g., for Labels, CommandButtons, CheckBoxes)
        On Error Resume Next
        controlInfo = controlInfo & "Caption: " & ctrl.Caption & vbCrLf
        On Error GoTo 0
        
        ' Add position and size properties
        controlInfo = controlInfo & "Height: " & ctrl.Height & vbCrLf
        controlInfo = controlInfo & "Left: " & ctrl.Left & vbCrLf
        controlInfo = controlInfo & "Top: " & ctrl.Top & vbCrLf
        controlInfo = controlInfo & "Width: " & ctrl.Width & vbCrLf
        controlInfo = controlInfo & String(50, "-")
        
        ' Write control info to file
        Print #fileNum, controlInfo
    Next ctrl
    
    ' Close the file
    Close #fileNum
    
    ' Notify user
    MsgBox "Control list exported to: " & filePath, vbInformation
End Sub
