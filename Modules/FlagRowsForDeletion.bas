Attribute VB_Name = "FlagRowsForDeletion"
Sub FlagRowsForDeletion()
    ' Purpose: Flags rows in a helper column based on blocks of identical values in a specified column.
    ' For "Prelim" rows, marks "Delete" if there's a "Valid" in the same block, else "Keep".
    ' Marks "Valid" rows as "Keep". Marks rows with empty or invalid inputs as "Invalid".
    ' Case-insensitive comparison for status text.
    ' Allows user-defined row range for processing.
    ' Block column can contain numbers or text; only empty cells are invalid.

    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim blockStart As Long
    Dim blockValue As Variant
    Dim hasValid As Boolean
    Dim isValidBlock As Boolean
    Dim isValidStatus As Boolean
    
    ' --- Customizable Variables ---
    ' Column containing the block identifiers (e.g., numbers like 1001 or text like "Group1")
    Const BlockColumn As String = "A"
    ' Column containing the status text (e.g., "Prelim" or "Valid")
    Const StatusColumn As String = "C"
    ' Column where the output (Delete/Keep/Invalid) will be written
    Const OutputColumn As String = "D"
    ' Text to check for "Prelim" status (case-insensitive)
    Const PrelimText As String = "preliminary"
    ' Text to check for "Valid" status (case-insensitive)
    Const ValidText As String = "validated"
    ' Output text for rows to be deleted
    Const DeleteOutput As String = "Delete"
    ' Output text for rows to be kept
    Const KeepOutput As String = "Keep"
    ' Output text for rows with empty or invalid inputs
    Const InvalidOutput As String = "Invalid"
    ' First row of data to process (e.g., 2 if header is in row 1)
    Const StartRow As Long = 3
    ' Last row of data to process (e.g., 4500 for rows 2 to 4500)
    Const EndRow As Long = 3433
    ' Header text for the output column
    Const OutputHeader As String = "Delete Check"
    ' -----------------------------

    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Validate row range
    If StartRow < 1 Or EndRow < StartRow Then
        MsgBox "Invalid row range. Ensure StartRow and EndRow are valid and EndRow >= StartRow.", vbCritical
        Exit Sub
    End If
    
    ' Validate column inputs
    If Len(BlockColumn) = 0 Or Len(StatusColumn) = 0 Or Len(OutputColumn) = 0 Then
        MsgBox "Invalid column settings. Ensure BlockColumn, StatusColumn, and OutputColumn are specified.", vbCritical
        Exit Sub
    End If
    
    ' Validate text inputs
    If Len(PrelimText) = 0 Or Len(ValidText) = 0 Or Len(DeleteOutput) = 0 Or Len(KeepOutput) = 0 Or Len(InvalidOutput) = 0 Then
        MsgBox "Invalid text settings. Ensure PrelimText, ValidText, DeleteOutput, KeepOutput, and InvalidOutput are specified.", vbCritical
        Exit Sub
    End If
    
    ' Add header for the output column
    ws.Range(OutputColumn & "1").Value = OutputHeader
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Initialize block start
    blockStart = StartRow
    
    ' Loop through rows to identify and process blocks
    For i = StartRow To EndRow + 1
        ' Check for end of block, end of data, or end of specified range
        If i > EndRow Or ws.Cells(i, BlockColumn).Value <> ws.Cells(blockStart, BlockColumn).Value Then
            ' Process the current block
            hasValid = False
            isValidBlock = True
            
            ' Check if the block is valid (all rows in block have non-empty values)
            For j = blockStart To i - 1
                If IsEmpty(ws.Cells(j, BlockColumn)) Then
                    isValidBlock = False
                    Exit For
                End If
            Next j
            
            ' Process the block if valid
            If isValidBlock Then
                ' Check if the block has any "Valid" rows
                For j = blockStart To i - 1
                    If UCase(ws.Cells(j, StatusColumn).Value) = UCase(ValidText) Then
                        hasValid = True
                        Exit For
                    End If
                Next j
                
                ' Flag rows in the block
                For j = blockStart To i - 1
                    ' Check if status is valid (not empty and matches Prelim or Valid)
                    isValidStatus = Not IsEmpty(ws.Cells(j, StatusColumn)) And _
                                   (UCase(ws.Cells(j, StatusColumn).Value) = UCase(PrelimText) Or _
                                    UCase(ws.Cells(j, StatusColumn).Value) = UCase(ValidText))
                    
                    If Not isValidStatus Then
                        ' Mark rows with empty or invalid status as Invalid
                        ws.Cells(j, OutputColumn).Value = InvalidOutput
                    ElseIf UCase(ws.Cells(j, StatusColumn).Value) = UCase(PrelimText) Then
                        ' Mark Prelim rows based on presence of Valid
                        ws.Cells(j, OutputColumn).Value = IIf(hasValid, DeleteOutput, KeepOutput)
                    ElseIf UCase(ws.Cells(j, StatusColumn).Value) = UCase(ValidText) Then
                        ' Mark Valid rows as Keep
                        ws.Cells(j, OutputColumn).Value = KeepOutput
                    End If
                Next j
            Else
                ' Mark all rows in an invalid block as Invalid
                For j = blockStart To i - 1
                    ws.Cells(j, OutputColumn).Value = InvalidOutput
                Next j
            End If
            
            ' Move to the next block
            blockStart = i
        End If
    Next i
    
    ' Restore screen updating
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Notify user
    MsgBox "Rows have been flagged in column " & OutputColumn & " for rows " & StartRow & " to " & EndRow & ". " & _
           "Filter for '" & DeleteOutput & "' to delete rows. Check '" & InvalidOutput & "' for errors.", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub
    
