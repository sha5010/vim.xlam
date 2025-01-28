Sub IncrementCellValue()
    On Error Resume Next
    If IsNumeric(ActiveCell.Value) Then
        Call RepeatRegister("IncrementCellValue")
        ActiveCell.Value = ActiveCell.Value + 1
    End If
    On Error GoTo 0 ' Restore normal error handling
End Sub
        
Sub DecrementCellValue()
    On Error Resume Next
    If IsNumeric(ActiveCell.Value) Then
        Call RepeatRegister("DecrementCellValue")
        ActiveCell.Value = ActiveCell.Value - 1
    End If
    On Error GoTo 0 ' Restore normal error handling
End Sub
