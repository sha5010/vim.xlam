Attribute VB_Name = "F_Paste"
Option Explicit

Function pasteSmart()
    If Application.CutCopyMode = 0 Then 'Empty
        Set gLastYanked = Nothing
    End If

    If gLastYanked Is Nothing Then
        Call paste_CtrlV
        Exit Function
    End If

    If gLastYanked.Rows.Count = gLastYanked.Parent.Rows.Count Then
        Call pasteColumns
    ElseIf gLastYanked.Columns.Count = gLastYanked.Parent.Columns.Count Then
        Call pasteRows
    Else
        Call paste_CtrlV
    End If
End Function

Private Function paste_CtrlV()
    Call keystroke(True, Ctrl_ + V_)
End Function

Private Function pasteRows()
    Dim yankedRows As Long
    Dim startRow As Long
    Dim endRow As Long

    yankedRows = gLastYanked.Rows.Count
    startRow = ActiveCell.Row
    endRow = startRow + yankedRows * gCount - 1

    With ActiveSheet
        If endRow > .Rows.Count Then
            endRow = startRow + WorksheetFunction.RoundDown((.Rows.Count + 1) / yankedRows, 0) - 1
        End If

        If gCount > 1 Then
            .Range(.Rows(startRow), .Rows(endRow)).Select
        End If

        Call keystroke(True, Ctrl_ + NumpadAdd_)
    End With

    gLastYanked.Copy
End Function

Private Function pasteColumns()
    Dim yankedColumns As Long
    Dim startColumn As Long
    Dim endColumn As Long

    yankedColumns = gLastYanked.Columns.Count
    startColumn = ActiveCell.Column
    endColumn = startColumn + yankedColumns * gCount - 1

    With ActiveSheet
        If endColumn > .Columns.Count Then
            endColumn = startColumn + WorksheetFunction.RoundDown((.Columns.Count + 1) / yankedColumns, 0) - 1
        End If

        If gCount > 1 Then
            .Range(.Columns(startColumn), .Columns(endColumn)).Select
        End If

        Call keystroke(True, Ctrl_ + NumpadAdd_)
    End With

    gLastYanked.Copy
End Function

Function pasteValue()
    Dim cb As Variant
    Dim cbType As Integer

    cb = Application.ClipboardFormats

    If cb(1) = -1 Then
        Exit Function
    End If
    cbType = cb(2)

    If Application.CutCopyMode > 0 Then 'Cells
        Call keystroke(True, Alt_ + H_, V_ + T_)

    ElseIf cbType = xlClipboardFormatText Then
        Call keystroke(True, Ctrl_ + V_)

    Else
        Call debugPrint("Unknown ClipboardType: " & cbType, "pasteValue")
    End If
End Function

Function pasteSpecial()
    If Application.ClipboardFormats(1) = -1 Then
        Call setStatusBarTemporarily("クリップボードが空です。", 2)
    Else
        Application.Dialogs(xlDialogPasteSpecial).Show
    End If
End Function
