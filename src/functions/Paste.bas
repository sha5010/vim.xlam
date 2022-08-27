Attribute VB_Name = "F_Paste"
Option Explicit

Function pasteSmart()
    Call repeatRegister("pasteSmart")

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

        .Range(.Rows(startRow), .Rows(endRow)).Select

        Call keystroke(True, Ctrl_ + NumpadAdd_)
    End With

    If Application.CutCopyMode = xlCopy Then
        gLastYanked.Copy
    End If
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

        .Range(.Columns(startColumn), .Columns(endColumn)).Select

        Call keystroke(True, Ctrl_ + NumpadAdd_)
    End With

    If Application.CutCopyMode = xlCopy Then
        gLastYanked.Copy
    End If
End Function

Function pasteValue()
    Call repeatRegister("pasteValue")

    Dim cb As Variant
    Dim cbType As Integer

    cb = Application.ClipboardFormats

    If cb(1) = -1 Then
        Exit Function
    End If
    cbType = cb(2)

    If Application.CutCopyMode > 0 Then 'Cells
        Call keystroke(True, Alt_ + H_, V_, V_)

    Else
        Select Case cbType
            Case xlClipboardFormatText
                Call keystroke(True, Ctrl_ + V_)
            Case xlClipboardFormatRTF
                Call keystroke(True, Alt_ + H_, V_, T_)
            Case xlHtml
                Call keystroke(True, Alt_ + H_, V_, S_, End_, Enter_)
            Case Else
                Call debugPrint("Unknown ClipboardType: " & cbType, "pasteValue")
        End Select
    End If
End Function

Function pasteSpecial()
    If Application.ClipboardFormats(1) = -1 Then
        Call setStatusBarTemporarily("クリップボードが空です。", 2)
    Else
        On Error Resume Next
        Application.Dialogs(xlDialogPasteSpecial).Show
    End If
End Function
