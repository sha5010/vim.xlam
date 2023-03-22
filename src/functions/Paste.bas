Attribute VB_Name = "F_Paste"
Option Explicit

Function pasteSmart(Optional ByVal PasteDirection As XlSearchDirection = xlNext)
    On Error GoTo Catch

    Call repeatRegister("pasteSmart")
    Call stopVisualMode

    If Application.CutCopyMode = 0 Then 'Empty
        Set gLastYanked = Nothing
    End If

    If gLastYanked Is Nothing Then
        Call paste_CtrlV
        Exit Function
    End If

    If gLastYanked.Rows.Count = gLastYanked.Parent.Rows.Count Then
        Call pasteColumns(PasteDirection)
    ElseIf gLastYanked.Columns.Count = gLastYanked.Parent.Columns.Count Then
        Call pasteRows(PasteDirection)
    Else
        Call paste_CtrlV
    End If
    Exit Function

Catch:
    Call errorHandler("pasteSmart")
End Function

Private Function paste_CtrlV()
    Call keystroke(True, Ctrl_ + V_)
End Function

Private Function pasteRows(ByVal PasteDirection As XlSearchDirection)
    On Error GoTo Catch

    Dim yankedRows As Long
    Dim startRow As Long
    Dim endRow As Long

    yankedRows = gLastYanked.Rows.Count
    startRow = ActiveCell.Row + IIf(PasteDirection = xlNext, 1, 0)
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
    Exit Function

Catch:
    Call errorHandler("pasteRows")
End Function

Private Function pasteColumns(ByVal PasteDirection As XlSearchDirection)
    On Error GoTo Catch

    Dim yankedColumns As Long
    Dim startColumn As Long
    Dim endColumn As Long

    yankedColumns = gLastYanked.Columns.Count
    startColumn = ActiveCell.Column + IIf(PasteDirection = xlNext, 1, 0)
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
    Exit Function

Catch:
    Call errorHandler("pasteColumns")
End Function

Function pasteValue()
    On Error GoTo Catch

    Call repeatRegister("pasteValue")
    Call stopVisualMode

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
    Exit Function

Catch:
    Call errorHandler("pasteValue")
End Function

Function pasteSpecial()
    On Error GoTo Catch

    Call stopVisualMode

    If Application.ClipboardFormats(1) = -1 Then
        Call setStatusBarTemporarily("クリップボードが空です。", 2)
    Else
        On Error Resume Next
        Application.Dialogs(xlDialogPasteSpecial).Show
    End If
    Exit Function

Catch:
    Call errorHandler("pasteSpecial")
End Function
