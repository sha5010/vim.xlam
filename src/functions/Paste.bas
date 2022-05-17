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

Function paste_CtrlV()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyV, 0, 0, 0
    keybd_event vbKeyV, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function pasteRows()
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

        Call keyupControlKeys
        Call releaseShiftKeys

        keybd_event vbKeyControl, 0, 0, 0
        keybd_event vbKeyAdd, 0, 0, 0
        keybd_event vbKeyAdd, 0, KEYUP, 0
        keybd_event vbKeyControl, 0, KEYUP, 0

        Call unkeyupControlKeys
    End With

    gLastYanked.Copy
End Function

Function pasteColumns()
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

        Call keyupControlKeys
        Call releaseShiftKeys

        keybd_event vbKeyControl, 0, 0, 0
        keybd_event vbKeyAdd, 0, 0, 0
        keybd_event vbKeyAdd, 0, KEYUP, 0
        keybd_event vbKeyControl, 0, KEYUP, 0

        Call unkeyupControlKeys
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

    Call keyupControlKeys
    Call releaseShiftKeys

    If Application.CutCopyMode > 0 Then 'Cells
        keybd_event vbKeyMenu, 0, 0, 0
        keybd_event vbKeyH, 0, 0, 0
        keybd_event vbKeyH, 0, KEYUP, 0
        keybd_event vbKeyMenu, 0, KEYUP, 0
        keybd_event vbKeyV, 0, 0, 0
        keybd_event vbKeyV, 0, KEYUP, 0
        keybd_event vbKeyT, 0, 0, 0
        keybd_event vbKeyT, 0, KEYUP, 0

    ElseIf cbType = xlClipboardFormatText Then
        keybd_event vbKeyControl, 0, 0, 0
        keybd_event vbKeyV, 0, 0, 0
        keybd_event vbKeyV, 0, KEYUP, 0
        keybd_event vbKeyControl, 0, KEYUP, 0

    Else
        Call debugPrint("Unknown ClipboardType: " & cbType, "pasteValue")
    End If

    Call unkeyupControlKeys
End Function

Function pasteSpecial()
    If Application.ClipboardFormats(1) = -1 Then
        Call setStatusBarTemporarily("クリップボードが空です。", 2)
    Else
        Application.Dialogs(xlDialogPasteSpecial).Show
    End If
End Function
