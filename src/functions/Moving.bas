Attribute VB_Name = "F_Moving"
Option Explicit
Option Private Module

Function MoveUp()
    Dim r As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row - gVim.Count1
        If r < 1 Then
            r = 1
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If
End Function

Function MoveDown()
    Dim r As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row + gVim.Count1
        If r > ActiveSheet.Rows.Count Then
            r = ActiveSheet.Rows.Count
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If
End Function

Function MoveLeft()
    Dim c As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column - gVim.Count1
        If c < 1 Then
            c = 1
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If
End Function

Function MoveRight()
    Dim c As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column + gVim.Count1
        If c > ActiveSheet.Columns.Count Then
            c = ActiveSheet.Columns.Count
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If
End Function

Private Function ResizeInner(Optional Up As Long = 0, _
                             Optional Down As Long = 0, _
                             Optional Left As Long = 0, _
                             Optional Right As Long = 0)
    On Error GoTo Catch

    Dim r As Long
    Dim c As Long
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim screenTop As Long
    Dim screenBottom As Long
    Dim screenLeft As Long
    Dim screenRight As Long

    Dim actCell As Range
    Dim baseRange As Range

    '値を取得
    r = Selection.Rows.Count
    c = Selection.Columns.Count

    firstRow = Selection.Item(1).Row
    firstColumn = Selection.Item(1).Column
    lastRow = Selection(Selection.Count).Row
    lastColumn = Selection(Selection.Count).Column

    screenTop = ActiveWindow.VisibleRange.Item(1).Row
    screenBottom = ActiveWindow.VisibleRange.Item(ActiveWindow.VisibleRange.Count).Row - 1
    screenLeft = ActiveWindow.VisibleRange.Item(1).Column
    screenRight = ActiveWindow.VisibleRange.Item(ActiveWindow.VisibleRange.Count).Column - 1

    'セル範囲を取得
    Set actCell = ActiveCell
    Set baseRange = Selection

    '飛び越える場合を計算
    If Up < 0 And -Up >= r Then
        Down = -(r + Up) + 1
        Up = 0
        Set baseRange = baseRange.Offset(RowOffset:=r - 1).Resize(RowSize:=1)
    ElseIf Down < 0 And -Down >= r Then
        Up = -(r + Down) + 1
        Down = 0
        Set baseRange = baseRange.Resize(RowSize:=1)
    ElseIf Left < 0 And -Left >= c Then
        Right = -(c + Left) + 1
        Left = 0
        Set baseRange = baseRange.Offset(ColumnOffset:=c - 1).Resize(ColumnSize:=1)
    ElseIf Right < 0 And -Right >= c Then
        Left = -(c + Right) + 1
        Right = 0
        Set baseRange = baseRange.Resize(ColumnSize:=1)
    End If

    '限界を越える場合は抑える
    If Up > 0 And firstRow <= Up Then
        Up = firstRow - 1
    ElseIf Down > 0 And lastRow + Down > ActiveSheet.Rows.Count Then
        Down = ActiveSheet.Rows.Count - lastRow
    ElseIf Left > 0 And firstColumn <= Left Then
        Left = firstColumn - 1
    ElseIf Right > 0 And lastColumn + Right > ActiveSheet.Columns.Count Then
        Right = ActiveSheet.Columns.Count - lastColumn
    End If

    '上方向が変化する場合
    If Up <> 0 Then
        baseRange.Offset(RowOffset:=-Up).Resize(RowSize:=baseRange.Rows.Count + Up).Select
        actCell.Activate

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenTop > firstRow - Up Then
            ActiveWindow.SmallScroll Up:=screenTop - (firstRow - Up)
        ElseIf screenBottom < firstRow - Up Then
            ActiveWindow.SmallScroll Down:=(firstRow - Up) - screenBottom
        End If

    '下方向が変化する場合
    ElseIf Down <> 0 Then
        baseRange.Resize(RowSize:=baseRange.Rows.Count + Down).Select
        actCell.Activate

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenTop > lastRow + Down Then
            ActiveWindow.SmallScroll Up:=screenTop - (lastRow + Down)
        ElseIf screenBottom < lastRow + Down Then
            ActiveWindow.SmallScroll Down:=(lastRow + Down) - screenBottom
        End If

    '左方向が変化する場合
    ElseIf Left <> 0 Then
        baseRange.Offset(ColumnOffset:=-Left).Resize(ColumnSize:=baseRange.Columns.Count + Left).Select
        actCell.Activate

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenLeft > firstColumn - Left Then
            ActiveWindow.SmallScroll ToLeft:=screenLeft - (firstColumn - Left)
        ElseIf screenRight < firstColumn - Left Then
            ActiveWindow.SmallScroll ToRight:=(firstColumn - Left) - screenRight
        End If

    '右方向が変化する場合
    ElseIf Right <> 0 Then
        baseRange.Resize(ColumnSize:=baseRange.Columns.Count + Right).Select
        actCell.Activate

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenLeft > lastColumn + Right Then
            ActiveWindow.SmallScroll ToLeft:=screenLeft - (lastColumn + Right)
        ElseIf screenRight < lastColumn + Right Then
            ActiveWindow.SmallScroll ToRight:=(lastColumn + Right) - screenRight
        End If

    End If

    Set actCell = Nothing
    Set baseRange = Nothing

    Exit Function

Catch:
    Call ErrorHandler("ResizeInner")
End Function

Function MoveUpWithShift()
    Dim r As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(1).Row = ActiveCell.Row Then
            Call ResizeInner(Down:=-gVim.Count1)
        Else
            Call ResizeInner(Up:=gVim.Count1)
        End If
    End If
End Function

Function MoveDownWithShift()
    Dim r As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(Selection.Count).Row = ActiveCell.Row Then
            Call ResizeInner(Up:=-gVim.Count1)
        Else
            Call ResizeInner(Down:=gVim.Count1)
        End If
    End If
End Function

Function MoveLeftWithShift()
    Dim c As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(1).Column = ActiveCell.Column Then
            Call ResizeInner(Right:=-gVim.Count1)
        Else
            Call ResizeInner(Left:=gVim.Count1)
        End If
    End If
End Function

Function MoveRightWithShift()
    Dim c As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(Selection.Count).Column = ActiveCell.Column Then
            Call ResizeInner(Left:=-gVim.Count1)
        Else
            Call ResizeInner(Right:=gVim.Count1)
        End If
    End If
End Function

Function MoveToTopRow(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gVim.Count = 0 Then
            .Cells(1, ActiveCell.Column).Select
        Else
            .Cells(gVim.Count1, ActiveCell.Column).Select
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToTopRow")
End Function

Function MoveToLastRow(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gVim.Count = 0 Then
            .Cells(.UsedRange.Item(.UsedRange.Count).Row, ActiveCell.Column).Select
        Else
            .Cells(gVim.Count1, ActiveCell.Column).Select
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToLastRow")
End Function

Function MoveToNthColumn(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    If gVim.Count1 > ActiveSheet.Columns.Count Then
        gVim.Count1 = ActiveSheet.Columns.Count
    End If

    ActiveSheet.Cells(ActiveCell.Row, gVim.Count1).Select
    Exit Function

Catch:
    Call ErrorHandler("MoveToNthColumn")
End Function

Function MoveToFirstColumn(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, 1).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToFirstColumn")
End Function

Function MoveToLeftEnd(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.Item(1).Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToLeftEnd")
End Function

Function MoveToRightEnd(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.Item(.UsedRange.Count).Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToRightEnd")
End Function

Function MoveToTopOfCurrentRegion(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    Dim targetRow As Long

    With ActiveWorkbook.ActiveSheet
        targetRow = ActiveCell.CurrentRegion.Item(1).Row
        If targetRow = ActiveCell.Row Then
            targetRow = ActiveCell.End(xlUp).Row
        End If

        .Cells(targetRow, ActiveCell.Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToTopOfCurrentRegion")
End Function

Function MoveToBottomOfCurrentRegion(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    Dim targetRow As Long

    With ActiveWorkbook.ActiveSheet
        targetRow = ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row
        If .Cells(targetRow, ActiveCell.Column).MergeArea.Row = ActiveCell.Row Then
            targetRow = ActiveCell.End(xlDown).Row
        End If

        .Cells(targetRow, ActiveCell.Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToBottomOfCurrentRegion")
End Function

Function MoveToA1(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(1, 1).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToA1")
End Function

Function MoveToSpecifiedCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim jumpAddress As String
    jumpAddress = UF_CmdLine.Launch("Jump to: ", "Jump to", False)

    If jumpAddress = CMDLINE_CANCELED Or Len(jumpAddress) = 0 Then
        Exit Function
    End If

    Dim jumpTarget As Range: Set jumpTarget = Nothing
    If RegExpMatch(jumpAddress, "^[0-9]{1,7}$") Then
        Set jumpTarget = ActiveSheet.Cells(CInt(jumpAddress), ActiveCell.Column)

    ElseIf RegExpMatch(jumpAddress, "^[a-z]{1,3}$", isIgnoreCase:=True) Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress & ActiveCell.Row)

    ElseIf RegExpMatch(jumpAddress, "^[a-z]{1,3}[0-9]{1,7}(:[a-z]{1,3}[0-9]{1,7})?$", isIgnoreCase:=True) Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress)

    ElseIf RegExpMatch(jumpAddress, "^[a-z]{1,3}:[a-z]{1,3}$", isIgnoreCase:=True) Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress)

    ElseIf RegExpMatch(jumpAddress, "[0-9]{1,7}:[0-9]{1,7}") Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress)

    End If

    If Not jumpTarget Is Nothing Then
        Call RecordToJumpList
        jumpTarget.Select
        Set jumpTarget = Nothing
    End If
    Exit Function

Catch:
    Call ErrorHandler("MoveToSpecifiedCell")
End Function

Function MoveToSpecifiedRow(Optional ByVal lineNum As String) As Boolean
    On Error GoTo Catch

    ' Set default return value to True (= Waiting for an argument)
    MoveToSpecifiedRow = True

    lineNum = Trim(lineNum)
    If RegExpMatch(lineNum, "^0*[1-9][0-9]{0,9}$") Then
        Dim n As Long: n = CLng(Right(lineNum, 10))

        If n > ActiveSheet.Rows.Count Then
            n = ActiveSheet.Rows.Count
        End If

        Call RecordToJumpList

        ActiveSheet.Cells(CLng(n), ActiveCell.Column).Select
        MoveToSpecifiedRow = False  ' Set return value to False (= Done)
    End If

    Exit Function

Catch:
    Call ErrorHandler("MoveToSpecifiedRow")
End Function
