Attribute VB_Name = "F_Moving"
Option Explicit
Option Private Module

Function moveUp()
    Dim r As Long
    If gCount = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row - gCount
        If r < 1 Then
            r = 1
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If
End Function

Function moveDown()
    Dim r As Long
    If gCount = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row + gCount
        If r > ActiveSheet.Rows.Count Then
            r = ActiveSheet.Rows.Count
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If
End Function

Function moveLeft()
    Dim c As Long
    If gCount = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column - gCount
        If c < 1 Then
            c = 1
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If
End Function

Function moveRight()
    Dim c As Long
    If gCount = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column + gCount
        If c > ActiveSheet.Columns.Count Then
            c = ActiveSheet.Columns.Count
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If
End Function

Private Function resizeAPI(Optional Up As Long = 0, _
                           Optional Down As Long = 0, _
                           Optional Left As Long = 0, _
                           Optional Right As Long = 0)
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

End Function

Function moveUpWithShift()
    Dim r As Long
    If gCount = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(1).Row = ActiveCell.Row Then
            Call resizeAPI(Down:=-gCount)
        Else
            Call resizeAPI(Up:=gCount)
        End If
    End If
End Function

Function moveDownWithShift()
    Dim r As Long
    If gCount = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(Selection.Count).Row = ActiveCell.Row Then
            Call resizeAPI(Up:=-gCount)
        Else
            Call resizeAPI(Down:=gCount)
        End If
    End If
End Function

Function moveLeftWithShift()
    Dim c As Long
    If gCount = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(1).Column = ActiveCell.Column Then
            Call resizeAPI(Right:=-gCount)
        Else
            Call resizeAPI(Left:=gCount)
        End If
    End If
End Function

Function moveRightWithShift()
    Dim c As Long
    If gCount = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.Item(Selection.Count).Column = ActiveCell.Column Then
            Call resizeAPI(Left:=-gCount)
        Else
            Call resizeAPI(Right:=gCount)
        End If
    End If
End Function

Function moveToTopRow()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gCount = 1 Then
            .Cells(1, ActiveCell.Column).Select
        Else
            .Cells(gCount, ActiveCell.Column).Select
        End If
    End With
End Function

Function moveToLastRow()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gCount = 1 Then
            .Cells(.UsedRange.Item(.UsedRange.Count).Row, ActiveCell.Column).Select
        Else
            .Cells(gCount, ActiveCell.Column).Select
        End If
    End With
End Function

Function moveToFirstColumn()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, 1).Select
    End With
End Function

Function moveToLeftEnd()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.Item(1).Column).Select
    End With
End Function

Function moveToRightEnd()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.Item(.UsedRange.Count).Column).Select
    End With
End Function

Function moveToTopOfCurrentRegion()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.CurrentRegion.Item(1).Row, ActiveCell.Column).Select
    End With
End Function

Function moveToBottomOfCurrentRegion()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row, ActiveCell.Column).Select
    End With
End Function

Function moveToA1()
    Call recordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(1, 1).Select
    End With
End Function

Function moveToSpecifiedCell(ByVal Address As String) As Boolean
    Call recordToJumpList

    On Error GoTo Catch
    Address = Trim(Address)

    If reMatch(Address, "^[0-9]{1,7}$") Then
        ActiveSheet.Cells(CInt(Address), ActiveCell.Column).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "^[a-z]{1,3}$", IgnoreCase:=True) Then
        ActiveSheet.Range(Address & ActiveCell.Row).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "^[a-z]{1,3}[0-9]{1,7}(:[a-z]{1,3}[0-9]{1,7})?$", IgnoreCase:=True) Then
        ActiveSheet.Range(Address).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "^[a-z]{1,3}:[a-z]{1,3}$", IgnoreCase:=True) Then
        ActiveSheet.Range(Address).Select
        moveToSpecifiedCell = True

    ElseIf reMatch(Address, "[0-9]{1,7}:[0-9]{1,7}") Then
        ActiveSheet.Range(Address).Select
        moveToSpecifiedCell = True

    End If

Catch:
End Function

Function moveToSpecifiedRow(ByVal n As String) As Boolean
    On Error GoTo Catch
    n = Trim(n)
    If reMatch(n, "^[0-9]{1,7}$") Then
        If CLng(n) > ActiveSheet.Rows.Count Then
            Exit Function
        End If

        Call recordToJumpList

        ActiveSheet.Cells(CLng(n), ActiveCell.Column).Select
        moveToSpecifiedRow = True
    End If

Catch:
End Function
