Attribute VB_Name = "F_Scrolling"
Option Explicit
Option Private Module

Private Enum rowSearchMode
    modeTop = -1
    modeMiddle = 0
    modeBottom = 1
End Enum

Private Enum columnSearchMode
    modeLeft = -1
    modeCenter = 0
    modeRight = 1
End Enum

Private Function activateCellInVisibleRange()
    On Error GoTo Catch

    Dim targetRow As Long
    Dim targetColumn As Long
    Dim visibleTop As Long, visibleBottom As Long
    Dim visibleLeft As Long, visibleRight As Long

    targetRow = ActiveCell.Row
    targetColumn = ActiveCell.Column

    With ActiveWindow.VisibleRange
        visibleTop = .Item(1).Row
        visibleBottom = pointToRow(.Item(.Count).Top - 1, xlNone)
        visibleLeft = .Item(1).Column
        visibleRight = pointToColumn(.Item(.Count).Left - 1, xlNone)
    End With

    If targetRow < visibleTop Then
        targetRow = visibleTop
    ElseIf targetRow > visibleBottom Then
        targetRow = visibleBottom
    End If

    If targetColumn < visibleLeft Then
        targetColumn = visibleLeft
    ElseIf targetColumn > visibleRight Then
        targetColumn = visibleRight
    End If

    If TypeName(Selection) = "Range" Then
        If ActiveCell.Row <> targetRow Or ActiveCell.Column <> targetColumn Then
            Cells(targetRow, targetColumn).Activate
        End If
    End If
    Exit Function

Catch:
    Call errorHandler("activateCellInVisibleRange")
End Function

Function scrollUpHalf()
    On Error GoTo Catch

    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    If gCount > 1 Then
        ActiveWindow.LargeScroll Up:=gCount ¥ 2
    End If

    If (gCount And 1) = 1 Then
        topRowVisible = ActiveWindow.VisibleRange.Row

        scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
        targetRow = topRowVisible - scrollWidth

        If targetRow < 1 Then
            targetRow = 1
        End If

        ActiveWindow.SmallScroll Up:=scrollWidth
    End If

    Call activateCellInVisibleRange
    Exit Function

Catch:
    Call errorHandler("scrollUpHalf")
End Function

Function scrollDownHalf()
    On Error GoTo Catch

    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    If gCount > 1 Then
        ActiveWindow.LargeScroll Down:=gCount ¥ 2
    End If

    If (gCount And 1) = 1 Then
        topRowVisible = ActiveWindow.VisibleRange.Row

        scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
        targetRow = topRowVisible + scrollWidth

        If targetRow > ActiveSheet.Rows.Count Then
            targetRow = ActiveSheet.Rows.Count
        End If

        ActiveWindow.SmallScroll Down:=scrollWidth
    End If

    Call activateCellInVisibleRange
    Exit Function

Catch:
    Call errorHandler("scrollDownHalf")
End Function

Function scrollUp()
    ActiveWindow.LargeScroll Up:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollDown()
    ActiveWindow.LargeScroll Down:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollLeft()
    ActiveWindow.LargeScroll ToLeft:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollRight()
    ActiveWindow.LargeScroll ToRight:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollUp1Row()
    ActiveWindow.SmallScroll Up:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollDown1Row()
    ActiveWindow.SmallScroll Down:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollLeft1Column()
    ActiveWindow.SmallScroll ToLeft:=gCount
    Call activateCellInVisibleRange
End Function

Function scrollRight1Column()
    ActiveWindow.SmallScroll ToRight:=gCount
    Call activateCellInVisibleRange
End Function

Private Function pointToRow(ByVal point As Double, ByVal searchMode As rowSearchMode) As Long
    On Error GoTo Catch

    Dim avg As Double
    Dim pred As Long
    Dim diff As Double
    Dim predTop As Double
    Dim i As Integer
    Dim l As Long
    Dim m As Long
    Dim h As Long
    Dim tmp As Long

    '範囲外のケースを省く
    If point > Rows(Rows.Count).Top Then
        pointToRow = Rows.Count
        Exit Function
    ElseIf point <= 0 Then
        pointToRow = 1
        Exit Function
    End If

    '見えている範囲の高さで平均を算出
    avg = ActiveWindow.VisibleRange.Height / ActiveWindow.VisibleRange.Rows.Count

    '平均から行を推測
    pred = CLng(point / avg) + 1
    If pred > Rows.Count Then
        pred = Rows.Count
    ElseIf pred < 1 Then
        pred = 1
    End If
    predTop = Rows(pred).Top

    '差分を取得
    diff = point - predTop

    '範囲を広げながら行が含まれる範囲を特定
    i = 0
    l = pred
    h = pred
    Do Until diff = 0
        tmp = CLng(diff / avg + 0.5) * 2 ^ i
        If tmp = 0 Then
            tmp = Sgn(diff) * 2 ^ i
        End If

        If tmp > Rows.Count Then
            tmp = Rows.Count
        Else
            tmp = pred + tmp
        End If

        If diff < 0 Then
            h = l
            If tmp < 1 Then
                l = 1
            Else
                l = tmp
            End If
        Else
            l = h
            If tmp > Rows.Count Then
                h = Rows.Count
            Else
                h = tmp
            End If
        End If

        If Rows(l).Top <= point And point < Rows(h).Top Then
            Exit Do
        End If

        i = i + 1
    Loop

    '二分探索で行を特定
    Do
        m = Round(l + (h - l) / 2 - 0.25)
        If h - l < 2 Then
            Exit Do
        End If

        predTop = Rows(m).Top
        If point < predTop Then
            h = m
        Else
            l = m
        End If
    Loop

    'モードに応じて処理分岐
    Select Case searchMode
        '中央寄せ
        Case modeMiddle
            '1行追加することで差分の絶対値が近くなる方を選ぶ
            If (point - Rows(m).Top) >= Rows(m).Height / 2 Then
                pointToRow = m + 1
            Else
                pointToRow = m
            End If

        '上寄せ
        Case modeTop
            'ピッタリでないなら1行追加
            If point > Rows(m).Top Then
                pointToRow = m + 1
            Else
                pointToRow = m
            End If

        '下寄せ
        Case modeBottom
            'SCROLL_OFFSET に収まらない範囲なら1行追加
            If point - SCROLL_OFFSET > Rows(m).Top Then
                pointToRow = m + 1
            Else
                pointToRow = m
            End If

        '例外
        Case Else
            pointToRow = m

    End Select
    Exit Function

Catch:
    Call errorHandler("pointToRow")
End Function

Private Function pointToColumn(ByVal point As Double, ByVal searchMode As columnSearchMode) As Long
    On Error GoTo Catch

    Dim avg As Double
    Dim pred As Long
    Dim diff As Double
    Dim predLeft As Double
    Dim i As Integer
    Dim l As Long
    Dim m As Long
    Dim h As Long
    Dim tmp As Long

    '範囲外のケースを省く
    If point > Columns(Columns.Count).Left Then
        pointToColumn = Columns.Count
        Exit Function
    ElseIf point <= 0 Then
        pointToColumn = 1
        Exit Function
    End If

    '見えている範囲の幅で平均を算出
    avg = ActiveWindow.VisibleRange.Width / ActiveWindow.VisibleRange.Columns.Count

    '平均から列を推測
    pred = CLng(point / avg) + 1
    If pred > Columns.Count Then
        pred = Columns.Count
    ElseIf pred < 1 Then
        pred = 1
    End If
    predLeft = Columns(pred).Left

    '差分を取得
    diff = point - predLeft

    '範囲を広げながら列が含まれる範囲を特定
    i = 0
    l = pred
    h = pred
    Do Until diff = 0
        tmp = CLng(diff / avg + 0.5) * 2 ^ i
        If tmp = 0 Then
            tmp = Sgn(diff) * 2 ^ i
        End If

        If tmp > Columns.Count Then
            tmp = Columns.Count
        Else
            tmp = pred + tmp
        End If

        If diff < 0 Then
            h = l
            If tmp < 1 Then
                l = 1
            Else
                l = tmp
            End If
        Else
            l = h
            If tmp > Columns.Count Then
                h = Columns.Count
            Else
                h = tmp
            End If
        End If

        If Columns(l).Left <= point And point < Columns(h).Left Then
            Exit Do
        End If

        i = i + 1
    Loop

    '二分探索で列を特定
    Do
        m = Round(l + (h - l) / 2 - 0.25)
        If h - l < 2 Then
            Exit Do
        End If

        predLeft = Columns(m).Left
        If point < predLeft Then
            h = m
        Else
            l = m
        End If
    Loop

    'モードに応じて処理分岐
    Select Case searchMode
        '中央寄せ
        Case modeCenter
            '1列追加することで差分の絶対値が近くなる方を選ぶ
            If (point - Columns(m).Left) >= Columns(m).Width / 2 Then
                pointToColumn = m + 1
            Else
                pointToColumn = m
            End If

        '左寄せ, 右寄せ
        Case modeLeft, modeRight
            'ピッタリでないなら1列追加
            If point > Columns(m).Left Then
                pointToColumn = m + 1
            Else
                pointToColumn = m
            End If

        '例外
        Case Else
            pointToColumn = m
    End Select
    Exit Function

Catch:
    Call errorHandler("pointToColumn")
End Function

Private Function getLengthWithZoomConsidered(ByVal Length As Double) As Double
    'Zoomを考慮した長さを取得
    Dim rate As Double

    If 90 < ActiveWindow.Zoom And ActiveWindow.Zoom < 110 Then
        rate = 1
    Else
        rate = 103.32 / ActiveWindow.Zoom - 0.05
    End If
    getLengthWithZoomConsidered = Length * rate
End Function

Private Function getRealUsableHeight() As Double
    If ActiveWindow.DisplayHeadings Then
        getRealUsableHeight = ActiveWindow.UsableHeight - ActiveSheet.StandardHeight
    Else
        getRealUsableHeight = ActiveWindow.UsableHeight
    End If
End Function

Private Function getRealUsableWidth() As Double
    Dim maxVisibleRow As Long
    Dim headingWidth As Double

    If ActiveWindow.DisplayHeadings Then
        maxVisibleRow = ActiveWindow.VisibleRange.Item(ActiveWindow.VisibleRange.Count).Row
        headingWidth = 25

        If maxVisibleRow >= 1000 Then
            headingWidth = headingWidth + 6.75 * (Len(CStr(maxVisibleRow)) - 3)
        End If

        getRealUsableWidth = ActiveWindow.UsableWidth - headingWidth
    Else
        getRealUsableWidth = ActiveWindow.UsableWidth
    End If
End Function

Function scrollCurrentTop()
    ActiveWindow.ScrollRow = pointToRow(ActiveCell.Top - getLengthWithZoomConsidered(SCROLL_OFFSET), modeTop)
End Function

Function scrollCurrentBottom()
    Dim uh As Double
    uh = getRealUsableHeight()

    ActiveWindow.ScrollRow = pointToRow(ActiveCell.Top + ActiveCell.Height - getLengthWithZoomConsidered(uh - SCROLL_OFFSET), modeBottom)
End Function

Function scrollCurrentMiddle()
    Dim uh As Double
    uh = getRealUsableHeight()

    ActiveWindow.ScrollRow = pointToRow(ActiveCell.Top + ActiveCell.Height / 2 - getLengthWithZoomConsidered(uh) / 2, modeMiddle)
End Function

Function scrollCurrentLeft()
    ActiveWindow.ScrollColumn = ActiveCell.Column
End Function

Function scrollCurrentRight()
    Dim uw As Double
    uw = getRealUsableWidth()

    ActiveWindow.ScrollColumn = pointToColumn(ActiveCell.Left + ActiveCell.Width - getLengthWithZoomConsidered(uw), modeRight)
End Function

Function scrollCurrentCenter()
    Dim uw As Double
    uw = getRealUsableWidth()

    ActiveWindow.ScrollColumn = pointToColumn(ActiveCell.Left + ActiveCell.Width / 2 - getLengthWithZoomConsidered(uw) / 2, modeCenter)
End Function
