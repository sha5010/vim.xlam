Attribute VB_Name = "F_Scrolling"
Option Explicit
Option Private Module

Function scrollUpHalf()
    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    topRowVisible = ActiveWindow.VisibleRange.Row

    scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
    targetRow = topRowVisible - scrollWidth

    If targetRow < 1 Then
        targetRow = 1
    End If

    ActiveWindow.SmallScroll Up:=scrollWidth

    Cells(targetRow, ActiveCell.Column).Activate
End Function

Function scrollDownHalf()
    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    topRowVisible = ActiveWindow.VisibleRange.Row

    scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
    targetRow = topRowVisible + scrollWidth

    If targetRow > ActiveSheet.Rows.Count Then
        targetRow = ActiveSheet.Rows.Count
    End If

    ActiveWindow.SmallScroll Down:=scrollWidth

    Cells(targetRow, ActiveCell.Column).Activate
End Function

Function scrollUp()
    Call keyupControlKeys

    keybd_event vbKeyPageUp, 0, EXTENDED_KEY, 0
    keybd_event vbKeyPageUp, 0, EXTENDED_KEY Or KEYUP, 0

    Call unkeyupControlKeys
End Function

Function scrollDown()
    Call keyupControlKeys

    keybd_event vbKeyPageDown, 0, EXTENDED_KEY, 0
    keybd_event vbKeyPageDown, 0, EXTENDED_KEY Or KEYUP, 0

    Call unkeyupControlKeys
End Function

Function scrollUp1Row()
    ActiveWindow.SmallScroll Up:=1
End Function

Function scrollDown1Row()
    ActiveWindow.SmallScroll Down:=1
End Function

Function scrollCurrentTop()
    With ActiveWindow
        .ScrollIntoView _
            Left:=.VisibleRange.Left * 3, _
            Top:=(ActiveCell.Top - SCROLL_OFFSET) * 3, _
            Width:=0, Height:=0

        While ActiveCell.Top - SCROLL_OFFSET > .VisibleRange.Top
            .SmallScroll Down:=1
        Wend
    End With
End Function

Function scrollCurrentBottom()
    With ActiveWindow
        .ScrollIntoView _
            Left:=.VisibleRange.Left * 3, _
            Top:=(ActiveCell.Top - .UsableHeight + 18 + SCROLL_OFFSET + ActiveCell.Height) * 3, _
            Width:=0, Height:=0

        While ActiveCell.Top + ActiveCell.Height > .VisibleRange.Top + .UsableHeight - 18
            .SmallScroll Down:=1
        Wend
    End With
End Function

Function scrollCurrentMiddle()
    With ActiveWindow
        .ScrollIntoView _
            Left:=.VisibleRange.Left * 3, _
            Top:=(ActiveCell.Top + ActiveCell.Height / 2 - (.UsableHeight - 18) / 2) * 3, _
            Width:=0, Height:=0

        While (ActiveCell.Top - .VisibleRange.Top + ActiveCell.Height / 2) - ((.UsableHeight - 18) / 2) > .VisibleRange.Rows(1).Height / 2
            .SmallScroll Down:=1
        Wend
    End With
End Function

Function scrollCurrentLeft()
    With ActiveWindow
        .ScrollIntoView _
            Left:=ActiveCell.Left * 3, _
            Top:=.VisibleRange.Top * 3, _
            Width:=0, Height:=0
    End With
End Function

Function scrollCurrentRight()
    With ActiveWindow
        .ScrollIntoView _
            Left:=(ActiveCell.Left - .UsableWidth + 22 + ActiveCell.Width) * 3, _
            Top:=.VisibleRange.Top * 3, _
            Width:=0, Height:=0

        While ActiveCell.Left + ActiveCell.Width > .VisibleRange.Left + .UsableWidth - 22
            .SmallScroll ToRight:=1
        Wend
    End With
End Function

Function scrollCurrentCenter()
    With ActiveWindow
        .ScrollIntoView _
            Left:=(ActiveCell.Left + ActiveCell.Width / 2 - (.UsableWidth - 22) / 2) * 3, _
            Top:=.VisibleRange.Top * 3, _
            Width:=0, Height:=0

        While (ActiveCell.Left - .VisibleRange.Left + ActiveCell.Width / 2) - ((.UsableWidth - 22) / 2) > .VisibleRange.Columns(1).Width / 2
            .SmallScroll ToRight:=1
        Wend
    End With
End Function
