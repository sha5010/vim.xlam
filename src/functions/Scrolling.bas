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
    Dim topRowVisible As Long
    Dim targetRow As Long

    topRowVisible = ActiveWindow.VisibleRange.Row
    targetRow = WorksheetFunction.Max(ActiveCell.Row - SCROLL_OFFSET, 1)

    ActiveWindow.SmallScroll Down:=targetRow - topRowVisible
End Function

Function scrollCurrentBottom()
    Dim bottomRowVisible As Long
    Dim targetRow As Long

    bottomRowVisible = ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count
    targetRow = WorksheetFunction.Max(ActiveCell.Row + SCROLL_OFFSET + 1, 1)

    ActiveWindow.SmallScroll Up:=bottomRowVisible - targetRow
End Function

Function scrollCurrentMiddle()
    Dim middleRowVisible As Long
    Dim targetRow As Long

    middleRowVisible = ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count / 2 - 1
    targetRow = ActiveCell.Row

    If middleRowVisible > targetRow Then
        ActiveWindow.SmallScroll Up:=middleRowVisible - targetRow
    ElseIf middleRowVisible < targetRow Then
        ActiveWindow.SmallScroll Down:=targetRow - middleRowVisible
    End If
End Function

Function scrollCurrentLeft()
    Dim leftColumnVisible As Long
    Dim targetColumn As Long

    leftColumnVisible = ActiveWindow.VisibleRange.Column
    targetColumn = ActiveCell.Column

    ActiveWindow.SmallScroll ToRight:=targetColumn - leftColumnVisible
End Function

Function scrollCurrentRight()
    Dim rightColumnVisible As Long
    Dim targetColumn As Long

    rightColumnVisible = ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count - 1
    targetColumn = ActiveCell.Column

    ActiveWindow.SmallScroll ToLeft:=rightColumnVisible - targetColumn
End Function

Function scrollCurrentCenter()
    Dim centerColumnVisible As Long
    Dim targetColumn As Long

    centerColumnVisible = ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count / 2 - 1
    targetColumn = ActiveCell.Column

    If centerColumnVisible > targetColumn Then
        ActiveWindow.SmallScroll ToLeft:=centerColumnVisible - targetColumn
    ElseIf centerColumnVisible < targetColumn Then
        ActiveWindow.SmallScroll ToRight:=targetColumn - centerColumnVisible
    End If
End Function
