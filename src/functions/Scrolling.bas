Attribute VB_Name = "F_Scrolling"
Option Explicit
Option Private Module

Private Function ActivateCellInVisibleRange()
    On Error GoTo Catch

    Dim targetRow As Long
    Dim targetColumn As Long
    Dim visibleTop As Long, visibleBottom As Long
    Dim visibleLeft As Long, visibleRight As Long

    targetRow = ActiveCell.Row
    targetColumn = ActiveCell.Column

    With ActiveWindow.VisibleRange
        visibleTop = .Item(1).Row
        visibleBottom = PointToRow(.Item(.Count).Top - 1, xlNone)
        visibleLeft = .Item(1).Column
        visibleRight = PointToColumn(.Item(.Count).Left - 1, xlNone)
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
            ActiveWindow.ScrollRow = visibleTop
            ActiveWindow.ScrollColumn = visibleLeft
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("ActivateCellInVisibleRange")
End Function

Function ScrollUpHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll Up:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        topRowVisible = ActiveWindow.VisibleRange.Row

        scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
        targetRow = topRowVisible - scrollWidth

        If targetRow < 1 Then
            targetRow = 1
        End If

        ActiveWindow.SmallScroll Up:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollUpHalf")
End Function

Function ScrollDownHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll Down:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        topRowVisible = ActiveWindow.VisibleRange.Row

        scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
        targetRow = topRowVisible + scrollWidth

        If targetRow > ActiveSheet.Rows.Count Then
            targetRow = ActiveSheet.Rows.Count
        End If

        ActiveWindow.SmallScroll Down:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollDownHalf")
End Function

Function ScrollLeftHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim leftColVisible As Long
    Dim scrollWidth As Integer
    Dim targetCol As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll ToLeft:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        leftColVisible = ActiveWindow.VisibleRange.Column

        scrollWidth = ActiveWindow.VisibleRange.Columns.Count / 2
        targetCol = leftColVisible - scrollWidth

        If targetCol < 1 Then
            targetCol = 1
        End If

        ActiveWindow.SmallScroll ToLeft:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollLeftHalf")
End Function

Function ScrollRightHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim leftColVisible As Long
    Dim scrollWidth As Integer
    Dim targetCol As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll ToRight:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        leftColVisible = ActiveWindow.VisibleRange.Column

        scrollWidth = ActiveWindow.VisibleRange.Columns.Count / 2
        targetCol = leftColVisible + scrollWidth

        If targetCol > ActiveSheet.Columns.Count Then
            targetCol = ActiveSheet.Columns.Count
        End If

        ActiveWindow.SmallScroll ToRight:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollRightHalf")
End Function


Function ScrollUp(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll Up:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollDown(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll Down:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollLeft(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll ToLeft:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollRight(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll ToRight:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollUp1Row(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll Up:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollDown1Row(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll Down:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollLeft1Column(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll ToLeft:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollRight1Column(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll ToRight:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollCurrentTop(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToSpecifiedRow(CStr(gVim.Count))
    End If
    ActiveWindow.ScrollRow = PointToRow(ActiveCell.Top - GetLengthWithZoomConsidered(gVim.Config.ScrollOffset), ModeTop)
End Function

Function ScrollCurrentBottom(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToSpecifiedRow(CStr(gVim.Count))
    End If

    Dim uh As Double
    uh = GetRealUsableHeight()

    ActiveWindow.ScrollRow = PointToRow(ActiveCell.Top + ActiveCell.Height - GetLengthWithZoomConsidered(uh - gVim.Config.ScrollOffset), ModeBottom)
End Function

Function ScrollCurrentMiddle(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToSpecifiedRow(CStr(gVim.Count))
    End If

    Dim uh As Double
    uh = GetRealUsableHeight()

    ActiveWindow.ScrollRow = PointToRow(ActiveCell.Top + ActiveCell.Height / 2 - GetLengthWithZoomConsidered(uh) / 2, ModeMiddle)
End Function

Function ScrollCurrentLeft(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToNthColumn
    End If

    ActiveWindow.ScrollColumn = ActiveCell.Column
End Function

Function ScrollCurrentRight(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToNthColumn
    End If

    Dim uw As Double
    uw = GetRealUsableWidth()

    ActiveWindow.ScrollColumn = PointToColumn(ActiveCell.Left + ActiveCell.Width - GetLengthWithZoomConsidered(uw), ModeRight)
End Function

Function ScrollCurrentCenter(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToNthColumn
    End If

    Dim uw As Double
    uw = GetRealUsableWidth()

    ActiveWindow.ScrollColumn = PointToColumn(ActiveCell.Left + ActiveCell.Width / 2 - GetLengthWithZoomConsidered(uw) / 2, ModeCenter)
End Function
