Attribute VB_Name = "F_UsefulCmd"
Option Explicit
Option Private Module

Function undo_CtrlZ()
    Call keystroke(True, Ctrl_ + Z_)
End Function

Function redoExecute()
    On Error Resume Next
    Application.CommandBars.ExecuteMso "Redo"
    On Error GoTo 0
End Function

Function toggleFreezePanes()
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
End Function

Function zoomIn()
    Dim afterZoomRate As Integer

    If gCount > 10 Then
        afterZoomRate = ActiveWindow.Zoom + gCount
    Else
        afterZoomRate = ActiveWindow.Zoom + gCount * 10
    End If

    If afterZoomRate > 400 Then
        afterZoomRate = 400
    End If

    ActiveWindow.Zoom = afterZoomRate
End Function

Function zoomOut()
    Dim afterZoomRate As Integer

    If gCount > 10 Then
        afterZoomRate = ActiveWindow.Zoom - gCount
    Else
        afterZoomRate = ActiveWindow.Zoom - gCount * 10
    End If

    If afterZoomRate < 10 Then
        afterZoomRate = 10
    End If

    ActiveWindow.Zoom = afterZoomRate
End Function

Function zoomSpecifiedScale()
    Dim zoomScale As Integer

    Select Case gCount
        Case 1
            zoomScale = 100
        Case 2
            zoomScale = 25
        Case 3
            zoomScale = 55
        Case 4
            zoomScale = 85
        Case 5
            zoomScale = 130
        Case 6
            zoomScale = 160
        Case 7
            zoomScale = 200
        Case 8
            zoomScale = 400
        Case 9
            ActiveWindow.Zoom = True
            Exit Function
        Case Is > 400
            zoomScale = 400
        Case Is <= 400
            zoomScale = gCount
    End Select

    ActiveWindow.Zoom = zoomScale
End Function

Function toggleFormulaBar()
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
End Function

Function showSummaryInfo()
    Application.Dialogs(xlDialogSummaryInfo).Show
End Function

Function jumpPrev()
    Dim t As Range
    Dim wb As Workbook
    Dim ws As Worksheet

    If Not JumpList Is Nothing Then
        Set t = JumpList.Back
        If Not t Is Nothing Then
            Set wb = t.Parent.Parent
            Set ws = t.Parent

            wb.Activate
            ws.Activate
            t.Select
        Else
            Call setStatusBarTemporarily("一番古い履歴です。", 1)
        End If
    End If
End Function

Function jumpNext()
    Dim t As Range
    Dim wb As Workbook
    Dim ws As Worksheet

    If Not JumpList Is Nothing Then
        Set t = JumpList.Forward
        If Not t Is Nothing Then
            Set wb = t.Parent.Parent
            Set ws = t.Parent

            wb.Activate
            ws.Activate
            t.Select
        Else
            Call setStatusBarTemporarily("一番新しい履歴です。", 1)
        End If
    End If
End Function

Function clearJumps()
    If Not JumpList Is Nothing Then
        Call JumpList.ClearAll
        Call setStatusBarTemporarily("ジャンプリストをクリアしました。", 2)
    End If
End Function

Function recordToJumpList(Optional Target As Range)
    'JumpList が利用できるか検証
    If JumpList Is Nothing Then
        Exit Function
    End If

    'Target が未指定の場合は選択中のセル
    If Target Is Nothing Then
        If TypeName(Selection) = "Range" Then
            Set Target = Selection
        ElseIf Not ActiveCell Is Nothing Then
            Set Target = ActiveCell
        Else
            Exit Function
        End If
    End If

    '最新の記録と完全に一致しないなら記録する
    If JumpList.Latest Is Nothing Then
        Call JumpList.Add(Target)
    ElseIf Target.Address <> JumpList.Latest.Address Then
        Call JumpList.Add(Target)
    ElseIf Target.Parent.Name <> JumpList.Latest.Parent.Name Then
        Call JumpList.Add(Target)
    ElseIf Target.Parent.Parent.FullName <> JumpList.Latest.Parent.Parent.FullName Then
        Call JumpList.Add(Target)
    End If
End Function

Function smartFillColor()
    If TypeName(Selection) = "Range" Then
        Call changeInteriorColor
    ElseIf VarType(Selection) = vbObject Then
        Call changeShapeFillColor
    End If
End Function

Function smartFontColor()
    If TypeName(Selection) = "Range" Then
        Call changeFontColor
    ElseIf VarType(Selection) = vbObject Then
        Call changeShapeFontColor
    End If
End Function

Function showContextMenu()
    'Send Shift+F10
    Call keystroke(True, Shift_ + F10_)
End Function

Function showMacroDialog()
    'Send Alt+F8
    Call keystroke(True, Alt_ + F8_, Tab_)
End Function
