Attribute VB_Name = "F_UsefulCmd"
Option Explicit
Option Private Module

Function undo_CtrlZ()
    Call keystroke(True, Ctrl_ + Z_)
End Function

Function repeat_F4()
    On Error Resume Next
    Application.Repeat
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
