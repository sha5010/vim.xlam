Attribute VB_Name = "F_UsefulCmd"
Option Explicit
Option Private Module

Function undo_CtrlZ()
    Call keystroke(True, Ctrl_ + Z_)
End Function

Function redoExecute()
    On Error Resume Next
    Application.CommandBars.ExecuteMso "Redo"
End Function

Function toggleFreezePanes()
    On Error GoTo Catch
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
    Exit Function

Catch:
    Call errorHandler("toggleFreezePanes")
End Function

Function zoomIn()
    On Error GoTo Catch

    Dim afterZoomRate As Integer

    If gVim.Count > 0 Then
        afterZoomRate = ActiveWindow.Zoom + gVim.Count
    Else
        afterZoomRate = ActiveWindow.Zoom + 10
    End If

    If afterZoomRate > 400 Then
        afterZoomRate = 400
    End If

    ActiveWindow.Zoom = afterZoomRate
    Exit Function

Catch:
    If errorHandler("zoomIn") Then
        Call keystroke(True, Ctrl_ + Shift_ + Alt_ + Minus_)
    End If
End Function

Function zoomOut()
    On Error GoTo Catch

    Dim afterZoomRate As Integer

    If gVim.Count > 0 Then
        afterZoomRate = ActiveWindow.Zoom - gVim.Count
    Else
        afterZoomRate = ActiveWindow.Zoom - 10
    End If

    If afterZoomRate < 10 Then
        afterZoomRate = 10
    End If

    ActiveWindow.Zoom = afterZoomRate
    Exit Function

Catch:
    If errorHandler("zoomOut") Then
        Call keystroke(True, Ctrl_ + Alt_ + Minus_)
    End If
End Function

Function zoomSpecifiedScale()
    On Error GoTo Catch

    Dim zoomScale As Integer

    Select Case gVim.Count1
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
            zoomScale = gVim.Count1
    End Select

    ActiveWindow.Zoom = zoomScale
    Exit Function

Catch:
    Call errorHandler("zoomSpecifiedScale")
End Function

Function toggleFormulaBar()
    On Error GoTo Catch
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    Exit Function

Catch:
    Call errorHandler("toggleFormulaBar")
End Function

Function showSummaryInfo()
    On Error GoTo Catch
    Application.Dialogs(xlDialogSummaryInfo).Show
    Exit Function

Catch:
    Call errorHandler("showSummaryInfo")
End Function

Function smartFillColor()
    Call stopVisualMode

    If TypeName(Selection) = "Range" Then
        Call changeInteriorColor
    ElseIf VarType(Selection) = vbObject Then
        Call changeShapeFillColor
    End If
End Function

Function smartFontColor()
    Call stopVisualMode

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

Function setPrintArea()
    Call stopVisualMode

    'Send Alt + P, R, S
    Call keystroke(True, Alt_ + P_, R_, S_)
End Function

Function clearPrintArea()
    Call stopVisualMode

    'Send Alt + P, R, C
    Call keystroke(True, Alt_ + P_, R_, C_)
End Function
