Attribute VB_Name = "F_UsefulCmd"
Option Explicit
Option Private Module

Function closeAskSaving()
    ActiveWorkbook.Close
End Function

Function closeWithoutSaving()
    ActiveWorkbook.Close False
End Function

Function closeWithSaving()
    ActiveWorkbook.Close True
End Function

Function saveWorkbook()
    If ActiveWorkbook.Path = "" Then
        Application.CommandBars.ExecuteMso "FileSaveAs"
    ElseIf ActiveWorkbook.ReadOnly Then
        Application.CommandBars.ExecuteMso "FileSaveAs"
    Else
        ActiveWorkbook.Save
    End If
End Function

Function openWorkbook()
    Application.CommandBars.ExecuteMso "FileOpenUsingBackstage"
End Function

Function reopenActiveWorkbook()
    Dim wbFullName As String
    Dim ret As VbMsgBoxResult
    
    If InStr(ActiveWorkbook.FullName, "¥") = 0 Then
        Exit Function
    End If

    If Not ActiveWorkbook.Saved Then
        ret = MsgBox("ファイルを開き直す前に、編集内容を保存しますか?", vbYesNoCancel + vbQuestion)
        If ret = vbCancel Then
            Exit Function
        ElseIf ret = vbNo Then
            ActiveWorkbook.Saved = True
        ElseIf ret = vbYes Then
            ActiveWorkbook.Save
        End If
    End If
    
    wbFullName = ActiveWorkbook.FullName
    
    ActiveWorkbook.Close
    Call Workbooks.Open(wbFullName)
End Function

Function undo_CtrlZ()
    Call keystroke(True, Ctrl_ + Z_)
End Function

Function repeat_F4()
'    Call keyupControlKeys
'    Call releaseShiftKeys
'
'    keybd_event vbKeyF4, 0, 0, 0
'    keybd_event vbKeyF4, 0, KEYUP, 0
'
'    Call unkeyupControlKeys

    On Error Resume Next
    Application.Repeat
    On Error GoTo 0
End Function

Function toggleFreezePanes()
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
End Function

Function zoomIn()
'    Dim i As Integer
'
'    Call keyupControlKeys
'    Call releaseShiftKeys
'
'    keybd_event vbKeyControl, 0, 0, 0
'    keybd_event vbKeyMenu, 0, 0, 0
'    keybd_event vbKeyShift, 0, 0, 0
'
'    For i = 1 To gCount
'        keybd_event &HBD, 0, 0, 0
'        keybd_event &HBD, 0, KEYUP, 0
'    Next i
'
'    keybd_event vbKeyShift, 0, KEYUP, 0
'    keybd_event vbKeyMenu, 0, KEYUP, 0
'    keybd_event vbKeyControl, 0, KEYUP, 0
'
'    Call unkeyupControlKeys

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
'    Dim i As Integer
'
'    Call keyupControlKeys
'    'Call releaseShiftKeys
'
'    keybd_event vbKeyControl, 0, 0, 0
'    keybd_event vbKeyMenu, 0, 0, 0
'
'    For i = 1 To gCount
'        keybd_event vbKeySubtract, 0, 0, 0
'        keybd_event vbKeySubtract, 0, KEYUP, 0
'    Next i
'
'    keybd_event vbKeyMenu, 0, KEYUP, 0
'    keybd_event vbKeyControl, 0, KEYUP, 0
    
    Call unkeyupControlKeys

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

Function activateWorkbook(ByVal n As String) As Boolean
    Dim idx As Integer
    Dim isForce As Boolean
    
    On Error GoTo Catch
    
    isForce = (InStr(n, "!") > 0)
    n = Replace(n, "!", "")
    
    If Not IsNumeric(n) Or InStr(n, ".") > 0 Then
        Exit Function
    End If
    
    idx = CInt(n)
    
    If idx < 1 Or Windows.Count < idx Then
        Exit Function
    End If
    
    With Windows(idx)
        If .Visible Or isForce Then
            .Visible = True
            .Activate
            activateWorkbook = True
        End If
    End With
    Exit Function
    
Catch:
    Call debugPrint("Cannot activate window. ErrNo: " & Err.Number & "  Description: " & Err.Description, "activateWorkbook")
End Function

Function nextWorkbook()
    Dim i As Integer
    Dim idx As Integer

    idx = getWorkbookIndex(ActiveWorkbook)
    For i = 1 To Workbooks.Count
        idx = (idx Mod Workbooks.Count) + 1
        If Windows(Workbooks(idx).Name).Visible Then
            Workbooks(idx).Activate
            Exit Function
        End If
    Next i
End Function

Function previousWorkbook()
    Dim i As Integer
    Dim idx As Integer

    idx = getWorkbookIndex(ActiveWorkbook)
    For i = 1 To Workbooks.Count
        idx = ((idx - 2 + Workbooks.Count) Mod Workbooks.Count) + 1
        If Windows(Workbooks(idx).Name).Visible Then
            Workbooks(idx).Activate
            Exit Function
        End If
    Next i
End Function

Function toggleReadOnly()
    Dim ret As VbMsgBoxResult
    
    If InStr(ActiveWorkbook.FullName, "¥") = 0 Then
        Exit Function
    End If
    
    If ActiveWorkbook.ReadOnly Then
        ActiveWorkbook.Saved = True
        Call ActiveWorkbook.ChangeFileAccess(xlReadWrite)
    Else
        If Not ActiveWorkbook.Saved Then
            ret = MsgBox("読み取り専用の切り替えを行う前に、編集内容を保存しますか?", vbYesNoCancel + vbQuestion)
            If ret = vbCancel Then
                Exit Function
            ElseIf ret = vbNo Then
                ActiveWorkbook.Saved = True
            ElseIf ret = vbYes Then
                ActiveWorkbook.Save
            End If
        End If
        
        Call ActiveWorkbook.ChangeFileAccess(xlReadOnly)
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

