Attribute VB_Name = "F_WBFunc"
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

Function saveAsNewWorkbook()
    Application.CommandBars.ExecuteMso "FileSaveAs"
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
