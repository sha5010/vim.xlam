Attribute VB_Name = "F_WBFunc"
Option Explicit
Option Private Module

Function closeAskSaving()
    On Error GoTo Catch
    ActiveWorkbook.Close
    Exit Function

Catch:
    Call errorHandler("closeAskSaving")
End Function

Function closeWithoutSaving()
    On Error GoTo Catch
    ActiveWorkbook.Close False
    Exit Function

Catch:
    Call errorHandler("closeWithoutSaving")
End Function

Function closeWithSaving()
    On Error GoTo Catch
    ActiveWorkbook.Close True
    Exit Function

Catch:
    Call errorHandler("closeWithSaving")
End Function

Function saveWorkbook()
    On Error GoTo Catch

    If ActiveWorkbook.Path = "" Then
        Application.CommandBars.ExecuteMso "FileSaveAs"
    ElseIf ActiveWorkbook.ReadOnly Then
        Application.CommandBars.ExecuteMso "FileSaveAs"
    Else
        ActiveWorkbook.Save
    End If
    Exit Function

Catch:
    Call errorHandler("saveWorkbook")
End Function

Function saveAsNewWorkbook()
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "FileSaveAs"
    Exit Function

Catch:
    Call errorHandler("saveAsNewWorkbook")
End Function

Function openWorkbook()
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "FileOpenUsingBackstage"
    Exit Function

Catch:
    Call errorHandler("openWorkbook")
End Function

Function reopenActiveWorkbook()
    On Error GoTo Catch

    Dim wbFullName As String
    Dim ret As VbMsgBoxResult

    If InStr(ActiveWorkbook.FullName, "¥") = 0 And InStr(ActiveWorkbook.FullName, "/") = 0 Then
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
    Exit Function

Catch:
    Call errorHandler("reopenActiveWorkbook")
End Function

Function ActivateWorkbook(Optional ByVal arg As String) As Boolean
    On Error GoTo Catch

    Dim idx As Long
    Dim isForce As Boolean

    isForce = (InStr(arg, "!") > 0)
    arg = Replace(arg, "!", "")

    If Len(arg) = 0 Or arg Like "*[!0-9]*" Then
        Exit Function
    End If

    idx = CLng(arg)

    If idx < 1 Then
        idx = 1
    ElseIf Windows.Count < idx Then
        idx = Windows.Count
    End If

    With Windows(idx)
        If .Visible Or isForce Then
            .Visible = True
            .Activate
            ActivateWorkbook = True
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("ActivateWorkbook")
End Function

Function nextWorkbook()
    On Error GoTo Catch

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
    Exit Function

Catch:
    Call errorHandler("nextWorkbook")
End Function

Function previousWorkbook()
    On Error GoTo Catch

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
    Exit Function

Catch:
    Call errorHandler("previousWorkbook")
End Function

Function toggleReadOnly()
    On Error GoTo Catch

    Dim ret As VbMsgBoxResult

    If InStr(ActiveWorkbook.FullName, "¥") = 0 And InStr(ActiveWorkbook.FullName, "/") = 0 Then
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
    Exit Function

Catch:
    Call errorHandler("toggleReadOnly")
End Function
