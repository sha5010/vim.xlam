Attribute VB_Name = "F_WBFunc"
Option Explicit
Option Private Module

Function CloseAskSaving(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWorkbook.Close
    Exit Function

Catch:
    Call ErrorHandler("CloseAskSaving")
End Function

Function CloseWithoutSaving(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWorkbook.Close False
    Exit Function

Catch:
    Call ErrorHandler("CloseWithoutSaving")
End Function

Function CloseWithSaving(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWorkbook.Close True
    Exit Function

Catch:
    Call ErrorHandler("CloseWithSaving")
End Function

Function SaveWorkbook(Optional ByVal g As String) As Boolean
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
    Call ErrorHandler("SaveWorkbook")
End Function

Function SaveAsNewWorkbook(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "FileSaveAs"
    Exit Function

Catch:
    Call ErrorHandler("SaveAsNewWorkbook")
End Function

Function OpenWorkbook(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "FileOpenUsingBackstage"
    Exit Function

Catch:
    Call ErrorHandler("OpenWorkbook")
End Function

Function ReopenActiveWorkbook(Optional ByVal g As String) As Boolean
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
    Call ErrorHandler("ReopenActiveWorkbook")
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

Function NextWorkbook(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer
    Dim idx As Integer

    idx = GetWorkbookIndex(ActiveWorkbook)
    For i = 1 To Workbooks.Count
        idx = (idx Mod Workbooks.Count) + 1
        If Windows(Workbooks(idx).Name).Visible Then
            Workbooks(idx).Activate
            Exit Function
        End If
    Next i
    Exit Function

Catch:
    Call ErrorHandler("NextWorkbook")
End Function

Function PreviousWorkbook(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Integer
    Dim idx As Integer

    idx = GetWorkbookIndex(ActiveWorkbook)
    For i = 1 To Workbooks.Count
        idx = ((idx - 2 + Workbooks.Count) Mod Workbooks.Count) + 1
        If Windows(Workbooks(idx).Name).Visible Then
            Workbooks(idx).Activate
            Exit Function
        End If
    Next i
    Exit Function

Catch:
    Call ErrorHandler("PreviousWorkbook")
End Function

Function ToggleReadOnly(Optional ByVal g As String) As Boolean
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
    Call ErrorHandler("ToggleReadOnly")
End Function
