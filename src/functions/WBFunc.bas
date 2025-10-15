Attribute VB_Name = "F_WBFunc"
Option Explicit
Option Private Module

Private Sub CheckAndQuitIfNoWorkbooks()
    On Error GoTo Catch
    If gVim.Config.QuitApp And Application.Workbooks.Count = 0 Then
        Application.Quit
    End If
    Exit Sub

Catch:
    Call ErrorHandler("CheckAndQuitIfNoWorkbooks")
End Sub

Private Function GetWorkbookByName(ByVal bookName As String) As Workbook
    On Error Resume Next
    If bookName = "" Then
        Set GetWorkbookByName = ActiveWorkbook
    Else
        Set GetWorkbookByName = Workbooks(bookName)
    End If
    If Err.Number <> 0 Then
        Set GetWorkbookByName = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

Function CloseAskSaving(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If
    targetWorkbook.Close
    Call CheckAndQuitIfNoWorkbooks
    Exit Function

Catch:
    Call ErrorHandler("CloseAskSaving")
End Function

Function CloseWithoutSaving(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If
    targetWorkbook.Close False
    Call CheckAndQuitIfNoWorkbooks
    Exit Function

Catch:
    Call ErrorHandler("CloseWithoutSaving")
End Function

Function CloseWithSaving(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If
    targetWorkbook.Close True
    Call CheckAndQuitIfNoWorkbooks
    Exit Function

Catch:
    Call ErrorHandler("CloseWithSaving")
End Function

Function SaveWorkbook(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If

    If targetWorkbook.Path = "" Then
        Application.CommandBars.ExecuteMso "FileSaveAs"
    ElseIf targetWorkbook.ReadOnly Then
        Application.CommandBars.ExecuteMso "FileSaveAs"
    Else
        targetWorkbook.Save
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

Function OpenWorkbook(Optional ByVal relPath As String) As Boolean
    On Error GoTo Catch

    If relPath = "" Then
        Application.CommandBars.ExecuteMso "FileOpenUsingBackstage"
        Exit Function
    End If

    Dim absPath As String
    absPath = ResolvePath(relPath)

    Workbooks.Open absPath
    Exit Function

Catch:
    Call ErrorHandler("OpenWorkbook")
End Function

Function ReopenActiveWorkbook(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch

    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If

    Dim wbFullName As String
    Dim ret As VbMsgBoxResult

    If InStr(targetWorkbook.FullName, "\") = 0 And InStr(targetWorkbook.FullName, "/") = 0 Then
        Exit Function
    End If

    If Not targetWorkbook.Saved Then
        ret = MsgBox(gVim.Msg.ConfirmToSaveBeforeReopening, vbYesNoCancel + vbQuestion)
        If ret = vbCancel Then
            Exit Function
        ElseIf ret = vbNo Then
            targetWorkbook.Saved = True
        ElseIf ret = vbYes Then
            targetWorkbook.Save
        End If
    End If

    wbFullName = targetWorkbook.FullName

    targetWorkbook.Close
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

    Dim i As Long: i = GetWorkbookIndex(ActiveWorkbook)
    Dim cnt As Long: cnt = gVim.Count1
    Dim currentIdx As Long: currentIdx = i

    Do While cnt > 0
        i = (i Mod Workbooks.Count) + 1
        If Windows(Workbooks(i).Name).Visible Then
            cnt = cnt - 1
        End If

        If i = currentIdx Then
            Dim visibleBooks As Long
            visibleBooks = gVim.Count1 - cnt
            cnt = cnt Mod visibleBooks
        End If
    Loop
    Workbooks(i).Activate
    Exit Function

Catch:
    Call ErrorHandler("NextWorkbook")
End Function

Function PreviousWorkbook(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long: i = GetWorkbookIndex(ActiveWorkbook)
    Dim cnt As Long: cnt = gVim.Count1
    Dim currentIdx As Long: currentIdx = i

    Do While cnt > 0
        i = ((i - 2 + Workbooks.Count) Mod Workbooks.Count) + 1
        If Windows(Workbooks(i).Name).Visible Then
            cnt = cnt - 1
        End If

        If i = currentIdx Then
            Dim visibleBooks As Long
            visibleBooks = gVim.Count1 - cnt
            cnt = cnt Mod visibleBooks
        End If
    Loop
    Workbooks(i).Activate
    Exit Function

Catch:
    Call ErrorHandler("PreviousWorkbook")
End Function

Function ToggleReadOnly(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If

    Dim ret As VbMsgBoxResult

    If InStr(targetWorkbook.FullName, "\") = 0 And InStr(targetWorkbook.FullName, "/") = 0 Then
        Exit Function
    End If

    If targetWorkbook.ReadOnly Then
        targetWorkbook.Saved = True
        Call targetWorkbook.ChangeFileAccess(xlReadWrite)
    Else
        If Not targetWorkbook.Saved Then
            ret = MsgBox(gVim.Msg.ConfirmToSaveBeforeSwitchReadonly, vbYesNoCancel + vbQuestion)
            If ret = vbCancel Then
                Exit Function
            ElseIf ret = vbNo Then
                targetWorkbook.Saved = True
            ElseIf ret = vbYes Then
                targetWorkbook.Save
            End If
        End If

        Call targetWorkbook.ChangeFileAccess(xlReadOnly)
    End If
    Exit Function

Catch:
    Call ErrorHandler("ToggleReadOnly")
End Function

Function OpenWorkbookDir(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If
    targetWorkbook.FollowHyperlink targetWorkbook.Path
    Exit Function

Catch:
    Call ErrorHandler("OpenWorkbookDir")
End Function

Function YankWorkbookPath(Optional ByVal bookName As String = "") As Boolean
    On Error GoTo Catch
    Dim targetWorkbook As Workbook
    Set targetWorkbook = GetWorkbookByName(bookName)
    If targetWorkbook Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.WorkbookNotFound & " (" & bookName & ")", 3000)
        Exit Function
    End If

    'Set to clipboard
    With New DataObject
        .SetText targetWorkbook.FullName
        .PutInClipboard
    End With

    Call SetStatusBarTemporarily(gVim.Msg.YankDone & " (" & targetWorkbook.FullName & ")", 3000)
    Exit Function

Catch:
    Call ErrorHandler("YankWorkbookPath")
End Function
