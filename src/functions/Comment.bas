Attribute VB_Name = "F_Comment"
Option Explicit
Option Private Module

Function editCellComment()
    Call repeatRegister("editCellComment")
    Call stopVisualMode

    If TypeName(Selection) = "Range" Then
        Call keystroke(True, Shift_ + F2_)
    End If
End Function

Function deleteCellComment()
    Call repeatRegister("deleteCellComment")
    Call stopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        Call keystroke(True, Alt_ + R_, D_)
    End If
End Function

Function deleteCellCommentAll()
    On Error GoTo Catch

    Dim cmt As Comment

    'アクティブシートにコメントがないなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    '確認メッセージ
    If MsgBox("アクティブシート上のすべてのコメントを削除します。よろしいですか?" & vbLf & _
              "　※この操作は取り消せません。", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Function
    End If

    '1つ1つ削除
    For Each cmt In ActiveSheet.Comments
        cmt.Delete
    Next cmt
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("deleteCellCommentAll")
    End If
End Function

Function toggleCellComment()
    On Error GoTo Catch

    Call repeatRegister("toggleCellComment")
    Call stopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        Application.CommandBars.ExecuteMso "ReviewShowOrHideComment"
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("toggleCellComment")
    End If
End Function

Function hideCellComment()
    On Error GoTo Catch

    Call repeatRegister("hideCellComment")
    Call stopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        ActiveCell.Comment.Visible = False
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("hideCellComment")
    End If
End Function

Function showCellComment()
    On Error GoTo Catch

    Call repeatRegister("showCellComment")
    Call stopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        ActiveCell.Comment.Visible = True
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("showCellComment")
    End If
End Function

Function toggleCellCommentAll()
    On Error GoTo Catch

    Application.CommandBars.ExecuteMso "ReviewShowAllComments"
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("toggleCellCommentAll")
    End If
End Function

Function hideCellCommentAll()
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
End Function

Function showCellCommentAll()
    Application.DisplayCommentIndicator = xlCommentAndIndicator
End Function

Function hideCellCommentIndicator()
    Application.DisplayCommentIndicator = xlNoIndicator
End Function

Function nextCommentedCell()
    On Error GoTo Catch

    Dim buf As Boolean

    'アクティブシートにコメントが無いなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    Call stopVisualMode

    'もともとの値を取得
    buf = Application.DisplayAlerts

    Application.DisplayAlerts = False
    Application.CommandBars.ExecuteMso "ReviewNextComment"
    Application.DisplayAlerts = buf
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("nextCommentedCell")
    End If
End Function

Function prevCommentedCell()
    On Error GoTo Catch

    Dim buf As Boolean

    'アクティブシートにコメントが無いなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    Call stopVisualMode

    'もともとの値を取得
    buf = Application.DisplayAlerts

    Application.DisplayAlerts = False
    Application.CommandBars.ExecuteMso "ReviewPreviousComment"
    Application.DisplayAlerts = buf
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("prevCommentedCell")
    End If
End Function
