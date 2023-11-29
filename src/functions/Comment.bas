Attribute VB_Name = "F_Comment"
Option Explicit
Option Private Module

Function EditCellComment(Optional ByVal g As String) As Boolean
    Call RepeatRegister("EditCellComment")
    Call StopVisualMode

    If TypeName(Selection) = "Range" Then
        Call KeyStroke(True, Shift_ + F2_)
    End If
End Function

Function DeleteCellComment(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteCellComment")
    Call StopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        Call KeyStroke(True, Alt_ + R_, D_)
    End If
End Function

Function DeleteCellCommentAll(Optional ByVal g As String) As Boolean
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
    Call ErrorHandler("DeleteCellCommentAll")
End Function

Function ToggleCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("ToggleCellComment")
    Call StopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        Application.CommandBars.ExecuteMso "ReviewShowOrHideComment"
    End If
    Exit Function

Catch:
    Call ErrorHandler("ToggleCellComment")
End Function

Function HideCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("HideCellComment")
    Call StopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        ActiveCell.Comment.Visible = False
    End If
    Exit Function

Catch:
    Call ErrorHandler("HideCellComment")
End Function

Function ShowCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("ShowCellComment")
    Call StopVisualMode

    If Not ActiveCell.Comment Is Nothing Then
        ActiveCell.Comment.Visible = True
    End If
    Exit Function

Catch:
    Call ErrorHandler("ShowCellComment")
End Function

Function ToggleCellCommentAll(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Application.CommandBars.ExecuteMso "ReviewShowAllComments"
    Exit Function

Catch:
    Call ErrorHandler("ToggleCellCommentAll")
End Function

Function HideCellCommentAll(Optional ByVal g As String) As Boolean
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
End Function

Function ShowCellCommentAll(Optional ByVal g As String) As Boolean
    Application.DisplayCommentIndicator = xlCommentAndIndicator
End Function

Function HideCellCommentIndicator(Optional ByVal g As String) As Boolean
    Application.DisplayCommentIndicator = xlNoIndicator
End Function

Function NextCommentedCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim buf As Boolean

    'アクティブシートにコメントが無いなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    Call StopVisualMode

    'もともとの値を取得
    buf = Application.DisplayAlerts

    Application.DisplayAlerts = False
    Application.CommandBars.ExecuteMso "ReviewNextComment"
    Application.DisplayAlerts = buf
    Exit Function

Catch:
    Call ErrorHandler("NextCommentedCell")
End Function

Function PrevCommentedCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim buf As Boolean

    'アクティブシートにコメントが無いなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    Call StopVisualMode

    'もともとの値を取得
    buf = Application.DisplayAlerts

    Application.DisplayAlerts = False
    Application.CommandBars.ExecuteMso "ReviewPreviousComment"
    Application.DisplayAlerts = buf
    Exit Function

Catch:
    Call ErrorHandler("PrevCommentedCell")
End Function
