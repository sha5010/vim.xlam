Attribute VB_Name = "F_Comment"
Option Explicit
Option Private Module

Function editCellComment()
    Call repeatRegister("editCellComment")

    If TypeName(Selection) = "Range" Then
        Call keystroke(True, Shift_ + F2_)
    End If
End Function

Function deleteCellComment()
    Call repeatRegister("deleteCellComment")

    If Not ActiveCell.Comment Is Nothing Then
        Call keystroke(True, Alt_ + R_, D_)
    End If
End Function

Function deleteCellCommentAll()
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
End Function

Function toggleCellComment()
    Call repeatRegister("toggleCellComment")

    If Not ActiveCell.Comment Is Nothing Then
        Application.CommandBars.ExecuteMso "ReviewShowOrHideComment"
    End If
End Function

Function hideCellComment()
    Call repeatRegister("hideCellComment")

    If Not ActiveCell.Comment Is Nothing Then
        ActiveCell.Comment.Visible = False
    End If
End Function

Function showCellComment()
    Call repeatRegister("showCellComment")

    If Not ActiveCell.Comment Is Nothing Then
        ActiveCell.Comment.Visible = True
    End If
End Function

Function toggleCellCommentAll()
    Application.CommandBars.ExecuteMso "ReviewShowAllComments"
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
    Dim buf As Boolean

    'アクティブシートにコメントが無いなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    'もともとの値を取得
    buf = Application.DisplayAlerts

    Application.DisplayAlerts = False
    Application.CommandBars.ExecuteMso "ReviewNextComment"
    Application.DisplayAlerts = buf
End Function

Function prevCommentedCell()
    Dim buf As Boolean

    'アクティブシートにコメントが無いなら何もしない
    If ActiveSheet.Comments.Count = 0 Then
        Exit Function
    End If

    'もともとの値を取得
    buf = Application.DisplayAlerts

    Application.DisplayAlerts = False
    Application.CommandBars.ExecuteMso "ReviewPreviousComment"
    Application.DisplayAlerts = buf
End Function
