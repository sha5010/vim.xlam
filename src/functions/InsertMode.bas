Attribute VB_Name = "F_InsertMode"
Option Explicit
Option Private Module

Function InsertWithIME(Optional ByVal g As String) As Boolean
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(Space_, BackSpace_, Ctrl_ + Home_, IME_On_)
    Else
        Call KeyStroke(F2_, Ctrl_ + Home_, IME_On_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function InsertWithoutIME(Optional ByVal g As String) As Boolean
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(Space_, BackSpace_, Ctrl_ + Home_)
    Else
        Call KeyStroke(F2_, Ctrl_ + Home_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function AppendWithIME(Optional ByVal g As String) As Boolean
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(Space_, BackSpace_, IME_On_)
    Else
        Call KeyStroke(F2_, IME_On_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function AppendWithoutIME(Optional ByVal g As String) As Boolean
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(Space_, BackSpace_)
    Else
        Call KeyStroke(F2_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function SubstituteWithIME(Optional ByVal g As String) As Boolean
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(Enter_, Delete_, IME_On_)
    Else
        Call KeyStroke(BackSpace_, F2_, IME_On_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function SubstituteWithoutIME(Optional ByVal g As String) As Boolean
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(Enter_, Delete_)
    Else
        Call KeyStroke(BackSpace_, F2_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function InsertFollowLangMode(Optional ByVal g As String) As Boolean
    If gVim.IsJapanese Then
        Call InsertWithIME
    Else
        Call InsertWithoutIME
    End If
End Function

Function InsertNotFollowLangMode(Optional ByVal g As String) As Boolean
    If Not gVim.IsJapanese Then
        Call InsertWithIME
    Else
        Call InsertWithoutIME
    End If
End Function

Function AppendFollowLangMode(Optional ByVal g As String) As Boolean
    If gVim.IsJapanese Then
        Call AppendWithIME
    Else
        Call AppendWithoutIME
    End If
End Function

Function AppendNotFollowLangMode(Optional ByVal g As String) As Boolean
    If Not gVim.IsJapanese Then
        Call AppendWithIME
    Else
        Call AppendWithoutIME
    End If
End Function

Function SubstituteFollowLangMode(Optional ByVal g As String) As Boolean
    If gVim.IsJapanese Then
        Call SubstituteWithIME
    Else
        Call SubstituteWithoutIME
    End If
End Function

Function SubstituteNotFollowLangMode(Optional ByVal g As String) As Boolean
    If Not gVim.IsJapanese Then
        Call SubstituteWithIME
    Else
        Call SubstituteWithoutIME
    End If
End Function

Private Sub StartEditing()
    gVim.Vars.FromInsertCmd = True
    Application.OnTime Now + (1 / 86400) * 0.1, "StopEditing"
End Sub

Private Sub StopEditing()
    Call StopVisualMode
    gVim.Vars.FromInsertCmd = False
End Sub
