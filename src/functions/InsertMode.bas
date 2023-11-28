Attribute VB_Name = "F_InsertMode"
Option Explicit
Option Private Module

Function insertWithIME()
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(True, Space_, BackSpace_, Ctrl_ + Home_, IME_On_)
    Else
        Call KeyStroke(True, F2_, Ctrl_ + Home_, IME_On_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function insertWithoutIME()
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(True, Space_, BackSpace_, Ctrl_ + Home_)
    Else
        Call KeyStroke(True, F2_, Ctrl_ + Home_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function appendWithIME()
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(True, Space_, BackSpace_, IME_On_)
    Else
        Call KeyStroke(True, F2_, IME_On_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function appendWithoutIME()
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(True, Space_, BackSpace_)
    Else
        Call KeyStroke(True, F2_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function substituteWithIME()
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(True, Enter_, Delete_, IME_On_)
    Else
        Call KeyStroke(True, BackSpace_, F2_, IME_On_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function substituteWithoutIME()
    If VarType(Selection) = vbObject Then
        Call ChangeToShapeInsertMode
        Call KeyStroke(True, Enter_, Delete_)
    Else
        Call KeyStroke(True, BackSpace_, F2_)
        Call StartEditing

        Application.OnTime Now + 0.1 / 86400, "DisableIME"
    End If
End Function

Function insertFollowLangMode()
    If gVim.IsJapanese Then
        Call insertWithIME
    Else
        Call insertWithoutIME
    End If
End Function

Function insertNotFollowLangMode()
    If Not gVim.IsJapanese Then
        Call insertWithIME
    Else
        Call insertWithoutIME
    End If
End Function

Function appendFollowLangMode()
    If gVim.IsJapanese Then
        Call appendWithIME
    Else
        Call appendWithoutIME
    End If
End Function

Function appendNotFollowLangMode()
    If Not gVim.IsJapanese Then
        Call appendWithIME
    Else
        Call appendWithoutIME
    End If
End Function

Function substituteFollowLangMode()
    If gVim.IsJapanese Then
        Call substituteWithIME
    Else
        Call substituteWithoutIME
    End If
End Function

Function substituteNotFollowLangMode()
    If Not gVim.IsJapanese Then
        Call substituteWithIME
    Else
        Call substituteWithoutIME
    End If
End Function

Private Sub StartEditing()
    gVim.Vars.FromInsertCmd = True
    Application.OnTime Now + (1 / 86400) * 0.1, "StopEditing"
End Sub

Private Sub StopEditing()
    Call stopVisualMode
    gVim.Vars.FromInsertCmd = False
End Sub
