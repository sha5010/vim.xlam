Attribute VB_Name = "F_InsertMode"
Option Explicit
Option Private Module

Function insertWithIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + Home_, Kanji_)
    Else
        Call keystroke(True, F2_, Ctrl_ + Home_, Kanji_)
        Call StartEditing
    End If
    Application.OnTime Now + 0.1 / 86400, "disableIME"
End Function

Function insertWithoutIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + Home_)
    Else
        Call keystroke(True, F2_, Ctrl_ + Home_)
        Call StartEditing
    End If
    Application.OnTime Now + 0.1 / 86400, "disableIME"
End Function

Function appendWithIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + End_, Kanji_)
    Else
        Call keystroke(True, F2_, Kanji_)
        Call StartEditing
    End If
    Application.OnTime Now + 0.1 / 86400, "disableIME"
End Function

Function appendWithoutIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + End_)
    Else
        Call keystroke(True, F2_)
        Call StartEditing
    End If
    Application.OnTime Now + 0.1 / 86400, "disableIME"
End Function

Function substituteWithIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Enter_, Delete_, Kanji_)
    Else
        Call keystroke(True, BackSpace_, F2_, Kanji_)
        Call StartEditing
    End If
    Application.OnTime Now + 0.1 / 86400, "disableIME"
End Function

Function substituteWithoutIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Enter_, Delete_)
    Else
        Call keystroke(True, BackSpace_, F2_)
        Call StartEditing
    End If
    Application.OnTime Now + 0.1 / 86400, "disableIME"
End Function

Function insertFollowLangMode()
    If gLangJa Then
        Call insertWithIME
    Else
        Call insertWithoutIME
    End If
End Function

Function insertNotFollowLangMode()
    If Not gLangJa Then
        Call insertWithIME
    Else
        Call insertWithoutIME
    End If
End Function

Function appendFollowLangMode()
    If gLangJa Then
        Call appendWithIME
    Else
        Call appendWithoutIME
    End If
End Function

Function appendNotFollowLangMode()
    If Not gLangJa Then
        Call appendWithIME
    Else
        Call appendWithoutIME
    End If
End Function

Function substituteFollowLangMode()
    If gLangJa Then
        Call substituteWithIME
    Else
        Call substituteWithoutIME
    End If
End Function

Function substituteNotFollowLangMode()
    If Not gLangJa Then
        Call substituteWithIME
    Else
        Call substituteWithoutIME
    End If
End Function

Sub StartEditing()
    Call X.StartEditing
    Application.OnTime Now + (1 / 86400) * 0.1, "StopEditing"
End Sub

Sub StopEditing()
    Call X.StopEditing
End Sub
