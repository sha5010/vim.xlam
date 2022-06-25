Attribute VB_Name = "F_InsertMode"
Option Explicit
Option Private Module

Function insertWithIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + Home_, Kanji_)
    Else
        Call keystroke(True, F2_, Ctrl_ + Home_, Kanji_)
    End If
End Function

Function insertWithoutIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + Home_)
    Else
        Call keystroke(True, F2_, Ctrl_ + Home_)
    End If
End Function

Function appendWithIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + End_, Kanji_)
    Else
        Call keystroke(True, F2_, Kanji_)
    End If
End Function

Function appendWithoutIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Space_, BackSpace_, Ctrl_ + End_)
    Else
        Call keystroke(True, F2_)
    End If
End Function

Function substituteWithIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Enter_, Delete_, Kanji_)
    Else
        Call keystroke(True, F2_, Ctrl_ + Shift_ + Home_, Delete_, Kanji_)
    End If
End Function

Function substituteWithoutIME()
    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        Call keystroke(True, Enter_, Delete_)
    Else
        Call keystroke(True, F2_, Ctrl_ + Shift_ + Home_, Delete_)
    End If
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
