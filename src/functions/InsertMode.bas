Attribute VB_Name = "F_InsertMode"
Option Explicit
Option Private Module

Function insertWithIME()
    Call keyupControlKeys
    Call releaseShiftKeys

    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        keybd_event vbKeyReturn, 0, 0, 0
        keybd_event vbKeyReturn, 0, KEYUP, 0
    Else
        keybd_event vbKeyF2, 0, 0, 0
        keybd_event vbKeyF2, 0, KEYUP, 0
    End If

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyHome, 0, EXTENDED_KEY Or 0, 0
    keybd_event vbKeyHome, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event KANJI, 0, 0, 0
    keybd_event KANJI, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function insertWithoutIME()
    Call keyupControlKeys
    Call releaseShiftKeys

    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        keybd_event vbKeyReturn, 0, 0, 0
        keybd_event vbKeyReturn, 0, KEYUP, 0
    Else
        keybd_event vbKeyF2, 0, 0, 0
        keybd_event vbKeyF2, 0, KEYUP, 0
    End If

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyHome, 0, EXTENDED_KEY Or 0, 0
    keybd_event vbKeyHome, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function appendWithIME()
    Call keyupControlKeys
    Call releaseShiftKeys

    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        keybd_event vbKeyReturn, 0, 0, 0
        keybd_event vbKeyReturn, 0, KEYUP, 0
        keybd_event vbKeyRight, 0, 0, 0
        keybd_event vbKeyRight, 0, KEYUP, 0
    Else
        keybd_event vbKeyF2, 0, 0, 0
        keybd_event vbKeyF2, 0, KEYUP, 0
    End If

    keybd_event KANJI, 0, 0, 0
    keybd_event KANJI, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function appendWithoutIME()
    Call keyupControlKeys
    Call releaseShiftKeys

    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        keybd_event vbKeyReturn, 0, 0, 0
        keybd_event vbKeyReturn, 0, KEYUP, 0
        keybd_event vbKeyRight, 0, 0, 0
        keybd_event vbKeyRight, 0, KEYUP, 0
    Else
        keybd_event vbKeyF2, 0, 0, 0
        keybd_event vbKeyF2, 0, KEYUP, 0
    End If

    Call unkeyupControlKeys
End Function

Function substituteWithIME()
    Call keyupControlKeys
    Call releaseShiftKeys

    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        keybd_event vbKeyReturn, 0, 0, 0
        keybd_event vbKeyReturn, 0, KEYUP, 0
    Else
        keybd_event vbKeyF2, 0, 0, 0
        keybd_event vbKeyF2, 0, KEYUP, 0
        keybd_event vbKeyShift, 0, 0, 0
        keybd_event vbKeyControl, 0, 0, 0
        keybd_event vbKeyHome, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyHome, 0, EXTENDED_KEY Or KEYUP, 0
        keybd_event vbKeyControl, 0, KEYUP, 0
        keybd_event vbKeyShift, 0, KEYUP, 0
    End If

    keybd_event vbKeyDelete, 0, EXTENDED_KEY Or 0, 0
    keybd_event vbKeyDelete, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event KANJI, 0, 0, 0
    keybd_event KANJI, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function substituteWithoutIME()
    Call keyupControlKeys
    Call releaseShiftKeys


    If VarType(Selection) = vbObject Then
        Call temporarilyDisableVim
        keybd_event vbKeyReturn, 0, 0, 0
        keybd_event vbKeyReturn, 0, KEYUP, 0
    Else
        keybd_event vbKeyF2, 0, 0, 0
        keybd_event vbKeyF2, 0, KEYUP, 0
        keybd_event vbKeyShift, 0, 0, 0
        keybd_event vbKeyControl, 0, 0, 0
        keybd_event vbKeyHome, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyHome, 0, EXTENDED_KEY Or KEYUP, 0
        keybd_event vbKeyControl, 0, KEYUP, 0
        keybd_event vbKeyShift, 0, KEYUP, 0
    End If

    keybd_event vbKeyDelete, 0, EXTENDED_KEY Or 0, 0
    keybd_event vbKeyDelete, 0, EXTENDED_KEY Or KEYUP, 0

    Call unkeyupControlKeys

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

