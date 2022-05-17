Attribute VB_Name = "F_Font"
Option Explicit
Option Private Module

Function increaseFontSize()
    On Error GoTo Catch
'    Selection.Font.Size = Selection.Font.Size + 1
    Call keyupControlKeys
    'Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyG, 0, 0, 0
    keybd_event vbKeyG, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
End Function

Function decreaseFontSize()
    On Error GoTo Catch
'    Selection.Font.Size = Selection.Font.Size - 1
    Call keyupControlKeys
    'Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyK, 0, 0, 0
    keybd_event vbKeyK, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
End Function

Function changeFontName()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function changeFontSize()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyS, 0, 0, 0
    keybd_event vbKeyS, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function alignLeft()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYUP, 0
    keybd_event vbKeyL, 0, 0, 0
    keybd_event vbKeyL, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function alignCenter()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYUP, 0
    keybd_event vbKeyC, 0, 0, 0
    keybd_event vbKeyC, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function alignRight()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYUP, 0
    keybd_event vbKeyR, 0, 0, 0
    keybd_event vbKeyR, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function alignTop()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYUP, 0
    keybd_event vbKeyT, 0, 0, 0
    keybd_event vbKeyT, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function alignMiddle()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYUP, 0
    keybd_event vbKeyM, 0, 0, 0
    keybd_event vbKeyM, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function alignBottom()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyA, 0, 0, 0
    keybd_event vbKeyA, 0, KEYUP, 0
    keybd_event vbKeyB, 0, 0, 0
    keybd_event vbKeyB, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function toggleBold()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKey2, 0, 0, 0
    keybd_event vbKey2, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function toggleItalic()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKey3, 0, 0, 0
    keybd_event vbKey3, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function toggleUnderline()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKey4, 0, 0, 0
    keybd_event vbKey4, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function showFontDialog()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKey1, 0, 0, 0
    keybd_event vbKey1, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function changeFontColor()
    Dim resultColor As cls_FontColor
    Dim colorTable As Variant

    colorTable = Array(2, 1, 4, 3, 5, 6, 7, 8, 9, 10)
    Set resultColor = UF_ColorPicker.showColorPicker()

    If Not resultColor Is Nothing Then
        With Selection.Font
            If resultColor.IsNull Then
                .ColorIndex = xlAutomatic
            ElseIf resultColor.IsThemeColor Then
                .ThemeColor = colorTable(resultColor.ThemeColor - 1)
                .TintAndShade = resultColor.TintAndShade
            Else
                .Color = resultColor.Color
            End If
        End With
    End If
End Function
