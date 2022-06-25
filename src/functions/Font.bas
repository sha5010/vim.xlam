Attribute VB_Name = "F_Font"
Option Explicit
Option Private Module

Function increaseFontSize()
    Call repeatRegister("increaseFontSize")
    Call keystroke(False, Alt_ + H_, F_, G_)
End Function

Function decreaseFontSize()
    Call repeatRegister("decreaseFontSize")
    Call keystroke(False, Alt_ + H_, F_, K_)
End Function

Function changeFontName()
    Call keystroke(True, Alt_ + H_, F_, F_)
End Function

Function changeFontSize()
    Call keystroke(True, Alt_ + H_, F_, S_)
End Function

Function alignLeft()
    Call repeatRegister("alignLeft")
    Call keystroke(True, Alt_ + H_, A_, L_)
End Function

Function alignCenter()
    Call repeatRegister("alignCenter")
    Call keystroke(True, Alt_ + H_, A_, C_)
End Function

Function alignRight()
    Call repeatRegister("alignRight")
    Call keystroke(True, Alt_ + H_, A_, R_)
End Function

Function alignTop()
    Call repeatRegister("alignTop")
    Call keystroke(True, Alt_ + H_, A_, T_)
End Function

Function alignMiddle()
    Call repeatRegister("alignMiddle")
    Call keystroke(True, Alt_ + H_, A_, M_)
End Function

Function alignBottom()
    Call repeatRegister("alignBottom")
    Call keystroke(True, Alt_ + H_, A_, B_)
End Function

Function toggleBold()
    Call repeatRegister("toggleBold")
    Call keystroke(True, Ctrl_ + k2_)
End Function

Function toggleItalic()
    Call repeatRegister("toggleItalic")
    Call keystroke(True, Ctrl_ + k3_)
End Function

Function toggleUnderline()
    Call repeatRegister("toggleUnderline")
    Call keystroke(True, Ctrl_ + k4_)
End Function

Function showFontDialog()
    Call keystroke(True, Ctrl_ + k1_)
End Function

Function changeFontColor(Optional ByVal resultColor As cls_FontColor)
    Dim colorTable As Variant

    If TypeName(Selection) = "Nothing" Then
        Exit Function
    End If

    colorTable = Array(2, 1, 4, 3, 5, 6, 7, 8, 9, 10)

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
    End If

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

        Call repeatRegister("changeFontColor", resultColor)
    End If
End Function
