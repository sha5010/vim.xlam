Attribute VB_Name = "F_Font"
Option Explicit
Option Private Module

Function increaseFontSize()
    Call repeatRegister("increaseFontSize")
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, F_, G_)
End Function

Function decreaseFontSize()
    Call repeatRegister("decreaseFontSize")
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, F_, K_)
End Function

Function changeFontName()
    Call keystroke(True, Alt_ + H_, F_, F_)
End Function

Function changeFontSize()
    Call keystroke(True, Alt_ + H_, F_, S_)
End Function

Function alignLeft()
    Call repeatRegister("alignLeft")
    Call stopVisualMode

    'Check excel version
    On Error GoTo Excel2016
    If CDbl(Application.Version) >= 16 Then
        'Raise error in Excel 2016 (Concat exists in Excel 2019 and later)
        WorksheetFunction.Concat ""
    End If

    'Default
    Call keystroke(True, Alt_ + H_, A_, L_)
    Exit Function

Excel2016:
    Call keystroke(True, Alt_ + H_, L_, k1_)
End Function

Function alignCenter()
    Call repeatRegister("alignCenter")
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, A_, C_)
End Function

Function alignRight()
    Call repeatRegister("alignRight")
    Call stopVisualMode

    'Check excel version
    On Error GoTo Excel2016
    If CDbl(Application.Version) >= 16 Then
        'Raise error in Excel 2016 (Concat exists in Excel 2019 and later)
        WorksheetFunction.Concat ""
    End If

    'Default
    Call keystroke(True, Alt_ + H_, A_, R_)
    Exit Function

Excel2016:
    'Excel 2013 and earlier
    Call keystroke(True, Alt_ + H_, R_)
End Function

Function alignTop()
    Call repeatRegister("alignTop")
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, A_, T_)
End Function

Function alignMiddle()
    Call repeatRegister("alignMiddle")
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, A_, M_)
End Function

Function alignBottom()
    Call repeatRegister("alignBottom")
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, A_, B_)
End Function

Function toggleBold()
    Call repeatRegister("toggleBold")
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + k2_)
End Function

Function toggleItalic()
    Call repeatRegister("toggleItalic")
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + k3_)
End Function

Function toggleUnderline()
    Call repeatRegister("toggleUnderline")
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + k4_)
End Function

Function toggleStrikethrough()
    Call repeatRegister("toggleStrikethrough")
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + k5_)
End Function

Function changeFormat()
    Call keystroke(True, Alt_ + H_, N_, Down_, Down_)
End Function

Function showFontDialog()
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + k1_)
End Function

Function changeFontColor(Optional ByVal resultColor As cls_FontColor)
    If TypeName(Selection) = "Nothing" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
    End If

    If Not resultColor Is Nothing Then
        With Selection.Font
            If resultColor.IsNull Then
                .ColorIndex = xlAutomatic
            ElseIf resultColor.IsThemeColor Then
                .ThemeColor = resultColor.ThemeColor
                .TintAndShade = resultColor.TintAndShade
            Else
                .Color = resultColor.Color
            End If
        End With

        Call repeatRegister("changeFontColor", resultColor)
        Call stopVisualMode
    End If
End Function
