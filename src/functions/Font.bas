Attribute VB_Name = "F_Font"
Option Explicit
Option Private Module

Function IncreaseFontSize(Optional ByVal g As String) As Boolean
    Call RepeatRegister("IncreaseFontSize")
    Call StopVisualMode

    Dim i As Long
    For i = 1 To gVim.Count1
        Application.CommandBars.ExecuteMso "GrowFont"
    Next i
End Function

Function DecreaseFontSize(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DecreaseFontSize")
    Call StopVisualMode

    Dim i As Long
    For i = 1 To gVim.Count1
        Application.CommandBars.ExecuteMso "ShrinkFont"
    Next i
End Function

Function ChangeFontName(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_, H_, F_, F_)
End Function

Function ChangeFontSize(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_, H_, F_, S_)
End Function

Function AlignLeft(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignLeft")
    Call StopVisualMode
    Selection.HorizontalAlignment = xlLeft
End Function

Function AlignCenter(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignCenter")
    Call StopVisualMode
    Selection.HorizontalAlignment = xlCenter
End Function

Function AlignRight(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignRight")
    Call StopVisualMode
    Selection.HorizontalAlignment = xlRight
End Function

Function AlignTop(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignTop")
    Call StopVisualMode
    Selection.VerticalAlignment = xlTop
End Function

Function AlignMiddle(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignMiddle")
    Call StopVisualMode
    Selection.VerticalAlignment = xlCenter
End Function

Function AlignBottom(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignBottom")
    Call StopVisualMode
    Selection.VerticalAlignment = xlBottom
End Function

Function ToggleBold(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleBold")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k2_)
End Function

Function ToggleItalic(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleItalic")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k3_)
End Function

Function ToggleUnderline(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleUnderline")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k4_)
End Function

Function ToggleStrikethrough(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleStrikethrough")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k5_)
End Function

Function ChangeFormat(Optional ByVal g As String) As Boolean
    Application.CommandBars.ExecuteMso "NumberFormatsDialog"
End Function

Function showFontDialog(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k1_)
End Function

Function ChangeFontColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) = "Nothing" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
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

        Call RepeatRegister("ChangeFontColor", resultColor)
        Call StopVisualMode
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeFontColor")
End Function
