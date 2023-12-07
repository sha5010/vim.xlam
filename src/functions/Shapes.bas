Attribute VB_Name = "F_Shapes"
Option Explicit
Option Private Module

Function ChangeShapeFillColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With shp.Fill
            If resultColor.IsNull Then
                .Visible = msoFalse
            ElseIf resultColor.IsThemeColor Then
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = resultColor.ObjectThemeColor
                .ForeColor.TintAndShade = resultColor.TintAndShade
            Else
                .Visible = msoTrue
                .ForeColor.RGB = resultColor.Color
            End If

            Call RepeatRegister("ChangeShapeFillColor", resultColor)
        End With
    End If

    Set shp = Nothing
    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeFillColor")
End Function

Function ChangeShapeFontColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With shp.TextFrame2.TextRange.Font.Fill.ForeColor
            If resultColor.IsNull Then
                .RGB = 0
            ElseIf resultColor.IsThemeColor Then
                .ObjectThemeColor = resultColor.ObjectThemeColor
                .TintAndShade = resultColor.TintAndShade
            Else
                .RGB = resultColor.Color
            End If

            Call RepeatRegister("ChangeShapeFontColor", resultColor)
        End With
    End If

    Set shp = Nothing
    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeFontColor")
End Function

Function ChangeShapeBorderColor(Optional garbage As String, _
                                Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        ChangeShapeBorderColor = True
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With shp.Line
            If resultColor.IsNull Then
                .Visible = msoFalse
            ElseIf resultColor.IsThemeColor Then
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = resultColor.ObjectThemeColor
                .ForeColor.TintAndShade = resultColor.TintAndShade
            Else
                .Visible = msoTrue
                .ForeColor.RGB = resultColor.Color
            End If

            Call RepeatRegister("ChangeShapeBorderColor", "", resultColor)
        End With
    End If

    Set shp = Nothing
    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeBorderColor")
End Function

Function NextShape(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long
    Dim cnt As Long
    Dim shp As Shape

    If VarType(Selection) = vbObject Then
        For i = 1 To gVim.Count1
            Call KeyStroke(Tab_)
        Next i
    Else
        cnt = ActiveSheet.Shapes.Count
        If cnt = 0 Then
            Exit Function
        End If
        ActiveSheet.Shapes((gVim.Count1 - 1) Mod cnt + 1).Select
    End If
    Exit Function

Catch:
    Call ErrorHandler("NextShape")
End Function

Function PrevShape(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long
    Dim cnt As Long
    Dim shp As Shape

    If VarType(Selection) = vbObject Then
        For i = 1 To gVim.Count1
            Call KeyStroke(Shift_ + Tab_)
        Next i
    Else
        cnt = ActiveSheet.Shapes.Count
        If cnt = 0 Then
            Exit Function
        End If
        ActiveSheet.Shapes(cnt - (gVim.Count1 - 1) Mod cnt).Select
    End If
    Exit Function

Catch:
    Call ErrorHandler("PrevShape")
End Function
