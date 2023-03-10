Attribute VB_Name = "F_Shapes"
Option Explicit
Option Private Module

Function changeShapeFillColor(Optional ByVal resultColor As cls_FontColor)
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
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

            Call repeatRegister("changeShapeFillColor", resultColor)
        End With
    End If

    Set shp = Nothing
    Exit Function

Catch:
    Call errorHandler("changeShapeFillColor")
End Function

Function changeShapeFontColor(Optional ByVal resultColor As cls_FontColor)
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
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

            Call repeatRegister("changeShapeFontColor", resultColor)
        End With
    End If

    Set shp = Nothing
    Exit Function

Catch:
    Call errorHandler("changeShapeFontColor")
End Function

Function changeShapeBorderColor(Optional garbage As String, _
                                Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
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

            Call repeatRegister("changeShapeBorderColor", "", resultColor)
        End With
    End If

    Set shp = Nothing
    changeShapeBorderColor = True
    Exit Function

Catch:
    Call errorHandler("changeShapeBorderColor")
End Function
