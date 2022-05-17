Attribute VB_Name = "F_Shapes"
Option Explicit
Option Private Module

Function changeShapeFillColor()
    Dim shp As ShapeRange
    Dim resultColor As cls_FontColor

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If
    
    Set shp = Selection.ShapeRange
    Set resultColor = UF_ColorPicker.showColorPicker()
    
    If Not resultColor Is Nothing Then
        With shp.Fill
            If resultColor.IsNull Then
                .Visible = msoFalse
            ElseIf resultColor.IsThemeColor Then
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = resultColor.ThemeColor
                .ForeColor.TintAndShade = resultColor.TintAndShade
            Else
                .Visible = msoTrue
                .ForeColor.RGB = resultColor.Color
            End If
        End With
    End If
    
    Set shp = Nothing
End Function

Function changeShapeFontColor()
    Dim shp As ShapeRange
    Dim resultColor As cls_FontColor
    
    If VarType(Selection) <> vbObject Then
        Exit Function
    End If
    
    Set shp = Selection.ShapeRange
    Set resultColor = UF_ColorPicker.showColorPicker()
    
    If Not resultColor Is Nothing Then
        With shp.TextFrame2.TextRange.Font.Fill.ForeColor
            If resultColor.IsNull Then
                .RGB = 0
            ElseIf resultColor.IsThemeColor Then
                .ObjectThemeColor = resultColor.ThemeColor
                .TintAndShade = resultColor.TintAndShade
            Else
                .RGB = resultColor.Color
            End If
        End With
    End If
    
    Set shp = Nothing
End Function

Function changeShapeBorderColor(Optional garbage As String) As Boolean
    Dim shp As ShapeRange
    Dim resultColor As cls_FontColor

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If
    
    Set shp = Selection.ShapeRange
    Set resultColor = UF_ColorPicker.showColorPicker()
    
    If Not resultColor Is Nothing Then
        With shp.Line
            If resultColor.IsNull Then
                .Visible = msoFalse
            ElseIf resultColor.IsThemeColor Then
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = resultColor.ThemeColor
                .ForeColor.TintAndShade = resultColor.TintAndShade
            Else
                .Visible = msoTrue
                .ForeColor.RGB = resultColor.Color
            End If
        End With
    End If
    
    Set shp = Nothing
    changeShapeBorderColor = True
End Function

