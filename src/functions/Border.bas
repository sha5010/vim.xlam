Attribute VB_Name = "F_Border"
Option Explicit
Option Private Module

Private Enum Mode
    APIsetBorder = 1
    APIdeleteBorder
    APItoggleBorder
End Enum

Private Function BorderAPI(ByVal OpMode As Mode, _
                           Optional ByVal Index As Variant = 0, _
                           Optional ByVal LineStyle As XlLineStyle = -1, _
                           Optional ByVal Weight As XlBorderWeight = -1, _
                           Optional ByVal ColorIndex As XlColorIndex = -1, _
                           Optional ByVal Color As Long = -1, _
                           Optional ByVal ThemeColor As XlThemeColor = -1, _
                           Optional ByVal TintAndShade As Double = 0)
    Dim arr As Variant
    Dim i As Variant
    Dim sameAll As Boolean

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If OpMode = APIdeleteBorder Then
        LineStyle = xlLineStyleNone
    End If

    If ColorIndex = -1 And Color = -1 And ThemeColor = -1 Then
        ColorIndex = xlColorIndexAutomatic

        If LineStyle = -1 Then
            LineStyle = xlContinuous
        End If

        If Weight = -1 Then
            Weight = xlThin
        End If
    End If

    Select Case TypeName(Index)
        Case "Variant()"
            arr = Index
        Case "Integer", "Long", "Byte"
            If Index = 0 Then
                arr = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            Else
                arr = Array(Index)
            End If
        Case Else
            Call debugPrint("Unexpected type: " & TypeName(Index), "BorderAPI")
    End Select

    If OpMode = APItoggleBorder Then
        sameAll = True

        For Each i In arr
            With Selection.Borders(i)
                sameAll = sameAll And (.LineStyle = LineStyle)
                sameAll = sameAll And (.Weight = Weight)
            End With
            If Not sameAll Then
                Exit For
            End If
        Next i

        If sameAll Then
            LineStyle = xlLineStyleNone
        End If
    End If

    For Each i In arr
        With Selection.Borders(i)
            If LineStyle <> -1 Then
                .LineStyle = LineStyle
            End If

            If .LineStyle <> xlLineStyleNone Then
                If Weight <> -1 Then
                    .Weight = Weight
                End If

                If ColorIndex <> -1 Then
                    .ColorIndex = ColorIndex
                ElseIf Color <> -1 Then
                    .Color = Color
                ElseIf ThemeColor <> -1 Then
                    .ThemeColor = ThemeColor
                    .TintAndShade = TintAndShade
                End If
            End If
        End With
    Next i
End Function

Function BorderColorAPI(Optional ByVal Index As Variant = 0, _
                        Optional ByVal resultColor As cls_FontColor)
    Dim colorTable As Variant

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    colorTable = Array(2, 1, 4, 3, 5, 6, 7, 8, 9, 10)

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
    End If

    If Not resultColor Is Nothing Then
        With resultColor
            If .IsNull Then
                Call BorderAPI(APIsetBorder, Index, ColorIndex:=xlColorIndexAutomatic)
            ElseIf .IsThemeColor Then
                Call BorderAPI(APIsetBorder, Index, ThemeColor:=colorTable(.ThemeColor - 1), TintAndShade:=.TintAndShade)
            Else
                Call BorderAPI(APIsetBorder, Index, Color:=.Color)
            End If
        End With

        Call repeatRegister("BorderColorAPI", Index, resultColor)
    End If
End Function

Function toggleBorderAround(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderAround", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, LineStyle:=LineStyle, Weight:=Weight)
End Function

Function toggleBorderLeft(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                          Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderLeft", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlEdgeLeft, LineStyle, Weight)
End Function

Function toggleBorderTop(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderTop", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlEdgeTop, LineStyle, Weight)
End Function

Function toggleBorderBottom(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderBottom", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlEdgeBottom, LineStyle, Weight)
End Function

Function toggleBorderRight(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderRight", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlEdgeRight, LineStyle, Weight)
End Function

Function toggleBorderInnerHorizontal(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                     Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderInnerHorizontal", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlInsideHorizontal, LineStyle, Weight)
End Function

Function toggleBorderInnerVertical(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                   Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderInnerVertical", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlInsideVertical, LineStyle, Weight)
End Function

Function toggleBorderInner(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderInner", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, Array(xlInsideHorizontal, xlInsideVertical), LineStyle, Weight)
End Function

Function toggleBorderDiagonalUp(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderDiagonalUp", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlDiagonalUp, LineStyle, Weight)
End Function

Function toggleBorderDiagonalDown(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                  Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderDiagonalDown", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, xlDiagonalDown, LineStyle, Weight)
End Function

Function toggleBorderAll(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin)
    Call repeatRegister("toggleBorderAll", LineStyle, Weight)
    Call BorderAPI(APItoggleBorder, Array( _
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
        xlInsideHorizontal, xlInsideVertical), LineStyle, Weight)
End Function


Function deleteBorderAround()
    Call repeatRegister("deleteBorderAround")
    Call BorderAPI(APIdeleteBorder)
End Function

Function deleteBorderLeft()
    Call repeatRegister("deleteBorderLeft")
    Call BorderAPI(APIdeleteBorder, xlEdgeLeft)
End Function

Function deleteBorderTop()
    Call repeatRegister("deleteBorderTop")
    Call BorderAPI(APIdeleteBorder, xlEdgeTop)
End Function

Function deleteBorderBottom()
    Call repeatRegister("deleteBorderBottom")
    Call BorderAPI(APIdeleteBorder, xlEdgeBottom)
End Function

Function deleteBorderRight()
    Call repeatRegister("deleteBorderRight")
    Call BorderAPI(APIdeleteBorder, xlEdgeRight)
End Function

Function deleteBorderInnerHorizontal()
    Call repeatRegister("deleteBorderInnerHorizontal")
    Call BorderAPI(APIdeleteBorder, xlInsideHorizontal)
End Function

Function deleteBorderInnerVertical()
    Call repeatRegister("deleteBorderInnerVertical")
    Call BorderAPI(APIdeleteBorder, xlInsideVertical)
End Function

Function deleteBorderInner()
    Call repeatRegister("deleteBorderInner")
    Call BorderAPI(APIdeleteBorder, Array(xlInsideHorizontal, xlInsideVertical))
End Function

Function deleteBorderDiagonalUp()
    Call repeatRegister("deleteBorderDiagonalUp")
    Call BorderAPI(APIdeleteBorder, xlDiagonalUp)
End Function

Function deleteBorderDiagonalDown()
    Call repeatRegister("deleteBorderDiagonalDown")
    Call BorderAPI(APIdeleteBorder, xlDiagonalDown)
End Function

Function deleteBorderAll()
    Call repeatRegister("deleteBorderAll")
    Call BorderAPI(APIdeleteBorder, Array( _
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
        xlInsideHorizontal, xlInsideVertical))
End Function


Function setBorderColorAround()
    Call BorderColorAPI
End Function

Function setBorderColorLeft()
    Call BorderColorAPI(xlEdgeLeft)
End Function

Function setBorderColorTop()
    Call BorderColorAPI(xlEdgeTop)
End Function

Function setBorderColorBottom()
    Call BorderColorAPI(xlEdgeBottom)
End Function

Function setBorderColorRight()
    Call BorderColorAPI(xlEdgeRight)
End Function

Function setBorderColorInnerHorizontal()
    Call BorderColorAPI(xlInsideHorizontal)
End Function

Function setBorderColorInnerVertical()
    Call BorderColorAPI(xlInsideVertical)
End Function

Function setBorderColorInner()
    Call BorderColorAPI(Array(xlInsideHorizontal, xlInsideVertical))
End Function

Function setBorderColorAll()
    Call BorderColorAPI(Array( _
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
        xlInsideHorizontal, xlInsideVertical))
End Function
