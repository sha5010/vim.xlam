Attribute VB_Name = "F_Border"
Option Explicit
Option Private Module

Private Enum eBorderMode
    BorderModeSet = 1
    BorderModeDelete
    BorderModeToggle
End Enum

Private Function BorderInner(ByVal OpMode As eBorderMode, _
                    Optional ByVal Index As Variant = 0, _
                    Optional ByVal LineStyle As XlLineStyle = -1, _
                    Optional ByVal Weight As XlBorderWeight = -1, _
                    Optional ByVal ColorIndex As XlColorIndex = -1, _
                    Optional ByVal Color As Long = -1, _
                    Optional ByVal ThemeColor As XlThemeColor = -1, _
                    Optional ByVal TintAndShade As Double = 0)
    On Error GoTo Catch

    Dim arr As Variant
    Dim i As Variant
    Dim sameAll As Boolean

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If OpMode = eBorderMode.BorderModeDelete Then
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
            Call DebugPrint("Unexpected type: " & TypeName(Index), "BorderInner")
    End Select

    If OpMode = eBorderMode.BorderModeToggle Then
        sameAll = True

        For Each i In arr
            If Not ((i = xlInsideHorizontal And Selection.Rows.Count = 1) Or (i = xlInsideVertical And Selection.Columns.Count = 1)) Then
                With Selection.Borders(i)
                    sameAll = sameAll And (.LineStyle = LineStyle)
                    sameAll = sameAll And (.Weight = Weight)
                End With
                If Not sameAll Then
                    Exit For
                End If
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
    Exit Function

Catch:
    Call ErrorHandler("BorderInner")
End Function

Private Function BorderColorInner(Optional ByVal Index As Variant = 0, _
                                  Optional ByVal resultColor As cls_FontColor)
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With resultColor
            If .IsNull Then
                Call BorderInner(eBorderMode.BorderModeSet, Index, ColorIndex:=xlColorIndexAutomatic)
            ElseIf .IsThemeColor Then
                Call BorderInner(eBorderMode.BorderModeSet, Index, ThemeColor:=.ThemeColor, TintAndShade:=.TintAndShade)
            Else
                Call BorderInner(eBorderMode.BorderModeSet, Index, Color:=.Color)
            End If
        End With

        Call RepeatRegister("BorderColorInner", Index, resultColor)
    End If
    Exit Function

Catch:
    Call ErrorHandler("BorderColorInner")
End Function

Function ToggleBorderAround(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderAround", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, LineStyle:=LineStyle, Weight:=Weight)
End Function

Function ToggleBorderLeft(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                          Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderLeft", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlEdgeLeft, LineStyle, Weight)
End Function

Function ToggleBorderTop(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderTop", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlEdgeTop, LineStyle, Weight)
End Function

Function ToggleBorderBottom(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderBottom", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlEdgeBottom, LineStyle, Weight)
End Function

Function ToggleBorderRight(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderRight", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlEdgeRight, LineStyle, Weight)
End Function

Function ToggleBorderInnerHorizontal(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                     Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderInnerHorizontal", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlInsideHorizontal, LineStyle, Weight)
End Function

Function ToggleBorderInnerVertical(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                   Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderInnerVertical", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlInsideVertical, LineStyle, Weight)
End Function

Function ToggleBorderInner(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderInner", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, Array(xlInsideHorizontal, xlInsideVertical), LineStyle, Weight)
End Function

Function ToggleBorderDiagonalUp(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderDiagonalUp", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlDiagonalUp, LineStyle, Weight)
End Function

Function ToggleBorderDiagonalDown(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                  Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderDiagonalDown", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, xlDiagonalDown, LineStyle, Weight)
End Function

Function ToggleBorderAll(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call RepeatRegister("ToggleBorderAll", LineStyle, Weight)
    Call BorderInner(eBorderMode.BorderModeToggle, Array( _
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
        xlInsideHorizontal, xlInsideVertical), LineStyle, Weight)
End Function


Function DeleteBorderAround(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderAround")
    Call BorderInner(eBorderMode.BorderModeDelete)
End Function

Function DeleteBorderLeft(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderLeft")
    Call BorderInner(eBorderMode.BorderModeDelete, xlEdgeLeft)
End Function

Function DeleteBorderTop(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderTop")
    Call BorderInner(eBorderMode.BorderModeDelete, xlEdgeTop)
End Function

Function DeleteBorderBottom(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderBottom")
    Call BorderInner(eBorderMode.BorderModeDelete, xlEdgeBottom)
End Function

Function DeleteBorderRight(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderRight")
    Call BorderInner(eBorderMode.BorderModeDelete, xlEdgeRight)
End Function

Function DeleteBorderInnerHorizontal(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderInnerHorizontal")
    Call BorderInner(eBorderMode.BorderModeDelete, xlInsideHorizontal)
End Function

Function DeleteBorderInnerVertical(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderInnerVertical")
    Call BorderInner(eBorderMode.BorderModeDelete, xlInsideVertical)
End Function

Function DeleteBorderInner(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderInner")
    Call BorderInner(eBorderMode.BorderModeDelete, Array(xlInsideHorizontal, xlInsideVertical))
End Function

Function DeleteBorderDiagonalUp(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderDiagonalUp")
    Call BorderInner(eBorderMode.BorderModeDelete, xlDiagonalUp)
End Function

Function DeleteBorderDiagonalDown(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderDiagonalDown")
    Call BorderInner(eBorderMode.BorderModeDelete, xlDiagonalDown)
End Function

Function DeleteBorderAll(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteBorderAll")
    Call BorderInner(eBorderMode.BorderModeDelete, Array( _
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
        xlInsideHorizontal, xlInsideVertical))
End Function


Function SetBorderColorAround(Optional ByVal g As String) As Boolean
    Call BorderColorInner
End Function

Function SetBorderColorLeft(Optional ByVal g As String) As Boolean
    Call BorderColorInner(xlEdgeLeft)
End Function

Function SetBorderColorTop(Optional ByVal g As String) As Boolean
    Call BorderColorInner(xlEdgeTop)
End Function

Function SetBorderColorBottom(Optional ByVal g As String) As Boolean
    Call BorderColorInner(xlEdgeBottom)
End Function

Function SetBorderColorRight(Optional ByVal g As String) As Boolean
    Call BorderColorInner(xlEdgeRight)
End Function

Function SetBorderColorInnerHorizontal(Optional ByVal g As String) As Boolean
    Call BorderColorInner(xlInsideHorizontal)
End Function

Function SetBorderColorInnerVertical(Optional ByVal g As String) As Boolean
    Call BorderColorInner(xlInsideVertical)
End Function

Function SetBorderColorInner(Optional ByVal g As String) As Boolean
    Call BorderColorInner(Array(xlInsideHorizontal, xlInsideVertical))
End Function

Function SetBorderColorAll(Optional ByVal g As String) As Boolean
    Call BorderColorInner(Array( _
        xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
        xlInsideHorizontal, xlInsideVertical))
End Function
