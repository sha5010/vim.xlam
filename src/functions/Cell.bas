Attribute VB_Name = "F_Cell"
Option Explicit
Option Private Module

Function cutCell()
    Call keystroke(True, Ctrl_ + X_)

    If TypeName(Selection) = "Range" Then
        Set gLastYanked = Selection
    End If
End Function

Function yankCell()
    Call keystroke(True, Ctrl_ + C_)

    If TypeName(Selection) = "Range" Then
        Set gLastYanked = Selection
    End If
End Function

Function yankFromUpCell()
    Call keystroke(True, Alt_ + H_, F_, I_, D_)
End Function

Function yankFromDownCell()
    Call keystroke(True, Alt_ + H_, F_, I_, U_)
End Function

Function yankFromLeftCell()
    Call keystroke(True, Alt_ + H_, F_, I_, R_)
End Function

Function yankFromRightCell()
    Call keystroke(True, Alt_ + H_, F_, I_, L_)
End Function

Function incrementText()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k6_)
    Next i

    Call unkeyupControlKeys
End Function

Function decrementText()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k5_)
    Next i

    Call unkeyupControlKeys
End Function

Function increaseDecimal()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k0_)
    Next i

    Call unkeyupControlKeys
End Function

Function decreaseDecimal()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k9_)
    Next i

    Call unkeyupControlKeys
End Function


Function insertCellsUp()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + Semicoron_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
End Function

Function insertCellsDown()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If Selection.Row < ActiveSheet.Rows.Count Then
        Selection.Offset(1, 0).Select
    End If

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + Semicoron_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
End Function

Function insertCellsLeft()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + Semicoron_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
End Function

Function insertCellsRight()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If Selection.Column < ActiveSheet.Columns.Count Then
        Selection.Offset(0, 1).Select
    End If

    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + Semicoron_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
End Function

Function deleteValue()
    Call keystroke(True, Delete_)
End Function

Function deleteToUp()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Minus_, U_, Enter_)

Catch:
    Application.ScreenUpdating = True
End Function

Function deleteToLeft()
    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Ctrl_ + Minus_, L_, Enter_)

Catch:
    Application.ScreenUpdating = True
End Function

Function toggleWrapText()
    Call keystroke(True, Alt_ + H_, W_)
End Function

Function toggleMergeCells()
    If TypeName(Selection) = "Range" Then
        If Not ActiveCell.MergeCells And Selection.Count = 1 Then
            Exit Function
        End If

        If ActiveCell.MergeCells Then
            Call keystroke(True, Alt_ + H_, M_, U_)
        Else
            Call keystroke(True, Alt_ + H_, M_, M_)
        End If
    End If
End Function

Function changeInteriorColor()
    Dim resultColor As cls_FontColor
    Dim colorTable As Variant

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    colorTable = Array(2, 1, 4, 3, 5, 6, 7, 8, 9, 10)
    Set resultColor = UF_ColorPicker.showColorPicker()

    If Not resultColor Is Nothing Then
        With Selection.Interior
            If resultColor.IsNull Then
                .ColorIndex = xlNone
            ElseIf resultColor.IsThemeColor Then
                .ThemeColor = colorTable(resultColor.ThemeColor - 1)
                .TintAndShade = resultColor.TintAndShade
            Else
                .Color = resultColor.Color
            End If
        End With
    End If
End Function

Function unionSelectCells()
    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If gExtendRange Is Nothing Then
        Set gExtendRange = Selection

    ElseIf Not gExtendRange.Parent Is ActiveSheet Then
        Call setStatusBarTemporarily("異なるシートで拡張選択はできないため、選択範囲は初期化されました。", 2)
        Set gExtendRange = Selection

    Else
        Set actCell = ActiveCell
        Set gExtendRange = Union2(gExtendRange, Selection)
        gExtendRange.Select
        actCell.Activate

    End If
End Function

Function exceptSelectCells()
    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If Not gExtendRange Is Nothing Then
        Set actCell = ActiveCell
        Set gExtendRange = Except2(gExtendRange, Selection)

        If Not gExtendRange Is Nothing Then
            gExtendRange.Select
        Else
            Call setStatusBarTemporarily("保存されている拡張選択範囲をクリアしました。", 2)
        End If
    End If
End Function

Function followHyperlinkOfActiveCell()
    On Error Resume Next

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If ActiveCell.Hyperlinks.Count > 0 Then
        ActiveCell.Hyperlinks(1).Follow
    ElseIf InStr(UCase(ActiveCell.Formula), "=HYPERLINK(") > 0 Then
        ActiveWorkbook.followHyperlink Split(ActiveCell.Formula, """")(1)
    End If
End Function
