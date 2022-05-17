Attribute VB_Name = "F_Cell"
Option Explicit
Option Private Module

Function cutCell()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyX, 0, 0, 0
    keybd_event vbKeyX, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys

    If TypeName(Selection) = "Range" Then
        Set gLastYanked = Selection
    End If
End Function

Function yankCell()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyC, 0, 0, 0
    keybd_event vbKeyC, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0

    Call unkeyupControlKeys

    If TypeName(Selection) = "Range" Then
        Set gLastYanked = Selection
    End If
End Function

Function yankFromUpCell()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyI, 0, 0, 0
    keybd_event vbKeyI, 0, KEYUP, 0
    keybd_event vbKeyD, 0, 0, 0
    keybd_event vbKeyD, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function yankFromDownCell()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyI, 0, 0, 0
    keybd_event vbKeyI, 0, KEYUP, 0
    keybd_event vbKeyU, 0, 0, 0
    keybd_event vbKeyU, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function yankFromLeftCell()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyI, 0, 0, 0
    keybd_event vbKeyI, 0, KEYUP, 0
    keybd_event vbKeyR, 0, 0, 0
    keybd_event vbKeyR, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function yankFromRightCell()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyF, 0, 0, 0
    keybd_event vbKeyF, 0, KEYUP, 0
    keybd_event vbKeyI, 0, 0, 0
    keybd_event vbKeyI, 0, KEYUP, 0
    keybd_event vbKeyL, 0, 0, 0
    keybd_event vbKeyL, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function incrementText()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        keybd_event vbKeyMenu, 0, 0, 0
        keybd_event vbKeyH, 0, 0, 0
        keybd_event vbKeyH, 0, KEYUP, 0
        keybd_event vbKeyMenu, 0, KEYUP, 0
        keybd_event vbKey6, 0, 0, 0
        keybd_event vbKey6, 0, KEYUP, 0
    Next i

    Call unkeyupControlKeys
End Function

Function decrementText()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        keybd_event vbKeyMenu, 0, 0, 0
        keybd_event vbKeyH, 0, 0, 0
        keybd_event vbKeyH, 0, KEYUP, 0
        keybd_event vbKeyMenu, 0, KEYUP, 0
        keybd_event vbKey5, 0, 0, 0
        keybd_event vbKey5, 0, KEYUP, 0
    Next i

    Call unkeyupControlKeys
End Function

Function increaseDecimal()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        keybd_event vbKeyMenu, 0, 0, 0
        keybd_event vbKeyH, 0, 0, 0
        keybd_event vbKeyH, 0, KEYUP, 0
        keybd_event vbKeyMenu, 0, KEYUP, 0
        keybd_event vbKey0, 0, 0, 0
        keybd_event vbKey0, 0, KEYUP, 0
    Next i

    Call unkeyupControlKeys
End Function

Function decreaseDecimal()
    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        keybd_event vbKeyMenu, 0, 0, 0
        keybd_event vbKeyH, 0, 0, 0
        keybd_event vbKeyH, 0, KEYUP, 0
        keybd_event vbKeyMenu, 0, KEYUP, 0
        keybd_event vbKey9, 0, 0, 0
        keybd_event vbKey9, 0, KEYUP, 0
    Next i

    Call unkeyupControlKeys
End Function


Function insertCellsUp()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyShift, 0, 0, 0
    keybd_event &HBB, 0, 0, 0
    keybd_event &HBB, 0, KEYUP, 0
    keybd_event vbKeyShift, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event vbKeyD, 0, 0, 0
    keybd_event vbKeyD, 0, KEYUP, 0
    keybd_event vbKeyReturn, 0, 0, 0
    keybd_event vbKeyReturn, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
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

    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyShift, 0, 0, 0
    keybd_event &HBB, 0, 0, 0
    keybd_event &HBB, 0, KEYUP, 0
    keybd_event vbKeyShift, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event vbKeyD, 0, 0, 0
    keybd_event vbKeyD, 0, KEYUP, 0
    keybd_event vbKeyReturn, 0, 0, 0
    keybd_event vbKeyReturn, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
    Application.ScreenUpdating = True
End Function

Function insertCellsLeft()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyShift, 0, 0, 0
    keybd_event &HBB, 0, 0, 0
    keybd_event &HBB, 0, KEYUP, 0
    keybd_event vbKeyShift, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event vbKeyI, 0, 0, 0
    keybd_event vbKeyI, 0, KEYUP, 0
    keybd_event vbKeyReturn, 0, 0, 0
    keybd_event vbKeyReturn, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
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

    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeyShift, 0, 0, 0
    keybd_event &HBB, 0, 0, 0
    keybd_event &HBB, 0, KEYUP, 0
    keybd_event vbKeyShift, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event vbKeyI, 0, 0, 0
    keybd_event vbKeyI, 0, KEYUP, 0
    keybd_event vbKeyReturn, 0, 0, 0
    keybd_event vbKeyReturn, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
    Application.ScreenUpdating = True
End Function

Function deleteValue()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyDelete, 0, 0, 0
    keybd_event vbKeyDelete, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function deleteToUp()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeySubtract, 0, 0, 0
    keybd_event vbKeySubtract, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event vbKeyU, 0, 0, 0
    keybd_event vbKeyU, 0, KEYUP, 0
    keybd_event vbKeyReturn, 0, 0, 0
    keybd_event vbKeyReturn, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
    Application.ScreenUpdating = True
End Function

Function deleteToLeft()
    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyControl, 0, 0, 0
    keybd_event vbKeySubtract, 0, 0, 0
    keybd_event vbKeySubtract, 0, KEYUP, 0
    keybd_event vbKeyControl, 0, KEYUP, 0
    keybd_event vbKeyL, 0, 0, 0
    keybd_event vbKeyL, 0, KEYUP, 0
    keybd_event vbKeyReturn, 0, 0, 0
    keybd_event vbKeyReturn, 0, KEYUP, 0

Catch:
    Call unkeyupControlKeys
    Application.ScreenUpdating = True
End Function

Function toggleWrapText()
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event vbKeyMenu, 0, 0, 0
    keybd_event vbKeyH, 0, 0, 0
    keybd_event vbKeyH, 0, KEYUP, 0
    keybd_event vbKeyMenu, 0, KEYUP, 0
    keybd_event vbKeyW, 0, 0, 0
    keybd_event vbKeyW, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function toggleMergeCells()
    If TypeName(Selection) = "Range" Then
        If Not ActiveCell.MergeCells And Selection.Count = 1 Then
            Exit Function
        End If

        Call keyupControlKeys
        Call releaseShiftKeys

        keybd_event vbKeyMenu, 0, 0, 0
        keybd_event vbKeyH, 0, 0, 0
        keybd_event vbKeyH, 0, KEYUP, 0
        keybd_event vbKeyMenu, 0, KEYUP, 0
        keybd_event vbKeyM, 0, 0, 0
        keybd_event vbKeyM, 0, KEYUP, 0

        If ActiveCell.MergeCells Then
            keybd_event vbKeyU, 0, 0, 0
            keybd_event vbKeyU, 0, KEYUP, 0
        Else
            keybd_event vbKeyM, 0, 0, 0
            keybd_event vbKeyM, 0, KEYUP, 0
        End If

        Call unkeyupControlKeys
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

