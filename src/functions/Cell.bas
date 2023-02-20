Attribute VB_Name = "F_Cell"
Option Explicit
Option Private Module

Private Enum searchMode
    TopToBottom = 1
    LeftToRight
    BottomToTop
    RightToLeft
End Enum

Function cutCell()
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + X_)

    If TypeName(Selection) = "Range" Then
        Set gLastYanked = Selection
    End If
End Function

Function yankCell()
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + C_)

    If TypeName(Selection) = "Range" Then
        Set gLastYanked = Selection
    End If
End Function

Function yankFromUpCell()
    Call repeatRegister("yankFromUpCell")
    Call keystroke(True, Alt_ + H_, F_, I_, D_)
End Function

Function yankFromDownCell()
    Call repeatRegister("yankFromDownCell")
    Call keystroke(True, Alt_ + H_, F_, I_, U_)
End Function

Function yankFromLeftCell()
    Call repeatRegister("yankFromLeftCell")
    Call keystroke(True, Alt_ + H_, F_, I_, R_)
End Function

Function yankFromRightCell()
    Call repeatRegister("yankFromRightCell")
    Call keystroke(True, Alt_ + H_, F_, I_, L_)
End Function

Function incrementText()
    Call repeatRegister("incrementText")
    Call stopVisualMode

    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k6_)
    Next i

    Call unkeyupControlKeys
End Function

Function decrementText()
    Call repeatRegister("decrementText")
    Call stopVisualMode

    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k5_)
    Next i

    Call unkeyupControlKeys
End Function

Function increaseDecimal()
    Call repeatRegister("increaseDecimal")
    Call stopVisualMode

    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gCount
        Call keystrokeWithoutKeyup(Alt_ + H_, k0_)
    Next i

    Call unkeyupControlKeys
End Function

Function decreaseDecimal()
    Call repeatRegister("decreaseDecimal")
    Call stopVisualMode

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

    Call repeatRegister("insertCellsUp")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + Semicoron_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("insertCellsUp")
    End If
End Function

Function insertCellsDown()
    On Error GoTo Catch

    Call repeatRegister("insertCellsDown")
    Call stopVisualMode

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
    If Err.Number <> 0 Then
        Call errorHandler("insertCellsDown")
    End If
End Function

Function insertCellsLeft()
    On Error GoTo Catch

    Call repeatRegister("insertCellsLeft")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + Semicoron_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("insertCellsLeft")
    End If
End Function

Function insertCellsRight()
    On Error GoTo Catch

    Call repeatRegister("insertCellsRight")
    Call stopVisualMode

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
    If Err.Number <> 0 Then
        Call errorHandler("insertCellsRight")
    End If
End Function

Function deleteValue()
    Call repeatRegister("deleteValue")
    Call stopVisualMode
    Call keystroke(True, Delete_)
End Function

Function deleteToUp()
    On Error GoTo Catch

    Call repeatRegister("deleteToUp")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Minus_, U_, Enter_)

Catch:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("deleteToUp")
    End If
End Function

Function deleteToLeft()
    On Error GoTo Catch

    Call repeatRegister("deleteToLeft")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Ctrl_ + Minus_, L_, Enter_)

Catch:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("deleteToLeft")
    End If
End Function

Function toggleWrapText()
    Call stopVisualMode
    Call keystroke(True, Alt_ + H_, W_)
End Function

Function toggleMergeCells()
    Call repeatRegister("toggleMergeCells")
    Call stopVisualMode

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

Function changeInteriorColor(Optional ByVal resultColor As cls_FontColor)
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.showColorPicker()
    End If

    If Not resultColor Is Nothing Then
        With Selection.Interior
            If resultColor.IsNull Then
                .ColorIndex = xlNone
            ElseIf resultColor.IsThemeColor Then
                .ThemeColor = resultColor.ThemeColor
                .TintAndShade = resultColor.TintAndShade
            Else
                .Color = resultColor.Color
            End If
        End With

        Call repeatRegister("changeInteriorColor", resultColor)
        Call stopVisualMode
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("changeInteriorColor")
    End If
End Function

Function unionSelectCells()
    On Error GoTo Catch

    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call stopVisualMode

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
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("unionSelectCells")
    End If
End Function

Function exceptSelectCells()
    On Error GoTo Catch

    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call stopVisualMode

    If Not gExtendRange Is Nothing Then
        Set actCell = ActiveCell
        Set gExtendRange = Except2(gExtendRange, Selection)

        If Not gExtendRange Is Nothing Then
            gExtendRange.Select
        Else
            Call setStatusBarTemporarily("保存されている拡張選択範囲をクリアしました。", 2)
        End If
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("exceptSelectCells")
    End If
End Function

Function followHyperlinkOfActiveCell()
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If ActiveCell.Hyperlinks.Count > 0 Then
        ActiveCell.Hyperlinks(1).Follow
    ElseIf InStr(UCase(ActiveCell.Formula), "=HYPERLINK(") > 0 Then
        ActiveWorkbook.followHyperlink Split(ActiveCell.Formula, """")(1)
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("followHyperlinkOfActiveCell")
    End If
End Function

Function changeSelectedCells(ByVal Value As String)
    On Error GoTo Catch

    Call stopVisualMode

    If TypeName(Selection) = "Range" Then
        Selection.Value = Value
    ElseIf Not ActiveCell Is Nothing Then
        ActiveCell.Value = Value
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("changeSelectedCells")
    End If
End Function

Function applyFlashFill()
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call repeatRegister("applyFlashFill")

    Selection.FlashFill

    Call stopVisualMode

    Exit Function
Catch:
    If Err.Number = 1004 Then
        Call applyAutoFill(fallback:=True)
    Else
        Call errorHandler("applyFlashFill")
    End If
End Function

Function applyAutoFill(Optional fallback As Boolean = False)
    On Error GoTo Catch

    Dim baseRange As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    ElseIf Selection.Count = 1 Then
        Exit Function
    End If

    If Not fallback Then
        Call repeatRegister("applyAutoFill")
    End If

    Set baseRange = determineBaseRange()
    If baseRange Is Nothing Then
        Exit Function
    End If

    If baseRange.Count = 1 And IsNumeric(baseRange.Formula) Then
        baseRange.AutoFill Selection, xlFillSeries
    Else
        baseRange.AutoFill Selection
    End If

    Call stopVisualMode
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("applyAutoFill")
    End If
End Function

Private Function determineBaseRange() As Range
    On Error GoTo Catch

    Dim avgTop As Double
    Dim avgLeft As Double
    Dim avgBottom As Double
    Dim avgRight As Double
    Dim avgMax As Double

    'n x n cells
    If Selection.Columns.Count > 1 And Selection.Rows.Count > 1 Then
        With Selection
            avgTop = WorksheetFunction.CountA(Range(.Item(1), Cells(.Item(1).Row, .Item(.Count).Column))) / .Columns.Count
            avgLeft = WorksheetFunction.CountA(Range(.Item(1), Cells(.Item(.Count).Row, .Item(1).Column))) / .Rows.Count
            avgBottom = WorksheetFunction.CountA(Range(Cells(.Item(.Count).Row, .Item(1).Column), .Item(.Count))) / .Columns.Count
            avgRight = WorksheetFunction.CountA(Range(Cells(.Item(1).Row, .Item(.Count).Column), .Item(.Count))) / .Rows.Count

            'x - -
            '- - -
            '- - -
            If .Item(1).Formula = "" Then
                avgTop = 0
                avgLeft = 0
            End If

            '- - -
            '- - -
            'x - -
            If Cells(.Item(.Count).Row, .Item(1).Column).Formula = "" Then
                avgLeft = 0
                avgBottom = 0
            End If

            '- - x
            '- - -
            '- - -
            If Cells(.Item(1).Row, .Item(.Count).Column).Formula = "" Then
                avgTop = 0
                avgRight = 0
            End If

            '- - -
            '- - -
            '- - x
            If .Item(.Count).Formula = "" Then
                avgBottom = 0
                avgRight = 0
            End If

            avgMax = WorksheetFunction.Max(avgTop, avgLeft, avgBottom, avgRight)

            Select Case avgMax
                Case 0
                    Call setStatusBarTemporarily("元となるデータを特定できません。", 3)
                    Exit Function
                Case avgTop
                    Set determineBaseRange = Range(.Item(1), Cells(.Item(1).Row, .Item(.Count).Column))
                    Set determineBaseRange = Range(determineBaseRange, innerDataSearch(determineBaseRange, TopToBottom, .Rows.Count - 1))
                Case avgLeft
                    Set determineBaseRange = Range(.Item(1), Cells(.Item(.Count).Row, .Item(1).Column))
                    Set determineBaseRange = Range(determineBaseRange, innerDataSearch(determineBaseRange, LeftToRight, .Columns.Count - 1))
                Case avgBottom
                    Set determineBaseRange = Range(Cells(.Item(.Count).Row, .Item(1).Column), .Item(.Count))
                    Set determineBaseRange = Range(determineBaseRange, innerDataSearch(determineBaseRange, BottomToTop, .Rows.Count - 1))
                Case avgRight
                    Set determineBaseRange = Range(Cells(.Item(1).Row, .Item(.Count).Column), .Item(.Count))
                    Set determineBaseRange = Range(determineBaseRange, innerDataSearch(determineBaseRange, RightToLeft, .Columns.Count - 1))
                Case Else
                    Call debugPrint("Unexpected values: " & avgMax & ", " & avgTop & ", " & avgLeft & ", " & avgBottom & ", " & avgRight, "determineBaseRange")
                    Exit Function
            End Select
        End With

    '1 x n or n x 1 cells
    Else
        If Selection.Item(1).Formula <> "" Then
            If Selection.Item(2).Formula <> "" Then
                If Selection.Columns.Count > 1 Then
                    Set determineBaseRange = Range(Selection.Item(1), Selection.Item(1).End(xlToRight))
                Else
                    Set determineBaseRange = Range(Selection.Item(1), Selection.Item(1).End(xlDown))
                End If
            Else
                Set determineBaseRange = Selection.Item(1)
            End If
        ElseIf Selection.Item(Selection.Count).Formula <> "" Then
            If Selection.Item(Selection.Count - 1).Formula <> "" Then
                If Selection.Columns.Count > 1 Then
                    Set determineBaseRange = Range(Selection.Item(Selection.Count).End(xlToLeft), Selection.Item(Selection.Count))
                Else
                    Set determineBaseRange = Range(Selection.Item(Selection.Count).End(xlUp), Selection.Item(Selection.Count))
                End If
            Else
                Set determineBaseRange = Selection.Item(Selection.Count)
            End If
        Else
            'there is no data at first and last
            Call setStatusBarTemporarily("選択セルの先頭、又は末尾にデータがありません。", 3)
            Exit Function
        End If
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("determineBaseRange")
    End If
End Function

Private Function innerDataSearch(ByVal targetRange As Range, _
                                 ByVal searchMode As searchMode, _
                                 ByVal searchLimit As Long, _
                                 Optional ByVal searchCount As Long = 0, _
                                 Optional ByVal expectCells As Long = 0) As Range
    On Error GoTo Catch

    Dim rowOff As Integer
    Dim columnOff As Integer
    Dim nonBlankCells As Long

    If searchCount > searchLimit Then
        Set innerDataSearch = targetRange
        Exit Function
    End If

    Select Case searchMode
        Case TopToBottom
            rowOff = 1
        Case LeftToRight
            columnOff = 1
        Case BottomToTop
            rowOff = -1
        Case RightToLeft
            columnOff = -1
    End Select

    nonBlankCells = WorksheetFunction.CountA(targetRange)

    If searchCount = 0 Or expectCells = nonBlankCells Then
        Set innerDataSearch = innerDataSearch(targetRange.Offset(rowOff, columnOff), searchMode, searchLimit, searchCount + 1, nonBlankCells)

        If innerDataSearch Is Nothing Then
            Set innerDataSearch = targetRange
        End If
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("innerDataSearch")
    End If
End Function

Function toggleVisualMode()
    On Error GoTo Catch

    If X.VisualMode Then
        Call stopVisualMode
    Else
        Call visualMap("toggleVisualMode")
        Call X.StartVisualMode
        Call setStatusBar(STATUS_PREFIX & "-- VISUAL (ESC to exit) --")
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("toggleVisualMode")
    End If
End Function

Function toggleVisualLine()
    On Error GoTo Catch

    If X.VisualLine Then
        Call stopVisualMode
    Else
        Call visualMap("toggleVisualLine")
        Call X.StartVisualLine
        Call setStatusBar(STATUS_PREFIX & "-- VISUAL LINE (ESC to exit) --")
    End If
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("toggleVisualLine")
    End If
End Function

Private Function visualMap(ByVal funcName As String)
    'Register toggleVisualMode to escape key
    Call temporaryMap("{ESC}", funcName)
    Call temporaryMap("^{[}", funcName)
    Call temporaryMap("o", "swapVisualBase")
End Function

Function swapVisualBase()
    If X.VisualMode Or X.VisualLine Then
        Call X.SwapBase
    End If
End Function

Function stopVisualMode()
    If X.VisualMode Or X.VisualLine Then
        'Unregister escape key
        Call restoreAppliedMap("{ESC}")
        Call restoreAppliedMap("^{[}")
        Call restoreAppliedMap("o")

        Call X.stopVisualMode

        Call setStatusBar("")
    End If
End Function
