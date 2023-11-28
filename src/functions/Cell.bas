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
        Set gVim.Vars.LastYanked = Selection
    End If
End Function

Function yankCell()
    Call stopVisualMode
    Call keystroke(True, Ctrl_ + C_)

    If TypeName(Selection) = "Range" Then
        Set gVim.Vars.LastYanked = Selection
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

Function yankAsPlaintext(Optional ByVal ColumnSpliter As String = vbTab)
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    'Error if too many cells selected
    If Selection.Count > 1048576 * 8 Then
        Err.Raise 6
    End If

    Call stopVisualMode

    Dim resultText As String
    Dim aryTarget As Variant
    Dim aryX() As String
    Dim aryY() As String
    Dim i As Long
    Dim j As Long
    Dim startTime As Double
    Dim currentTime As Double

    'Exit if all selected cells are blank
    If WorksheetFunction.CountBlank(Selection) = Selection.Count Then
        Exit Function
    End If

    If Selection.Count = 1 Then
        resultText = Selection.Value

    ElseIf Selection.Columns.Count = 1 Then
        aryTarget = Selection
        aryTarget = WorksheetFunction.Transpose(aryTarget)
        resultText = Join(aryTarget, vbCrLf)

    ElseIf Selection.Rows.Count = 1 Then
        aryTarget = Selection

        'Array dimensionality reduction
        aryTarget = WorksheetFunction.Transpose(aryTarget)
        aryTarget = WorksheetFunction.Transpose(aryTarget)

        resultText = Join(aryTarget, ColumnSpliter)

    Else
fallback:
        startTime = Timer
        aryTarget = Selection
        ReDim aryX(LBound(aryTarget, 1) To UBound(aryTarget, 1))
        ReDim aryY(LBound(aryTarget, 2) To UBound(aryTarget, 2))

        For i = LBound(aryX) To UBound(aryX)
            For j = LBound(aryY) To UBound(aryY)
                aryY(j) = aryTarget(i, j)
            Next j
            aryX(i) = Join(aryY, ColumnSpliter)

            'Avoid freeze
            If (i And &HFFF) = 0 Then
                'Show progress bar in status bar
                Call SetStatusBar("テキストをコピーしています...", _
                                 currentCount:=i, maximumCount:=UBound(aryX), progressBar:=True)

                currentTime = Timer
                If currentTime < startTime Or currentTime - startTime > 2 Then
                    DoEvents
                    startTime = currentTime
                End If
            End If
        Next i
        resultText = Join(aryX, vbCrLf)
        Call SetStatusBar
    End If

    'Set to clipboard
    With New DataObject
        .SetText resultText
        .PutInClipboard
    End With

    Call SetStatusBarTemporarily("クリップボードにコピーしました。(" & _
                                 LenB(StrConv(resultText, vbFromUnicode)) & " Bytes)", 3000)
    Exit Function

Catch:
    If Err.Number = 6 Then
        Call SetStatusBarTemporarily("選択セル数が多すぎます", 3000)
    ElseIf Err.Number = 13 Then
        'Error from WorksheetFunction.Transpose
        Resume fallback
    Else
        Call errorHandler("yankAsPlaintext")
    End If
End Function

Function incrementText()
    Call repeatRegister("incrementText")
    Call stopVisualMode

    Dim i As Integer

    Call keyupControlKeys
    Call releaseShiftKeys

    For i = 1 To gVim.Count1
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

    For i = 1 To gVim.Count1
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

    For i = 1 To gVim.Count1
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

    For i = 1 To gVim.Count1
        Call keystrokeWithoutKeyup(Alt_ + H_, k9_)
    Next i

    Call unkeyupControlKeys
End Function

Function insertCellsUp()
    On Error GoTo Catch

    Call repeatRegister("insertCellsUp")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(True, Ctrl_ + Shift_ + Semicoron_JIS_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("insertCellsUp")
End Function

Function insertCellsDown()
    On Error GoTo Catch

    Call repeatRegister("insertCellsDown")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If Selection.Row < ActiveSheet.Rows.Count Then
        Selection.Offset(1, 0).Select
    End If

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(True, Ctrl_ + Shift_ + Semicoron_JIS_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("insertCellsDown")
End Function

Function insertCellsLeft()
    On Error GoTo Catch

    Call repeatRegister("insertCellsLeft")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(True, Ctrl_ + Shift_ + Semicoron_JIS_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("insertCellsLeft")
End Function

Function insertCellsRight()
    On Error GoTo Catch

    Call repeatRegister("insertCellsRight")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If Selection.Column < ActiveSheet.Columns.Count Then
        Selection.Offset(0, 1).Select
    End If

    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(True, Ctrl_ + Shift_ + Semicoron_JIS_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("insertCellsRight")
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
    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Minus_, U_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("deleteToUp")
End Function

Function deleteToLeft()
    On Error GoTo Catch

    Call repeatRegister("deleteToLeft")
    Call stopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call keystroke(True, Ctrl_ + Minus_, L_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("deleteToLeft")
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

Function applyCommaStyle()
    Call repeatRegister("applyCommaStyle")
    Call stopVisualMode

    Call keystroke(True, Alt_ + H_, K_)
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
    Call errorHandler("changeInteriorColor")
End Function

Function unionSelectCells()
    On Error GoTo Catch

    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call stopVisualMode

    If gVim.Vars.ExtendRange Is Nothing Then
        Set gVim.Vars.ExtendRange = Selection

    ElseIf Not gVim.Vars.ExtendRange.Parent Is ActiveSheet Then
        Call SetStatusBarTemporarily("異なるシートで拡張選択はできないため、選択範囲は初期化されました。", 2000)
        Set gVim.Vars.ExtendRange = Selection

    Else
        Set actCell = ActiveCell
        Set gVim.Vars.ExtendRange = Union2(gVim.Vars.ExtendRange, Selection)
        gVim.Vars.ExtendRange.Select
        actCell.Activate

    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Selection
    Else
        Call errorHandler("unionSelectCells")
    End If
End Function

Function exceptSelectCells()
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call stopVisualMode

    If Not gVim.Vars.ExtendRange Is Nothing Then
        If Selection.Address = gVim.Vars.ExtendRange.Address Then
            Set gVim.Vars.ExtendRange = Except2(gVim.Vars.ExtendRange, ActiveCell)
        Else
            Set gVim.Vars.ExtendRange = Except2(gVim.Vars.ExtendRange, Selection)
        End If

        If Not gVim.Vars.ExtendRange Is Nothing Then
            gVim.Vars.ExtendRange.Select
        Else
            Call SetStatusBarTemporarily("保存されている拡張選択範囲をクリアしました。", 2000)
        End If
    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Nothing
    Else
        Call errorHandler("exceptSelectCells")
    End If
End Function

Function clearSelectCells()
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call stopVisualMode

    If Not gVim.Vars.ExtendRange Is Nothing Then
        If Selection.Address = gVim.Vars.ExtendRange.Address Then
            Set gVim.Vars.ExtendRange = Nothing
            Call SetStatusBarTemporarily("保存されている拡張選択範囲をクリアしました。", 2000)
            Exit Function
        End If
    End If

    If Selection.Columns.Count > 1 Or Selection.Rows.Count > 1 Or Selection.Areas.Count > 1 Then
        ActiveCell.Select
    ElseIf Not gVim.Vars.ExtendRange Is Nothing Then
        gVim.Vars.ExtendRange.Select
    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Nothing
    Else
        Call errorHandler("clearSelectCells")
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
    Call errorHandler("followHyperlinkOfActiveCell")
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
    Call errorHandler("changeSelectedCells")
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
        Call ApplyAutoFillInner(fallback:=True)
    Else
        Call errorHandler("applyFlashFill")
    End If
End Function

Function ApplyAutoFill(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    ElseIf Selection.Count = 1 Then
        Exit Function
    End If

    Call RepeatRegister("ApplyAutoFill")

    Call ApplyAutoFillInner

    Exit Function

Catch:
    Call ErrorHandler("ApplyAutoFill")
End Function

Function ApplyAutoFillInner(Optional fallback As Boolean = False)
    On Error GoTo Catch

    Dim baseRange As Range

    Set baseRange = DetermineBaseRange()
    If baseRange Is Nothing Then
        Exit Function
    End If

    If baseRange.Count = 1 And IsNumeric(baseRange.Formula) Then
        baseRange.AutoFill Selection, xlFillSeries
    Else
        baseRange.AutoFill Selection
    End If

    Call StopVisualMode
    Exit Function

Catch:
    Call ErrorHandler("ApplyAutoFillInner")
End Function

Private Function DetermineBaseRange() As Range
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
                    Call SetStatusBarTemporarily("元となるデータを特定できません。", 3000)
                    Exit Function
                Case avgTop
                    Set DetermineBaseRange = Range(.Item(1), Cells(.Item(1).Row, .Item(.Count).Column))
                    Set DetermineBaseRange = Range(DetermineBaseRange, InnerDataSearch(DetermineBaseRange, TopToBottom, .Rows.Count - 1))
                Case avgLeft
                    Set DetermineBaseRange = Range(.Item(1), Cells(.Item(.Count).Row, .Item(1).Column))
                    Set DetermineBaseRange = Range(DetermineBaseRange, InnerDataSearch(DetermineBaseRange, LeftToRight, .Columns.Count - 1))
                Case avgBottom
                    Set DetermineBaseRange = Range(Cells(.Item(.Count).Row, .Item(1).Column), .Item(.Count))
                    Set DetermineBaseRange = Range(DetermineBaseRange, InnerDataSearch(DetermineBaseRange, BottomToTop, .Rows.Count - 1))
                Case avgRight
                    Set DetermineBaseRange = Range(Cells(.Item(1).Row, .Item(.Count).Column), .Item(.Count))
                    Set DetermineBaseRange = Range(DetermineBaseRange, InnerDataSearch(DetermineBaseRange, RightToLeft, .Columns.Count - 1))
                Case Else
                    Call DebugPrint("Unexpected values: " & avgMax & ", " & avgTop & ", " & avgLeft & ", " & avgBottom & ", " & avgRight, "determineBaseRange")
                    Exit Function
            End Select
        End With

    '1 x n or n x 1 cells
    Else
        If Selection.Item(1).Formula <> "" Then
            If Selection.Item(2).Formula <> "" Then
                If Selection.Columns.Count > 1 Then
                    Set DetermineBaseRange = Range(Selection.Item(1), Selection.Item(1).End(xlToRight))
                Else
                    Set DetermineBaseRange = Range(Selection.Item(1), Selection.Item(1).End(xlDown))
                End If
            Else
                Set DetermineBaseRange = Selection.Item(1)
            End If
        ElseIf Selection.Item(Selection.Count).Formula <> "" Then
            If Selection.Item(Selection.Count - 1).Formula <> "" Then
                If Selection.Columns.Count > 1 Then
                    Set DetermineBaseRange = Range(Selection.Item(Selection.Count).End(xlToLeft), Selection.Item(Selection.Count))
                Else
                    Set DetermineBaseRange = Range(Selection.Item(Selection.Count).End(xlUp), Selection.Item(Selection.Count))
                End If
            Else
                Set DetermineBaseRange = Selection.Item(Selection.Count)
            End If
        Else
            'there is no data at first and last
            Call SetStatusBarTemporarily("選択セルの先頭、又は末尾にデータがありません。", 3000)
            Exit Function
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("DetermineBaseRange")
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
    Call errorHandler("innerDataSearch")
End Function

Private Function autoSumInner(ByVal lastKey As Long)
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call keystroke(True, Alt_ + M_, U_, lastKey)

    Exit Function
Catch:
    Call errorHandler("autoSumInner")
End Function

Function autoSum()
    Call autoSumInner(S_)
End Function

Function autoAverage()
    Call autoSumInner(A_)
End Function

Function autoCount()
    Call autoSumInner(C_)
End Function

Function autoMax()
    Call autoSumInner(M_)
End Function

Function autoMin()
    Call autoSumInner(I_)
End Function

Function insertFunction()
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "AutoSumMoreFunctions"
    Exit Function
Catch:
    Call errorHandler("insertFunction")
End Function
