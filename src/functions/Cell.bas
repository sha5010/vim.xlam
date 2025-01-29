Attribute VB_Name = "F_Cell"
Option Explicit
Option Private Module

Private Enum eSearchMode
    TopToBottom = 1
    LeftToRight
    BottomToTop
    RightToLeft
End Enum

Function CutCell(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + X_)

    If TypeName(Selection) = "Range" Then
        Set gVim.Vars.LastYanked = Selection
    End If
End Function

Function YankCell(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + C_)

    If TypeName(Selection) = "Range" Then
        Set gVim.Vars.LastYanked = Selection
    End If
End Function

Function YankFromUpCell(Optional ByVal g As String) As Boolean
    Call RepeatRegister("YankFromUpCell")
    Call KeyStroke(Alt_ + H_, F_, I_, D_)
End Function

Function YankFromDownCell(Optional ByVal g As String) As Boolean
    Call RepeatRegister("YankFromDownCell")
    Call KeyStroke(Alt_ + H_, F_, I_, U_)
End Function

Function YankFromLeftCell(Optional ByVal g As String) As Boolean
    Call RepeatRegister("YankFromLeftCell")
    Call KeyStroke(Alt_ + H_, F_, I_, R_)
End Function

Function YankFromRightCell(Optional ByVal g As String) As Boolean
    Call RepeatRegister("YankFromRightCell")
    Call KeyStroke(Alt_ + H_, F_, I_, L_)
End Function

Function YankAsPlaintext(Optional ByVal ColumnSpliter As String = vbTab) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    'Error if too many cells selected
    If Selection.Count > 1048576 * 8 Then
        Err.Raise 6
    End If

    Call StopVisualMode

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
                Call SetStatusBar(gVim.Msg.YankInProgress, _
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

    Call SetStatusBarTemporarily(gVim.Msg.YankDone & "(" & _
                                 LenB(StrConv(resultText, vbFromUnicode)) & " Bytes)", 3000)
    Exit Function

Catch:
    If Err.Number = 6 Then
        Call SetStatusBarTemporarily(gVim.Msg.TooManyCells, 3000)
    ElseIf Err.Number = 13 Then
        'Error from WorksheetFunction.Transpose
        Resume fallback
    Else
        Call ErrorHandler("YankAsPlaintext")
    End If
End Function

Function IncrementText(Optional ByVal g As String) As Boolean
    Call RepeatRegister("IncrementText")
    Call StopVisualMode

    Dim i As Integer

    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_ + H_, k6_)
    Next i
End Function

Function DecrementText(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DecrementText")
    Call StopVisualMode

    Dim i As Integer

    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_ + H_, k5_)
    Next i
End Function

Function IncreaseDecimal(Optional ByVal g As String) As Boolean
    Call RepeatRegister("IncreaseDecimal")
    Call StopVisualMode

    Dim i As Integer

    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_ + H_, k0_)
    Next i
End Function

Function DecreaseDecimal(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DecreaseDecimal")
    Call StopVisualMode

    Dim i As Integer

    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_ + H_, k9_)
    Next i
End Function

Function InsertCellsUp(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsUp")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsUp")
End Function

Function InsertCellsDown(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsDown")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If Selection.Row < ActiveSheet.Rows.Count Then
        Selection.Offset(1, 0).Select
    End If

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsDown")
End Function

Function InsertCellsLeft(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsLeft")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsLeft")
End Function

Function InsertCellsRight(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsRight")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If Selection.Column < ActiveSheet.Columns.Count Then
        Selection.Offset(0, 1).Select
    End If

    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsRight")
End Function

Function DeleteValue(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteValue")
    Call StopVisualMode
    Call KeyStroke(Delete_)
End Function

Function DeleteToUp(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("DeleteToUp")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Ctrl_ + Minus_, U_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("DeleteToUp")
End Function

Function DeleteToLeft(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("DeleteToLeft")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Ctrl_ + Minus_, L_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("DeleteToLeft")
End Function

Function ToggleWrapText(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Alt_ + H_, W_)
End Function

Function ToggleMergeCells(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleMergeCells")
    Call StopVisualMode

    If TypeName(Selection) = "Range" Then
        If Not ActiveCell.MergeCells And Selection.Count = 1 Then
            Exit Function
        End If

        If ActiveCell.MergeCells Then
            Call KeyStroke(Alt_ + H_, M_, U_)
        Else
            Call KeyStroke(Alt_ + H_, M_, M_)
        End If
    End If
End Function

Function ApplyCommaStyle(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ApplyCommaStyle")
    Call StopVisualMode

    Call KeyStroke(Alt_ + H_, K_)
End Function

Function ChangeInteriorColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
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

        Call RepeatRegister("ChangeInteriorColor", resultColor)
        Call StopVisualMode
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeInteriorColor")
End Function

Function UnionSelectCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call StopVisualMode

    If gVim.Vars.ExtendRange Is Nothing Then
        Set gVim.Vars.ExtendRange = Selection

    ElseIf Not gVim.Vars.ExtendRange.Parent Is ActiveSheet Then
        Call SetStatusBarTemporarily(gVim.Msg.InitializedExtendedSelection, 2000)
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
        Call ErrorHandler("UnionSelectCells")
    End If
End Function

Function ExceptSelectCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call StopVisualMode

    If Not gVim.Vars.ExtendRange Is Nothing Then
        If Selection.Address = gVim.Vars.ExtendRange.Address Then
            Set gVim.Vars.ExtendRange = Except2(gVim.Vars.ExtendRange, ActiveCell)
        Else
            Set gVim.Vars.ExtendRange = Except2(gVim.Vars.ExtendRange, Selection)
        End If

        If Not gVim.Vars.ExtendRange Is Nothing Then
            gVim.Vars.ExtendRange.Select
        Else
            Call SetStatusBarTemporarily(gVim.Msg.ClearedExtendedSelection, 2000)
        End If
    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Nothing
    Else
        Call ErrorHandler("ExceptSelectCells")
    End If
End Function

Function ClearSelectCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call StopVisualMode

    If Not gVim.Vars.ExtendRange Is Nothing Then
        If Selection.Address = gVim.Vars.ExtendRange.Address Then
            Set gVim.Vars.ExtendRange = Nothing
            Call SetStatusBarTemporarily(gVim.Msg.ClearedExtendedSelection, 2000)
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
        Call ErrorHandler("ClearSelectCells")
    End If
End Function

Function FollowHyperlinkOfActiveCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    With ActiveCell
        If .Hyperlinks.Count > 0 Then
            .Hyperlinks(1).Follow
        ElseIf .Formula <> .Value And InStr(.Formula, "HYPERLINK") > 0 Then
            Dim linkAddr As String
            linkAddr = Application.Evaluate(Replace(.Formula, "HYPERLINK", "IFERROR"))
            ActiveWorkbook.FollowHyperlink linkAddr
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("FollowHyperlinkOfActiveCell")
End Function

Function ChangeSelectedCells(ByVal Value As String) As Boolean
    On Error GoTo Catch

    Call StopVisualMode

    If TypeName(Selection) = "Range" Then
        Selection.Value = Value
    ElseIf Not ActiveCell Is Nothing Then
        ActiveCell.Value = Value
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeSelectedCells")
End Function

Function ApplyFlashFill(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call RepeatRegister("ApplyFlashFill")

    Selection.FlashFill

    Call StopVisualMode

    Exit Function
Catch:
    If Err.Number = 1004 Then
        Call ApplyAutoFillInner(fallback:=True)
    Else
        Call ErrorHandler("ApplyFlashFill")
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
                    Call SetStatusBarTemporarily(gVim.Msg.UnableIdentifySource, 3000)
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
            Call SetStatusBarTemporarily(gVim.Msg.NoDataInSelectedCells, 3000)
            Exit Function
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("DetermineBaseRange")
End Function

Private Function InnerDataSearch(ByVal targetRange As Range, _
                                 ByVal searchMode As eSearchMode, _
                                 ByVal searchLimit As Long, _
                                 Optional ByVal searchCount As Long = 0, _
                                 Optional ByVal expectCells As Long = 0) As Range
    On Error GoTo Catch

    Dim rowOff As Integer
    Dim columnOff As Integer
    Dim nonBlankCells As Long

    If searchCount > searchLimit Then
        Set InnerDataSearch = targetRange
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
        Set InnerDataSearch = InnerDataSearch(targetRange.Offset(rowOff, columnOff), searchMode, searchLimit, searchCount + 1, nonBlankCells)

        If InnerDataSearch Is Nothing Then
            Set InnerDataSearch = targetRange
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("InnerDataSearch")
End Function

Private Function AutoSumInner(ByVal lastKey As Long)
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call KeyStroke(Alt_ + M_, U_, lastKey)

    Exit Function
Catch:
    Call ErrorHandler("AutoSumInner")
End Function

Function AutoSum(Optional ByVal g As String) As Boolean
    Call AutoSumInner(S_)
End Function

Function AutoAverage(Optional ByVal g As String) As Boolean
    Call AutoSumInner(A_)
End Function

Function AutoCount(Optional ByVal g As String) As Boolean
    Call AutoSumInner(C_)
End Function

Function AutoMax(Optional ByVal g As String) As Boolean
    Call AutoSumInner(M_)
End Function

Function AutoMin(Optional ByVal g As String) As Boolean
    Call AutoSumInner(I_)
End Function

Function InsertFunction(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "AutoSumMoreFunctions"
    Exit Function
Catch:
    Call ErrorHandler("InsertFunction")
End Function
