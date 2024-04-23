Attribute VB_Name = "F_Find_Replace"
Option Explicit
Option Private Module

Function ShowFindFollowLang(Optional ByVal g As String) As Boolean
    Dim searchStr As String
    searchStr = UF_CmdLine.Launch("/", "Find", gVim.IsJapanese)

    If searchStr <> CMDLINE_CANCELED Then
        Call FindInner(searchStr)
    End If
End Function

Function ShowFindNotFollowLang(Optional ByVal g As String) As Boolean
    Dim searchStr As String
    searchStr = UF_CmdLine.Launch("/", "Find", Not gVim.IsJapanese)

    If searchStr <> CMDLINE_CANCELED Then
        Call FindInner(searchStr)
    End If
End Function

Private Sub FindInner(ByVal findString As String)
    Dim t As Range

    If findString = "" Then
        Call NextFoundCell
        Exit Sub
    End If

    Set t = ActiveSheet.Cells.Find(What:=findString, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   MatchByte:=False)
    If Not t Is Nothing Then
        Call RecordToJumpList

        ActiveWorkbook.ActiveSheet.Activate
        t.Activate
    End If
End Sub

Function NextFoundCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim i As Integer

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
    End If

    Call RecordToJumpList

    For i = gVim.Count1 To 1 Step -1
        If gVim.Count1 = 1 Then
            Application.ScreenUpdating = True
        End If

        Set t = Cells.FindNext(After:=ActiveCell)
        If Not t Is Nothing Then
            t.Activate
        Else
            Application.ScreenUpdating = True
            Exit Function
        End If

    Next i
    Exit Function

Catch:
    Call ErrorHandler("NextFoundCell")
End Function

Function PreviousFoundCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim i As Integer

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
    End If

    Call RecordToJumpList

    For i = gVim.Count1 To 1 Step -1
        If i = 1 Then
            Application.ScreenUpdating = True
        End If

        Set t = Cells.FindPrevious(After:=ActiveCell)
        If Not t Is Nothing Then
            t.Activate
        Else
            Application.ScreenUpdating = True
            Exit Function
        End If

    Next i
    Exit Function

Catch:
    Call ErrorHandler("PreviousFoundCell")
End Function

Function ShowReplaceWindow(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_ + E_, E_)
End Function

Function FindActiveValueNext(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim findText As String

    If ActiveCell Is Nothing Then
        Exit Function
    End If

    findText = ActiveCell.Value

    If findText = "" Then
        Exit Function
    End If

    Set t = ActiveSheet.Cells.Find(What:=findText, _
                                   After:=ActiveCell, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   MatchByte:=False)

    If Not t Is Nothing Then
        Call RecordToJumpList
        Call NextFoundCell
    End If

    Call SetStatusBarTemporarily("/" & findText, 2000, disablePrefix:=True)
    Exit Function

Catch:
    Call ErrorHandler("FindActiveValueNext")
End Function

Function FindActiveValuePrev(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim findText As String

    If ActiveCell Is Nothing Then
        Exit Function
    End If

    findText = ActiveCell.Value

    If findText = "" Then
        Exit Function
    End If

    Set t = ActiveSheet.Cells.Find(What:=findText, _
                                   After:=ActiveCell, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   MatchByte:=False)

    If Not t Is Nothing Then
        Call RecordToJumpList
        Call PreviousFoundCell
    End If

    Call SetStatusBarTemporarily("?" & findText, 2000, disablePrefix:=True)
    Exit Function

Catch:
    Call ErrorHandler("FindActiveValuePrev")
End Function

Function NextSpecialCells(ByVal TypeValue As XlCellType, Optional SearchOrder As XlSearchOrder = xlByColumns) As Boolean
    On Error GoTo Catch

    Dim rngSpecialCells As Range
    Dim rngResultCell As Range
    Dim i As Long

    Call RecordToJumpList

    'Raise error if target cell does not exists
    Set rngSpecialCells = ActiveSheet.UsedRange.SpecialCells(TypeValue)

    'Calculate next cell
    Set rngResultCell = ActiveCell
    For i = 1 To (gVim.Count1 - 1) Mod rngSpecialCells.Count + 1
        Set rngResultCell = DetermineCell(rngResultCell, rngSpecialCells, TypeValue, SearchOrder, xlNext)
    Next i

    If Not rngResultCell Is Nothing Then
        rngResultCell.Activate
    End If

    Exit Function

Catch:
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoMatchingCell, 2000)
    Else
        Call ErrorHandler("NextSpecialCells")
    End If
End Function

Function PrevSpecialCells(ByVal TypeValue As XlCellType, Optional SearchOrder As XlSearchOrder = xlByColumns) As Boolean
    On Error GoTo Catch

    Dim rngSpecialCells As Range
    Dim rngResultCell As Range
    Dim i As Long

    Call RecordToJumpList

    'Raise error if target cell does not exists
    Set rngSpecialCells = ActiveSheet.UsedRange.SpecialCells(TypeValue)

    'Calculate next cell
    Set rngResultCell = ActiveCell
    For i = 1 To (gVim.Count1 - 1) Mod rngSpecialCells.Count + 1
        Set rngResultCell = DetermineCell(rngResultCell, rngSpecialCells, TypeValue, SearchOrder, xlPrevious)
    Next i

    If Not rngResultCell Is Nothing Then
        rngResultCell.Activate
    End If
    Exit Function

Catch:
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoMatchingCell, 2000)
    Else
        Call ErrorHandler("PrevSpecialCells")
    End If
End Function

Private Function DetermineCell(ByRef BaseCell As Range, _
                               ByRef FoundCells As Range, _
                               ByVal TypeValue As XlCellType, _
                               ByVal SearchOrder As XlSearchOrder, _
                               ByVal SearchDirection As XlSearchDirection) As Range
    On Error GoTo Catch

    Dim rngCheckCells As Range
    Dim rngResultCells As Range
    Dim rngFoundCells As Range
    Dim lngRow As Long
    Dim lngCol As Long
    Dim minRow As Long
    Dim minCol As Long
    Dim maxRow As Long
    Dim maxCol As Long

    If BaseCell Is Nothing Or FoundCells Is Nothing Then
        Exit Function
    End If

    lngRow = BaseCell.Row
    lngCol = BaseCell.Column

    With ActiveSheet.UsedRange
        minRow = .Item(1).Row
        minCol = .Item(1).Column
        maxRow = .Item(.Count).Row
        maxCol = .Item(.Count).Column
    End With

    Set rngCheckCells = Nothing
    Set rngResultCells = Nothing

    'Step 1
    If SearchOrder = xlByColumns Then
        If SearchDirection = xlNext Then
            If lngRow < maxRow Then
                lngRow = lngRow + 1
                Set rngCheckCells = Range(Cells(lngRow, lngCol), Cells(maxRow, lngCol))
            End If

        ElseIf SearchDirection = xlPrevious Then
            If lngRow > minRow Then
                lngRow = lngRow - 1
                Set rngCheckCells = Range(Cells(minRow, lngCol), Cells(lngRow, lngCol))
            End If

        End If

    ElseIf SearchOrder = xlByRows Then
        If SearchDirection = xlNext Then
            If lngCol < maxCol Then
                lngCol = lngCol + 1
                Set rngCheckCells = Range(Cells(lngRow, lngCol), Cells(lngRow, maxCol))
            End If

        ElseIf SearchDirection = xlPrevious Then
            If lngCol > minCol Then
                lngCol = lngCol - 1
                Set rngCheckCells = Range(Cells(lngRow, minCol), Cells(lngRow, lngCol))
            End If
        End If
    End If

    If Not rngCheckCells Is Nothing Then
        Set rngResultCells = Nothing

        On Error Resume Next
        Set rngResultCells = Intersect(rngCheckCells.SpecialCells(TypeValue), rngCheckCells)
        On Error GoTo Catch

        If Not rngResultCells Is Nothing Then
            Set DetermineCell = ClosestSearch(rngResultCells, SearchOrder, SearchDirection, TypeValue = xlCellTypeBlanks)
            If Not DetermineCell Is Nothing Then
                Exit Function
            End If
        End If
    End If

    'Step 2
    Set rngCheckCells = Nothing
    Set rngFoundCells = Nothing
    If SearchOrder = xlByColumns Then
        If SearchDirection = xlNext Then
            If lngCol < maxCol Then
                lngCol = lngCol + 1
                Set rngCheckCells = Range(Cells(minRow, lngCol), Cells(maxRow, maxCol))
            End If

        ElseIf SearchDirection = xlPrevious Then
            If lngCol > minCol Then
                lngCol = lngCol - 1
                Set rngCheckCells = Range(Cells(minRow, minCol), Cells(maxRow, lngCol))
            End If

        End If

        If Not rngCheckCells Is Nothing Then
            On Error Resume Next
            Set rngFoundCells = rngCheckCells.SpecialCells(TypeValue)
            On Error GoTo Catch
        End If

    ElseIf SearchOrder = xlByRows Then
        If SearchDirection = xlNext Then
            If lngRow < maxRow Then
                lngRow = lngRow + 1
                Set rngCheckCells = Range(Cells(lngRow, minCol), Cells(maxRow, maxCol))
            End If

        ElseIf SearchDirection = xlPrevious Then
            If lngRow > 1 Then
                lngRow = lngRow - 1
                Set rngCheckCells = Range(Cells(minRow, minCol), Cells(lngRow, maxCol))
            End If
        End If

        If Not rngCheckCells Is Nothing Then
            On Error Resume Next
            Set rngFoundCells = rngCheckCells.SpecialCells(TypeValue)
            On Error GoTo Catch
        End If
    End If

    If Not rngFoundCells Is Nothing Then
        Set rngResultCells = Intersect(rngCheckCells, rngFoundCells)
        If Not rngResultCells Is Nothing Then
            Set DetermineCell = ClosestSearch(rngResultCells, SearchOrder, SearchDirection, TypeValue = xlCellTypeBlanks)
            If Not DetermineCell Is Nothing Then
                Exit Function
            End If
        End If
    End If

    'Step 3
    Set rngCheckCells = Range(Cells(minRow, minCol), Cells(maxRow, maxCol))
    Set rngResultCells = Intersect(rngCheckCells, FoundCells)
    If Not rngResultCells Is Nothing Then
        Set DetermineCell = ClosestSearch(rngResultCells, SearchOrder, SearchDirection, TypeValue = xlCellTypeBlanks)
    End If

    Exit Function

Catch:
    Call ErrorHandler("DetermineCell")
End Function

Private Function ClosestSearch(ByRef rngResultCells As Range, _
                               ByVal SearchOrder As XlSearchOrder, _
                               ByVal SearchDirection As XlSearchDirection, _
                               ByVal checkIsBlankMergedCell As Boolean) As Range
    On Error GoTo Catch

    Dim r As Range
    Dim tmp As Range

    If rngResultCells Is Nothing Then
        Exit Function
    End If

    For Each r In rngResultCells.Areas
        'Search smallest
        If SearchDirection = xlNext Then
            If r.Item(1).MergeCells Then
                'Consider merged cell
                If checkIsBlankMergedCell And r.Item(1).MergeArea.Item(1).Value <> "" Then
                    Set tmp = Nothing
                Else
                    Set tmp = Intersect(r, r.Item(1).MergeArea.Item(1))
                End If

                Set tmp = Union2(tmp, Except2(r, r.Item(1).MergeArea))
                If tmp Is Nothing Then
                    GoTo Continue
                ElseIf tmp.Address <> r.Address And tmp.Count > 1 Then
                    Set tmp = ClosestSearch(tmp, SearchOrder, SearchDirection, checkIsBlankMergedCell)
                Else
                    Set tmp = tmp.Item(1).MergeArea.Item(1)
                End If
            Else
                Set tmp = r.Item(1)
            End If

            If tmp Is Nothing Then
                GoTo Continue
            End If

            If ClosestSearch Is Nothing Then
                Set ClosestSearch = tmp
            ElseIf SearchOrder = xlByColumns Then
                If ClosestSearch.Column > tmp.Column Or (ClosestSearch.Column = tmp.Column And ClosestSearch.Row > tmp.Row) Then
                    Set ClosestSearch = tmp
                End If
            ElseIf SearchOrder = xlByRows Then
                If ClosestSearch.Row > tmp.Row Or (ClosestSearch.Row = tmp.Row And ClosestSearch.Column > tmp.Column) Then
                    Set ClosestSearch = tmp
                End If
            End If

        'Search biggest
        ElseIf SearchDirection = xlPrevious Then
            If r.Item(r.Count).MergeCells Then
                'Consider merged cell
                If checkIsBlankMergedCell And r.Item(r.Count).MergeArea.Item(1).Value <> "" Then
                    Set tmp = Nothing
                Else
                    Set tmp = Intersect(r, r.Item(r.Count).MergeArea.Item(1))
                End If

                Set tmp = Union2(tmp, Except2(r, r.Item(r.Count).MergeArea))
                If tmp Is Nothing Then
                    GoTo Continue
                ElseIf tmp.Address <> r.Address And tmp.Count > 1 Then
                    Set tmp = ClosestSearch(tmp, SearchOrder, SearchDirection, checkIsBlankMergedCell)
                Else
                    Set tmp = tmp.Item(tmp.Count).MergeArea.Item(1)
                End If
            Else
                Set tmp = r.Item(r.Count).MergeArea.Item(1)
            End If

            If tmp Is Nothing Then
                GoTo Continue
            End If

            If ClosestSearch Is Nothing Then
                Set ClosestSearch = tmp
            ElseIf SearchOrder = xlByColumns Then
                If ClosestSearch.Column < tmp.Column Or (ClosestSearch.Column = tmp.Column And ClosestSearch.Row < tmp.Row) Then
                    Set ClosestSearch = tmp
                End If
            ElseIf SearchOrder = xlByRows Then
                If ClosestSearch.Row < tmp.Row Or (ClosestSearch.Row = tmp.Row And ClosestSearch.Column < tmp.Column) Then
                    Set ClosestSearch = tmp
                End If
            End If

        End If
Continue:
    Next r
    Exit Function

Catch:
    Call ErrorHandler("ClosestSearch")
End Function
