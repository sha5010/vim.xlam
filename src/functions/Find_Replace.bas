Attribute VB_Name = "F_Find_Replace"
Option Explicit
Option Private Module

Function showFindFollowLang()
    UF_FindForm.Show
End Function

Function showFindNotFollowLang()
    gLangJa = Not gLangJa
    UF_FindForm.Show
    gLangJa = Not gLangJa
End Function

Function nextFoundCell()
    On Error GoTo Catch

    Dim t As Range
    Dim i As Integer

    If gCount > 1 Then
        Application.ScreenUpdating = False
    End If

    Call recordToJumpList

    For i = gCount To 1 Step -1
        If gCount = 1 Then
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
    Call errorHandler("nextFoundCell")
End Function

Function previousFoundCell()
    On Error GoTo Catch

    Dim t As Range
    Dim i As Integer

    If gCount > 1 Then
        Application.ScreenUpdating = False
    End If

    Call recordToJumpList

    For i = gCount To 1 Step -1
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
    Call errorHandler("previousFoundCell")
End Function

Function showReplaceWindow()
    Call keystroke(True, Alt_ + E_, E_)
End Function

Function findActiveValueNext()
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
        Call recordToJumpList

        ActiveWorkbook.ActiveSheet.Activate
        t.Activate
    End If

    Call setStatusBarTemporarily("/" & findText, 2, disablePrefix:=True)
    Exit Function

Catch:
    Call errorHandler("findActiveValueNext")
End Function

Function findActiveValuePrev()
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
        Call recordToJumpList

        ActiveWorkbook.ActiveSheet.Activate
        Set t = Cells.FindPrevious(After:=ActiveCell)
        t.Activate
    End If

    Call setStatusBarTemporarily("?" & findText, 2, disablePrefix:=True)
    Exit Function

Catch:
    Call errorHandler("findActiveValuePrev")
End Function

Function nextSpecialCells(ByVal TypeValue As XlCellType, Optional SearchOrder As XlSearchOrder = xlByColumns)
    On Error GoTo Catch

    Dim rngSpecialCells As Range
    Dim rngResultCell As Range
    Dim i As Long

    'Raise error if target cell does not exists
    Set rngSpecialCells = ActiveSheet.UsedRange.SpecialCells(TypeValue)

    'Calculate next cell
    Set rngResultCell = ActiveCell
    For i = 1 To (gCount - 1) Mod rngSpecialCells.Count + 1
        Set rngResultCell = determineCell(rngResultCell, rngSpecialCells, SearchOrder, xlNext)
    Next i

    If Not rngResultCell Is Nothing Then
        rngResultCell.Activate
    End If

    Exit Function

Catch:
    If Err.Number = 1004 Then
        Call setStatusBarTemporarily("該当するセルが見つかりません。", 2)
    Else
        Call errorHandler("nextSpecialCells")
    End If
End Function

Function prevSpecialCells(ByVal TypeValue As XlCellType, Optional SearchOrder As XlSearchOrder = xlByColumns)
    On Error GoTo Catch

    Dim rngSpecialCells As Range
    Dim rngResultCell As Range
    Dim i As Long

    'Raise error if target cell does not exists
    Set rngSpecialCells = ActiveSheet.UsedRange.SpecialCells(TypeValue)

    'Calculate next cell
    Set rngResultCell = ActiveCell
    For i = 1 To (gCount - 1) Mod rngSpecialCells.Count + 1
        Set rngResultCell = determineCell(rngResultCell, rngSpecialCells, SearchOrder, xlPrevious)
    Next i

    If Not rngResultCell Is Nothing Then
        rngResultCell.Activate
    End If
    Exit Function

Catch:
    If Err.Number = 1004 Then
        Call setStatusBarTemporarily("該当するセルが見つかりません。", 2)
    Else
        Call errorHandler("prevSpecialCells")
    End If
End Function

Private Function determineCell(ByRef BaseCell As Range, _
                               ByRef FoundCells As Range, _
                               ByVal SearchOrder As XlSearchOrder, _
                               ByVal SearchDirection As XlSearchDirection) As Range
    On Error GoTo Catch

    Dim rngCheckCells As Range
    Dim rngResultCells As Range
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
        Set rngResultCells = Intersect(rngCheckCells, FoundCells)
        If Not rngResultCells Is Nothing Then
            Set determineCell = closestSearch(rngResultCells, SearchDirection)
            Exit Function
        End If
    End If

    'Step 2
    If SearchOrder = xlByColumns Then
        If SearchDirection = xlNext Then
            If lngCol < maxCol Then
                lngCol = lngCol + 1
                Set rngCheckCells = Intersect(Range(Columns(lngCol), Columns(maxCol)), FoundCells.EntireColumn)
                Set rngCheckCells = closestSearch(rngCheckCells, SearchDirection).EntireColumn
            End If

        ElseIf SearchDirection = xlPrevious Then
            If lngCol > minCol Then
                lngCol = lngCol - 1
                Set rngCheckCells = Intersect(Range(Columns(minCol), Columns(lngCol)), FoundCells.EntireColumn)
                Set rngCheckCells = closestSearch(rngCheckCells, SearchDirection).EntireColumn
            End If

        End If

    ElseIf SearchOrder = xlByRows Then
        If SearchDirection = xlNext Then
            If lngRow < maxRow Then
                lngRow = lngRow + 1
                Set rngCheckCells = Intersect(Range(Rows(lngRow), Rows(maxRow)), FoundCells.EntireRow)
                Set rngCheckCells = closestSearch(rngCheckCells, SearchDirection).EntireRow
            End If

        ElseIf SearchDirection = xlPrevious Then
            If lngRow > 1 Then
                lngRow = lngRow - 1
                Set rngCheckCells = Intersect(Range(Rows(minRow), Rows(lngRow)), FoundCells.EntireRow)
                Set rngCheckCells = closestSearch(rngCheckCells, SearchDirection).EntireRow
            End If
        End If
    End If

    If Not rngCheckCells Is Nothing Then
        Set rngResultCells = Intersect(rngCheckCells, FoundCells)
        If Not rngResultCells Is Nothing Then
            Set determineCell = closestSearch(rngResultCells, SearchDirection)
            Exit Function
        End If
    End If

    'Step 3
    Set rngCheckCells = Range(Cells(minRow, minCol), Cells(maxRow, maxCol))
    If SearchOrder = xlByColumns Then
        Set rngCheckCells = Intersect(rngCheckCells.EntireColumn, FoundCells.EntireColumn)
        Set rngCheckCells = closestSearch(rngCheckCells, SearchDirection).EntireColumn
    ElseIf SearchOrder = xlByRows Then
        Set rngCheckCells = Intersect(rngCheckCells.EntireRow, FoundCells.EntireRow)
        Set rngCheckCells = closestSearch(rngCheckCells, SearchDirection).EntireRow
    End If

    Set rngResultCells = Intersect(rngCheckCells, FoundCells)
    If Not rngResultCells Is Nothing Then
        Set determineCell = closestSearch(rngResultCells, SearchDirection)
    End If
    Exit Function

Catch:
    Call errorHandler("determineCell")
End Function

Private Function closestSearch(ByRef rngResultCells As Range, _
                               ByVal SearchDirection As XlSearchDirection) As Range
    On Error GoTo Catch

    Dim r As Range

    For Each r In rngResultCells.Areas
        'Search smallest
        If SearchDirection = xlNext Then
            If closestSearch Is Nothing Then
                Set closestSearch = r.Item(1)
            ElseIf closestSearch.Row > r.Item(1).Row Or closestSearch.Column > r.Item(1).Column Then
                Set closestSearch = r.Item(1)
            End If

        'Search biggest
        ElseIf SearchDirection = xlPrevious Then
            If closestSearch Is Nothing Then
                Set closestSearch = r.Item(r.Count)
            ElseIf closestSearch.Row < r.Item(r.Count).Row Or closestSearch.Column < r.Item(r.Count).Column Then
                Set closestSearch = r.Item(r.Count)
            End If

        End If
    Next r
    Exit Function

Catch:
    Call errorHandler("closestSearch")
End Function
