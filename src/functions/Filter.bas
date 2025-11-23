Attribute VB_Name = "F_Filter"
Option Explicit
Option Private Module

Private Enum eFilterCriteriaType
    eEquals
    eNotEquals
    eContains
    eNotContains
    eLessThan
    eGreaterThan
    eLessEqual
    eGreaterEqual
    eBeginsWith
    eEndsWith
    eNotBeginsWith
    eNotEndsWith
    eBlanks
    eNotBlanks
    eAboveAverage
    eBelowAverage
    eTop10Items
    eBottom10Items
    eTop10Percent
    eBottom10Percent
End Enum

'/*
' * Checks if the active sheet is a Worksheet.
' *
' * @returns {Boolean} - True if the active sheet is a Worksheet, False otherwise.
' */
Private Function IsTargetWorksheet() As Boolean
    If TypeOf ActiveSheet Is Worksheet Then
        IsTargetWorksheet = True
        Exit Function
    End If

    Call SetStatusBarTemporarily(gVim.Msg.NotWorksheet, 3000)
    IsTargetWorksheet = False
End Function

'/*
' * Gets the AutoFilter object for the active cell (ListObject or Worksheet).
' *
' * @param {Worksheet} ws - The active worksheet.
' * @returns {AutoFilter} - The AutoFilter object, or Nothing if not applicable.
' */
Private Function GetTargetAutoFilter(ByRef ws As Worksheet) As AutoFilter
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ActiveCell.ListObject

    If Not lo Is Nothing Then
        Set GetTargetAutoFilter = lo.AutoFilter
    Else
        Set GetTargetAutoFilter = ws.AutoFilter
    End If
    On Error GoTo 0
End Function

'/*
' * Checks if AutoFilter is currently on for the active context (ListObject or Worksheet).
' *
' * @param {Worksheet} ws - The active worksheet.
' * @returns {Boolean} - True if AutoFilter is on.
' */
Private Function IsAutoFilterOn(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ActiveCell.ListObject

    If Not lo Is Nothing Then
        IsAutoFilterOn = lo.ShowAutoFilterDropDown
    Else
        IsAutoFilterOn = ws.AutoFilterMode
    End If
    On Error GoTo 0
End Function

'/*
' * Toggles the AutoFilter on the active worksheet.
' *
' * If AutoFilter is currently active, it will be turned off.
' * If AutoFilter is currently inactive, it will be applied to the UsedRange.
' * A status message will be displayed in case of no data or an error.
' */
Function ToggleAutoFilter(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lo As ListObject
    Set lo = ActiveCell.ListObject

    If Not lo Is Nothing Then
        ' Table context
        Call KeyStroke(Alt_ + J_, T_, B_)

        If lo.ShowAutoFilterDropDown Then
            Call SetStatusBarTemporarily(gVim.Msg.AutoFilterOff, 2000)
        Else
            Call SetStatusBarTemporarily(gVim.Msg.AutoFilterOn, 2000)
        End If
    Else
        ' Worksheet context
        Call KeyStroke(Alt_ + A_, T_)

        If ws.AutoFilterMode Then
            Call SetStatusBarTemporarily(gVim.Msg.AutoFilterOff, 2000)
        Else
            Call SetStatusBarTemporarily(gVim.Msg.AutoFilterOn, 2000)
        End If
    End If

    Exit Function

Catch:
    Call ErrorHandler("ToggleAutoFilter")
End Function

'/*
' * Clears all filters currently applied to the active worksheet.
' *
' * If filters are active, they will be removed. If no filters are applied,
' * a status message will be displayed.
' */
Function ClearAllFilters(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lo As ListObject
    Set lo = ActiveCell.ListObject
    Dim cleared As Boolean
    cleared = False

    If Not lo Is Nothing Then
        If lo.ShowAutoFilter Then
            ' Check if any filter is applied
            If Not lo.AutoFilter Is Nothing Then
                If lo.AutoFilter.FilterMode Then
                    Call KeyStroke(Alt_ + A_, C_)
                    cleared = True
                End If
            End If
        End If
    Else
        If ws.FilterMode Then
            Call KeyStroke(Alt_ + A_, C_)
            cleared = True
        End If
    End If

    If cleared Then
        Call SetStatusBarTemporarily(gVim.Msg.FiltersCleared, 2000)
    Else
        Call SetStatusBarTemporarily(gVim.Msg.NoFilterApplied, 3000)
    End If

    Exit Function

Catch:
    Call ErrorHandler("ClearAllFilters")
End Function



'/*
' * Filters the active column by the value of the active cell.
' *
' * If AutoFilter is not active, it will be applied first.
' * If the active cell is empty, no filter will be applied and a message will be displayed.
' * The filter applies to the column of the active cell, using a partial match (contains).
' */
Function FilterByActiveCellValue(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Dim filterValue As String
    filterValue = CStr(ActiveCell.Value)

    If Trim(filterValue) = "" Then
        Call SetStatusBarTemporarily(gVim.Msg.NoDataInSelectedCells, 3000)
        Exit Function
    End If

    If Not IsAutoFilterOn(ws) Then
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
        Exit Function
    End If

    Dim af As AutoFilter
    Set af = GetTargetAutoFilter(ws)

    If af Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
        Exit Function
    End If

    ' Ensure the active cell is within the filter range
    If Not Intersect(ActiveCell, af.Range) Is Nothing Then
        ' Adjust field index relative to the range
        Dim fieldIndex As Long
        fieldIndex = activeCol - af.Range.Column + 1

        ' Determine filter type based on active cell value type
        ' Use Excel's IsNumber to respect how Excel sees the data (e.g. '123 is text)
        Dim isNumericOrDate As Boolean
        isNumericOrDate = Application.WorksheetFunction.IsNumber(ActiveCell)

        If isNumericOrDate Then
            ' Exact match for numeric/date types
            af.Range.AutoFilter Field:=fieldIndex, Criteria1:="=" & filterValue, Operator:=xlAnd
        Else
            ' Partial match for strings and others
            af.Range.AutoFilter Field:=fieldIndex, Criteria1:="*" & filterValue & "*", Operator:=xlAnd
        End If
    Else
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
    End If

    Exit Function

Catch:
    Call ErrorHandler("FilterByActiveCellValue")
End Function

Private Function ApplyAutoFilter(ByVal Field As Long, ByVal Criteria1 As Variant, Optional ByVal Operator As XlAutoFilterOperator = xlAnd, Optional ByVal Criteria2 As Variant) As Boolean
    If Not IsTargetWorksheet() Then
        ApplyAutoFilter = False
        Exit Function
    End If

    On Error GoTo Catch

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If Not IsAutoFilterOn(ws) Then
        ' If AutoFilter is not active, apply it to the current region.
        ' This assumes the active cell is within the data range.
        Call KeyStroke(Alt_ + A_, T_)
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterOn, 2000)
    End If

    Dim af As AutoFilter
    Set af = GetTargetAutoFilter(ws)

    If af Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
        ApplyAutoFilter = False
        Exit Function
    End If

    ' Ensure the active cell is within the filter range before applying filter
    If Not Intersect(ActiveCell, af.Range) Is Nothing Then
        ' Adjust field index relative to the range
        Dim fieldIndex As Long
        fieldIndex = Field - af.Range.Column + 1

        If IsMissing(Criteria2) Then
            af.Range.AutoFilter Field:=fieldIndex, Criteria1:=Criteria1, Operator:=Operator
        Else
            af.Range.AutoFilter Field:=fieldIndex, Criteria1:=Criteria1, Operator:=Operator, Criteria2:=Criteria2
        End If
        ApplyAutoFilter = True
    Else
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
        ApplyAutoFilter = False
    End If

    Exit Function

Catch:
    Call ErrorHandler("ApplyAutoFilter")
    ApplyAutoFilter = False
End Function

Private Function FilterByInput(criteriaType As eFilterCriteriaType) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Dim filterValue As String
    filterValue = InputBox(gVim.Msg.InputFilterValuePrompt)

    If Trim(filterValue) = "" Then
        Call SetStatusBarTemporarily(gVim.Msg.NoDataInSelectedCells, 3000)
        Exit Function
    End If

    Dim af As AutoFilter
    Set af = GetTargetAutoFilter(ActiveSheet)

    If af Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
        Exit Function
    End If

    If Not Intersect(ActiveCell, af.Range) Is Nothing Then
        Select Case criteriaType
            Case eEquals
                Call ApplyAutoFilter(activeCol, filterValue)
            Case eNotEquals
                Call ApplyAutoFilter(activeCol, "<>" & filterValue)
            Case eContains
                Call ApplyAutoFilter(activeCol, "*" & filterValue & "*")
            Case eNotContains
                Call ApplyAutoFilter(activeCol, "<>*" & filterValue & "*")
        End Select
    Else
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
    End If

    Exit Function

Catch:
    Call ErrorHandler("FilterByInput")
    FilterByInput = False
End Function

Public Function FilterByInputEquals(Optional ByVal g As String) As Boolean
    FilterByInputEquals = FilterByInput(eEquals)
End Function

Public Function FilterByInputNotEquals(Optional ByVal g As String) As Boolean
    FilterByInputNotEquals = FilterByInput(eNotEquals)
End Function

Public Function FilterByInputContains(Optional ByVal g As String) As Boolean
    FilterByInputContains = FilterByInput(eContains)
End Function

Public Function FilterByInputNotContains(Optional ByVal g As String) As Boolean
    FilterByInputNotContains = FilterByInput(eNotContains)
End Function

'/*
' * Filters the active column to show only blank cells.
' */
Public Function FilterByBlanks(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Call ApplyAutoFilter(activeCol, "=")

    Exit Function

Catch:
    Call ErrorHandler("FilterByBlanks")
End Function

'/*
' * Filters the active column to show only non-blank cells.
' */
Public Function FilterByNonBlanks(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Call ApplyAutoFilter(activeCol, "<>")

    Exit Function

Catch:
    Call ErrorHandler("FilterByNonBlanks")
End Function

'/*
' * Clears the filter applied to the active column.
' * If no filter is applied to the column, a message will be displayed.
' */
Public Function ClearCurrentColumnFilter(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim af As AutoFilter
    Set af = GetTargetAutoFilter(ws)

    If af Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.NoFilterApplied, 3000)
        Exit Function
    End If

    ' Identify all columns in the selection that are part of the AutoFilter range
    Dim fieldsToClear As New Dictionary
    Dim area As Range
    Dim isect As Range
    Dim c As Long
    Dim fieldIndex As Long

    For Each area In Selection.Areas
        Set isect = Intersect2(area, af.Range)
        If Not isect Is Nothing Then
            For c = isect.Column To isect.Column + isect.Columns.Count - 1
                fieldIndex = c - af.Range.Column + 1
                If Not fieldsToClear.Exists(fieldIndex) Then
                    fieldsToClear.Add fieldIndex, fieldIndex
                End If
            Next c
        End If
    Next area

    If fieldsToClear.Count = 0 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoFilterAppliedToColumn, 3000)
        Exit Function
    End If

    ' Identify all currently filtered columns
    Dim filteredFields As New Dictionary
    Dim i As Long
    For i = 1 To af.Filters.Count
        If af.Filters(i).On Then
            filteredFields.Add i, i
        End If
    Next i

    If filteredFields.Count = 0 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoFilterApplied, 3000)
        Exit Function
    End If

    ' Check if all filtered columns are included in the selection
    Dim allFilteredSelected As Boolean
    allFilteredSelected = True
    Dim filteredField As Variant

    For Each filteredField In filteredFields.Keys
        If Not fieldsToClear.Exists(filteredField) Then
            allFilteredSelected = False
            Exit For
        End If
    Next filteredField

    ' If all filtered columns are selected, execute ClearAllFilters
    If allFilteredSelected Then
        Call ClearAllFilters
        Exit Function
    End If

    ' Clear filter for each selected column
    Dim cleared As Boolean
    Dim clearField As Variant

    For Each clearField In fieldsToClear.Keys
        fieldIndex = CLng(clearField)
        If fieldIndex <= af.Filters.Count Then
            If af.Filters(fieldIndex).On Then
                af.Range.AutoFilter Field:=fieldIndex
                cleared = True
            End If
        End If
    Next clearField

    If cleared Then
        Call SetStatusBarTemporarily(gVim.Msg.FilterClearedForColumn, 2000)
    Else
        Call SetStatusBarTemporarily(gVim.Msg.NoFilterAppliedToColumn, 3000)
    End If

    Exit Function

Catch:
    Call ErrorHandler("ClearCurrentColumnFilter")
End Function

'/*
' * Filters the active column based on an input value and a specified operator prefix.
' * This is a helper function for comparison-based filters.
' */
Private Function FilterByInputWithOperator(ByVal operatorPrefix As String, ByVal prompt As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Dim filterValue As String
    filterValue = InputBox(prompt)

    If Trim(filterValue) = "" Then
        Call SetStatusBarTemporarily(gVim.Msg.NoDataInSelectedCells, 3000)
        Exit Function
    End If

    Call ApplyAutoFilter(activeCol, operatorPrefix & filterValue)

    Exit Function

Catch:
    Call ErrorHandler("FilterByInputWithOperator")
    FilterByInputWithOperator = False
End Function

'/*
' * Filters the active column for values less than the input.
' */
Public Function FilterByInputLessThan(Optional ByVal g As String) As Boolean
    FilterByInputLessThan = FilterByInputWithOperator("<", gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values greater than the input.
' */
Public Function FilterByInputGreaterThan(Optional ByVal g As String) As Boolean
    FilterByInputGreaterThan = FilterByInputWithOperator(">", gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values less than or equal to the input.
' */
Public Function FilterByInputLessEqual(Optional ByVal g As String) As Boolean
    FilterByInputLessEqual = FilterByInputWithOperator("<=", gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values greater than or equal to the input.
' */
Public Function FilterByInputGreaterEqual(Optional ByVal g As String) As Boolean
    FilterByInputGreaterEqual = FilterByInputWithOperator(">=", gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column based on an input value and a specified wildcard pattern.
' * This is a helper function for wildcard-based filters (e.g., begins with, ends with).
' */
Private Function FilterByInputWithWildcard(ByVal criteriaType As eFilterCriteriaType, ByVal prompt As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Dim filterValue As String
    filterValue = InputBox(prompt)

    If Trim(filterValue) = "" Then
        Call SetStatusBarTemporarily(gVim.Msg.NoDataInSelectedCells, 3000)
        Exit Function
    End If

    Dim criteria As String
    Select Case criteriaType
        Case eBeginsWith
            criteria = filterValue & "*"
        Case eEndsWith
            criteria = "*" & filterValue
        Case eNotBeginsWith
            criteria = "<>" & filterValue & "*"
        Case eNotEndsWith
            criteria = "<>" & "*" & filterValue
        Case Else
            ' Should not happen
            FilterByInputWithWildcard = False
            Exit Function
    End Select

    Call ApplyAutoFilter(activeCol, criteria)
    Exit Function

Catch:
    Call ErrorHandler("FilterByInputWithWildcard")
End Function

'/*
' * Filters the active column for values that begin with the input.
' */
Public Function FilterByInputBeginsWith(Optional ByVal g As String) As Boolean
    FilterByInputBeginsWith = FilterByInputWithWildcard(eBeginsWith, gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values that end with the input.
' */
Public Function FilterByInputEndsWith(Optional ByVal g As String) As Boolean
    FilterByInputEndsWith = FilterByInputWithWildcard(eEndsWith, gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values that do not begin with the input.
' */
Public Function FilterByInputNotBeginsWith(Optional ByVal g As String) As Boolean
    FilterByInputNotBeginsWith = FilterByInputWithWildcard(eNotBeginsWith, gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values that do not end with the input.
' */
Public Function FilterByInputNotEndsWith(Optional ByVal g As String) As Boolean
    FilterByInputNotEndsWith = FilterByInputWithWildcard(eNotEndsWith, gVim.Msg.InputFilterValuePrompt)
End Function

'/*
' * Filters the active column for values above the average.
' */
Public Function FilterByAboveAverage(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Call ApplyAutoFilter(activeCol, xlFilterAboveAverage, xlFilterDynamic)

    Exit Function

Catch:
    Call ErrorHandler("FilterByAboveAverage")
End Function

'/*
' * Filters the active column for values below the average.
' */
Public Function FilterByBelowAverage(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Call ApplyAutoFilter(activeCol, xlFilterBelowAverage, xlFilterDynamic)

    Exit Function

Catch:
    Call ErrorHandler("FilterByBelowAverage")
End Function

'/*
' * Filters the active column based on a top/bottom N value or percentage.
' * This is a helper function for top/bottom N filters.
' */
Private Function FilterByTopBottomN(ByVal op As XlAutoFilterOperator, ByVal prompt As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Dim nValue As String
    nValue = InputBox(prompt, Default:="10")

    If Trim(nValue) = "" Then
        Call SetStatusBarTemporarily(gVim.Msg.NoDataInSelectedCells, 3000)
        Exit Function
    End If

    If Not IsNumeric(nValue) Then
        Call SetStatusBarTemporarily(gVim.Msg.InvalidNumberInput, 3000)
        Exit Function
    End If

    Call ApplyAutoFilter(activeCol, CInt(nValue), op)

    Exit Function

Catch:
    Call ErrorHandler("FilterByTopBottomN")
End Function

'/*
' * Filters the active column for the top N items.
' */
Public Function FilterByTopNItems(Optional ByVal g As String) As Boolean
    FilterByTopNItems = FilterByTopBottomN(xlTop10Items, gVim.Msg.InputTopNPrompt)
End Function

'/*
' * Filters the active column for the bottom N items.
' */
Public Function FilterByBottomNItems(Optional ByVal g As String) As Boolean
    FilterByBottomNItems = FilterByTopBottomN(xlBottom10Items, gVim.Msg.InputBottomNPrompt)
End Function

'/*
' * Filters the active column for the top N percent.
' */
Public Function FilterByTopNPercent(Optional ByVal g As String) As Boolean
    FilterByTopNPercent = FilterByTopBottomN(xlTop10Percent, gVim.Msg.InputTopNPercentPrompt)
End Function

'/*
' * Filters the active column for the bottom N percent.
' */
Public Function FilterByBottomNPercent(Optional ByVal g As String) As Boolean
    FilterByBottomNPercent = FilterByTopBottomN(xlBottom10Percent, gVim.Msg.InputBottomNPercentPrompt)
End Function

'/*
' * Shows the Excel built-in filter dialog and focuses the search box.
' * Selects the header cell of the current column, sends Alt+Down, then E.
' */
Public Function ShowFilterDialog(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim af As AutoFilter
    Set af = GetTargetAutoFilter(ws)

    If af Is Nothing Then
        Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
        Exit Function
    End If

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    ' Ensure active cell is within filter range (columns)
    If activeCol < af.Range.Column Or activeCol > af.Range.Column + af.Range.Columns.Count - 1 Then
         Call SetStatusBarTemporarily(gVim.Msg.AutoFilterRequiresData, 3000)
         Exit Function
    End If

    Dim fieldIndex As Long
    fieldIndex = activeCol - af.Range.Column + 1

    ' Use the row of the AutoFilter range and the active column to identify the header cell
    ws.Cells(af.Range.Row, activeCol).Select

    ' Alt + Down, then E
    Call KeyStroke(Alt_ + Down_, E_)

    Exit Function

Catch:
    Call ErrorHandler("ShowFilterDialog")
End Function

'/*
' * Sorts the active column by the cell color of the active cell.
' */
Public Function SortColor(Optional ByVal g As String) As Boolean
    If Not IsTargetWorksheet() Then
        Exit Function
    End If

    On Error GoTo Catch

    Dim activeCol As Long
    activeCol = ActiveCell.Column

    Dim activeCellColor As Long
    activeCellColor = ActiveCell.Interior.Color

    ' Check if the active cell has a fill color
    If activeCellColor = xlNone Then
        Call SetStatusBarTemporarily(gVim.Msg.NoColorInActiveCell, 3000)
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim af As AutoFilter
    Set af = GetTargetAutoFilter(ws)

    ' If AutoFilter is active, use it to sort
    If Not af Is Nothing Then
        ' Check if the active cell is within the AutoFilter range
        If Not Intersect(ActiveCell, af.Range) Is Nothing Then
             af.Sort.SortFields.Clear
             af.Sort.SortFields.Add(af.Range.Columns(activeCol - af.Range.Column + 1), _
                xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = activeCellColor

            With af.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Exit Function
        End If
    End If

    ' If AutoFilter is not active or active cell is outside, try standard sort (less reliable for color without range detection, but let's try current region)
    ' For safety and consistency with other functions, we might want to enforce AutoFilter or just Sort the current region.
    ' Let's use the Sort method on the CurrentRegion or Selection.

    Dim targetRange As Range
    If Selection.Cells.Count > 1 Then
        Set targetRange = Selection
    Else
        Set targetRange = ActiveCell.CurrentRegion
    End If

    ' Sort by color
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=Intersect(targetRange, ws.Columns(activeCol)), _
            SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields(1).SortOnValue.Color = activeCellColor

        .SetRange targetRange
        .Header = xlYes ' Assume header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Exit Function

Catch:
    Call ErrorHandler("SortColor")
End Function
