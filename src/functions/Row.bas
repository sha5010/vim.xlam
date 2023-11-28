Attribute VB_Name = "F_Row"
Option Explicit
Option Private Module

Enum TargetRowType
    Entire
    ToTopRows
    ToBottomRows
    ToTopOfCurrentRegionRows
    ToBottomOfCurrentRegionRows
End Enum

Private Function getTargetRows(ByVal TargetType As TargetRowType) As Range
    'Error handling
    On Error GoTo Catch

    'Return Nothing when selection is not Range
    If TypeName(Selection) <> "Range" Then
        Set getTargetRows = Nothing
        Exit Function
    End If

    Dim rngSelection As Range
    Dim startRow     As Long
    Dim endRow       As Long

    Set rngSelection = Selection

    'Entire
    If TargetType = Entire Then
        With rngSelection
            If .Rows.Count > 1 Or gVim.Count1 = 1 Then
                Set getTargetRows = .EntireRow
                Exit Function
            ElseIf gVim.Count1 > 1 Then
                startRow = .Row
                endRow = .Row + gVim.Count1 - 1
            End If
        End With

    'ToTopRows
    ElseIf TargetType = ToTopRows Then
        startRow = ActiveSheet.UsedRange.Row
        endRow = ActiveCell.Row

        'Out of range
        If startRow > endRow Then
            Set getTargetRows = Nothing
            Exit Function
        End If

    'ToBottomRows
    ElseIf TargetType = ToBottomRows Then
        With ActiveSheet.UsedRange
            startRow = ActiveCell.Row
            endRow = .Rows(.Rows.Count).Row
        End With

        'Out of range
        If startRow > endRow Then
            Set getTargetRows = Nothing
            Exit Function
        End If

    'ToTopOfCurrentRegionRows
    ElseIf TargetType = ToTopOfCurrentRegionRows Then
        startRow = ActiveCell.CurrentRegion.Row
        endRow = ActiveCell.Row

        'Out of range
        If startRow > endRow Then
            Set getTargetRows = Nothing
            Exit Function
        End If

    'ToBottomOfCurrentRegionRows
    ElseIf TargetType = ToBottomOfCurrentRegionRows Then
        With ActiveCell.CurrentRegion
            startRow = ActiveCell.Row
            endRow = .Rows(.Rows.Count).Row
        End With

        'Out of range
        If startRow > endRow Then
            Set getTargetRows = Nothing
            Exit Function
        End If

    End If

    With ActiveSheet
        If endRow > .Rows.Count Then
            endRow = .Rows.Count
        End If

        Set getTargetRows = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    Set getTargetRows = Nothing
    Call errorHandler("getTargetRows")
End Function

Private Function selectRowsInternal(ByVal TargetType As TargetRowType) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = getTargetRows(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call stopVisualMode

    Set savedCell = ActiveCell

    Target.Select
    savedCell.Activate

    selectRowsInternal = True
    Exit Function

Catch:
    Call errorHandler("selectRowsInternal")
End Function

Function selectRows(Optional ByVal TargetType As TargetRowType = Entire)
    On Error GoTo Catch

    Call selectRowsInternal(TargetType)
    Exit Function

Catch:
    Call errorHandler("selectRows")
End Function

Function insertRows()
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = getTargetRows(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call repeatRegister("insertRows")
    Call stopVisualMode

    Application.ScreenUpdating = False

    Set savedCell = ActiveCell
    Target.Select
    savedCell.Activate

    Call keystroke(True, Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("insertRows")
End Function

Function appendRows()
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = getTargetRows(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call repeatRegister("appendRows")
    Call stopVisualMode

    Set savedCell = ActiveCell

    If Target.Item(Target.Count).Row < ActiveSheet.Rows.Count Then
        Set Target = Target.Offset(1, 0)
        Set savedCell = savedCell.Offset(1, 0)
    End If

    Application.ScreenUpdating = False

    Target.Select
    savedCell.Activate

    Call keystroke(True, Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("appendRows")
End Function

Function deleteRows(Optional ByVal TargetType As TargetRowType = Entire)
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If selectRowsInternal(TargetType) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call repeatRegister("deleteRows")
    Call stopVisualMode

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("deleteRows")
End Function

Function yankRows(Optional ByVal TargetType As TargetRowType = Entire)
    On Error GoTo Catch

    Dim Target As Range

    Set Target = getTargetRows(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call stopVisualMode

    Target.Copy
    Set gVim.Vars.LastYanked = Target

    Exit Function

Catch:
    Call errorHandler("yankRows")
End Function

Function cutRows(Optional ByVal TargetType As TargetRowType = Entire)
    On Error GoTo Catch

    Dim Target As Range

    Set Target = getTargetRows(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call stopVisualMode

    Target.Cut
    Set gVim.Vars.LastYanked = Target

    Exit Function

Catch:
    Call errorHandler("cutRows")
End Function

Function hideRows()
    On Error GoTo Catch

    If selectRowsInternal(Entire) = False Then
        Exit Function
    End If

    Call repeatRegister("hideRows")
    Call stopVisualMode

    Call keystroke(True, Ctrl_ + k9_)

Catch:
    Call errorHandler("hideRows")
End Function

Function unhideRows()
    On Error GoTo Catch

    If selectRowsInternal(Entire) = False Then
        Exit Function
    End If

    Call repeatRegister("unhideRows")
    Call stopVisualMode

    Call keystroke(True, Ctrl_ + Shift_ + k9_)
    Exit Function

Catch:
    Call errorHandler("unhideRows")
End Function

Function groupRows()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If selectRowsInternal(Entire) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call repeatRegister("groupRows")
    Call stopVisualMode

    Call keystroke(True, Alt_ + Shift_ + Right_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("groupRows")
End Function

Function ungroupRows()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If selectRowsInternal(Entire) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call repeatRegister("ungroupRows")
    Call stopVisualMode

    Call keystroke(True, Alt_ + Shift_ + Left_)

Catch:
    Application.ScreenUpdating = True
    Call errorHandler("ungroupRows")
End Function

Function foldRowsGroup()
    On Error GoTo Catch

    Call repeatRegister("foldRowsGroup")
    Call stopVisualMode

    Dim targetRow As Long
    Dim i As Integer

    targetRow = ActiveCell.Row

    For i = 1 To gVim.Count1
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",FALSE)")
    Next i
    Exit Function

Catch:
    Call errorHandler("foldRowsGroup")
End Function

Function spreadRowsGroup()
    On Error GoTo Catch

    Call repeatRegister("spreadRowsGroup")
    Call stopVisualMode

    Dim targetRow As Long
    Dim i As Integer

    targetRow = ActiveCell.Row

    For i = 1 To gVim.Count1
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",TRUE)")
    Next i
    Exit Function

Catch:
    Call errorHandler("spreadRowsGroup")
End Function

Function adjustRowsHeight()
    On Error GoTo Catch

    Call repeatRegister("adjustRowsHeight")
    Call stopVisualMode

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, A_)
    Exit Function

Catch:
    Call errorHandler("adjustRowsHeight")
End Function

Function setRowsHeight()
    On Error GoTo Catch

    Call stopVisualMode

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, H_)
    Exit Function

Catch:
    Call errorHandler("setRowsHeight")
End Function

Function narrowRowsHeight()
    On Error GoTo Catch

    Call repeatRegister("narrowRowsHeight")
    Call stopVisualMode

    Dim currentHeight As Double
    Dim targetRows As Range

    If TypeName(Selection) = "Range" Then
        If Not IsNull(Selection.EntireRow.RowHeight) Then
            currentHeight = Selection.EntireRow.RowHeight
        Else
            currentHeight = ActiveCell.EntireRow.RowHeight
        End If
        Set targetRows = Selection.EntireRow
    Else
        currentHeight = ActiveCell.EntireRow.RowHeight
        Set targetRows = ActiveCell.EntireRow
    End If

    If currentHeight - gVim.Count1 < 0 Then
        targetRows.EntireRow.RowHeight = 0
    Else
        targetRows.EntireRow.RowHeight = currentHeight - gVim.Count1
    End If

    Set targetRows = Nothing
    Exit Function

Catch:
    Call errorHandler("narrowRowsHeight")
End Function

Function wideRowsHeight()
    On Error GoTo Catch

    Call repeatRegister("wideRowsHeight")
    Call stopVisualMode

    Dim currentHeight As Double
    Dim targetRows As Range

    If TypeName(Selection) = "Range" Then
        If Not IsNull(Selection.EntireRow.RowHeight) Then
            currentHeight = Selection.EntireRow.RowHeight
        Else
            currentHeight = ActiveCell.EntireRow.RowHeight
        End If
        Set targetRows = Selection.EntireRow
    Else
        currentHeight = ActiveCell.EntireRow.RowHeight
        Set targetRows = ActiveCell.EntireRow
    End If

    If currentHeight + gVim.Count1 > 409.5 Then
        targetRows.EntireRow.RowHeight = 409.5
    Else
        targetRows.EntireRow.RowHeight = currentHeight + gVim.Count1
    End If

    Set targetRows = Nothing
    Exit Function

Catch:
    Call errorHandler("wideRowsHeight")
End Function
