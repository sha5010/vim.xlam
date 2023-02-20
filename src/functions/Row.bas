Attribute VB_Name = "F_Row"
Option Explicit
Option Private Module

Function selectRows()
    On Error GoTo Catch

    Dim t As Range
    Dim startRow As Long
    Dim endRow As Long

    With ActiveWorkbook.ActiveSheet
        Set t = ActiveCell

        If gCount = 1 And TypeName(Selection) = "Range" Then
            If Selection.Rows.Count > 1 Then
                Selection.EntireRow.Select
                Exit Function
            End If
        End If

        startRow = t.Row
        endRow = startRow + gCount - 1

        If endRow > .Rows.Count Then
            endRow = .Rows.Count
            startRow = endRow - gCount + 1
        End If

        .Range(.Rows(startRow), .Rows(endRow)).Select
        t.Activate
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("selectRows")
    End If
End Function

Function insertRows()
    On Error GoTo Catch

    Call repeatRegister("insertRows")
    Call stopVisualMode

    Dim savedColumn As Long

    savedColumn = ActiveCell.Column

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).EntireRow.Select
    Else
        Selection.EntireRow.Select
    End If
    Cells(ActiveCell.Row, savedColumn).Activate

    Call keystroke(True, Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("insertRows")
    End If
End Function

Function appendRows()
    On Error GoTo Catch

    Call repeatRegister("appendRows")
    Call stopVisualMode

    Dim savedColumn As Long

    savedColumn = ActiveCell.Column

    Application.ScreenUpdating = False
    If Selection.Row < ActiveSheet.Rows.Count Then
        Selection.Offset(1, 0).Select
    End If

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).EntireRow.Select
    Else
        Selection.EntireRow.Select
    End If
    Cells(ActiveCell.Row, savedColumn).Activate

    Call keystroke(True, Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("appendRows")
    End If
End Function

Function deleteRows()
    On Error GoTo Catch

    Call repeatRegister("deleteRows")
    Call stopVisualMode

    Dim t As Range

    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Then
        ActiveCell.Select
    End If

    Set t = ActiveCell

    With ActiveSheet
        If gCount > 1 Then
            .Range(.Rows(Selection.Row), .Rows(WorksheetFunction.Min(Selection.Row + gCount - 1, .Rows.Count))).Select
        Else
            Selection.EntireRow.Select
        End If
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
    t.Activate
    Set t = Nothing

    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("deleteRows")
    End If
End Function

Function deleteToTopRows()
    On Error GoTo Catch

    Call repeatRegister("deleteToTopRows")
    Call stopVisualMode

    With ActiveSheet
        .Range(.Rows(1), .Rows(ActiveCell.Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("deleteToTopRows")
    End If
End Function

Function deleteToBottomRows()
    On Error GoTo Catch

    Call repeatRegister("deleteToBottomRows")
    Call stopVisualMode

    With ActiveSheet
        If ActiveCell.Row > .UsedRange.Item(.UsedRange.Count).Row Then
            Exit Function
        End If

        .Range(.Rows(ActiveCell.Row), .Rows(.UsedRange.Item(.UsedRange.Count).Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("deleteToBottomRows")
    End If
End Function

Function deleteToTopOfCurrentRegionRows()
    On Error GoTo Catch

    Call repeatRegister("deleteToTopOfCurrentRegionRows")
    Call stopVisualMode

    With ActiveSheet
        .Range(.Rows(ActiveCell.CurrentRegion.Item(1).Row), .Rows(ActiveCell.Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("deleteToTopOfCurrentRegionRows")
    End If
End Function

Function deleteToBottomOfCurrentRegionRows()
    On Error GoTo Catch

    Call repeatRegister("deleteToBottomOfCurrentRegionRows")
    Call stopVisualMode

    With ActiveSheet
        .Range(.Rows(ActiveCell.Row), .Rows(ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("deleteToBottomOfCurrentRegionRows")
    End If
End Function

Function yankRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = ActiveCell.Row
    endRow = WorksheetFunction.Min(ActiveCell.Row + gCount - 1, ActiveSheet.Rows.Count)

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Copy
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("yankRows")
    End If
End Function

Function yankToTopRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = 1
    endRow = ActiveCell.Row

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Copy
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("yankToTopRows")
    End If
End Function

Function yankToBottomRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    With ActiveSheet
        startRow = ActiveCell.Row
        endRow = .UsedRange.Item(.UsedRange.Count).Row

        If startRow > endRow Then
            Exit Function
        End If

        .Range(.Rows(startRow), .Rows(endRow)).Copy
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("yankToBottomRows")
    End If
End Function

Function yankToTopOfCurrentRegionRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = ActiveCell.CurrentRegion.Item(1).Row
    endRow = ActiveCell.Row

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Copy
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("yankToTopOfCurrentRegionRows")
    End If
End Function

Function yankToBottomOfCurrentRegionRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = ActiveCell.Row
    endRow = ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Copy
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("yankToBottomOfCurrentRegionRows")
    End If
End Function

Function cutRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = ActiveCell.Row
    endRow = WorksheetFunction.Min(ActiveCell.Row + gCount - 1, ActiveSheet.Rows.Count)

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Cut
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("cutRows")
    End If
End Function

Function cutToTopRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = 1
    endRow = ActiveCell.Row

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Cut
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("cutToTopRows")
    End If
End Function

Function cutToBottomRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    With ActiveSheet
        startRow = ActiveCell.Row
        endRow = .UsedRange.Item(.UsedRange.Count).Row

        If startRow > endRow Then
            Exit Function
        End If

        .Range(.Rows(startRow), .Rows(endRow)).Cut
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("cutToBottomRows")
    End If
End Function

Function cutToTopOfCurrentRegionRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = ActiveCell.CurrentRegion.Item(1).Row
    endRow = ActiveCell.Row

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Cut
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("cutToTopOfCurrentRegionRows")
    End If
End Function

Function cutToBottomOfCurrentRegionRows()
    On Error GoTo Catch

    Call stopVisualMode

    Dim startRow As Long
    Dim endRow As Long

    startRow = ActiveCell.Row
    endRow = ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row

    With ActiveSheet
        .Range(.Rows(startRow), .Rows(endRow)).Cut
        Set gLastYanked = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("cutToBottomOfCurrentRegionRows")
    End If
End Function

Function hideRows()
    On Error GoTo Catch

    Call repeatRegister("hideRows")
    Call stopVisualMode

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + k9_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("hideRows")
    End If
End Function

Function unhideRows()
    On Error GoTo Catch

    Call repeatRegister("unhideRows")
    Call stopVisualMode

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + k9_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("unhideRows")
    End If
End Function

Function groupRows()
    On Error GoTo Catch

    Call repeatRegister("groupRows")
    Call stopVisualMode

    Dim t As Range

    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Then
        ActiveCell.Select
    End If

    Set t = ActiveCell

    With ActiveSheet
        If gCount > 1 Then
            .Range(.Rows(Selection.Row), .Rows(WorksheetFunction.Min(Selection.Row + gCount - 1, .Rows.Count))).Select
        Else
            Selection.EntireRow.Select
        End If
    End With

    Call keystroke(True, Alt_ + Shift_ + Right_)

Catch:
    t.Activate
    Set t = Nothing

    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("groupRows")
    End If
End Function

Function ungroupRows()
    On Error GoTo Catch

    Call repeatRegister("ungroupRows")
    Call stopVisualMode

    Dim t As Range

    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Then
        ActiveCell.Select
    End If

    Set t = ActiveCell

    With ActiveSheet
        If gCount > 1 Then
            .Range(.Rows(Selection.Row), .Rows(WorksheetFunction.Min(Selection.Row + gCount - 1, .Rows.Count))).Select
        Else
            Selection.EntireRow.Select
        End If
    End With

    Call keystroke(True, Alt_ + Shift_ + Left_)

Catch:
    t.Activate
    Set t = Nothing

    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Call errorHandler("ungroupRows")
    End If
End Function

Function foldRowsGroup()
    On Error GoTo Catch

    Call repeatRegister("foldRowsGroup")
    Call stopVisualMode

    Dim targetRow As Long
    Dim i As Integer

    targetRow = ActiveCell.Row

    For i = 1 To gCount
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",FALSE)")
    Next i
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("foldRowsGroup")
    End If
End Function

Function spreadRowsGroup()
    On Error GoTo Catch

    Call repeatRegister("spreadRowsGroup")
    Call stopVisualMode

    Dim targetRow As Long
    Dim i As Integer

    targetRow = ActiveCell.Row

    For i = 1 To gCount
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",TRUE)")
    Next i
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("spreadRowsGroup")
    End If
End Function

Function adjustRowsHeight()
    On Error GoTo Catch

    Call repeatRegister("adjustRowsHeight")
    Call stopVisualMode

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, A_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("adjustRowsHeight")
    End If
End Function

Function setRowsHeight()
    On Error GoTo Catch

    Call stopVisualMode

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, H_)
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("setRowsHeight")
    End If
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

    If currentHeight - gCount < 0 Then
        targetRows.EntireRow.RowHeight = 0
    Else
        targetRows.EntireRow.RowHeight = currentHeight - gCount
    End If

    Set targetRows = Nothing
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("narrowRowsHeight")
    End If
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

    If currentHeight + gCount > 409.5 Then
        targetRows.EntireRow.RowHeight = 409.5
    Else
        targetRows.EntireRow.RowHeight = currentHeight + gCount
    End If

    Set targetRows = Nothing
    Exit Function

Catch:
    If Err.Number <> 0 Then
        Call errorHandler("wideRowsHeight")
    End If
End Function
