Attribute VB_Name = "F_Row"
Option Explicit
Option Private Module

Function selectRows()
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
End Function

Function insertRows()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).EntireRow.Select
    Else
        Selection.EntireRow.Select
    End If

    Call keystroke(True, Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
End Function

Function appendRows()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If Selection.Row < ActiveSheet.Rows.Count Then
        Selection.Offset(1, 0).Select
    End If

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).EntireRow.Select
    Else
        Selection.EntireRow.Select
    End If

    Call keystroke(True, Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
End Function

Function deleteRows()
    Dim t As Range
    On Error GoTo Catch

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
End Function

Function deleteToTopRows()
    On Error GoTo Catch

    With ActiveSheet
        .Range(.Rows(1), .Rows(ActiveCell.Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function deleteToBottomRows()
    On Error GoTo Catch

    With ActiveSheet
        If ActiveCell.Row > .UsedRange.Item(.UsedRange.Count).Row Then
            Exit Function
        End If

        .Range(.Rows(ActiveCell.Row), .Rows(.UsedRange.Item(.UsedRange.Count).Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function deleteToTopOfCurrentRegionRows()
    On Error GoTo Catch

    With ActiveSheet
        .Range(.Rows(ActiveCell.CurrentRegion.Item(1).Row), .Rows(ActiveCell.Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function deleteToBottomOfCurrentRegionRows()
    On Error GoTo Catch

    With ActiveSheet
        .Range(.Rows(ActiveCell.Row), .Rows(ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Row)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function yankRows()
    On Error GoTo Catch

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
End Function

Function yankToTopRows()
    On Error GoTo Catch

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
End Function

Function yankToBottomRows()
    On Error GoTo Catch

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
End Function

Function yankToTopOfCurrentRegionRows()
    On Error GoTo Catch

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
End Function

Function yankToBottomOfCurrentRegionRows()
    On Error GoTo Catch

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
End Function

Function cutRows()
    On Error GoTo Catch

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
End Function

Function cutToTopRows()
    On Error GoTo Catch

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
End Function

Function cutToBottomRows()
    On Error GoTo Catch

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
End Function

Function cutToTopOfCurrentRegionRows()
    On Error GoTo Catch

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
End Function

Function cutToBottomOfCurrentRegionRows()
    On Error GoTo Catch

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
End Function

Function hideRows()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + k9_)

Catch:
End Function


Function unhideRows()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Ctrl_ + Shift_ + k9_)

Catch:
End Function

Function groupRows()
    Dim t As Range
    On Error GoTo Catch

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
End Function

Function ungroupRows()
    Dim t As Range
    On Error GoTo Catch

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
End Function

Function foldRowsGroup()
    Dim targetRow As Long
    Dim i As Integer

    On Error GoTo Catch

    targetRow = ActiveCell.Row

    For i = 1 To gCount
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",FALSE)")
    Next i
    Exit Function

Catch:
End Function


Function spreadRowsGroup()
    Dim targetRow As Long
    Dim i As Integer

    On Error GoTo Catch

    targetRow = ActiveCell.Row

    For i = 1 To gCount
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",TRUE)")
    Next i
    Exit Function

Catch:
End Function

Function adjustRowsHeight()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, A_)

Catch:
End Function

Function setRowsHeight()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(gCount, Selection.Columns.Count).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, H_)

Catch:
End Function

Function narrowRowsHeight()
    On Error GoTo Catch

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
End Function

Function wideRowsHeight()
    On Error GoTo Catch

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
End Function
