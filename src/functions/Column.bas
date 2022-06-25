Attribute VB_Name = "F_Column"
Option Explicit
Option Private Module

Function selectColumns()
    Dim t As Range
    Dim startColumn As Long
    Dim endColumn As Long

    With ActiveWorkbook.ActiveSheet
        Set t = ActiveCell

        If gCount = 1 And TypeName(Selection) = "Range" Then
            If Selection.Columns.Count > 1 Then
                Selection.EntireColumn.Select
                Exit Function
            End If
        End If

        startColumn = t.Column
        endColumn = startColumn + gCount - 1

        If endColumn > .Columns.Count Then
            endColumn = .Columns.Count
            startColumn = endColumn - gCount + 1
        End If

        .Range(.Columns(startColumn), .Columns(endColumn)).Select
        t.Activate
    End With
End Function

Function insertColumns()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).EntireColumn.Select
    Else
        Selection.EntireColumn.Select
    End If

    Call keystroke(True, Alt_ + I_, C_)

Catch:
    Application.ScreenUpdating = True
End Function

Function appendColumns()
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If Selection.Column < ActiveSheet.Columns.Count Then
        Selection.Offset(0, 1).Select
    End If

    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).EntireColumn.Select
    Else
        Selection.EntireColumn.Select
    End If

    Call keystroke(True, Alt_ + I_, C_)

Catch:
    Application.ScreenUpdating = True
End Function

Function deleteColumns()
    Dim t As Range
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Then
        ActiveCell.Select
    End If
    Set t = ActiveCell

    With ActiveSheet
        If gCount > 1 Then
            .Range(.Columns(Selection.Column), .Columns(WorksheetFunction.Min(Selection.Column + gCount - 1, .Columns.Count))).Select
        Else
            Selection.EntireColumn.Select
        End If
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
    t.Activate
    Set t = Nothing

    Application.ScreenUpdating = True
End Function

Function deleteToLeftEndColumns()
    On Error GoTo Catch

    With ActiveSheet
        .Range(.Columns(1), .Columns(ActiveCell.Columns)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function deleteToRightEndColumns()
    On Error GoTo Catch

    With ActiveSheet
        If ActiveCell.Column > .UsedRange.Item(.UsedRange.Count).Column Then
            Exit Function
        End If

        .Range(.Columns(ActiveCell.Column), .Columns(.UsedRange.Item(.UsedRange.Count).Column)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function deleteToLeftOfCurrentRegionColumns()
    On Error GoTo Catch

    With ActiveSheet
        .Range(.Columns(ActiveCell.CurrentRegion.Item(1).Column), .Columns(ActiveCell.Column)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function deleteToRightOfCurrentRegionColumns()
    On Error GoTo Catch

    With ActiveSheet
        .Range(.Columns(ActiveCell.Column), .Columns(ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Column)).Select
    End With

    Call keystroke(True, Ctrl_ + Minus_)

Catch:
End Function

Function yankColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = ActiveCell.Column
    endColumn = WorksheetFunction.Min(startColumn + gCount - 1, ActiveSheet.Columns.Count)

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Copy
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function yankToLeftEndColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = 1
    endColumn = ActiveCell.Column

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Copy
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function yankToRightEndColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    With ActiveSheet
        startColumn = ActiveCell.Column
        endColumn = .UsedRange.Item(.UsedRange.Count).Column

        If startColumn > endColumn Then
            Exit Function
        End If

        .Range(.Columns(startColumn), .Columns(endColumn)).Copy
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With

Catch:
End Function

Function yankToLeftOfCurrentRegionColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = ActiveCell.CurrentRegion.Item(1).Column
    endColumn = ActiveCell.Column

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Copy
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function yankToRightOfCurrentRegionColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = ActiveCell.Column
    endColumn = ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Column

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Copy
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function cutColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = ActiveCell.Column
    endColumn = WorksheetFunction.Min(startColumn + gCount - 1, ActiveSheet.Columns.Count)

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Cut
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function cutToLeftEndColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = 1
    endColumn = ActiveCell.Column

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Cut
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function cutToRightEndColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    With ActiveSheet
        startColumn = ActiveCell.Column
        endColumn = .UsedRange.Item(.UsedRange.Count).Column

        If startColumn > endColumn Then
            Exit Function
        End If

        .Range(.Columns(startColumn), .Columns(endColumn)).Cut
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function cutToLeftOfCurrentRegionColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = ActiveCell.CurrentRegion.Item(1).Column
    endColumn = ActiveCell.Column

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Cut
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function

Function cutToRightOfCurrentRegionColumns()
    On Error GoTo Catch

    Dim startColumn As Long
    Dim endColumn As Long

    startColumn = ActiveCell.Column
    endColumn = ActiveCell.CurrentRegion.Item(ActiveCell.CurrentRegion.Count).Column

    With ActiveSheet
        .Range(.Columns(startColumn), .Columns(endColumn)).Cut
        Set gLastYanked = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
End Function


Function hideColumns()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Ctrl_ + k0_)

Catch:
End Function


Function unhideColumns()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    'ref: https://excel.nj-clucker.com/ctrl-shift-0-not-working/
    Call keystroke(True, Ctrl_ + Shift_ + k0_)

Catch:
End Function

Function groupColumns()
    Dim t As Range
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Then
        ActiveCell.Select
    End If

    Set t = ActiveCell

    With ActiveSheet
        If gCount > 1 Then
            .Range(.Columns(Selection.Column), .Columns(WorksheetFunction.Min(Selection.Column + gCount - 1, .Columns.Count))).Select
        Else
            Selection.EntireColumn.Select
        End If
    End With

    Call keystroke(True, Alt_ + Shift_ + Right_)

Catch:
    t.Activate
    Set t = Nothing

    Application.ScreenUpdating = True
End Function

Function ungroupColumns()
    Dim t As Range
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Then
        ActiveCell.Select
    End If

    Set t = ActiveCell

    With ActiveSheet
        If gCount > 1 Then
            .Range(.Columns(Selection.Column), .Columns(WorksheetFunction.Min(Selection.Column + gCount - 1, .Columns.Count))).Select
        Else
            Selection.EntireColumn.Select
        End If
    End With

    Call keystroke(True, Alt_ + Shift_ + Left_)

Catch:
    t.Activate
    Set t = Nothing

    Application.ScreenUpdating = True
End Function

Function foldColumnsGroup()
    Dim targetColumn As Long
    Dim i As Integer

    On Error GoTo Catch

    targetColumn = ActiveCell.Column

    For i = 1 To gCount
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(2," & targetColumn & ",FALSE)")
    Next i
    Exit Function

Catch:
End Function


Function spreadColumnsGroup()
    Dim targetColumn As Long
    Dim i As Integer

    On Error GoTo Catch

    targetColumn = ActiveCell.Column

    For i = 1 To gCount
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(2," & targetColumn & ",TRUE)")
    Next i
    Exit Function

Catch:
End Function

Function adjustColumnsWidth()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, I_)

Catch:
End Function

Function setColumnsWidth()
    On Error GoTo Catch

    If gCount > 1 Then
        Selection.Resize(Selection.Rows.Count, gCount).Select
    End If

    Call keystroke(True, Alt_ + H_, O_, W_)

Catch:
End Function


Function narrowColumnsWidth()
    On Error GoTo Catch

    Dim currentWidth As Double
    Dim targetColumns As Range

    If TypeName(Selection) = "Range" Then
        If Not IsNull(Selection.EntireColumn.ColumnWidth) Then
            currentWidth = Selection.EntireColumn.ColumnWidth
        Else
            currentWidth = ActiveCell.EntireColumn.ColumnWidth
        End If
        Set targetColumns = Selection.EntireColumn
    Else
        currentWidth = ActiveCell.EntireColumn.ColumnWidth
        Set targetColumns = ActiveCell.EntireColumn
    End If

    If currentWidth - gCount < 0 Then
        targetColumns.EntireColumn.ColumnWidth = 0
    Else
        targetColumns.EntireColumn.ColumnWidth = currentWidth - gCount
    End If

    Set targetColumns = Nothing
    Exit Function

Catch:
End Function

Function wideColumnsWidth()
    On Error GoTo Catch

    Dim currentWidth As Double
    Dim targetColumns As Range

    If TypeName(Selection) = "Range" Then
        If Not IsNull(Selection.EntireColumn.ColumnWidth) Then
            currentWidth = Selection.EntireColumn.ColumnWidth
        Else
            currentWidth = ActiveCell.EntireColumn.ColumnWidth
        End If
        Set targetColumns = Selection.EntireColumn
    Else
        currentWidth = ActiveCell.EntireColumn.ColumnWidth
        Set targetColumns = ActiveCell.EntireColumn
    End If

    If currentWidth + gCount > 255 Then
        targetColumns.EntireColumn.ColumnWidth = 255
    Else
        targetColumns.EntireColumn.ColumnWidth = currentWidth + gCount
    End If

    Set targetColumns = Nothing
    Exit Function

Catch:
End Function
