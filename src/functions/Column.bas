Attribute VB_Name = "F_Column"
Option Explicit
Option Private Module

Enum eTargetColumnType
    Entire
    ToLeftEndColumns
    ToRightEndColumns
    ToLeftOfCurrentRegionColumns
    ToRightOfCurrentRegionColumns
End Enum

Private Function GetTargetColumns(ByVal TargetType As eTargetColumnType) As Range
    'Error handling
    On Error GoTo Catch

    'Return Nothing when selection is not Range
    If TypeName(Selection) <> "Range" Then
        Set GetTargetColumns = Nothing
        Exit Function
    End If

    Dim rngSelection As Range
    Dim startColumn  As Long
    Dim endColumn    As Long

    Set rngSelection = Selection

    'Entire
    If TargetType = Entire Then
        With rngSelection
            If .Columns.Count > 1 Or gVim.Count1 = 1 Then
                Set GetTargetColumns = .EntireColumn
                Exit Function
            ElseIf gVim.Count1 > 1 Then
                startColumn = .Column
                endColumn = .Column + gVim.Count1 - 1
            End If
        End With

    'ToLeftEndColumns
    ElseIf TargetType = ToLeftEndColumns Then
        startColumn = ActiveSheet.UsedRange.Column
        endColumn = ActiveCell.Column

        'Out of range
        If startColumn > endColumn Then
            Set GetTargetColumns = Nothing
            Exit Function
        End If

    'ToRightEndColumns
    ElseIf TargetType = ToRightEndColumns Then
        With ActiveSheet.UsedRange
            startColumn = ActiveCell.Column
            endColumn = .Columns(.Columns.Count).Column
        End With

        'Out of range
        If startColumn > endColumn Then
            Set GetTargetColumns = Nothing
            Exit Function
        End If

    'ToLeftOfCurrentRegionColumns
    ElseIf TargetType = ToLeftOfCurrentRegionColumns Then
        startColumn = ActiveCell.CurrentRegion.Column
        endColumn = ActiveCell.Column

        'Out of range
        If startColumn > endColumn Then
            Set GetTargetColumns = Nothing
            Exit Function
        End If

    'ToRightOfCurrentRegionColumns
    ElseIf TargetType = ToRightOfCurrentRegionColumns Then
        With ActiveCell.CurrentRegion
            startColumn = ActiveCell.Column
            endColumn = .Columns(.Columns.Count).Column
        End With

        'Out of range
        If startColumn > endColumn Then
            Set GetTargetColumns = Nothing
            Exit Function
        End If

    End If

    With ActiveSheet
        If endColumn > .Columns.Count Then
            endColumn = .Columns.Count
        End If

        Set GetTargetColumns = .Range(.Columns(startColumn), .Columns(endColumn))
    End With
    Exit Function

Catch:
    Set GetTargetColumns = Nothing
    Call ErrorHandler("GetTargetColumns")
End Function

Private Function SelectColumnsInternal(ByVal TargetType As eTargetColumnType) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = GetTargetColumns(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Set savedCell = ActiveCell

    Target.Select
    savedCell.Activate

    SelectColumnsInternal = True
    Exit Function

Catch:
    Call ErrorHandler("SelectColumnsInternal")
End Function

Function selectColumns(Optional ByVal TargetType As eTargetColumnType = Entire) As Boolean
    On Error GoTo Catch

    Call SelectColumnsInternal(TargetType)
    Exit Function

Catch:
    Call ErrorHandler("SelectColumns")
End Function

Function InsertColumns(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = GetTargetColumns(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call RepeatRegister("InsertColumns")
    Call StopVisualMode

    Application.ScreenUpdating = False

    Set savedCell = ActiveCell
    Target.Select
    savedCell.Activate

    Call KeyStroke(Alt_ + I_, C_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertColumns")
End Function

Function AppendColumns(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = GetTargetColumns(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call RepeatRegister("AppendColumns")
    Call StopVisualMode

    Set savedCell = ActiveCell

    If Target.Item(Target.Count).Column < ActiveSheet.Columns.Count Then
        Set Target = Target.Offset(0, 1)
        Set savedCell = savedCell.Offset(0, 1)
    End If

    Application.ScreenUpdating = False

    Target.Select
    savedCell.Activate

    Call KeyStroke(Alt_ + I_, C_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("AppendColumns")
End Function

Function DeleteColumns(Optional ByVal TargetType As eTargetColumnType = Entire) As Boolean
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If SelectColumnsInternal(TargetType) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call RepeatRegister("DeleteColumns")
    Call StopVisualMode

    Call KeyStroke(Ctrl_ + Minus_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("DeleteColumns")
End Function

Function YankColumns(Optional ByVal TargetType As eTargetColumnType = Entire) As Boolean
    On Error GoTo Catch

    Dim Target As Range

    Set Target = GetTargetColumns(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Target.Copy
    Set gVim.Vars.LastYanked = Target

    Exit Function

Catch:
    Call ErrorHandler("YankColumns")
End Function

Function CutColumns(Optional ByVal TargetType As eTargetColumnType = Entire) As Boolean
    On Error GoTo Catch

    Dim Target As Range

    Set Target = GetTargetColumns(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Target.Cut
    Set gVim.Vars.LastYanked = Target

    Exit Function

Catch:
    Call ErrorHandler("CutColumns")
End Function

Function HideColumns(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If SelectColumnsInternal(Entire) = False Then
        Exit Function
    End If

    Call RepeatRegister("HideColumns")
    Call StopVisualMode

    Call KeyStroke(Ctrl_ + k0_)

Catch:
    Call ErrorHandler("HideColumns")
End Function

Function UnhideColumns(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If SelectColumnsInternal(Entire) = False Then
        Exit Function
    End If

    Call RepeatRegister("UnhideColumns")
    Call StopVisualMode

    'ref: https://excel.nj-clucker.com/ctrl-shift-0-not-working/
    'Call KeyStroke(Ctrl_ + Shift_ + k0_)
    Call KeyStroke(Alt_ + H_, O_, U_, L_)
    Exit Function

Catch:
    Call ErrorHandler("UnhideColumns")
End Function

Function GroupColumns(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If SelectColumnsInternal(Entire) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call RepeatRegister("GroupColumns")
    Call StopVisualMode

    Call KeyStroke(Alt_ + Shift_ + Right_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("GroupColumns")
End Function

Function UngroupColumns(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If SelectColumnsInternal(Entire) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call RepeatRegister("UngroupColumns")
    Call StopVisualMode

    Call KeyStroke(Alt_ + Shift_ + Left_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("UngroupColumns")
End Function

Function FoldColumnsGroup(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("FoldColumnsGroup")
    Call StopVisualMode

    Dim targetColumn As Long
    Dim i As Integer

    targetColumn = ActiveCell.Column

    For i = 1 To gVim.Count1
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(2," & targetColumn & ",FALSE)")
    Next i
    Exit Function

Catch:
    Call ErrorHandler("FoldColumnsGroup")
End Function

Function SpreadColumnsGroup(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("SpreadColumnsGroup")
    Call StopVisualMode

    Dim targetColumn As Long
    Dim i As Integer

    targetColumn = ActiveCell.Column

    For i = 1 To gVim.Count1
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(2," & targetColumn & ",TRUE)")
    Next i
    Exit Function

Catch:
    Call ErrorHandler("SpreadColumnsGroup")
End Function

Function AdjustColumnsWidth(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("AdjustColumnsWidth")
    Call StopVisualMode

    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Alt_ + H_, O_, I_)
    Exit Function

Catch:
    Call ErrorHandler("AdjustColumnsWidth")
End Function

Function SetColumnsWidth(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call StopVisualMode

    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Alt_ + H_, O_, W_)
    Exit Function

Catch:
    Call ErrorHandler("SetColumnsWidth")
End Function

Function NarrowColumnsWidth(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("NarrowColumnsWidth")
    Call StopVisualMode

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

    If currentWidth - gVim.Count1 < 0 Then
        targetColumns.EntireColumn.ColumnWidth = 0
    Else
        targetColumns.EntireColumn.ColumnWidth = currentWidth - gVim.Count1
    End If

    Set targetColumns = Nothing
    Exit Function

Catch:
    Call ErrorHandler("NarrowColumnsWidth")
End Function

Function WideColumnsWidth(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("WideColumnsWidth")
    Call StopVisualMode

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

    If currentWidth + gVim.Count1 > 255 Then
        targetColumns.EntireColumn.ColumnWidth = 255
    Else
        targetColumns.EntireColumn.ColumnWidth = currentWidth + gVim.Count1
    End If

    Set targetColumns = Nothing
    Exit Function

Catch:
    Call ErrorHandler("WideColumnsWidth")
End Function

Function ApplyColumnsLock(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim Target As Range

    Set Target = GetTargetColumns(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    With Target
        gVim.Vars.SetLockedColumns .Item(1).Column, .Item(.Count).Column
    End With

    Call gVim.Mode.Normal.ApplySelectionLock
    Call SetStatusBar(gVim.Msg.LockingRange & gVim.Vars.GetLockedRange())
    Exit Function

Catch:
    Call ErrorHandler("ApplyColumnsLock")
End Function

Function ClearColumnsLock(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    gVim.Vars.SetLockedColumns 0, 0

    Dim lockedRange As String
    lockedRange = gVim.Vars.GetLockedRange()

    If lockedRange = "" Then
        Call SetStatusBar
        Call SetStatusBarTemporarily(gVim.Msg.ClearedSelectionLock, 2000)
    Else
        Call SetStatusBar(gVim.Msg.LockingRange & gVim.Vars.GetLockedRange())
    End If
    Exit Function

Catch:
    Call ErrorHandler("ClearColumnsLock")
End Function
