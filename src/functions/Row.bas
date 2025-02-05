Attribute VB_Name = "F_Row"
Option Explicit
Option Private Module

Enum eTargetRowType
    Entire
    ToTopRows
    ToBottomRows
    ToTopOfCurrentRegionRows
    ToBottomOfCurrentRegionRows
End Enum

Private Function GetTargetRows(ByVal TargetType As eTargetRowType) As Range
    'Error handling
    On Error GoTo Catch

    'Return Nothing when selection is not Range
    If TypeName(Selection) <> "Range" Then
        Set GetTargetRows = Nothing
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
                Set GetTargetRows = .EntireRow
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
            Set GetTargetRows = Nothing
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
            Set GetTargetRows = Nothing
            Exit Function
        End If

    'ToTopOfCurrentRegionRows
    ElseIf TargetType = ToTopOfCurrentRegionRows Then
        startRow = ActiveCell.CurrentRegion.Row
        endRow = ActiveCell.Row

        'Out of range
        If startRow > endRow Then
            Set GetTargetRows = Nothing
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
            Set GetTargetRows = Nothing
            Exit Function
        End If

    End If

    With ActiveSheet
        If endRow > .Rows.Count Then
            endRow = .Rows.Count
        End If

        Set GetTargetRows = .Range(.Rows(startRow), .Rows(endRow))
    End With
    Exit Function

Catch:
    Set GetTargetRows = Nothing
    Call ErrorHandler("GetTargetRows")
End Function

Private Function SelectRowsInternal(ByVal TargetType As eTargetRowType) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = GetTargetRows(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Set savedCell = ActiveCell

    Target.Select
    savedCell.Activate

    SelectRowsInternal = True
    Exit Function

Catch:
    Call ErrorHandler("SelectRowsInternal")
End Function

Function SelectRows(Optional ByVal TargetType As eTargetRowType = Entire) As Boolean
    On Error GoTo Catch

    Call SelectRowsInternal(TargetType)
    Exit Function

Catch:
    Call ErrorHandler("SelectRows")
End Function

Function InsertRows(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = GetTargetRows(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call RepeatRegister("InsertRows")
    Call StopVisualMode

    Application.ScreenUpdating = False

    Set savedCell = ActiveCell
    Target.Select
    savedCell.Activate

    Call KeyStroke(Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertRows")
End Function

Function AppendRows(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim savedCell As Range
    Dim Target As Range

    Set Target = GetTargetRows(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call RepeatRegister("AppendRows")
    Call StopVisualMode

    Set savedCell = ActiveCell

    If Target.Item(Target.Count).Row < ActiveSheet.Rows.Count Then
        Set Target = Target.Offset(1, 0)
        Set savedCell = savedCell.Offset(1, 0)
    End If

    Application.ScreenUpdating = False

    Target.Select
    savedCell.Activate

    Call KeyStroke(Alt_ + I_, R_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("AppendRows")
End Function

Function DeleteRows(Optional ByVal TargetType As eTargetRowType = Entire) As Boolean
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If SelectRowsInternal(TargetType) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call RepeatRegister("DeleteRows")
    Call StopVisualMode

    Call KeyStroke(Ctrl_ + Minus_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("DeleteRows")
End Function

Function YankRows(Optional ByVal TargetType As eTargetRowType = Entire) As Boolean
    On Error GoTo Catch

    Dim Target As Range

    Set Target = GetTargetRows(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Target.Copy
    Set gVim.Vars.LastYanked = Target

    Exit Function

Catch:
    Call ErrorHandler("YankRows")
End Function

Function CutRows(Optional ByVal TargetType As eTargetRowType = Entire) As Boolean
    On Error GoTo Catch

    Dim Target As Range

    Set Target = GetTargetRows(TargetType)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Target.Cut
    Set gVim.Vars.LastYanked = Target

    Exit Function

Catch:
    Call ErrorHandler("CutRows")
End Function

Function HideRows(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If SelectRowsInternal(Entire) = False Then
        Exit Function
    End If

    Call RepeatRegister("HideRows")
    Call StopVisualMode

    Call KeyStroke(Ctrl_ + k9_)

Catch:
    Call ErrorHandler("HideRows")
End Function

Function UnhideRows(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If SelectRowsInternal(Entire) = False Then
        Exit Function
    End If

    Call RepeatRegister("UnhideRows")
    Call StopVisualMode

    Call KeyStroke(Ctrl_ + Shift_ + k9_)
    Exit Function

Catch:
    Call ErrorHandler("UnhideRows")
End Function

Function GroupRows(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If SelectRowsInternal(Entire) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call RepeatRegister("GroupRows")
    Call StopVisualMode

    Call KeyStroke(Alt_ + Shift_ + Right_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("GroupRows")
End Function

Function UngroupRows(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Application.ScreenUpdating = False
    If SelectRowsInternal(Entire) = False Then
        Application.ScreenUpdating = True
        Exit Function
    End If

    Call RepeatRegister("UngroupRows")
    Call StopVisualMode

    Call KeyStroke(Alt_ + Shift_ + Left_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("UngroupRows")
End Function

Function FoldRowsGroup(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("FoldRowsGroup")
    Call StopVisualMode

    Dim targetRow As Long
    Dim i As Integer

    targetRow = ActiveCell.Row

    For i = 1 To gVim.Count1
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",FALSE)")
    Next i
    Exit Function

Catch:
    Call ErrorHandler("FoldRowsGroup")
End Function

Function SpreadRowsGroup(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("SpreadRowsGroup")
    Call StopVisualMode

    Dim targetRow As Long
    Dim i As Integer

    targetRow = ActiveCell.Row

    For i = 1 To gVim.Count1
        Call Application.ExecuteExcel4Macro("SHOW.DETAIL(1," & targetRow & ",TRUE)")
    Next i
    Exit Function

Catch:
    Call ErrorHandler("SpreadRowsGroup")
End Function

Function AdjustRowsHeight(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("AdjustRowsHeight")
    Call StopVisualMode

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Alt_ + H_, O_, A_)
    Exit Function

Catch:
    Call ErrorHandler("AdjustRowsHeight")
End Function

Function SetRowsHeight(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call StopVisualMode

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Alt_ + H_, O_, H_)
    Exit Function

Catch:
    Call ErrorHandler("SetRowsHeight")
End Function

Function NarrowRowsHeight(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("NarrowRowsHeight")
    Call StopVisualMode

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
    Call ErrorHandler("NarrowRowsHeight")
End Function

Function WideRowsHeight(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("WideRowsHeight")
    Call StopVisualMode

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
    Call ErrorHandler("WideRowsHeight")
End Function

Function ApplyRowsLock(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim Target As Range

    Set Target = GetTargetRows(Entire)
    If Target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    With Target
        gVim.Vars.SetLockedRows .Item(1).Row, .Item(.Count).Row
    End With

    Call gVim.Mode.Normal.ApplySelectionLock
    Call SetStatusBar(gVim.Msg.LockingRange & gVim.Vars.GetLockedRange())
    Exit Function

Catch:
    Call ErrorHandler("ApplyRowsLock")
End Function

Function ClearRowsLock(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    gVim.Vars.SetLockedRows 0, 0

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
    Call ErrorHandler("ClearRowsLock")
End Function
