VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ModeNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private WithEvents AppForLock As Application
Attribute AppForLock.VB_VarHelpID = -1

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If gVim.Enabled Then
        Call DisableIME
        Call RecordToJumpList(Target)

        If gVim.Vars.FromInsertCmd Then
            ' Note: Undo command will no longer be available
            Call RepeatRegister("ChangeSelectedCells", Target.Item(1).Formula)
        End If
    End If
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    If gVim.Enabled Then
        On Error Resume Next

        Unload UF_Cmd
        Unload UF_CmdLine
        Unload UF_ColorPicker
        Unload UF_SheetPicker

        Load UF_Cmd
        Load UF_CmdLine

        On Error GoTo 0
    End If
End Sub

Public Sub ApplySelectionLock()
    If AppForLock Is Nothing Then
        Set AppForLock = Application
        Set gVim.Vars.LockedSheet = ActiveSheet
    End If
End Sub

Private Sub ClearSelectionLock()
    Set AppForLock = Nothing
    With gVim.Vars
        Set .LockedSheet = Nothing
        .SetLockedRows 0, 0
        .SetLockedColumns 0, 0
        Call SetStatusBar
    End With
End Sub

Private Sub AppForLock_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Catch

    If Not gVim.Vars.LockedSheet Is Sh Then
        Call ClearSelectionLock
        Exit Sub
    End If

    With gVim.Vars
        Dim lockedRange As Range
        If .LockedRowsBegin > 0 And .LockedColumnsBegin > 0 Then
            Set lockedRange = .LockedSheet.Range(.LockedSheet.Cells(.LockedRowsBegin, .LockedColumnsBegin), .LockedSheet.Cells(.LockedRowsEnd, .LockedColumnsEnd))
        ElseIf .LockedRowsBegin = 0 And .LockedColumnsBegin > 0 Then
            Set lockedRange = .LockedSheet.Range(.LockedSheet.Columns(.LockedColumnsBegin), .LockedSheet.Columns(.LockedColumnsEnd))
        ElseIf .LockedRowsBegin > 0 And .LockedColumnsBegin = 0 Then
            Set lockedRange = .LockedSheet.Range(.LockedSheet.Rows(.LockedRowsBegin), .LockedSheet.Rows(.LockedRowsEnd))
        Else
            Call ClearSelectionLock
            Exit Sub
        End If

        Dim newRange As Range
        Set newRange = Intersect2(Target, lockedRange)

        If Not newRange Is Nothing Then
            Application.EnableEvents = False
            newRange.Select
            Application.EnableEvents = True
            Exit Sub
        End If

        Dim newRow As Long
        Dim newCol As Long
        If .LockedColumnsBegin > 0 Then
            If .LockedColumnsBegin > Target(Target.Count).Column Then
                newCol = .LockedColumnsBegin
            ElseIf .LockedColumnsEnd < Target(1).Column Then
                newCol = .LockedColumnsEnd
            End If
        End If

        If .LockedRowsBegin > 0 Then
            If .LockedRowsBegin > Target(Target.Count).Row Then
                newRow = .LockedRowsBegin
            ElseIf .LockedRowsEnd < Target(1).Row Then
                newRow = .LockedRowsEnd
            End If
        End If

        If newRow > 0 And newCol > 0 Then
            Set newRange = .LockedSheet.Cells(newRow, newCol)
        ElseIf newRow > 0 And newCol = 0 Then
            Set newRange = Intersect2(Target.EntireColumn, .LockedSheet.Rows(newRow))
        ElseIf newRow = 0 And newCol > 0 Then
            Set newRange = Intersect2(Target.EntireRow, .LockedSheet.Columns(newCol))
        End If

        If Not newRange Is Nothing Then
            Application.EnableEvents = False
            newRange.Select
            Application.EnableEvents = True
        End If
    End With
    Exit Sub

Catch:
    Call ErrorHandler("ModeNormal.AppForLock_SheetSelectionChange")
End Sub

Private Sub AppForLock_SheetActivate(ByVal Sh As Object)
    If Not gVim.Vars.LockedSheet Is Sh Then
        Call ClearSelectionLock
    End If
End Sub

Private Sub AppForLock_WorkbookDeactivate(ByVal Wb As Workbook)
    If gVim.Vars.LockedSheet.Parent Is Wb Then
        Call ClearSelectionLock
    End If
End Sub

Private Sub Class_Initialize()
    Set App = Application
    If Not gVim.Vars.LockedSheet Is Nothing Then
        If gVim.Vars.LockedSheet Is ActiveSheet Then
            Call ApplySelectionLock
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
    Set AppForLock = Nothing
End Sub
