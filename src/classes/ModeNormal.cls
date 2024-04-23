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

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If gVim.Enabled Then
        Call DisableIME
        Call RecordToJumpList(Target)

        If gVim.Vars.FromInsertCmd Then
            ' Note: Undo command will no longer be available
            Call RepeatRegister("ChangeSelectedCells", Target.Item(1).Text)
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

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
