VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_EventHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Public WithEvents TempApp As Application
Attribute TempApp.VB_VarHelpID = -1

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If gVimMode Then
        If IMEStatus <> vbIMEModeOff Then
            SendKeys "{kanji}"
        End If

        If Not JumpList Is Nothing Then
            If JumpList.Latest Is Nothing Then
                Call JumpList.Add(Target)
            ElseIf Target.Address <> JumpList.Latest.Address Then
                Call JumpList.Add(Target)
            ElseIf Target.Parent.Name <> JumpList.Latest.Parent.Name Then
                Call JumpList.Add(Target)
            ElseIf Target.Parent.Parent.FullName <> JumpList.Latest.Parent.Parent.FullName Then
                Call JumpList.Add(Target)
            End If
        End If
    End If
End Sub

Private Sub App_WorkbookActivate(ByVal wb As Workbook)
    If gVimMode Then
        On Error Resume Next

        Unload UF_Cmd
        Unload UF_FindForm
        Unload UF_ColorPicker

        Load UF_Cmd
        Load UF_FindForm
        Load UF_ColorPicker

        On Error GoTo 0
    End If
End Sub

Private Sub TempApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Call resumeVim
End Sub
