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
Private WithEvents VisualApp As Application
Attribute VisualApp.VB_VarHelpID = -1

Public VisualMode As Boolean
Public VisualLine As Boolean
Private EditFlag As Boolean
Private BaseCell As Range
Private LastCell As Range

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If gVimMode Then
        Call disableIME
        Call recordToJumpList(Target)

        If EditFlag Then
            Call repeatRegister("changeSelectedCells", Target.Item(1).Text)   'Note: Undo command will no longer be available.
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

Public Sub StartEditing()
    EditFlag = True
End Sub

Public Sub StopEditing()
    EditFlag = False
End Sub

Private Sub VisualApp_SheetActivate(ByVal Sh As Object)
    If Not BaseCell.Parent Is Sh Then
        Call stopVisualMode
    End If
End Sub

Private Sub VisualApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim rowVisible As Double
    Dim colVisible As Double

    If BaseCell.Parent Is Sh Then
        Application.EnableEvents = False

        Set LastCell = ActiveCell
        rowVisible = ActiveWindow.ScrollRow
        colVisible = ActiveWindow.ScrollColumn

        If VisualMode Then
            Range(BaseCell, ActiveCell).Select
        ElseIf VisualLine Then
            Range(Rows(BaseCell.Row), Rows(ActiveCell.Row)).Select
        End If
        LastCell.Activate

        If ActiveWindow.ScrollRow <> rowVisible Then
            ActiveWindow.ScrollRow = rowVisible
        End If
        If ActiveWindow.ScrollColumn <> colVisible Then
            ActiveWindow.ScrollColumn = colVisible
        End If

        Application.EnableEvents = True
    Else
        Call Me.QuitVisualMode
    End If
End Sub

Private Sub SetupVisualMode()
    Set VisualApp = Application

    If TypeName(Selection) <> "Range" Then
        Set BaseCell = ActiveCell
        Set LastCell = ActiveCell
        Exit Sub
    End If

    Set BaseCell = Selection.Item(1)
    Set LastCell = Selection.Item(Selection.Count)

    If LastCell.Address = ActiveCell.Address Then
        Call SwapBase
    End If

    LastCell.Activate
End Sub

Public Sub StartVisualMode()
    Dim t As Range

    If VisualApp Is Nothing Then
        Call SetupVisualMode
    Else
        Application.EnableEvents = False

        Set t = ActiveCell
        Range(BaseCell, LastCell).Select
        t.Activate
        Set t = Nothing

        Application.EnableEvents = True
    End If

    VisualMode = True
    VisualLine = False
End Sub

Public Sub StartVisualLine()
    Dim t As Range

    If VisualApp Is Nothing Then
        Call SetupVisualMode
        Call selectRows
    Else
        Application.EnableEvents = False

        Set t = ActiveCell
        Range(Rows(BaseCell.Row), Rows(ActiveCell.Row)).Select
        t.Activate
        Set t = Nothing

        Application.EnableEvents = True
    End If

    VisualMode = False
    VisualLine = True
End Sub

Public Sub SwapBase()
    Dim t As Range

    Set t = BaseCell
    Set BaseCell = LastCell
    Set LastCell = t

    t.Activate
End Sub

Public Sub QuitVisualMode()
    Set VisualApp = Nothing
    Set BaseCell = Nothing
    VisualMode = False
    VisualLine = False
End Sub
