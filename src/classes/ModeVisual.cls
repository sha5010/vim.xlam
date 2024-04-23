VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ModeVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private NormalApp As cls_ModeNormal

Private cIsVisualLine As Boolean    ' Visual Line mode (True) / Visual mode (False)
Private cBaseCell As Range
Private cLastCell As Range

Private Sub App_SheetActivate(ByVal Sh As Object)
    If Not cBaseCell.Parent Is Sh Then
        Call StopVisualMode
    End If
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If cBaseCell.Parent Is Sh Then
        Call VisualSelect
    Else
        Call StopVisualMode
    End If
End Sub

Private Sub VisualSelect()
    Dim rowVisible As Double
    Dim colVisible As Double

    Application.EnableEvents = False

    Set cLastCell = ActiveCell
    rowVisible = ActiveWindow.ScrollRow
    colVisible = ActiveWindow.ScrollColumn

    If cIsVisualLine Then
        Range(Rows(cBaseCell.Row), Rows(ActiveCell.Row)).Select
    Else
        Range(cBaseCell, ActiveCell).Select
    End If
    cLastCell.Activate

    If ActiveWindow.ScrollRow <> rowVisible Then
        ActiveWindow.ScrollRow = rowVisible
    End If
    If ActiveWindow.ScrollColumn <> colVisible Then
        ActiveWindow.ScrollColumn = colVisible
    End If

    Application.EnableEvents = True

End Sub

Public Sub SwapBase()
    Dim t As Range

    Set t = cBaseCell
    Set cBaseCell = cLastCell
    Set cLastCell = t

    t.Activate
End Sub

Public Property Get IsVisualLine() As Boolean
    IsVisualLine = cIsVisualLine
End Property

Public Property Let IsVisualLine(ByVal arg As Boolean)
    cIsVisualLine = arg
    Call VisualSelect

    Dim statusMsg As String: statusMsg = gVim.Config.StatusPrefix
    If cIsVisualLine Then
        statusMsg = statusMsg & "-- VISUAL LINE (ESC to exit) --"
    Else
        statusMsg = statusMsg & "-- VISUAL (ESC to exit) --"
    End If

    Call SetStatusBar(statusMsg)
End Property

Private Sub Class_Initialize()
    ' Initialize application instances
    Set App = Application
    Set NormalApp = New cls_ModeNormal

    If Not TypeOf Selection Is Range Then
        Set cBaseCell = ActiveCell
        Set cLastCell = ActiveCell
        Exit Sub
    End If

    Set cBaseCell = Selection.Item(1)
    Set cLastCell = Selection.Item(Selection.Count)

    If cLastCell.Address = ActiveCell.Address Then
        Call SwapBase
    End If

    cLastCell.Activate
End Sub

Private Sub Class_Terminate()
    Set NormalApp = Nothing
    Set App = Nothing
    Set cBaseCell = Nothing
    Set cLastCell = Nothing
End Sub
