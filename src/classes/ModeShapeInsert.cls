VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ModeShapeInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Call ChangeToNormalMode
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    Call ChangeToNormalMode
End Sub

Private Sub Class_Initialize()
    ' Initialize application instances
    Set App = Application
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
