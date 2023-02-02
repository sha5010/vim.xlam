VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AddinUninstall()
    Call mapResetAll
End Sub

Private Sub Workbook_Open()
    gDebugMode = Not ThisWorkbook.IsAddin
    Call C_Core.startVim
End Sub
