VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Vars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cls_Vim

Public LastYanked As Range      ' Hold target cells with Copy, Cut command
Public ExtendRange As Range     ' Hold multi selected cells

Public FromInsertCmd As Boolean ' Whether started by insert mode
