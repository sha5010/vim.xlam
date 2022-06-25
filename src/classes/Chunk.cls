VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Chunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Fd_ As cls_Chunk
Private Bk_ As cls_Chunk
Private Content_ As Object

Private Sub Class_Initialize()
    Set Fd_ = Nothing
    Set Bk_ = Nothing
End Sub

Private Sub Class_Terminate()
    Set Fd_ = Nothing
    Set Bk_ = Nothing
    Set Content_ = Nothing
End Sub

Public Property Get Fd() As cls_Chunk
    Set Fd = Fd_
End Property

Public Property Set Fd(ByRef Fd As cls_Chunk)
    Set Fd_ = Fd
End Property

Public Property Get Bk() As cls_Chunk
    Set Bk = Bk_
End Property

Public Property Set Bk(ByRef Bk As cls_Chunk)
    Set Bk_ = Bk
End Property

Public Property Get Content() As Variant
    Set Content = Content_
End Property

Public Property Set Content(ByRef Content As Variant)
    Set Content_ = Content
End Property
