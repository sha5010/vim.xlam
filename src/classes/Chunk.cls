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
Private Content_ As Variant

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
    If IsObject(Content_) Or VarType(Content_) = vbDataObject Then
        Set Content = Content_
    Else
        Content = Content_
    End If
End Property

Public Sub SetContent(ByRef Content As Variant)
    If IsObject(Content) Or VarType(Content) = vbDataObject Then
        Set Content_ = Content
    Else
        Content_ = Content
    End If
End Sub

Public Function IsSame(ByRef Content As Variant) As Boolean
    If TypeName(Content_) <> TypeName(Content) Then
        IsSame = False
        Exit Function
    End If

    On Error GoTo Catch

    If TypeOf Content Is Range Then
        IsSame = True
        IsSame = IsSame And (Content_.Address = Content.Address)
        IsSame = IsSame And (Content_.Parent Is Content.Parent)
        IsSame = IsSame And (Content_.Parent.Parent Is Content.Parent.Parent)
    ElseIf IsObject(Content) Or VarType(Content) = vbDataObject Then
        IsSame = (Content_ Is Content)
    Else
        IsSame = (Content_ = Content)
    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        IsSame = False
        Err.Clear
    Else
        Call ErrorHandler("IsSame")
    End If
End Function
