VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_DblCircList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MainArena As cls_Chunk
Private CurrentChunk As cls_Chunk
Private Count_ As Long
Private StoreMax As Long

'                         Oldest                          Newest
'    ############# Fd   +--------+ Fd   +--------+ Fd   +--------+ Fd   ############# Fd
' -> #           # ---> |        | ---> |        | ---> |        | ---> #           # ->
'    # MainArena #      | Chunk1 |      | Chunk2 |      | Chunk3 |      # MainArena #
' <- #           # <--- |        | <--- |        | <--- |        | <--- #           # <-
' Bk #############   Bk +--------+   Bk +--------+   Bk +--------+   Bk #############

Private Sub Class_Initialize()
    Set MainArena = New cls_Chunk
    Set MainArena.Fd = MainArena
    Set MainArena.Bk = MainArena

    Set CurrentChunk = MainArena

    Count_ = 0
    StoreMax = 100
End Sub

Private Sub Class_Terminate()
    Call ClearAll

    Set MainArena.Fd = Nothing
    Set MainArena.Bk = Nothing
    Set MainArena = Nothing
    Set CurrentChunk = Nothing
End Sub

Private Sub Unlink(ByRef chunk As cls_Chunk)
    Dim Fd As cls_Chunk
    Dim Bk As cls_Chunk

    If chunk Is MainArena Then
        Exit Sub
    ElseIf chunk Is Nothing Then
        Exit Sub
    End If

    Set Fd = chunk.Fd
    Set Bk = chunk.Bk

    Set Fd.Bk = Bk
    Set Bk.Fd = Fd

    Set chunk = Nothing

    Count_ = Count_ - 1
End Sub

Private Sub UnlinkDeadObjects()
    Dim chunk As cls_Chunk
    Set chunk = MainArena.Fd

    While Not chunk Is MainArena
        Set chunk = chunk.Fd
        If Not IsAlive(chunk.Bk) Then
            Call Unlink(chunk.Bk)
        End If
    Wend
End Sub

Private Sub CleanUp(Optional ByVal isAddLater As Boolean = False)
    Dim buf As Long: buf = 0
    If isAddLater Then
        buf = 1
    End If

    If Count_ + buf > StoreMax Then
        Call UnlinkDeadObjects
    End If

    While Count_ + buf > StoreMax
        Call Unlink(MainArena.Fd)  'Unlink oldest chunk
    Wend
End Sub

Private Sub UnlinkAllFd(ByRef chunk As cls_Chunk)
    While Not chunk.Fd Is MainArena
        Call Unlink(chunk.Fd)
    Wend
End Sub

Private Sub UnlinkAllBk(ByRef chunk As cls_Chunk)
    While Not chunk.Bk Is MainArena
        Call Unlink(chunk.Bk)
    Wend
End Sub

Private Function IsAlive(ByRef chunk As cls_Chunk) As Boolean
    On Error GoTo Catch

    If chunk Is MainArena Then
        IsAlive = True
        Exit Function
    ElseIf Not (IsObject(chunk.Content) Or VarType(chunk.Content) = vbDataObject) Then
        IsAlive = True
        Exit Function
    End If

    'Access random property
    IsAlive = (TypeName(chunk.Content.Application.Application) = "")

Catch:
    If Err.Number = 424 Then
        IsAlive = False  'No substance
    ElseIf Err.Number = 438 Or Err.Number = 0 Then
        IsAlive = True   'Has substance
    Else
        IsAlive = True
        Call DebugPrint("Unexpected error type: " & Err.Number & "  Desc:" & Err.Description, "cls_MainArena.AccessCheck")
    End If
    On Error GoTo 0
End Function

Private Function Search(Content As Variant) As cls_Chunk
    Dim chunk As cls_Chunk
    Set chunk = MainArena.Bk

    While Not chunk Is MainArena
        If Not IsAlive(chunk) Then
            Set chunk = chunk.Bk
            Call Unlink(chunk.Fd)
        ElseIf chunk.IsSame(Content) Then
            Set Search = chunk
            Exit Function
        Else
            Set chunk = chunk.Bk
        End If
    Wend

    Set Search = Nothing
End Function

Public Function InList(Content As Variant) As Boolean
    InList = (Not Search(Content) Is Nothing)
End Function

Public Sub Add(Content As Variant, Optional ByVal CurrentToLatest As Boolean = True, _
                                   Optional ByVal AvoidDuplication As Boolean = True)

    Dim newChunk As cls_Chunk
    Dim latestChunk As cls_Chunk

    If AvoidDuplication Then
        ' Unlink if same content is exists
        Call Unlink(Search(Content))
    End If

    Call CleanUp(isAddLater:=True)

    Set newChunk = New cls_Chunk
    newChunk.SetContent Content

    Set latestChunk = MainArena.Bk
    Set latestChunk.Fd = newChunk
    Set newChunk.Bk = latestChunk

    Set newChunk.Fd = MainArena
    Set MainArena.Bk = newChunk

    If CurrentToLatest Then
        Set CurrentChunk = MainArena
    End If

    Count_ = Count_ + 1
End Sub

Public Sub ClearAll()
    Call UnlinkAllFd(MainArena)

    If Count_ <> 0 Then
        Call DebugPrint("Assertation Warning: Count is " & Count_ & " after ClearAll", "cls_MainArena.ClearAll")
    End If

    Count_ = 0
End Sub

Public Function Back() As Variant
    Do While Not CurrentChunk.Bk Is MainArena
        If Not IsAlive(CurrentChunk.Bk) Then
            Call Unlink(CurrentChunk.Bk)
        Else
            Set CurrentChunk = CurrentChunk.Bk
            With CurrentChunk
                If IsObject(.Content) Or VarType(.Content) = vbDataObject Then
                    Set Back = .Content
                Else
                    Back = .Content
                End If
            End With
            Exit Function
        End If
    Loop

    Set Back = Nothing
End Function

Public Property Get Current() As Variant
    If CurrentChunk Is MainArena Then
        Current = Empty
    ElseIf Not IsAlive(CurrentChunk) Then
        Current = Empty
        Call Unlink(CurrentChunk)
        Set CurrentChunk = MainArena
    Else
        With CurrentChunk
            If IsObject(.Content) Or VarType(.Content) = vbDataObject Then
                Set Current = .Content
            Else
                Current = .Content
            End If
        End With
    End If
End Property

Public Function Forward() As Variant
    If CurrentChunk Is MainArena Then
        Set Forward = Nothing
        Exit Function
    End If

    Do While Not CurrentChunk.Fd Is MainArena
        If Not IsAlive(CurrentChunk.Fd) Then
            Call Unlink(CurrentChunk.Fd)
        Else
            Set CurrentChunk = CurrentChunk.Fd
            With CurrentChunk
                If IsObject(.Content) Or VarType(.Content) = vbDataObject Then
                    Set Forward = .Content
                Else
                    Forward = .Content
                End If
            End With
            Exit Function
        End If
    Loop

    Set Forward = Nothing
End Function

Public Property Get Latest() As Variant
    Dim chunk As cls_Chunk

    Set chunk = MainArena.Bk
    While Not chunk Is MainArena
        If Not IsAlive(chunk) Then
            Set chunk = chunk.Bk
            Call Unlink(chunk.Fd)
        Else
            With chunk
                If IsObject(.Content) Or VarType(.Content) = vbDataObject Then
                    Set Latest = .Content
                Else
                    Latest = .Content
                End If
            End With
            Exit Property
        End If
    Wend

    Set Latest = Nothing
End Property

Public Property Get Count() As Long
    Count = Count_
End Property

Public Property Get Max() As Long
    Max = StoreMax
End Property

Public Property Let Max(ByVal arg As Long)
    If arg < 0 Then
        arg = 0
    ElseIf arg > 1000 Then
        arg = 1000
    End If
    StoreMax = arg
End Property



'Debug purpose only
Private Sub DescribeArena()
    Dim chunkFd As cls_Chunk
    Dim chunkBk As cls_Chunk
    Dim cnt As Integer
    Dim ret As String, appendString As String

    ret = "Fd"

    Set chunkFd = MainArena.Fd
    While Not chunkFd Is MainArena And cnt < 3
        If TypeName(chunkFd.Content) = "Range" Then
            appendString = chunkFd.Content.Address(False, False)
        Else
            appendString = TypeName(chunkFd.Content)
        End If

        If chunkFd Is CurrentChunk Then
            appendString = "*" & appendString & "*"
        End If

        ret = ret & " > " & appendString
        Set chunkFd = chunkFd.Fd

        cnt = cnt + 1
    Wend

    Set chunkBk = MainArena
    For cnt = cnt To 1 Step -1
        Set chunkBk = chunkBk.Bk
        If chunkBk Is chunkFd Then
            Exit For
        End If
    Next cnt

    If cnt = 0 And Not chunkBk Is chunkFd Then
        ret = ret & " ... "
    End If

    While Not chunkBk Is MainArena
        If TypeName(chunkBk.Content) = "Range" Then
            appendString = chunkBk.Content.Address(False, False)
        Else
            appendString = TypeName(chunkBk.Content)
        End If

        If chunkBk Is CurrentChunk Then
            appendString = "*" & appendString & "*"
        End If

        If InStr(ret, "...") > 0 Then
            ret = ret & appendString & " < "
        Else
            ret = ret & " > " & appendString
        End If

        Set chunkBk = chunkBk.Fd
    Wend

    If InStr(ret, "...") = 0 Then
        ret = ret & " < Bk"
    Else
        ret = ret & "Bk"
    End If

    Call DebugPrint(ret, "MainArena")
End Sub
