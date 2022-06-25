VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_MainArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MainArena As cls_Chunk
Private CurrentChunk As cls_Chunk
Private Count_ As Long
Private StoreMax As Long

Private Sub Class_Initialize()
    Set MainArena = New cls_Chunk
    Set MainArena.Fd = MainArena
    Set MainArena.Bk = MainArena

    Set CurrentChunk = MainArena

    Count_ = 0
    StoreMax = 0  'Infinity
End Sub

Private Sub Class_Terminate()
    Set MainArena = Nothing
    Set CurrentChunk = Nothing
End Sub

Private Sub Unlink(ByRef chunk As cls_Chunk)
    Dim Fd As cls_Chunk
    Dim Bk As cls_Chunk

    If chunk Is MainArena Then
        Exit Sub
    End If

    Set Fd = chunk.Fd
    Set Bk = chunk.Bk

    Set Fd.Bk = Bk
    Set Bk.Fd = Fd

    Set chunk = Nothing

    Count_ = Count_ - 1
End Sub

Private Sub CleanUp()
    Dim chunk As cls_Chunk

    Set chunk = MainArena.Fd

    While Not chunk.Fd Is Nothing And Not chunk Is MainArena
        If chunk.Content Is Nothing Then
            Call Unlink(chunk)
        End If
        Set chunk = chunk.Fd
    Wend
End Sub

Private Sub UnlinkAllFd(ByRef chunk As cls_Chunk)
    While Not chunk.Fd Is MainArena
        Call Unlink(chunk.Fd)
    Wend
End Sub

Private Function AccessCheck(ByRef chunk As cls_Chunk) As Boolean
    On Error GoTo Catch

    'Access random property
    AccessCheck = (TypeName(chunk.Content.Application.Application) = "")

Catch:
    If Err.Number = 424 Then
        AccessCheck = False  'No substance
    ElseIf Err.Number = 438 Or Err.Number = 0 Then
        AccessCheck = True   'Has substance
    Else
        AccessCheck = True
        Call debugPrint("Unexpected error type: " & Err.Number & "  Desc:" & Err.Description, "cls_MainArena.AccessCheck")
    End If
    On Error GoTo 0
End Function

Public Sub Add(ByRef Content As Object)
    Dim newChunk As cls_Chunk
    Dim latestChunk As cls_Chunk

    Set newChunk = New cls_Chunk
    Set newChunk.Content = Content

    If Count_ > 0 And Count_ >= StoreMax Then
        Call CleanUp

        While Count_ >= StoreMax
            Call Unlink(MainArena.Fd)  'Unlink oldest chunk
        Wend
    End If

    Set latestChunk = MainArena.Bk
    Set latestChunk.Fd = newChunk
    Set newChunk.Bk = latestChunk

    Set newChunk.Fd = MainArena
    Set MainArena.Bk = newChunk

    Set CurrentChunk = MainArena

    Count_ = Count_ + 1
End Sub

Public Sub ClearAll()
    Call UnlinkAllFd(MainArena)

    If Count_ <> 0 Then
        Call debugPrint("Assertation Warning: Count is " & Count_ & " after ClearAll", "cls_MainArena.ClearAll")
    End If

    Count_ = 0
End Sub

Public Function Back() As Object
    Do While Not CurrentChunk.Bk Is MainArena
        If CurrentChunk.Bk.Content Is Nothing Then
            Call Unlink(CurrentChunk.Bk)
        ElseIf Not AccessCheck(CurrentChunk.Bk) Then
            Call Unlink(CurrentChunk.Bk)
        Else
            Set CurrentChunk = CurrentChunk.Bk
            Set Back = CurrentChunk.Content
            Exit Function
        End If
    Loop

    Set Back = Nothing
End Function

Public Property Get Current() As Object
    If CurrentChunk Is MainArena Then
        Set Current = Nothing
    Else
        Set Current = CurrentChunk.Content
    End If
End Property

Public Function Forward() As Object
    Do While Not CurrentChunk.Fd Is MainArena
        If CurrentChunk.Fd.Content Is Nothing Then
            Call Unlink(CurrentChunk.Fd)
        ElseIf Not AccessCheck(CurrentChunk.Fd) Then
            Call Unlink(CurrentChunk.Fd)
        Else
            Set CurrentChunk = CurrentChunk.Fd
            Set Forward = CurrentChunk.Content
            Exit Function
        End If
    Loop

    Set Forward = Nothing
End Function

Public Property Get Latest() As Object
    Dim chunk As cls_Chunk

    Set chunk = MainArena.Bk
    While Not chunk Is MainArena
        If Not AccessCheck(chunk) Then
            Set chunk = chunk.Bk
            Call Unlink(chunk.Fd)
        Else
            Set Latest = chunk.Content
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

Public Sub SetMax(ByVal Max As Long)
    If StoreMax = 0 Then
        StoreMax = Max
    Else
        Err.Raise 50000, Description:="äiî[ç≈ëÂêîÇÃïœçXÇÕÇ≈Ç´Ç‹ÇπÇÒÅB"
    End If
End Sub



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

    Call debugPrint(ret, "MainArena")
End Sub
