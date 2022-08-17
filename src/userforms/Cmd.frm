VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Cmd 
   Caption         =   "Command"
   ClientHeight    =   408
   ClientLeft      =   42
   ClientTop       =   330
   ClientWidth     =   1596
   OleObjectBlob   =   "Cmd.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LONG_MAX As Long = 2147483647

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Top = 0
        .Left = 0
    End With

    With Me.Label_Text
        .WordWrap = False
        .AutoSize = True
    End With
End Sub

Private Sub UserForm_Activate()
    Me.Move Application.Left + 4, Application.Top + Application.Height - Me.Height - 4
    Me.Label_Text = gCmdBuf
End Sub


Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim cmd As String

    If KeyAscii = 27 Then  'Escape
        gCmdBuf = ""
        gCount = 1
        Me.Hide
    ElseIf KeyAscii = 13 Then  'Enter
        cmd = enterCmd()
        gCmdBuf = ""
        Me.Hide
        If cmd <> "" Then
            Call runCmd(cmd, returnOnly:=True)
        End If
        gCount = 1
    ElseIf KeyAscii = 8 Then  'Backspace
        gCmdBuf = Left(gCmdBuf, Len(gCmdBuf) - 1)
        If gCmdBuf = "" Then
            Me.Hide
        End If
    Else
        gCmdBuf = gCmdBuf + ChrW(KeyAscii)
        cmd = checkCmd()
        If cmd <> "" Then
            Call runCmd(cmd)
        End If
        gCount = 1
    End If

    On Error Resume Next
    Me.Label_Text = gCmdBuf
    Me.Width = Me.Label_Text.Left + Me.Label_Text.Width + 12
End Sub

Private Function getCmd(Optional ByVal countFirstOnly As Boolean = False) As String
    Dim i As Integer
    Dim char As Integer
    Dim cmd As String
    Dim numFlag As Boolean

    If countFirstOnly Then
        For i = 1 To Len(gCmdBuf)
            char = Asc(Mid(gCmdBuf, i, 1))
            If 47 < char And char < 58 Then  '0-9
                If numFlag Then
                    If gCount < LONG_MAX / 10 - 1 Then
                        gCount = gCount * 10 + (char - 48)
                    Else
                        gCount = LONG_MAX
                    End If
                Else
                    gCount = (char - 48)
                End If
                numFlag = True
            Else
                numFlag = False
                Exit For
            End If
        Next i
    Else
        i = 1
    End If

    For i = i To Len(gCmdBuf)
        char = Asc(Mid(gCmdBuf, i, 1))

        If (Not countFirstOnly) And 47 < char And char < 58 Then  '0-9
            If numFlag Then
                If gCount < LONG_MAX / 10 - 1 Then
                    gCount = gCount * 10 + (char - 48)
                Else
                    gCount = LONG_MAX
                End If
            Else
                gCount = (char - 48)
            End If
            numFlag = True
        Else
            numFlag = False
            cmd = cmd & Chr(char)
        End If
    Next i

    getCmd = cmd
End Function

Private Function checkCmd() As String
    Dim cmd As String
    Dim buf As String
    Dim i As Integer

    cmd = getCmd(countFirstOnly:=True)
    If gKeyMap(KEY_NORMAL).Exists(cmd) Then
        If gKeyMap(KEY_NORMAL)(cmd) <> "showCmdForm" Then
            checkCmd = cmd
            Exit Function
        End If
    End If

    cmd = getCmd()
    If gKeyMap(KEY_NORMAL).Exists(cmd) Then
        If gKeyMap(KEY_NORMAL)(cmd) <> "showCmdForm" Then
            checkCmd = cmd
            Exit Function
        End If
    End If

    For i = Len(gCmdBuf) To 1 Step -1
        buf = Mid(gCmdBuf, 1, i)
        If gKeyMap(KEY_NORMAL_ARG).Exists(buf) Then
            checkCmd = buf & " " & Mid(gCmdBuf, i + 1)
            Exit Function
        End If
    Next i
End Function

Private Function enterCmd() As String
    Dim cmd As String
    Dim buf As String
    Dim i As Integer

    cmd = getCmd(countFirstOnly:=True)
    If gKeyMap(KEY_RETURNONLY).Exists(cmd) Then
        enterCmd = cmd
        Exit Function
    End If

    cmd = getCmd()
    If gKeyMap(KEY_RETURNONLY).Exists(cmd) Then
        enterCmd = cmd
        Exit Function
    End If

    For i = Len(gCmdBuf) To 1 Step -1
        buf = Mid(gCmdBuf, 1, i)
        If gKeyMap(KEY_RETURNONLY_ARG).Exists(buf) Then
            enterCmd = buf & " " & Mid(gCmdBuf, i + 1)
            Exit Function
        End If
    Next i
End Function

Private Function runCmd(ByVal cmd As String, Optional ByVal returnOnly As Boolean = False)
    Dim hasArgs As Boolean
    Dim buf As Variant
    Dim ret As Variant

    hasArgs = InStr(cmd, " ") > 0
    buf = Split(cmd, " ", 2)

    If returnOnly Then
        Me.Hide
        If hasArgs Then
            ret = Application.Run(gKeyMap(KEY_RETURNONLY_ARG)(buf(0)), Trim(buf(1)))
        Else
            ret = Application.Run(gKeyMap(KEY_RETURNONLY)(buf(0)))
        End If
    Else
        If hasArgs Then
            ret = Application.Run(gKeyMap(KEY_NORMAL_ARG)(buf(0)), Trim(buf(1)))

            'コマンドが完了していない場合はフォームを閉じない
            If TypeName(ret) = "Boolean" Then
                If ret = True Then
                    Me.Hide
                Else
                    Exit Function
                End If
            End If
        Else
            Me.Hide
            ret = Application.Run(gKeyMap(KEY_NORMAL)(buf(0)))
        End If

    End If

    gCmdBuf = ""
    gCount = 1
End Function
