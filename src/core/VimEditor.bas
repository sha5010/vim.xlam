Attribute VB_Name = "C_VimEditor"
Option Explicit
Option Private Module

Public gVimEditorKeymap As Dictionary

Sub VimEditorKeyInit()
    If gVimEditorKeymap Is Nothing Then
        Set gVimEditorKeymap = New Dictionary
    End If

    Call EditorMap("n", "<ESC>", "NORMAL_ClearBuffer")
    Call EditorMap("n", "<C-[>", "NORMAL_ClearBuffer")
    Call EditorMap("n", "h", "keystroke", True, Left_)
    Call EditorMap("n", "j", "keystroke", True, Down_)
    Call EditorMap("n", "k", "keystroke", True, Up_)
    Call EditorMap("n", "l", "keystroke", True, Right_)
    Call EditorMap("n", "gg", "NORMAL_JumpTop")
    Call EditorMap("n", "G", "NORMAL_JumpButtom")
    Call EditorMap("n", "0", "NORMAL_GoToFirst")
    Call EditorMap("n", "^", "NORMAL_GoToNonBlankFirst")
    Call EditorMap("n", "$", "NORMAL_GoToLast")
    Call EditorMap("n", "<C-h>", "NORMAL_GoToLeftEdge")
    Call EditorMap("n", "<C-l>", "NORMAL_GoToRightEdge")
    Call EditorMap("n", "a", "NORMAL_EnterInsertMode", True)
    Call EditorMap("n", "A", "NORMAL_AppendFromLast")
    Call EditorMap("n", "i", "NORMAL_EnterinsertMode")
    Call EditorMap("n", "I", "NORMAL_InsertFromFirst")
    Call EditorMap("n", "n", "Nop")

    Call EditorMap("n", "ZQ", "NORMAL_Quit")

    Call EditorMap("i", "<ESC>", "INSERT_Leave")
    Call EditorMap("i", "<C-[>", "INSERT_Leave")
End Sub

Private Sub EditorMap(ByVal Mode As String, _
                      ByVal key As String, _
                      ByVal funcName As String, _
                      ParamArray args() As Variant)

    key = Join(ParseKeys(key), "_")
    funcName = "'" & funcName
    funcName = funcName & ParseArgs(args) & "'"

    Select Case LCase(Left(Mode, 1))
        Case "i": Mode = "INSERT"
        Case "v": Mode = "VISUAL"
        Case "c": Mode = "COMMAND"
        Case Else: Mode = "NORMAL"
    End Select

    key = Mode & "_" & key
    If gVimEditorKeymap.Exists(key) Then
        gVimEditorKeymap(key) = funcName
    Else
        gVimEditorKeymap.Add key, funcName
    End If
End Sub

Private Function ParseArgs(ByVal args As Variant) As String
    Dim i As Integer
    Dim u As Integer

    u = UBound(args)
    For i = 0 To u
        Select Case TypeName(args(i))
            Case "String"
                ParseArgs = ParseArgs & " """ & args(i) & ""","
            Case "Integer", "Long", "LongLong", "Double", "Single", "Byte"
                ParseArgs = ParseArgs & " " & args(i) & ","
            Case "Boolean"
                ParseArgs = ParseArgs & " " & CStr(args(i)) & ","
            Case Else
                Call debugPrint("Unsupport argument type: " & TypeName(args(i)), "parseArg")
        End Select
    Next i

    If Len(ParseArgs) > 0 Then
        ParseArgs = Left(ParseArgs, Len(ParseArgs) - 1)
    End If
End Function

Private Function ParseKeys(ByVal key As String) As Variant
    Dim i As Integer
    Dim u As Integer
    Dim j As Integer
    Dim char As String
    Dim buf As Integer
    Dim components As Variant
    Dim keys As String

    For i = 1 To Len(key)
        char = Mid(key, i, 1)
        If char <> "<" Then
            keys = keys & CStr(Asc(char))
        Else
            u = InStr(i, key, ">")
            If u > i Then
                components = Split(Mid(key, i + 1, u - i - 1), "-")
                i = u
                u = UBound(components)

                If u > 0 Then
                    For j = 0 To u - 1
                        Select Case LCase(components(j))
                            Case "c"
                                buf = buf Or Ctrl_
                            Case "s"
                                buf = buf Or Shift_
                            Case "a", "m"
                                buf = buf Or Alt_
                        End Select
                    Next j
                    buf = buf + Asc(UCase(components(u)))
                Else
                    Select Case LCase(components(0))
                        Case "bs"
                            buf = buf + BackSpace_
                        Case "tab"
                            buf = buf + Tab_
                        Case "cr", "return", "enter"
                            buf = buf + Enter_
                        Case "esc"
                            buf = buf + Escape_
                        Case "space"
                            buf = buf + Space_
                        Case "del"
                            buf = buf + Delete_
                        Case "up"
                            buf = buf + Up_
                        Case "down"
                            buf = buf + Down_
                        Case "left"
                            buf = buf + Left_
                        Case "right"
                            buf = buf + Right_
                        Case "home"
                            buf = buf + Home_
                        Case "end"
                            buf = buf + End_
                        Case "pageup"
                            buf = buf + PageUp_
                        Case "pagedown"
                            buf = buf + PageDown_
                    End Select
                End If
                keys = keys & buf
            Else
                keys = keys & CStr(Asc(buf))
            End If
        End If

        keys = keys & ","
    Next i

    keys = Left(keys, Len(keys) - 1)
    ParseKeys = Split(keys, ",")
End Function

Function NORMAL_EnterInsertMode(Optional IsAppend As Boolean = False)
    With UF_VimEditor.TextArea
        .SelLength = 0
        .SelStart = .SelStart + -CInt(IsAppend)
        Call UF_VimEditor.ChangeMode("INSERT")
    End With
End Function

Function NORMAL_InsertFromFirst()
    Call NORMAL_GoToNonBlankFirst
    Call NORMAL_EnterInsertMode
End Function

Function NORMAL_AppendFromLast()
    Call NORMAL_GoToLast
    Call NORMAL_EnterInsertMode(True)
End Function

Function NORMAL_ClearBuffer()
    Call UF_VimEditor.ClearCommandBuffer
End Function

Function NORMAL_JumpTop()
    Call UF_VimEditor.SetPos(BaseY:=UF_VimEditor.gCount)
End Function

Function NORMAL_JumpButtom()
    With UF_VimEditor
        If .gCount = 1 Then
            Call .SetPos(BaseY:=.MaxY)
        Else
            Call .SetPos(BaseY:=.gCount)
        End If
    End With
End Function

Function NORMAL_GoToFirst()
    Call UF_VimEditor.SetPos(BaseX:=1)
End Function

Function NORMAL_GoToNonBlankFirst()
    With UF_VimEditor
        Call .SetPos(BaseX:=1)
        Select Case Mid(UF_VimEditor.Buffer, .TextArea.SelStart + 1, 1)
            Case " ", vbTab
                Call keystroke(True, Ctrl_ + Right_)
        End Select
    End With
End Function

Function NORMAL_GoToLast()
    Call UF_VimEditor.SetPos(BaseX:=2147483647)
End Function

Function NORMAL_GoToLeftEdge()
    Call keystroke(True, Home_)
End Function

Function NORMAL_GoToRightEdge()
    Call keystroke(True, End_)
End Function

Function NORMAL_Quit()
    Unload UF_VimEditor
End Function

Function NORMAL_Del1Char()
    Call keystroke(True, Delete_)
End Function

Function INSERT_Leave()
    With UF_VimEditor.TextArea
        If .SelStart > 0 Then
            .SelStart = .SelStart - 1
        End If
        Call UF_VimEditor.ChangeMode("NORMAL")
    End With
End Function

Function Nop()
    Call UF_VimEditor.SetPos(54, 30)
End Function
