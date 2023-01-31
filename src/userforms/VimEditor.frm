VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_VimEditor 
   Caption         =   "Editor - VIM"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "VimEditor.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_VimEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// OPTIONS
Const VIMEDITOR_PADDING As Byte = 3
Const FONT_NAME As String = "PlemolJP Console NF"
Const FONT_SIZE As Single = 10

'// Colors
Const COLOR_BG As Long = &H282828
Const COLOR_BG1 As Long = &H36383C
Const COLOR_RED As Long = &H1D24CC
Const COLOR_GREEN As Long = &H1A9798
Const COLOR_YELLOW As Long = &H2199D7
Const COLOR_BLUE As Long = &H888545
Const COLOR_PURPLE As Long = &H8662B1
Const COLOR_AQUA As Long = &H6A9D68
Const COLOR_GRAY As Long = &H8489A8
Const COLOR_ORANGE As Long = &HE5DD6
Const COLOR_FG As Long = &HB2DBEB

'// Maximum history count
Const TEXT_BUFFER_HISTORY As Byte = 50

Private VimEditorMode As String
Private CommandBuffer As String
Private VimEditorCount As Long

Private TextBuffers(TEXT_BUFFER_HISTORY - 1) As String
Private TextBufferMax As Byte
Private TextBufferCur As Byte
Private TextBufferRotate As Boolean
Private TextBufferLock As Boolean   '// do not change the history if true

Private savedPosX As Long

Private IsLastIMEModeOn As Boolean

Private Sub Text_Command_Change()
    Label_Command.Caption = Text_Command.Text
End Sub

Private Sub UserForm_Initialize()
    Me.BackColor = COLOR_BG
    VimEditorCount = 1

    With TextArea
        .ForeColor = COLOR_FG
        .BackStyle = fmBackStyleTransparent
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .Multiline = True
        .EnterKeyBehavior = True
        .SelectionMargin = False
        .IMEMode = fmIMEModeDisable
    End With

    With Text_Command
        .ForeColor = COLOR_FG
        .BackStyle = fmBackStyleOpaque
        .BackColor = COLOR_BG
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .SelectionMargin = False
        .Visible = False
    End With

    With Label_Command
        .ForeColor = COLOR_FG
        .BackStyle = fmBackStyleOpaque
        .BackColor = COLOR_BG
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
    End With

    With Label_Mode
        .ForeColor = COLOR_BG
        .TextAlign = fmTextAlignCenter
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .Caption = "12345678"
        .Font.Bold = True
        .WordWrap = False
        .AutoSize = True
        .AutoSize = False
    End With

    With Label_Status
        .BackStyle = fmBackStyleOpaque
        .BackColor = COLOR_BG1
        .ForeColor = COLOR_FG
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .TextAlign = fmTextAlignRight
        .WordWrap = False
        .Height = Label_Mode.Height
    End With

    Call ChangeMode("NORMAL")
    Call Resize(300, 400)

    Call TextArea_Change    '// delete when Run method will be implemented
    Call VimEditorKeyInit
End Sub

Private Sub TextArea_Change()
    Dim buf As String

    '// do not change the history if locked
    If TextBufferLock Then
        Exit Sub
    End If

    '// evaluate except INSERT mode
    If VimEditorMode <> "INSERT" Then
        '// crlf -> lf
        buf = Replace(TextArea.Text, vbCr, "")

        '// check if they are completely same
        If TextBuffers(TextBufferCur) <> buf Then
            TextBufferCur = (TextBufferCur + 1) Mod TEXT_BUFFER_HISTORY
            TextBufferMax = TextBufferCur

            '// make circulable
            If TextBufferCur = 0 Then
                TextBufferRotate = True
            End If

            '// append
            TextBuffers(TextBufferCur) = buf
        End If
    End If
End Sub

Private Sub TextArea_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    '// Do not allow to lose focus expect in COMMAND mode
    If VimEditorMode <> "COMMAND" Then
        TextArea.SetFocus
    End If
End Sub

Private Sub TextArea_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '// KeyCode is the same as Common.vKey
    Dim key As String
    Dim code As Integer

    '// Ignore keys (will be typed)
    Select Case KeyCode
        'Ignore Ctrl, Shift, Alt
        Case 16 To 18
            Exit Sub

        'Ignore PgUp, PgDown, End, Home, Left, Up, Down, Right
        Case 33 To 40
            Exit Sub

        'Ignore Delete
        Case 46
            Exit Sub
    End Select

    code = code Or (Sgn(Shift And 2) * Ctrl_)
    code = code Or (Sgn(Shift And 4) * Alt_)

    'Not INSERT mode and pressed special a key
    If KeyCode < 48 Then
        key = VimEditorMode & "_" & CStr(KeyCode)
        If gVimEditorKeymap.Exists(key) Then
            Application.Run gVimEditorKeymap(key)
        End If

        If VimEditorMode <> "INSERT" Then
            CommandBuffer = ""
            KeyCode = 0         '// prevent default
        End If

    'Ctrl or Alt key is pressed
    ElseIf code > 0 Then
        code = code Or (Sgn(Shift And 1) * Shift_)

        key = VimEditorMode & "_" & CStr(code Or KeyCode)
        If gVimEditorKeymap.Exists(key) Then
            Application.Run gVimEditorKeymap(key)
        End If

        If VimEditorMode <> "INSERT" Then
            CommandBuffer = ""
            KeyCode = 0         '// prevent default
        End If

    'Only Shift is pressed
    ElseIf (Shift And 1) > 0 Then
        code = code Or Shift_

        key = VimEditorMode & "_" & CStr(code Or KeyCode)
        If gVimEditorKeymap.Exists(key) Then
            Application.Run gVimEditorKeymap(key)
        End If
    End If
End Sub

Private Function KeyToDictKey(ByVal keys As String) As String
    Dim i As Integer
    Dim u As Integer

    u = Len(keys)

    For i = 1 To u
        KeyToDictKey = KeyToDictKey & "_" & CStr(Asc(Mid(keys, i, 1)))
    Next i
End Function

Private Sub TextArea_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim key As String
    Dim char As String

    key = VimEditorMode & KeyToDictKey(CommandBuffer) & "_" & KeyAscii

    If 32 <= KeyAscii And KeyAscii < 128 Then
        char = Chr(KeyAscii)
    End If

    '// Prevent defaults except in INSERT mode
    If VimEditorMode <> "INSERT" Then
        KeyAscii = 0
    End If

    If gVimEditorKeymap.Exists(key) Then
        Application.Run gVimEditorKeymap(key)
        CommandBuffer = ""
    ElseIf VimEditorMode <> "INSERT" Then
        If Len(CommandBuffer) > 3 Then
            CommandBuffer = ""
        Else
            CommandBuffer = CommandBuffer & char
        End If
    End If
End Sub

Private Sub TextArea_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If VimEditorMode = "NORMAL" Then
        If KeyCode = Left_ And TextArea.SelStart > 0 Then
            TextArea.SelStart = TextArea.SelStart - 1
        End If
        TextArea.SelLength = -1
    End If

    Call Redraw
End Sub

Private Sub Redraw()
    Dim position As String
    Dim persent As String

    With TextArea
        position = Format(PosY, "#0") & ":" & Format(PosX, "0")
        If Len(TextBuffers(TextBufferCur)) = 0 Then
            persent = Format(0, "#0%")
        Else
            persent = Format(.SelStart / Len(TextBuffers(TextBufferCur)), "#0%")
        End If
    End With

    Label_Status.Caption = CommandBuffer & " | " & persent & " | " & position & " |"
    DoEvents
End Sub

Public Sub Resize(ByVal Height As Double, ByVal Width As Double)
    Const TITLE_WIDTH As Double = 30

    With Me
        .Height = Height
        .Width = Width

        .Label_Padding.Height = Height

        .TextArea.Left = VIMEDITOR_PADDING
        .TextArea.Top = 6
        .TextArea.Height = Height - VIMEDITOR_PADDING - .Label_Status.Height * 2 - TITLE_WIDTH
        .TextArea.Width = Width - VIMEDITOR_PADDING * 2 - 6

        .Label_Mode.Top = .TextArea.Top + .TextArea.Height
        .Label_Mode.Left = VIMEDITOR_PADDING

        .Label_Status.Top = .Label_Mode.Top
        .Label_Status.Left = .Label_Mode.Left + .Label_Mode.Width
        .Label_Status.Width = .TextArea.Width - .Label_Mode.Width

        .Label_Command.Top = .Label_Mode.Top + .Label_Mode.Height
        .Label_Command.Left = .TextArea.Left
        .Label_Command.Width = .TextArea.Width
        .Label_Command.Height = .Label_Mode.Height

        .Text_Command.Top = .Label_Command.Top - 1.5
        .Text_Command.Left = .Label_Command.Left - 1.5
        .Text_Command.Width = .Label_Command.Width + 1.5
        .Text_Command.Height = .Label_Command.Height + 1.5
    End With
End Sub

Public Sub ChangeMode(ByVal Mode As String, Optional ByVal CommandPrefix As String = "")
    Me.TextArea.Locked = (Mode <> "INSERT")
    Me.Text_Command.Visible = (Mode = "COMMAND")

    With Me.Label_Mode
        Select Case Mode
            Case "INSERT"
                .Caption = Mode
                .BackColor = COLOR_BLUE

                If IsLastIMEModeOn Then
                    Me.TextArea.IMEMode = fmIMEModeOn
                Else
                    Me.TextArea.IMEMode = fmIMEModeOff
                End If
            Case "VISUAL", "V-LINE"
                .Caption = Mode
                .BackColor = COLOR_ORANGE
            Case "REPLACE"
                .Caption = Mode
                .BackColor = COLOR_RED
            Case "COMMAND"
                .Caption = Mode
                .BackColor = COLOR_AQUA
                Me.Text_Command.Text = CommandPrefix
            Case Else
                '// when return from INSERT
                If .Caption = "INSERT" Then
                    VimEditorMode = "NORMAL"
                    Call TextArea_Change
                End If

                .Caption = "NORMAL"
                .BackColor = COLOR_GRAY
                IsLastIMEModeOn = (Me.TextArea.IMEMode <> fmIMEModeOff)
                Me.TextArea.IMEMode = fmIMEModeDisable

                If Mode <> "NORMAL" Then
                    Call debugPrint("Unsupported mode: " & Mode, "VimEditor|ChangeVimEditorMode")
                End If
        End Select

        VimEditorMode = .Caption
    End With
End Sub

Public Property Get Buffer() As String
    Buffer = TextBuffers(TextBufferCur)
End Property

Public Sub ClearCommandBuffer()
    CommandBuffer = ""
End Sub

Private Property Get HeadIndex() As Long
    If Me.TextArea.CurLine = 0 Then
        HeadIndex = 1
    Else
        HeadIndex = InStrRev(Buffer, vbLf, Me.TextArea.SelStart) + 1
    End If
End Property

Public Property Get PosX() As Long
    Dim head As Long

    With Me.TextArea
        If .SelStart = 0 Then
            PosX = 1
        ElseIf .CurLine = 0 Then
            PosX = LenB(StrConv(Mid(Buffer, 1, .SelStart), vbFromUnicode)) + 1
        Else
            head = HeadIndex
            PosX = LenB(StrConv(Mid(Buffer, head, .SelStart - head + 1), vbFromUnicode)) + 1
        End If
    End With
End Property

Public Property Get PosY() As Long
    PosY = StrCount(Left(Buffer, Me.TextArea.SelStart), vbLf) + 1
End Property

Public Property Get MaxY() As Long
    MaxY = StrCount(Buffer, vbLf) + 1
End Property

Public Sub SetPos(Optional BaseY As Long = 0, Optional BaseX As Long = 0, _
                  Optional TargetY As Long = 0, Optional TargetX As Long = 0, _
                  Optional MoveLR As Long = 0)

    Dim head As Long
    Dim tail As Long
    Dim changeLR As Boolean

    changeLR = (BaseX > 0 Or MoveLR <> 0)

    With Me.TextArea
        If BaseY > 0 And BaseY <> Me.PosY Then
            If Me.MaxY < BaseY Then
                .CurLine = .LineCount - 1
            Else
                .SelStart = StrNPos(Buffer, vbLf, BaseY - 1)
            End If

            If BaseX = 0 Then
                BaseX = savedPosX
            End If
        End If

        If (BaseX > 0 And BaseX <> Me.PosX) Or MoveLR <> 0 Then
            head = HeadIndex
            tail = InStr(HeadIndex, Buffer, vbLf)
            '// Last line
            If tail = 0 Then
                tail = Len(Buffer)
            End If

            If head = tail Then
                Exit Sub
            End If

            If MoveLR <> 0 Then
                head = head - 1
                tail = tail - 2 - (Me.PosY = Me.MaxY)
                MoveLR = .SelStart + MoveLR

                If MoveLR < head Then
                    MoveLR = head
                ElseIf MoveLR > tail Then
                    MoveLR = tail
                End If

                .SelStart = MoveLR
            Else
                If tail < head Then
                    tail = head
                Else
                    tail = LenB(StrConv(Mid(Buffer, head, tail - head), vbFromUnicode))
                End If

                If BaseX > tail Then
                    BaseX = Len(StrConv(LeftB(StrConv(Mid(Buffer, head), vbFromUnicode), tail), vbUnicode)) - (Me.PosY = Me.MaxY)
                Else
                    BaseX = Len(StrConv(LeftB(StrConv(Mid(Buffer, head), vbFromUnicode), BaseX), vbUnicode))
                End If

                .SelStart = head + BaseX - 2
            End If

            If changeLR Then
                savedPosX = Me.PosX
            End If
        End If

        .SelLength = 1
        .SetFocus
    End With
    DoEvents
End Sub

Public Sub UpdateSavedPosX()
    savedPosX = PosX
End Sub

Public Property Get gCount() As Long
    gCount = VimEditorCount
End Property

Public Function VimEditor_Undo() As Boolean
    Dim curPosX As Long
    Dim curPosY As Long

    '// check if it can be undone
    If TextBufferCur = 0 And Not TextBufferRotate Then
        Call VimEditor_SetStatus("Already at oldest change")
        Exit Function
    ElseIf TextBufferRotate And (TEXT_BUFFER_HISTORY + TextBufferCur - 1) Mod TEXT_BUFFER_HISTORY = TextBufferMax Then
        Call VimEditor_SetStatus("Already at oldest change")
        Exit Function
    End If

    '// lock buffers
    TextBufferLock = True

    '// undo
    curPosX = Me.PosX
    curPosY = Me.PosY
    TextBufferCur = (TEXT_BUFFER_HISTORY + TextBufferCur - 1) Mod TEXT_BUFFER_HISTORY
    TextArea.Text = TextBuffers(TextBufferCur)
    Call Me.SetPos(curPosY, curPosX)

    '// unlock buffers
    TextBufferLock = False

    VimEditor_Undo = True
End Function

Public Function VimEditor_Redo() As Boolean
    Dim curPosX As Long
    Dim curPosY As Long

    '// check if it can be redone
    If TextBufferCur = TextBufferMax Then
        Call VimEditor_SetStatus("Already at newest change", True)
        Exit Function
    ElseIf Not TextBufferRotate And TextBufferCur > TextBufferMax Then
        Call debugPrint("Unexpected situation: " & TextBufferCur & " > " & TextBufferMax, "VimEditor_Redo")
        Exit Function
    End If

    '// lock buffers
    TextBufferLock = True

    '// redo
    curPosX = Me.PosX
    curPosY = Me.PosY
    TextBufferCur = (TextBufferCur + 1) Mod TEXT_BUFFER_HISTORY
    TextArea.Text = TextBuffers(TextBufferCur)
    Call Me.SetPos(curPosY, curPosX)

    '// unlock buffers
    TextBufferLock = False

    VimEditor_Redo = True
End Function

'// update status bar
Public Sub VimEditor_SetStatus(msg As String, Optional Error As Boolean = False)
    With Me.Label_Command
        .Caption = msg

        If Error Then
            .BackColor = COLOR_RED
            .ForeColor = COLOR_FG
        Else
            .BackColor = COLOR_BG
            .ForeColor = COLOR_FG
        End If
    End With
End Sub
