VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ColorPicker 
   Caption         =   "ColorPicker"
   ClientHeight    =   3126
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   3432
   OleObjectBlob   =   "ColorPicker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LUMINANCES As String = "0;80;60;40;-25;-50"
Private Const LUMINANCES_WHITE As String = "0;-5;-15;-25;-35;-50"
Private Const LUMINANCES_BLACK As String = "0;50;35;25;15;5"
Private Const LUMINANCES_GRAY As String = "0;-10;-25;-50;-75;-90"

Private Const BOX_SIZE As Byte = 12
Private Const BOX_GAP As Byte = 3
Private Const KEY_LIST1 As String = "asdfghjkl;"
Private Const KEY_LIST2 As String = "qwertyuiop"
Private Const KEY_DETAIL As String = "12345"
Private Const KEY_NULL As String = "n"
Private Const TEXT_PREFIX As String = " "
Private Const IDX_LIST1 As Integer = 19
Private Const IDX_LIST2 As Integer = 3
Private Const IDX_LIST_TOP As Integer = 4
Private Const PLACEHOLDER As String = "* Type "" "" or # to use color code"

Private colorTable As Collection
Private colorObject As Collection
Private labelTable As Collection
Private resultLabel As MSForms.Label
Private textLabel As MSForms.Label
Private cmdBuf As String
Private isSuccess As Boolean
Private X As Integer
Private Y As Integer

Private Function GetLabelTitle(ByVal X As Integer, ByVal Y As Integer) As String
    GetLabelTitle = "Label_" & Format(X, "00") & Format(Y, "00")
End Function

Private Function PutLabel(ByVal X As Integer, ByVal Y As Integer, _
                     Optional ByVal AssosiateColor As cls_FontColor, _
                     Optional ByVal BackColor As Long, _
                     Optional ByVal Caption As String = "", _
                     Optional ByVal Visible As Boolean = True) As String

    Dim genLabel As MSForms.Label

    PutLabel = GetLabelTitle(X, Y)
    Set genLabel = Me.Controls.Add("Forms.Label.1", PutLabel, True)

    With genLabel
        .Left = (BOX_SIZE / 3) + X * (BOX_SIZE + BOX_GAP)
        .Top = 3 + (BOX_SIZE / 2) * Y
        .Height = BOX_SIZE
        .Width = BOX_SIZE
        .BorderStyle = 0
        .Enabled = Visible

        If Caption = "" Then
            If Not AssosiateColor Is Nothing Then
                BackColor = AssosiateColor.Color
            End If

            .Tag = BackColor
            '.BackColor = BackColor
            .ForeColor = BackColor
            .Caption = ChrW(&H2588)
            .Font.Size = BOX_SIZE * 2
            '.TextAlign = fmTextAlignCenter
        Else
            .Caption = Caption
            .TextAlign = fmTextAlignCenter
            .Font.Name = "Consolas"
            .BackStyle = fmBackStyleTransparent
            .Font.Size = BOX_SIZE / 4 * 3
        End If
    End With

    If Caption = "" Then
        colorTable.Add genLabel, PutLabel
        colorObject.Add AssosiateColor, PutLabel
    Else
        labelTable.Add genLabel, PutLabel
    End If
End Function

Private Sub ChangeAll(ByVal Enabled As Boolean, _
                      Optional ByVal Expect As String)
    Dim Label As MSForms.Label

    For Each Label In colorTable
        If Label.Name <> resultLabel.Name Then
            Label.Enabled = Enabled Xor Label.Name Like Expect
        End If
    Next Label

    For Each Label In labelTable
        If Label.Name <> textLabel.Name Then
            Label.Enabled = Enabled Xor Label.Name Like Expect
        End If
    Next Label
End Sub

Private Sub ChangeSpecific(ByVal Enabled As Boolean, _
                           Optional ByVal Specific As String)

    Dim Label As MSForms.Label

    For Each Label In colorTable
        If Label.Name <> resultLabel.Name And Label.Name Like Specific Then
            Label.Enabled = Enabled
        End If
    Next Label

    For Each Label In labelTable
        If Label.Name <> textLabel.Name And Label.Name Like Specific Then
            Label.Enabled = Enabled
        End If
    Next Label
End Sub

Private Function HexColorCodeToLong(ByVal colorCode As String) As Long
    If Len(colorCode) = 3 Then
        HexColorCodeToLong = Val("&H" & Mid(colorCode, 3, 1) & Mid(colorCode, 3, 1) & _
            Mid(colorCode, 2, 1) & Mid(colorCode, 2, 1) & Mid(colorCode, 1, 1) & Mid(colorCode, 1, 1) & "&")
    ElseIf Len(colorCode) = 6 Then
        HexColorCodeToLong = Val("&H" & Mid(colorCode, 5, 2) & Mid(colorCode, 3, 2) & Mid(colorCode, 1, 2) & "&")
    Else
        HexColorCodeToLong = -1
    End If
End Function


Private Sub checkCmd()
    Dim colorCode As String
    Dim colorValue As Long

    If InStr(cmdBuf, "#") = 1 Then
        If cmdBuf = "#" Then
            Call ChangeAll(False)
        End If

        colorCode = Mid(cmdBuf, 2)
        colorValue = HexColorCodeToLong(colorCode)

        If colorValue < 0 Then
            resultLabel.ForeColor = Me.BackColor
        Else
            resultLabel.ForeColor = colorValue
        End If

        Exit Sub
    End If

    If Len(cmdBuf) = 0 Then
        Call ChangeAll(True)
        Call ChangeSpecific(False, "Label_" & Format(0, "00") & "*")
        Call ChangeSpecific(False, "Label_" & Format(11, "00") & "*")
        resultLabel.ForeColor = Me.BackColor

    ElseIf Len(cmdBuf) = 1 Then
        If cmdBuf = KEY_NULL Then
            Call ChangeAll(False, "Label_n")
            Exit Sub
        End If

        X = InStr(KEY_LIST1, cmdBuf)
        If X > 0 Then
            Call ChangeAll(False)
            With Me.Controls(GetLabelTitle(X, IDX_LIST1))
                .Enabled = True
                resultLabel.ForeColor = .Tag
            End With
            Me.Controls(GetLabelTitle(X, IDX_LIST1 - 2)).Enabled = True

            Y = IDX_LIST1
            Exit Sub
        End If

        X = InStr(KEY_LIST2, cmdBuf)
        If X > 0 Then
            Call ChangeAll(False, "Label_" & Format(X, "00") & "*")
            Call ChangeSpecific(True, "Label_" & Format(0, "00") & "*")
            Call ChangeSpecific(True, "Label_" & Format(11, "00") & "*")
            Me.Controls(GetLabelTitle(X, IDX_LIST1)).Enabled = False
            Me.Controls(GetLabelTitle(X, IDX_LIST1 - 2)).Enabled = False

            resultLabel.ForeColor = Me.Controls(GetLabelTitle(X, IDX_LIST2)).Tag

            Y = IDX_LIST2
            Exit Sub
        End If
    ElseIf Len(cmdBuf) = 2 Then
        X = InStr(KEY_LIST2, Left(cmdBuf, 1))
        Y = InStr(KEY_DETAIL, Right(cmdBuf, 1)) * 2 + IDX_LIST_TOP
        If X > 0 And Y > 0 Then
            Call ChangeAll(False, "Label_" & Format(X, "00") & Format(Y, "00"))
            resultLabel.ForeColor = Me.Controls(GetLabelTitle(X, Y)).Tag

            Exit Sub
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    X = 0
    Y = 0
    isSuccess = False
    cmdBuf = ""
    Call checkCmd
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer, j As Integer
    Dim l As Variant, lw As Variant, lb As Variant, lg As Variant
    Dim Color As cls_FontColor
    Dim defaultColor As Variant
    Dim cnt As Integer

    'コレクションを初期化
    Set colorTable = New Collection
    Set labelTable = New Collection
    Set colorObject = New Collection

    cmdBuf = ""

    cnt = 1
    defaultColor = Array(192, 255, 49407, 65535, 5296274, 5287936, 15773696, 12611584, 6299648, 10498160)

    l = Split(LUMINANCES, ";")
    lw = Split(LUMINANCES_WHITE, ";")
    lb = Split(LUMINANCES_BLACK, ";")
    lg = Split(LUMINANCES_GRAY, ";")

    For j = 1 To 5
        Call PutLabel(0, IDX_LIST_TOP + j * 2, Caption:=Mid(KEY_DETAIL, j, 1), Visible:=False)
        Call PutLabel(11, IDX_LIST_TOP + j * 2, Caption:=Mid(KEY_DETAIL, j, 1), Visible:=False)
    Next j

    For i = 0 To 9
        Set Color = New cls_FontColor
        Call Color.Setup(msoThemeColorIndex:=i + 1)

        Call PutLabel(i + 1, 1, Caption:=Mid(KEY_LIST2, i + 1, 1))
        Call PutLabel(i + 1, IDX_LIST2, Color)

        For j = 1 To 5
            Set Color = New cls_FontColor
            Call Color.Setup(msoThemeColorIndex:=i + 1)

            Select Case i + 1
                Case 1
                    Color.Luminance = CInt(lw(j))
                Case 2
                    Color.Luminance = CInt(lb(j))
                Case 3
                    Color.Luminance = CInt(lg(j))
                Case Else
                    Color.Luminance = CInt(l(j))
            End Select

            Call PutLabel(i + 1, IDX_LIST_TOP + j * 2, Color)
        Next j

        Call PutLabel(i + 1, 17, Caption:=Mid(KEY_LIST1, i + 1, 1))

        Set Color = New cls_FontColor
        Call Color.Setup(colorCode:=defaultColor(i))

        Call PutLabel(i + 1, IDX_LIST1, Color)
    Next i

    With Me.Controls(PutLabel(1, 22, Caption:="n: Automatic or Null"))
        .Width = BOX_SIZE * 10 + BOX_GAP * 9
        .Name = "Label_n"
        .BorderStyle = 1
        .BorderColor = &HA0A0A0
    End With

    Set resultLabel = Me.Controls(PutLabel(0, 25, BackColor:=Me.BackColor))
    Set textLabel = Me.Controls(PutLabel(1, 25, Caption:=TEXT_PREFIX))
    textLabel.Width = BOX_SIZE * 11 + BOX_GAP * 10
    textLabel.TextAlign = fmTextAlignLeft
    textLabel.Caption = PLACEHOLDER

    Me.Width = BOX_SIZE * 13 + BOX_GAP * 13
    Me.Height = 36 + BOX_SIZE * 13 + BOX_SIZE / 50 * BOX_SIZE

    Set Color = Nothing
End Sub


Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim addChar As String

    If KeyAscii = 27 Then 'Escape
        textLabel.Caption = PLACEHOLDER
        Me.Hide

    ElseIf KeyAscii = 8 Then 'BackSpace
        If cmdBuf <> "" Then
            cmdBuf = Left(cmdBuf, Len(cmdBuf) - 1)

            If cmdBuf = "" Then
                textLabel.Caption = PLACEHOLDER
            Else
                textLabel.Caption = TEXT_PREFIX & cmdBuf
            End If
        End If

        Call checkCmd

    ElseIf KeyAscii = 13 Then  'Enter
        If resultLabel.ForeColor <> Me.BackColor Or cmdBuf = KEY_NULL Then
            isSuccess = True
            textLabel.Caption = PLACEHOLDER
            Me.Hide
        End If

    ElseIf KeyAscii > 31 Then
        addChar = LCase(Chr(KeyAscii))

        If Len(cmdBuf) = 0 Then
            If addChar = " " Or addChar = "#" Or InStr(KEY_LIST1 & KEY_LIST2 & KEY_NULL, addChar) > 0 Then
                If addChar = " " Then
                    cmdBuf = cmdBuf & "#"
                Else
                    cmdBuf = cmdBuf & addChar
                End If
                textLabel.Caption = TEXT_PREFIX & cmdBuf

                Call checkCmd
            End If
        ElseIf Left(cmdBuf, 1) = "#" And InStr("0123456789abcdefABCDEF", addChar) > 0 And Len(cmdBuf) < 7 Then
            cmdBuf = cmdBuf & addChar
            textLabel.Caption = TEXT_PREFIX & cmdBuf

            Call checkCmd

        ElseIf Len(cmdBuf) = 1 And InStr(KEY_DETAIL, addChar) > 0 Then
            cmdBuf = cmdBuf & addChar
            textLabel.Caption = TEXT_PREFIX & cmdBuf

            Call checkCmd

        End If
    End If
End Sub

Public Function ShowColorPicker() As cls_FontColor
    Dim colorCode As String
    Dim colorValue As Long

    UF_Cmd.Hide
    Me.Show
    If isSuccess Then
        If cmdBuf = KEY_NULL Then
            Set ShowColorPicker = New cls_FontColor
            Call ShowColorPicker.Setup(colorCode:=0)  'dummy
            ShowColorPicker.IsNull = True
        ElseIf InStr(cmdBuf, "#") = 1 Then
            colorCode = Mid(cmdBuf, 2)
            colorValue = HexColorCodeToLong(colorCode)

            Set ShowColorPicker = New cls_FontColor
            Call ShowColorPicker.Setup(colorCode:=colorValue)
        Else
            Set ShowColorPicker = colorObject(GetLabelTitle(X, Y))
        End If
    Else
        Set ShowColorPicker = Nothing
    End If
End Function
