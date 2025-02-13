VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_CmdLine 
   Caption         =   "Cmdline"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1950
   OleObjectBlob   =   "CmdLine.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_CmdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cUserForm As cls_UFKeyReceiver
Attribute cUserForm.VB_VarHelpID = -1
Private cReturn As String

Private Sub cUserForm_KeyPressWithSendKeys(ByVal key As String)
    If key Like "*{ENTER}" Then
        Me.Hide
        cReturn = Me.TextBox.Text
        Exit Sub
    End If

    Dim cmd As String
    cmd = gVim.KeyMap.Get_(key)
    If cmd <> "" And Not cmd Like "'" & SHOW_CMD_PROCEDURE & " ""*" Then
        On Error Resume Next
        Application.Run cmd
    ElseIf key = "{ESC}" Or key = "^{[}" Or key = "^{c}" Then
        Me.Hide
        cReturn = CMDLINE_CANCELED
    End If
End Sub

Private Sub TextBox_Change()
    If Not Me.Visible Then
        Exit Sub
    End If

    With Me.TextBox
        Dim newWidth As Double
        newWidth = .Left + .Width + 12

        Dim minWidth As Double: minWidth = 100
        Dim maxWidth As Double: maxWidth = Application.Width - 12

        If newWidth < minWidth Then
            newWidth = minWidth
        ElseIf newWidth > maxWidth Then
            newWidth = maxWidth
        End If

        Me.Width = newWidth

        .Text = Replace(.Text, vbTab, " ")
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Cmdline"
        .StartUpPosition = 0
    End With

    With Me.TextBox
        .Value = ""
        .Multiline = False
        .EnterKeyBehavior = True
        .AutoSize = True
        .WordWrap = False
        .TabKeyBehavior = False
    End With

    Set cUserForm = New cls_UFKeyReceiver
    Set cUserForm.TextBox = Me.TextBox
End Sub

Private Sub UserForm_Activate()
    Me.Move Application.Left + 4, Application.Top + Application.Height - Me.Height - 4

    With Me.Label_Prefix
        .AutoSize = False
        .WordWrap = False
        .Width = 500
        .AutoSize = True
    End With

    With Me.TextBox
        .SetFocus
        .Left = Me.Label_Prefix.Left + Me.Label_Prefix.Width - 7.5
    End With
    Call TextBox_Change
End Sub

Public Function Launch(Optional ByVal prefix As String = ":", _
                       Optional ByVal formCaption As String = "Cmdline", _
                       Optional ByVal enableIME As Boolean = False) As String

    ' Set the label text with the prefix key
    Me.Label_Prefix.Caption = prefix

    With Me.TextBox
        .Text = ""
        If enableIME Then
            .IMEMode = fmIMEModeHiragana
        Else
            .IMEMode = fmIMEModeOff
        End If
    End With

    Me.Caption = formCaption
    cReturn = ""

    Dim currentMode As String: currentMode = gVim.Mode.Current

    Call gVim.Mode.Change(MODE_CMDLINE)
    Me.Show
    Call gVim.Mode.Change(currentMode)

    Launch = cReturn
End Function
