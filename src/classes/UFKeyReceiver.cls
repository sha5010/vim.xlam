VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_UFKeyReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*
' * Represents a user form with key press events.
' */
Private WithEvents cUserForm As MSForms.UserForm
Attribute cUserForm.VB_VarHelpID = -1
Private WithEvents cTextBox As MSForms.TextBox
Attribute cTextBox.VB_VarHelpID = -1

' Event declarations
Public Event KeyPressWithSendKeys(ByVal key As String)
Public Event KeyPressWithString(ByVal str As String)

'/*
' * Handles the KeyDown event of the user form.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' */
Private Sub cUserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OnKeyDown(KeyCode, Shift)
End Sub

'/*
' * Handles the KeyDown event of the textbox.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' */
Private Sub cTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OnKeyDown(KeyCode, Shift)
End Sub

'/*
' * Handles the KeyDown event.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' */
Private Sub OnKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Convert key code and shift state to SendKeys style
    Dim sendKeysStyle As String
    sendKeysStyle = gVim.KeyMap.KeycodeToSendKeys(KeyCode, Shift)

    ' If a valid SendKeys style is obtained, raise the events
    If sendKeysStyle <> "" Then
        RaiseEvent KeyPressWithString(gVim.KeyMap.SendKeysToDisplayText(sendKeysStyle))
        RaiseEvent KeyPressWithSendKeys(sendKeysStyle)
    End If
End Sub

'/*
' * Sets the user form for which key press events are handled.
' *
' * @param {MSForms.UserForm} objUserForm - The user form to set.
' */
Public Property Set Form(ByRef objUserForm As MSForms.UserForm)
    Set cUserForm = objUserForm
End Property

'/*
' * Gets the user form for which key press events are handled.
' *
' * @returns {MSForms.UserForm} - The user form.
' */
Public Property Get Form() As MSForms.UserForm
    Set Form = cUserForm
End Property

Public Property Set TextBox(ByRef objTextBox As MSForms.TextBox)
    Set cTextBox = objTextBox
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = cTextBox
End Property
