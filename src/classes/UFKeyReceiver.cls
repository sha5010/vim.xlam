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
Private WithEvents cListBox As MSForms.ListBox
Attribute cListBox.VB_VarHelpID = -1

' Event declarations
Public Event KeyPressWithSendKeys(ByVal key As String, src As Object)
Public Event KeyPressWithString(ByVal str As String, src As Object)

' Flag to disable keydown
Public IsIgnoreKeydown As Boolean


'/*
' * Handles the KeyDown event of the user form.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' */
Private Sub cUserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OnKeyDown(KeyCode, Shift, cUserForm)
    If IsIgnoreKeydown Then
        KeyCode = 0
    End If
End Sub

'/*
' * Handles the KeyDown event of the textbox.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' */
Private Sub cTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OnKeyDown(KeyCode, Shift, cTextBox)
    If IsIgnoreKeydown Then
        KeyCode = 0
    End If
End Sub

'/*
' * Handles the KeyDown event of the listbox.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' */
Private Sub cListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call OnKeyDown(KeyCode, Shift, cListBox)
    If IsIgnoreKeydown Then
        KeyCode = 0
    End If
End Sub

'/*
' * Handles the KeyDown event.
' *
' * @param {MSForms.ReturnInteger} KeyCode - The key code of the pressed key.
' * @param {Integer} Shift - The shift state during the key press.
' * @param {Object} EventSource - The source object of the KeyDown event
' */
Private Sub OnKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer, ByRef EventSource As Object)
    IsIgnoreKeydown = False

    ' Convert key code and shift state to SendKeys style
    Dim sendKeysStyle As String
    sendKeysStyle = gVim.KeyMap.KeycodeToSendKeys(KeyCode, Shift)

    ' If a valid SendKeys style is obtained, raise the events
    If sendKeysStyle <> "" Then
        RaiseEvent KeyPressWithString(gVim.KeyMap.SendKeysToDisplayText(sendKeysStyle), EventSource)
        RaiseEvent KeyPressWithSendKeys(sendKeysStyle, EventSource)
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

'/*
' * Sets the textbox for which key press events are handled.
' *
' * @param {MSForms.TextBox} objTextBox - The textbox to set.
' */
Public Property Set TextBox(ByRef objTextBox As MSForms.TextBox)
    Set cTextBox = objTextBox
End Property

'/*
' * Gets the textbox for which key press events are handled.
' *
' * @returns {MSForms.TextBox} - The textbox.
' */
Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = cTextBox
End Property

'/*
' * Sets the listbox for which key press events are handled.
' *
' * @param {MSForms.ListBox} objListBox - The listbox to set.
' */
Public Property Set ListBox(ByRef objListBox As MSForms.ListBox)
    Set cListBox = objListBox
End Property

'/*
' * Gets the listbox for which key press events are handled.
' *
' * @returns {MSForms.ListBox} - The listbox.
' */
Public Property Get ListBox() As MSForms.ListBox
    Set ListBox = cListBox
End Property
