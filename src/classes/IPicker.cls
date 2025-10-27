VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_IPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*
' * Returns the caption of the picker.
' *
' * @returns {String} The caption of the picker.
' */
Public Property Get Caption() As String
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' *
' * @returns {Dictionary[String, String]} A dictionary where keys are display values and items are associated data.
' */
Public Property Get Item() As Dictionary
End Property

'/*
' * Returns the default selected item.
' *
' * @returns {String} The default selected item's key.
' */
Public Property Get Default() As String
End Property

'/*
' * Return characters of keys that are bound to specific actions.
' *
' * @returns {String} A string of bound key characters.
' */
Public Property Get BoundKeyChars() As String
End Property

'/*
' * Returns the width of the key column in the picker.
' * Specify 0 if you do not want to display the key column.
' *
' * @returns {Double} The width of the key column.
' */
Public Property Get KeyWidth() As Double
End Property

'/*
' * Returns the preferred width of the picker form.
' * Specify 0 if you want to use the default width based on KeyWidth.
' *
' * @returns {Double} The preferred width of the picker form.
' */
Public Property Get PickerFormWidth() As Double
End Property

'/*
' * Handles the event when an item is entered (selected).
' *
' * @param {String} selectedKey The key of the selected item.
' */
Public Sub OnEnter(ByVal selectedKey As String)
End Sub

'/*
' * Handles the event when the selection changes.
' *
' * @param {String} selectedKey The key of the newly selected item.
' */
Public Sub OnChange(ByVal selectedKey As String)
End Sub

'/*
' * Attaches the picker to UI elements.
' *
' * @param {MSForms.UserForm} objUserForm The user form object.
' * @param {MSForms.ListBox} objListBox The list box object.
' * @param {MSForms.TextBox} objTextBox The text box object.
' * @param {MSForms.Label} objLabel The label object.
' */
Public Sub Attach(ByRef objUserForm As MSForms.UserForm, ByRef objListBox As MSForms.ListBox, ByRef objTextBox As MSForms.TextBox, ByRef objLabel As MSForms.Label)
End Sub

'/*
' * Handles specific actions triggered by the picker.
' *
' * @param {String} selectedKey The key of the selected item.
' * @param {Long} ActionType The type of action to handle.
' * @returns {Boolean} True if the action was handled, False otherwise.
' */
Public Function HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
End Function

'/*
' * Handles the event when the picker UserForm is closed.
' * Implement cleanup logic here to release references and prevent memory leaks.
' */
Public Sub OnPickerClose()
End Sub
