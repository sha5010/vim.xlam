VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_GeneralPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_IPicker

'/*
' * Implements the IPicker interface to create a custom picker for general string selection.
' */

'/*
' * Provides virtual inheritance for base picker functionalities.
' * In VBA, since direct inheritance is not supported, this acts as a base class proxy.
' */
Private Super As cls_PickerBase

'/* The name of the picker, displayed in the form title. */
Private mCaption As String
'/* The width of the key column in the picker. Specify 0 to hide the key column. */
Private mKeyColumnWidth As Double
'/* The preferred width of the picker form. Specify 0 to use the default width based on KeyWidth. */
Private mPickerFormWidth As Double

'/* The items to be displayed in the picker. */
Private mItems As Dictionary
'/* The default selected item. */
Private mDefault As String
'/* The value selected by the user. */
Private mSelectedValue As String

'/*
' * Initializes the GeneralPicker object.
' * Sets up the base picker.
' */
Private Sub Class_Initialize()
    ' Instantiate the virtual base class
    Set Super = New cls_PickerBase
    Call Super.Init(Me)

    ' Initialize default values
    mCaption = "GeneralPicker" ' Default caption
    mKeyColumnWidth = 0 ' Default: hide key column
    mPickerFormWidth = 0 ' Default: auto width
    Set mItems = New Dictionary
    mDefault = ""
    mSelectedValue = "" ' Initialize selected value
End Sub

'/*
' * Cleans up object references when the class is terminated.
' */
Private Sub Class_Terminate()
    Set Super = Nothing
    Set mItems = Nothing
End Sub

'/*
' * Returns the caption for the picker form.
' *
' * @returns {String} - The caption of the picker.
' */
Public Property Get cls_IPicker_Caption() As String
    cls_IPicker_Caption = mCaption
End Property

'/*
' * Sets the caption for the picker form.
' */
Public Property Let Caption(ByVal newValue As String)
    mCaption = newValue
End Property

'/*
' * Returns a string of single-character keys that are bound to specific actions within this picker.
' *
' * @returns {String} - A string containing all bound single-character keys.
' */
Public Property Get cls_IPicker_BoundKeyChars() As String
    cls_IPicker_BoundKeyChars = Super.BoundKeyChars
End Property

'/*
' * Returns the preferred width of the key column in the picker.
' *
' * @returns {Double} - The width of the key column. Specify 0 to hide the key column.
' */
Public Property Get cls_IPicker_KeyWidth() As Double
    cls_IPicker_KeyWidth = mKeyColumnWidth
End Property

'/*
' * Sets the preferred width of the key column in the picker.
' */
Public Property Let KeyColumnWidth(ByVal newValue As Double)
    mKeyColumnWidth = newValue
End Property

'/*
' * Returns the preferred width of the picker form.
' * Specify 0 to use the default width based on KeyWidth.
' *
' * @returns {Double} - The preferred width of the picker form. Specify 0 to use the default width based on KeyWidth.
' */
Public Property Get cls_IPicker_PickerFormWidth() As Double
    cls_IPicker_PickerFormWidth = mPickerFormWidth
End Property

'/*
' * Sets the preferred width of the picker form.
' */
Public Property Let PickerFormWidth(ByVal newValue As Double)
    mPickerFormWidth = newValue
End Property

'/*
' * Attaches the picker to the specified UI elements.
' *
' * @param {MSForms.UserForm} objUserForm - The user form object to attach to.
' * @param {MSForms.ListBox} objListBox - The list box control within the user form.
' * @param {MSForms.TextBox} objTextBox - The text box control within the user form.
' * @param {MSForms.Label} objLabel - The label control for status messages.
' */
Public Sub cls_IPicker_Attach(ByRef objUserForm As MSForms.UserForm, ByRef objListBox As MSForms.ListBox, ByRef objTextBox As MSForms.TextBox, ByRef objLabel As MSForms.Label)
    Call Super.Attach(objUserForm, objListBox, objTextBox, objLabel)
End Sub

'/*
' * Returns the key of the default selected item when the picker is initialized.
' *
' * @returns {String} - The default selected item's key.
' */
Private Property Get cls_IPicker_Default() As String
    cls_IPicker_Default = mDefault
End Property

'/*
' * Sets the default selected item.
' */
Public Property Let Default(ByVal newValue As String)
    mDefault = newValue
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' * The dictionary keys are display values and items are associated data.
' *
' * @returns {Dictionary[String, String]} - A dictionary of picker items, where keys and items are both Strings.
' */
Public Property Get cls_IPicker_Item() As Dictionary
    Set cls_IPicker_Item = mItems
End Property

'/*
' * Sets the items for the picker.
' */
Public Property Set Items(ByVal newValue As Dictionary)
    Set mItems = newValue
End Property

'/*
' * Handles custom actions executed within the picker UI.
' * For GeneralPicker, no custom actions are handled.
' *
' * @param {String} selectedKey - The key of the currently selected item.
' * @param {Long} ActionType - The type of action to handle.
' * @returns {Boolean} - Always False, indicating no custom actions are handled by this class.
' */
Public Function cls_IPicker_HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    cls_IPicker_HandleAction = False
End Function

'/*
' * Called when the selection in the listbox changes.
' *
' * @param {String} selectedKey - The key of the newly selected item.
' */
Public Sub cls_IPicker_OnChange(ByVal selectedKey As String)
    ' No specific preview needed for GeneralPicker
End Sub

'/*
' * Called when an item is confirmed (e.g., by pressing Enter).
' * The form will automatically close after this method executes, so no UI considerations are needed.
' *
' * @param {String} selectedKey - The key of the confirmed item.
' */
Public Sub cls_IPicker_OnEnter(ByVal selectedKey As String)
    mSelectedValue = selectedKey
End Sub

'/*
' * Returns the value selected by the user in the picker.
' *
' * @returns {String} - The selected value, or an empty string if no item was selected or the picker was cancelled.
' */
Public Property Get SelectedValue() As String
    SelectedValue = mSelectedValue
End Property

'/*
' * Called when the picker UserForm is closed.
' * This is used to notify the base class to release references.
' */
Public Sub cls_IPicker_OnPickerClose()
    If Not Super Is Nothing Then
        Call Super.NotifyFormClosed
    End If
End Sub
