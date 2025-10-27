VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_PickerTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_IPicker

'/*
' * Implements the IPicker interface to create a custom picker.
' * This class serves as a template for new picker implementations.
' */

'/*
' * Provides virtual inheritance for base picker functionalities.
' * In VBA, since direct inheritance is not supported, this acts as a base class proxy.
' */
Private Super As cls_PickerBase

'/* The name of the picker, displayed in the form title. */
Private Const PICKER_NAME As String = "TemplatePicker"

'/* The width of the key column in the picker. Specify 0 to hide the key column. */
Private Const KEY_COLUMN_WIDTH As Double = 0

'/* The preferred width of the picker form. Specify 0 to use the default width based on KeyWidth. */
Private Const PICKER_FORM_WIDTH As Double = 0

'/*
' * Defines custom actions specific to this picker template.
' */
Private Enum TemplateAction
    '/* A sample action for demonstration purposes. */
    SampleAction
End Enum

'/*
' * Initializes the PickerTemplate object.
' * Sets up the base picker and declares custom actions.
' */
Private Sub Class_Initialize()
    ' Instantiate the virtual base class
    Set Super = New cls_PickerBase
    Call Super.Init(Me)

    ' Add custom functionalities here
    '/*
    ' * Declares a new custom action for the picker.
    ' * @param {eActionType} ActionType - The type of action (e.g., Func for custom function).
    ' * @param {String} lhs - The left-hand side key binding (Vim-style, e.g., "s" for 's' key).
    ' * @param {String} helpTextJa - Japanese help text displayed for this action.
    ' * @param {String} helpTextEn - English help text displayed for this action.
    ' * @param {TemplateAction} registerAction - The specific action enum value to be registered.
    ' */
    Call Super.DeclareAction(Func, "s", "日本語モードで表示されるヘルプ文字列", "Help text in English mode", TemplateAction.SampleAction)
End Sub

'/*
' * Cleans up object references when the class is terminated.
' */
Private Sub Class_Terminate()
    Set Super = Nothing
End Sub

'/*
' * Returns the caption for the picker form.
' *
' * @returns {String} - The caption of the picker.
' */
Public Property Get cls_IPicker_Caption() As String
    cls_IPicker_Caption = PICKER_NAME
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
    cls_IPicker_KeyWidth = KEY_COLUMN_WIDTH
End Property

'/*
' * Returns the preferred width of the picker form.
' * Specify 0 to use the default width based on KeyWidth.
' *
' * @returns {Double} - The preferred width of the picker form. Specify 0 to use the default width based on KeyWidth.
' */
Public Property Get cls_IPicker_PickerFormWidth() As Double
    cls_IPicker_PickerFormWidth = PICKER_FORM_WIDTH
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
    cls_IPicker_Default = "Sample1"
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' * The dictionary keys are display values and items are associated data.
' *
' * @returns {Dictionary[String, String]} - A dictionary of picker items, where keys and items are both Strings.
' */
Public Property Get cls_IPicker_Item() As Dictionary
    Set cls_IPicker_Item = New Dictionary
    cls_IPicker_Item.Add "Sample1", "Sample1"
    cls_IPicker_Item.Add "Sample2", "Sample2"
    cls_IPicker_Item.Add "Sample3", "Sample3"
End Property

'/*
' * Handles custom actions executed within the picker UI.
' *
' * @param {String} selectedKey - The key of the currently selected item.
' * @param {Long} ActionType - The type of action to handle (from TemplateAction enum).
' * @returns {Boolean} - True if the action was handled by this class, False otherwise to pass to default implementation.
' */
Public Function cls_IPicker_HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    cls_IPicker_HandleAction = True

    Select Case ActionType
        Case TemplateAction.SampleAction
            MsgBox "SampleAction was executed.", vbInformation

        Case Else
            ' If the executed action is not defined here, return False to pass it to the default implementation.
            cls_IPicker_HandleAction = False

    End Select
End Function

'/*
' * Called when the selection in the listbox changes.
' * Implement preview functionalities or other side effects here.
' *
' * @param {String} selectedKey - The key of the newly selected item.
' */
Public Sub cls_IPicker_OnChange(ByVal selectedKey As String)
    ' Implement preview features or other logic here.
End Sub

'/*
' * Called when an item is confirmed (e.g., by pressing Enter).
' * The form will automatically close after this method executes, so no UI considerations are needed.
' *
' * @param {String} selectedKey - The key of the confirmed item.
' */
Public Sub cls_IPicker_OnEnter(ByVal selectedKey As String)
    ' Implement final selection logic here.
End Sub

'----- Define custom actions and helper functions below this line -----
