VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_BookPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_IPicker

'/*
' * Implements the IPicker interface to create a custom workbook picker.
' * This class allows users to select and navigate through open workbooks.
' */

'/*
' * Provides virtual inheritance for base picker functionalities.
' * In VBA, since direct inheritance is not supported, this acts as a base class proxy.
' */
Private Super As cls_PickerBase

'/* The name of the picker, displayed in the form title. */
Private Const PICKER_NAME As String = "BookPicker"

'/* The width of the key column in the picker. Specify 0 to hide the key column. */
Private Const KEY_COLUMN_WIDTH As Double = 0

'/* The preferred width of the picker form. Specify 0 to use the default width based on KeyWidth. */
Private Const PICKER_FORM_WIDTH As Double = 0

'/*
' * Flag to indicate if the picker is in preview mode.
' * When true, selecting a workbook in the listbox will activate it.
' */
Private mPreviewMode As Boolean

'/*
' * Defines custom actions specific to this workbook picker.
' */
Private Enum BookAction
    PreviewBook
    CloseBook
    SaveBook
    SaveBookAs
    ToggleReadOnlyMode
    CopyPath
    OpenPath
End Enum

'/*
' * Initializes the BookPicker object.
' * Sets up the base picker and declares custom actions for workbook management.
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
    ' * @param {BookAction} registerAction - The specific action enum value to be registered.
    ' */
    Call Super.DeclareAction(Func, "Q", "ブックを閉じる", "Close the workbook", BookAction.CloseBook)
    Call Super.DeclareAction(Func, "S", "ブックを保存", "Save the workbook", BookAction.SaveBook)
    Call Super.DeclareAction(Func, "W", "名前を付けて保存", "Save the workbook as", BookAction.SaveBookAs)
    Call Super.DeclareAction(Func, "R", "読み取り専用切替", "Toggle read-only mode", BookAction.ToggleReadOnlyMode)
    Call Super.DeclareAction(Func, "l", "ブックをプレビュー", "Preview the workbook", BookAction.PreviewBook)
    Call Super.DeclareAction(Func, "Y", "ブックのパスをコピー", "Copy workbook path", BookAction.CopyPath)
    Call Super.DeclareAction(Func, "O", "ブックのパスを開く", "Open workbook directory", BookAction.OpenPath)
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
' * This is typically the active workbook's name.
' *
' * @returns {String} - The default selected item's key.
' */
Private Property Get cls_IPicker_Default() As String
    cls_IPicker_Default = ActiveWorkbook.Name
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' * The dictionary keys are workbook names and items are display names (which are also workbook names).
' *
' * @returns {Dictionary[String, String]} - A dictionary of picker items, where keys and items are both String (workbook names).
' */
Public Property Get cls_IPicker_Item() As Dictionary
    Set cls_IPicker_Item = New Dictionary

    Dim b As Workbook
    For Each b In Workbooks
        cls_IPicker_Item.Add b.Name, b.Name
    Next
End Property

'/*
' * Handles custom actions executed within the picker UI.
' *
' * @param {String} selectedKey - The key (workbook name) of the currently selected item.
' * @param {Long} ActionType - The type of action to handle (from BookAction enum).
' * @returns {Boolean} - True if the action was handled by this class, False otherwise to pass to default implementation.
' */
Public Function cls_IPicker_HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    cls_IPicker_HandleAction = True
    Select Case ActionType
        Case BookAction.PreviewBook
            Call ShowBook(selectedKey)
        Case BookAction.CloseBook
            Call F_WBFunc.CloseAskSaving(selectedKey)
        Case BookAction.SaveBook
            Call F_WBFunc.SaveWorkbook(selectedKey)
        Case BookAction.SaveBookAs
            Call F_WBFunc.SaveAsNewWorkbook(selectedKey)
        Case BookAction.ToggleReadOnlyMode
            Call F_WBFunc.ToggleReadOnly(selectedKey)
        Case BookAction.CopyPath
            Call F_WBFunc.YankWorkbookPath(selectedKey)
        Case BookAction.OpenPath
            Call F_WBFunc.OpenWorkbookDir(selectedKey)

        Case Else
            ' If the executed action is not defined here, return False to pass to the default implementation.
            cls_IPicker_HandleAction = False
    End Select
End Function

'/*
' * Called when the selection in the listbox changes.
' * If preview mode is active, it will show the selected workbook.
' *
' * @param {String} selectedKey - The key (workbook name) of the newly selected item.
' */
Public Sub cls_IPicker_OnChange(ByVal selectedKey As String)
    If mPreviewMode Then
        Call ShowBook(selectedKey)
    End If
End Sub

'/*
' * Called when an item is confirmed (e.g., by pressing Enter).
' * Ensures the selected workbook is visible and then activates it. The form will automatically close after this method executes.
' *
' * @param {String} selectedKey - The key (workbook name) of the confirmed item.
' */
Public Sub cls_IPicker_OnEnter(ByVal selectedKey As String)
    On Error Resume Next ' In case the workbook is protected or non-existent
    Windows(selectedKey).Visible = True
    Windows(selectedKey).Activate
End Sub

'/*
' * Called when the picker UserForm is closed.
' * This is used to notify the base class to release references.
' */
Public Sub cls_IPicker_OnPickerClose()
    If Not Super Is Nothing Then
        Call Super.NotifyFormClosed
    End If
End Sub

'----- Define custom actions and helper functions below this line -----

'/*
' * Activates the workbook with the given name.
' *
' * @param {String} bookName - The name of the workbook to activate.
' */
Private Sub ShowBook(ByVal bookName As String)
    If Not mPreviewMode Then
        UF_Picker.Hide
    End If
    Windows(bookName).Activate
    Application.OnTime Now + 0.01 / 86400, "ActivateWorkbook"
    If Not mPreviewMode Then
        Unload UF_Picker
    End If
End Sub
