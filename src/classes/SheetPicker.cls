VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_SheetPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_IPicker

'/*
' * Implements the IPicker interface to create a custom sheet picker.
' * This class allows users to select, manage, and navigate through worksheets.
' */

'/*
' * Provides virtual inheritance for base picker functionalities.
' * In VBA, since direct inheritance is not supported, this acts as a base class proxy.
' */
Private Super As cls_PickerBase

'/* The name of the picker, displayed in the form title. */
Private Const PICKER_NAME As String = "SheetPicker"

'/* The width of the key column in the picker. Specify 0 to hide the key column. */
Private Const KEY_COLUMN_WIDTH As Double = 0

'/*
' * Flag to indicate if the picker is in preview mode.
' * When true, selecting a sheet in the listbox will activate it.
' */
Private mPreviewMode As Boolean

'/*
' * Defines custom actions specific to this sheet picker.
' */
Private Enum SheetAction
    TogglePreviewMode
    ToggleHidden
    ToggleVeryHidden
    RenameSheet
    DeleteSheet
    MoveSheetForward
    MoveSheetBackward
    PreviewSheet
End Enum

'/*
' * Initializes the SheetPicker object.
' * Sets up the base picker and declares custom actions for sheet management.
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
    ' * @param {SheetAction} registerAction - The specific action enum value to be registered.
    ' */
    Call Super.DeclareAction(Func, "P", "プレビューモードを切り替え", "Toggle preview mode", SheetAction.TogglePreviewMode)
    Call Super.DeclareAction(Func, "h", "表示/非表示の切り替え", "Toggle sheet visible", SheetAction.ToggleHidden)
    Call Super.DeclareAction(Func, "H", "表示/非表示の切り替え(Very hidden)", "Toggle sheet visible (very hidden)", SheetAction.ToggleVeryHidden)
    Call Super.DeclareAction(Func, "R", "シート名を変更", "Rename sheet name", SheetAction.RenameSheet)
    Call Super.DeclareAction(Func, "D", "シートを削除", "Delete sheet", SheetAction.DeleteSheet)
    Call Super.DeclareAction(Func, "X", "シートを削除", "Delete sheet", SheetAction.DeleteSheet) ' Alias for DeleteSheet
    Call Super.DeclareAction(Func, "J", "シートを次に移動", "Move sheet forward", SheetAction.MoveSheetForward)
    Call Super.DeclareAction(Func, "K", "シートを前に移動", "Move sheet backward", SheetAction.MoveSheetBackward)
    Call Super.DeclareAction(Func, "l", "シートを表示", "Preview sheet", SheetAction.PreviewSheet)
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
' * This is typically the active sheet's name.
' *
' * @returns {String} - The default selected item's key.
' */
Private Property Get cls_IPicker_Default() As String
    cls_IPicker_Default = ActiveSheet.Name
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' * The dictionary keys are sheet names and items are display names (including visibility status).
' *
' * @returns {Dictionary[String, String]} - A dictionary of picker items, where keys are sheet names and items are display strings.
' */
Public Property Get cls_IPicker_Item() As Dictionary
    Dim i As Long

    Set cls_IPicker_Item = New Dictionary
    For i = 1 To Sheets.Count
        Dim displayName As String
        displayName = Sheets(i).Name

        If Sheets(i).Visible = XlSheetVisibility.xlSheetHidden Then
            displayName = displayName & " (hidden)"
        ElseIf Sheets(i).Visible = XlSheetVisibility.xlSheetVeryHidden Then
            displayName = displayName & " (HIDDEN)"
        End If
        cls_IPicker_Item.Add Sheets(i).Name, displayName
    Next
End Property

'/*
' * Handles custom actions executed within the picker UI.
' *
' * @param {String} selectedKey - The key (sheet name) of the currently selected item.
' * @param {Long} ActionType - The type of action to handle (from SheetAction enum).
' * @returns {Boolean} - True if the action was handled by this class, False otherwise to pass to default implementation.
' */
Public Function cls_IPicker_HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    cls_IPicker_HandleAction = True
    Select Case ActionType
        Case SheetAction.TogglePreviewMode
            mPreviewMode = Not mPreviewMode
            Call Super.UpdateCaption(IIf(mPreviewMode, "Preview mode", ""))
            Call cls_IPicker_OnChange(selectedKey)
        Case SheetAction.ToggleHidden
            Call ToggleSheetVisible(selectedKey)
        Case SheetAction.ToggleVeryHidden
            Call ToggleSheetVisible(selectedKey, True)
        Case SheetAction.RenameSheet
            Call F_Sheets.RenameSheet(selectedKey)
            Call Super.RegenerateList
        Case SheetAction.DeleteSheet
            Call F_Sheets.DeleteSheet(selectedKey)
            Call Super.RegenerateList
        Case SheetAction.MoveSheetForward
            Call F_Sheets.MoveSheetForward(selectedKey)
            Call Super.RegenerateList
        Case SheetAction.MoveSheetBackward
            Call F_Sheets.MoveSheetBack(selectedKey)
            Call Super.RegenerateList
        Case SheetAction.PreviewSheet
            Call ShowSheet(selectedKey)

        Case Else
            ' If the executed action is not defined here, return False to pass it to the default implementation.
            cls_IPicker_HandleAction = False
    End Select
End Function

'/*
' * Called when the selection in the listbox changes.
' * If preview mode is active, it will show the selected sheet.
' *
' * @param {String} selectedKey - The key (sheet name) of the newly selected item.
' */
Public Sub cls_IPicker_OnChange(ByVal selectedKey As String)
    If mPreviewMode Then
        Call ShowSheet(selectedKey)
    End If
End Sub

'/*
' * Called when an item is confirmed (e.g., by pressing Enter).
' * Ensures the selected sheet is visible and then activates it. The form will automatically close after this method executes.
' *
' * @param {String} selectedKey - The key (sheet name) of the confirmed item.
' */
Public Sub cls_IPicker_OnEnter(ByVal selectedKey As String)
    On Error Resume Next ' In case the sheet is protected or non-existent
    Sheets(selectedKey).Visible = XlSheetVisibility.xlSheetVisible
    Call ShowSheet(selectedKey)
End Sub

'----- Define custom actions and helper functions below this line -----

'/*
' * Activates the sheet with the given name.
' *
' * @param {String} sheetName - The name of the sheet to activate.
' */
Private Sub ShowSheet(ByVal sheetName As String)
    On Error Resume Next ' In case the sheet is protected or non-existent
    Sheets(sheetName).Activate
End Sub

'/*
' * Toggles the visibility of a specified sheet.
' *
' * @param {String} sheetName - The name of the sheet whose visibility is to be toggled.
' * @param {Boolean} [isVeryHidden=False] - If true, toggles between visible and xlSheetVeryHidden; otherwise, between visible and xlSheetHidden.
' */
Private Sub ToggleSheetVisible(ByVal sheetName As String, Optional ByVal isVeryHidden As Boolean = False)
    On Error GoTo Catch

    Dim sheetVisibility As XlSheetVisibility

    sheetVisibility = IIf(isVeryHidden, xlVeryHidden, xlSheetHidden)

    Dim i As Long
    Dim cnt As Long
    With ActiveWorkbook.Sheets(sheetName)
        If .Visible <> sheetVisibility Then
            ' Check the number of visible sheets
            For i = 1 To ActiveWorkbook.Sheets.Count
                If ActiveWorkbook.Sheets(i).Visible = xlSheetVisible Then
                    cnt = cnt + 1
                    If cnt > 1 Then
                        Exit For
                    End If
                End If
            Next i

            ' Not all sheets can be hidden. Ensure at least one visible sheet remains.
            If cnt = 1 And .Visible = xlSheetVisible Then
                MsgBox gVim.Msg.HideAllSheets, vbExclamation
                Exit Sub
            End If

            .Visible = sheetVisibility
        Else
            .Visible = xlSheetVisible
        End If
    End With
    Call Super.RegenerateList
Catch:
    ' Error handling can be added here if needed.
End Sub
