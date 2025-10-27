VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_CellPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_IPicker

'/*
' * Implements the IPicker interface to create a custom cell picker.
' * This class allows users to select and navigate to cells within the active worksheet.
' */

'/*
' * Provides virtual inheritance for base picker functionalities.
' * In VBA, since direct inheritance is not supported, this acts as a base class proxy.
' */
Private Super As cls_PickerBase

'/* The name of the picker, displayed in the form title. */
Private Const PICKER_NAME As String = "CellPicker"

'/* The width of the key column in the picker. Specify 0 to hide the key column. */
Private Const KEY_COLUMN_WIDTH As Double = 72

'/* Maximum number of cells to process to avoid performance issues. */
Private Const MAX_CELLS_COUNT As Long = 1048576

'/*
' * Flag to indicate if the picker is in preview mode.
' * When true, selecting a cell in the listbox will activate it.
' */
Private mPreviewMode As Boolean

Private Type OriginalState
    WorkbookName As String
    sheetName As String
    SelectionAddress As String
    ActiveCellAddress As String
    ScrollColumn As Long
    ScrollRow As Long
End Type
Private mOriginalState As OriginalState
Private mIsConfirmed As Boolean ' Flag to indicate if the picker was confirmed

'/*
' * Defines custom actions specific to this cell picker.
' */
Private Enum CellAction
    TogglePreviewMode
    PreviewCell
End Enum

'/*
' * Initializes the CellPicker object.
' * Sets up the base picker and declares custom actions for cell navigation.
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
    ' * @param {CellAction} registerAction - The specific action enum value to be registered.
    ' */
    Call Super.DeclareAction(Func, "P", "プレビューモードを切り替え", "Toggle preview mode", CellAction.TogglePreviewMode)
    Call Super.DeclareAction(Func, "l", "セルを表示", "Preview the cell", CellAction.PreviewCell)
End Sub

'/*
' * Cleans up object references when the class is terminated.
' * Also restores the original selection if the picker was cancelled.
' */
Private Sub Class_Terminate()
    ' Only restore if the picker was cancelled (not confirmed) and original info is available
    If Not mIsConfirmed And mOriginalState.WorkbookName <> "" And mOriginalState.sheetName <> "" And mOriginalState.SelectionAddress <> "" Then
        On Error Resume Next ' Handle potential errors during restoration
        Dim Wb As Workbook
        Set Wb = Workbooks(mOriginalState.WorkbookName)
        Wb.Activate

        Dim ws As Worksheet
        Set ws = Wb.Worksheets(mOriginalState.sheetName)
        ws.Activate

        ' Restore selection
        ws.Range(mOriginalState.SelectionAddress).Select

        ' Restore ActiveCell
        If mOriginalState.ActiveCellAddress <> "" Then
            ws.Range(mOriginalState.ActiveCellAddress).Activate
        End If

        ' Restore scroll position
        With ActiveWindow
            .ScrollColumn = mOriginalState.ScrollColumn
            .ScrollRow = mOriginalState.ScrollRow
        End With
        On Error GoTo 0 ' Reset error handling
    End If

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

    ' Record the initial selection information
    If Not ActiveWorkbook Is Nothing Then
        mOriginalState.WorkbookName = ActiveWorkbook.Name
        If Not ActiveSheet Is Nothing Then
            mOriginalState.sheetName = ActiveSheet.Name

            ' Record Selection address
            If Not Selection Is Nothing Then
                mOriginalState.SelectionAddress = Selection.Address(External:=False)
            End If

            ' Record ActiveCell address
            If Not ActiveCell Is Nothing Then
                mOriginalState.ActiveCellAddress = ActiveCell.Address(External:=False)
            End If

            ' Record scroll position
            With ActiveWindow
                mOriginalState.ScrollColumn = .ScrollColumn
                mOriginalState.ScrollRow = .ScrollRow
            End With
        End If
    End If
    mIsConfirmed = False ' Initialize as not confirmed
End Sub

'/*
' * Returns the key of the default selected item when the picker is initialized.
' * This is typically the active cell's address.
' *
' * @returns {String} - The default selected item's key (cell address).
' */
Private Property Get cls_IPicker_Default() As String
    cls_IPicker_Default = ActiveCell.Address(False, False)
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' * The dictionary keys are cell addresses and items are their values.
' *
' * @returns {Dictionary[String, String]} - A dictionary of picker items, where keys are cell addresses (String) and items are cell values (String).
' * @throws {Error} - If the active sheet is not a Worksheet or if there are too many cells in UsedRange.
' */
Public Property Get cls_IPicker_Item() As Dictionary
    If Not TypeOf ActiveSheet Is Worksheet Then
        Err.Raise 50000, Description:="CellPicker must run on Worksheet"
        Exit Property
    End If

    If ActiveSheet.UsedRange.Count > MAX_CELLS_COUNT Then
        Err.Raise 50000, Description:="Too many cells to process in UsedRange. Please narrow down your selection."
        Exit Property
    End If

    Set cls_IPicker_Item = New Dictionary

    Dim rng As Range
    Dim ary As Variant
    ary = ActiveSheet.UsedRange

    If Not IsArray(ary) Then
        ' Handle single cell UsedRange case
        cls_IPicker_Item.Add ActiveSheet.UsedRange.Address(False, False), CStr(ary)
        Exit Property
    End If

    ' Loop through the 2D array to get cell addresses and values
    ' Adjust indices based on UsedRange's starting row and column
    Dim startRow As Long
    Dim startCol As Long
    Dim i As Long
    Dim j As Long
    Dim key As String

    With ActiveSheet.UsedRange.Item(1)
        startRow = .Row
        startCol = .Column
    End With

    For i = LBound(ary, 1) To UBound(ary, 1)
        For j = LBound(ary, 2) To UBound(ary, 2)
            If Not IsEmpty(ary(i, j)) Then
                key = GetColumnCode(j + startCol - 1) & CStr(i + startRow - 1)
                cls_IPicker_Item.Add key, CStr(ary(i, j))
            End If
        Next j
    Next i
End Property

'/*
' * Handles custom actions executed within the picker UI.
' *
' * @param {String} selectedKey - The key (cell address) of the currently selected item.
' * @param {Long} ActionType - The type of action to handle (from CellAction enum).
' * @returns {Boolean} - True if the action was handled by this class, False otherwise to pass to default implementation.
' */
Public Function cls_IPicker_HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    cls_IPicker_HandleAction = True
    Select Case ActionType
        Case CellAction.TogglePreviewMode
            mPreviewMode = Not mPreviewMode
            Call Super.UpdateCaption(IIf(mPreviewMode, "Preview mode", ""))
            Call cls_IPicker_OnChange(selectedKey)
        Case CellAction.PreviewCell
            Call ShowCell(selectedKey)

        Case Else
            ' If the executed action is not defined here, return False to pass it to the default implementation.
            cls_IPicker_HandleAction = False
    End Select
End Function

'/*
' * Called when the selection in the listbox changes.
' * If preview mode is active, it will show the selected cell.
' *
' * @param {String} selectedKey - The key (cell address) of the newly selected item.
' */
Public Sub cls_IPicker_OnChange(ByVal selectedKey As String)
    If mPreviewMode Then
        Call ShowCell(selectedKey)
    End If
End Sub

'/*
' * Called when an item is confirmed (e.g., by pressing Enter).
' * Activates the selected cell. The form will automatically close after this method executes.
' *
' * @param {String} selectedKey - The key (cell address) of the confirmed item.
' */
Public Sub cls_IPicker_OnEnter(ByVal selectedKey As String)
    ' Record the original position to the JumpList before activating the new cell.
    ' Use mOriginalState.SelectionAddress to account for possible ActiveCell changes in preview mode.
    If mOriginalState.SelectionAddress <> "" Then
        Dim originalRange As Range
        On Error Resume Next ' Handle potential errors if the original range is invalid
        Set originalRange = ActiveSheet.Range(mOriginalState.SelectionAddress)
        On Error GoTo 0 ' Reset error handling
        If Not originalRange Is Nothing Then
            Call RecordToJumpList(Target:=originalRange)
        End If
    End If

    Call ShowCell(selectedKey)
    mIsConfirmed = True ' Set flag to confirmed
End Sub

'/*
' * Converts a column index to its corresponding Excel column letter code.
' *
' * @param {Long} ColumnIndex - The 1-based column index (e.g., 1 for "A", 2 for "B").
' * @returns {String} - The Excel column letter code (e.g., "A", "AA").
' */
Private Function GetColumnCode(ByVal ColumnIndex As Long) As String
    Const ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim idx As Long
    Dim res As Long
    res = ColumnIndex
    While res > 0
        idx = (res - 1) Mod 26 + 1
        res = (res - 1) \ 26
        GetColumnCode = Mid(ALPHA, idx, 1) & GetColumnCode
    Wend
End Function

'----- Define custom actions and helper functions below this line -----

'/*
' * Activates the cell with the given address.
' *
' * @param {String} cellAddress - The address of the cell to activate (e.g., "A1").
' */
Private Sub ShowCell(ByVal cellAddress As String)
    On Error Resume Next ' In case the cell address is invalid or sheet is protected
    Range(cellAddress).Activate
End Sub
