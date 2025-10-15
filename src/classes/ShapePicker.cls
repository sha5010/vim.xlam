VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ShapePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cls_IPicker

'/*
' * Implements the IPicker interface to create a custom picker for shapes.
' * This class allows users to select and manage shapes within the active workbook.
' */

'/*
' * Provides virtual inheritance for base picker functionalities.
' * In VBA, since direct inheritance is not supported, this acts as a base class proxy.
' */
Private Super As cls_PickerBase

'/* The name of the picker, displayed in the form title. */
Private Const PICKER_NAME As String = "ShapePicker"

'/* The width of the key column in the picker. Specify 0 to hide the key column. */
Private Const KEY_COLUMN_WIDTH As Double = 100

'/* Flag to indicate if the picker is in preview mode. */
Private mPreviewMode As Boolean

'/*
' * Defines custom actions specific to this shape picker.
' */
Private Enum ShapeAction
    TogglePreviewMode
    PreviewShape
End Enum

'/*
' * Initializes the ShapePicker object.
' * Sets up the base picker and declares custom actions.
' */
Private Sub Class_Initialize()
    ' Instantiate the virtual base class
    Set Super = New cls_PickerBase
    Call Super.Init(Me)

    ' Add custom functionalities here
    Call Super.DeclareAction(Func, "P", "プレビューモードを切り替える", "Toggle preview mode", ShapeAction.TogglePreviewMode)
    Call Super.DeclareAction(Func, "l", "シェイプをアクティブにする", "Activate shape", ShapeAction.PreviewShape)
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
' *
' * @returns {String} - The default selected item's key.
' */
Private Property Get cls_IPicker_Default() As String
    cls_IPicker_Default = "" ' No default selected item for shapes initially
End Property

'/*
' * Returns a dictionary of items to be displayed in the picker.
' * The dictionary keys are shape names and items are display names (including value/alt text).
' *
' * @returns {Dictionary[String, String]} - A dictionary of picker items, where keys are shape names and items are display strings.
' */
Public Property Get cls_IPicker_Item() As Dictionary
    Set cls_IPicker_Item = New Dictionary

    Dim ws As Worksheet
    If ActiveWorkbook Is Nothing Then Exit Property

    For Each ws In ActiveWorkbook.Sheets
        Call AddShapesRecursively(ws.Shapes, ws.Name, cls_IPicker_Item)
    Next ws
End Property

'/*
' * Handles custom actions executed within the picker UI.
' *
' * @param {String} selectedKey - The key (shape name) of the currently selected item.
' * @param {Long} ActionType - The type of action to handle (from ShapeAction enum).
' * @returns {Boolean} - True if the action was handled by this class, False otherwise to pass to default implementation.
' */
Public Function cls_IPicker_HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    cls_IPicker_HandleAction = True
    Select Case ActionType
        Case ShapeAction.TogglePreviewMode
            mPreviewMode = Not mPreviewMode
            Call Super.UpdateCaption(IIf(mPreviewMode, "Preview mode", ""))
            Call cls_IPicker_OnChange(selectedKey)
        Case ShapeAction.PreviewShape
            Call ShowShape(selectedKey)
        Case Else
            ' If the executed action is not defined here, return False to pass it to the default implementation.
            cls_IPicker_HandleAction = False
    End Select
End Function

'/*
' * Called when the selection in the listbox changes.
' * If preview mode is active, it will show the selected shape.
' *
' * @param {String} selectedKey - The key (shape name) of the newly selected item.
' */
Public Sub cls_IPicker_OnChange(ByVal selectedKey As String)
    If mPreviewMode Then
        Call ShowShape(selectedKey)
    End If
End Sub

'/*
' * Called when an item is confirmed (e.g., by pressing Enter).
' * Activates the selected shape. The form will automatically close after this method executes.
' *
' * @param {String} selectedKey - The key (shape name) of the confirmed item.
' */
Public Sub cls_IPicker_OnEnter(ByVal selectedKey As String)
    Call ShowShape(selectedKey)
End Sub

'----- Define custom actions and helper functions below this line -----

'/*
' * Selects the shape with the given name.
' *
' * @param {String} shapeName - The name of the shape to select.
' */
Private Sub ShowShape(ByVal shapeName As String)
    On Error Resume Next
    Dim parts() As String
    Dim sheetName As String
    Dim actualShapeName As String

    On Error Resume Next
    parts = Split(shapeName, "!")
    If UBound(parts) = 1 Then
        sheetName = parts(0)
        actualShapeName = parts(1)

        If ActiveWorkbook Is Nothing Then Exit Sub

        With ActiveWorkbook
            .Sheets(sheetName).Activate
            .Sheets(sheetName).Shapes(actualShapeName).Select Replace:=True
        End With
    End If
    On Error GoTo 0
End Sub

'/*
' * Retrieves the text content of a shape if it has text.
' *
' * @param {Shape} Sh - The shape object.
' * @returns {String} - The text content of the shape.
' */
Private Function GetShapeValue(ByVal Sh As Shape) As String
    On Error Resume Next
    If Sh.HasTextFrame Then
        If Sh.TextFrame.HasText Then
            GetShapeValue = Sh.TextFrame.Characters.Text
        End If
    End If
    On Error GoTo 0
End Function

'/*
' * Retrieves the alternative text (description) of a shape.
' *
' * @param {Shape} Sh - The shape object.
' * @returns {String} - The alternative text of the shape.
' */
Private Function GetShapeDescription(ByVal Sh As Shape) As String
    On Error Resume Next
    GetShapeDescription = Sh.AlternativeText
    On Error GoTo 0
End Function

'/*
' * Recursively adds shapes from a ShapeRange to the picker's item dictionary.
' *
' * @param {Object} shapesCollection - The collection of shapes to process (can be Shapes or GroupItems).
' * @param {String} currentSheetName - The name of the sheet the shapes belong to.
' */
Private Sub AddShapesRecursively(ByVal shapesCollection As Object, ByVal currentSheetName As String, ByRef outDict As Dictionary)
    Dim Sh As Shape
    Dim shapeKey As String
    Dim shapeValue As String

    For Each Sh In shapesCollection
        If Sh.Type = msoGroup Then
            ' Recursively process shapes within the group
            Call AddShapesRecursively(Sh.GroupItems, currentSheetName, outDict)
        Else
            ' Add individual shape to the dictionary
            shapeKey = currentSheetName & "!" & Sh.Name
            shapeValue = GetShapeValue(Sh)

            If Len(shapeValue) = 0 Then
                shapeValue = GetShapeDescription(Sh)
            End If

            If Len(shapeKey) > 0 Then
                outDict.Add shapeKey, shapeValue
            End If
        End If
    Next Sh
End Sub
