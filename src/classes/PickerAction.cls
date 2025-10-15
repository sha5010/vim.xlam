VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_PickerAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*
' * Defines the types of actions that can be performed in a picker.
' */
Public Enum eActionType
    Move            ' Represents a navigation action.
    Func            ' Represents a custom function action.
    Enter           ' Represents an action to confirm a selection.
    NotDisplayed    ' Represents an action that is not displayed in the help message.
End Enum

Private mActionType As eActionType  ' The type of action.
Private mHelpTextJa As String       ' The Japanese help text for the action.
Private mHelpTextEn As String       ' The English help text for the action.
Private mKeyBind As String          ' The key binding associated with the action.
Private mActionCode As Long         ' The numeric code identifying the action.

'/*
' * Returns the localized name of the action type.
' *
' * @returns {String} - The localized action type name.
' */
Public Property Get ActionTypeName() As String
    Select Case mActionType
        Case eActionType.Move
            ActionTypeName = IIf(gVim.IsJapanese, "移動", "Move")
        Case eActionType.Func
            ActionTypeName = IIf(gVim.IsJapanese, "機能", "Function")
        Case eActionType.Enter
            ActionTypeName = IIf(gVim.IsJapanese, "確定", "Enter")
    End Select
End Property

'/*
' * Returns the localized help message for the action, including its key binding.
' *
' * @returns {String} - The formatted help message.
' */
Public Property Get HelpMessage() As String
    If gVim.IsJapanese Then
        HelpMessage = mHelpTextJa
    Else
        HelpMessage = mHelpTextEn
    End If

    Dim displayKey As String
    displayKey = gVim.KeyMap.SendKeysToDisplayText(mKeyBind)
    displayKey = Replace(displayKey, "^", "Ctrl-")
    displayKey = Replace(displayKey, "%", "Alt-")
    If displayKey Like "<?*>" Then
        displayKey = StrConv(Mid(displayKey, 2, Len(displayKey) - 2), vbProperCase)
    End If

    HelpMessage = displayKey & vbTab & HelpMessage
End Property

'/*
' * Sets up the PickerAction object with its properties.
' *
' * @param {eActionType} ActionType - The type of action.
' * @param {String} key - The key binding for the action.
' * @param {String} helpTextJa - The Japanese help text.
' * @param {String} helpTextEn - The English help text.
' * @param {Long} registerAction - The numeric code for the action.
' * @returns {cls_PickerAction} - The initialized PickerAction object.
' */
Public Function Setup(ByVal ActionType As eActionType, ByVal key As String, ByVal helpTextJa As String, ByVal helpTextEn As String, ByVal registerAction As Long) As cls_PickerAction
    mActionType = ActionType
    mKeyBind = key
    mHelpTextJa = helpTextJa
    mHelpTextEn = helpTextEn
    mActionCode = registerAction

    Set Setup = Me
End Function

'/*
' * Returns the key binding associated with this action.
' *
' * @returns {String} - The key binding.
' */
Public Property Get KeyBind() As String
    KeyBind = mKeyBind
End Property

'/*
' * Returns the type of the action.
' *
' * @returns {Long} - The action type (from eActionType enum).
' */
Public Property Get ActionType() As Long
    ActionType = mActionType
End Property

'/*
' * Returns the numeric code for the action.
' *
' * @returns {Long} - The action code.
' */
Public Property Get ActionCode() As Long
    ActionCode = mActionCode
End Property
