VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cls_Vim

'config dictionary
Private cConfig As Dictionary

'keyname
Private Const STATUS_PREFIX = "statusprefix"
Private Const VIM_TOGGLE_KEY = "togglekey"
Private Const SCROLL_OFFSET = "scrolloff"
Private Const MAX_HISTORIES = "jumplisthistory"
Private Const LANG_JAPANESE = "japanese"
Private Const JIS_KEYBOARD = "jiskeyboard"
Private Const DEBUG_MODE = "debug"

'properties
Public Property Get StatusPrefix() As String
    StatusPrefix = cConfig(STATUS_PREFIX)
End Property
Public Property Let StatusPrefix(ByVal arg As String)
    If Len(arg) > 256 Then
        arg = Left(arg, 256)
    End If
    cConfig(STATUS_PREFIX) = arg
End Property
Public Property Get VimToggleKey() As String
    VimToggleKey = cConfig(VIM_TOGGLE_KEY)
End Property
Public Property Let VimToggleKey(ByVal arg As String)
    On Error Resume Next

    ' Convert from vim key to SendKeys style and store value
    cConfig(VIM_TOGGLE_KEY) = Parent.KeyMap.VimToVBA(arg, KEY_SEPARATOR, True, VBASendkeys)
End Property
Public Property Get ScrollOffset() As Double
    ScrollOffset = cConfig(SCROLL_OFFSET)
End Property
Public Property Let ScrollOffset(ByVal arg As Double)
    If arg < 0 Then
        arg = 0
    ElseIf arg > 500 Then
        arg = 500
    End If
    cConfig(SCROLL_OFFSET) = arg
End Property
Public Property Get MaxHistories() As Long
    MaxHistories = cConfig(MAX_HISTORIES)
End Property
Public Property Let MaxHistories(ByVal arg As Long)
    If arg < 0 Then
        arg = 0
    ElseIf arg > 1000 Then
        arg = 1000
    End If
    cConfig(MAX_HISTORIES) = arg
End Property
Public Property Get LangJapanese() As Boolean
    LangJapanese = cConfig(LANG_JAPANESE)
End Property
Public Property Let LangJapanese(ByVal arg As Boolean)
    cConfig(LANG_JAPANESE) = arg
End Property
Public Property Get IsJisKeyboard() As Boolean
    IsJisKeyboard = cConfig(JIS_KEYBOARD)
End Property
Public Property Let IsJisKeyboard(ByVal arg As Boolean)
    cConfig(JIS_KEYBOARD) = arg
End Property
Public Property Get DebugMode() As Boolean
    DebugMode = cConfig(DEBUG_MODE)
End Property
Public Property Let DebugMode(ByVal arg As Boolean)
    cConfig(DEBUG_MODE) = arg
End Property

Private Sub Class_Initialize()
    Set cConfig = New Dictionary

    'defined keys
    With cConfig
        .Add STATUS_PREFIX, "vim.xlam: "
        .Add VIM_TOGGLE_KEY, "^{m}"
        .Add SCROLL_OFFSET, 54#
        .Add MAX_HISTORIES, 100
        .Add LANG_JAPANESE, True
        .Add JIS_KEYBOARD, True
        .Add DEBUG_MODE, Not ThisWorkbook.IsAddin
    End With
End Sub

Private Sub Class_Terminate()
    Set cConfig = Nothing
End Sub

'/**
' * Get config value
' *
' * @param {String} keyName
' * @param {Variant} defaultValue - use this value when keyName does not exists
' * @return {Variant} - config value or default value
' */
Public Function Get_(ByRef keyName As String, Optional ByVal defaultValue As Variant) As Variant
    If cConfig.Exists(keyName) Then
        Get_ = cConfig(keyName)
    ElseIf Not defaultValue = Empty Then
        Get_ = defaultValue
    Else
        Err.Raise 50000, Description:="Config key does not exists"
    End If
End Function

'/**
' * Set config value
' *
' * @param {String} keyName
' * @param {Variant} arg - config value
' */
Public Sub Set_(ByRef keyName As String, ByVal arg As Variant)
    Select Case LCase(keyName)
        Case STATUS_PREFIX: StatusPrefix = arg
        Case VIM_TOGGLE_KEY: VimToggleKey = arg
        Case SCROLL_OFFSET: ScrollOffset = arg
        Case MAX_HISTORIES: MaxHistories = arg
        Case LANG_JAPANESE: LangJapanese = arg
        Case JIS_KEYBOARD: IsJisKeyboard = arg
        Case DEBUG_MODE: DebugMode = arg

        Case Else
            Err.Raise 50000, Description:="Config key does not exists"
    End Select
End Sub
