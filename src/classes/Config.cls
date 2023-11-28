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
Private cConfKeys As Collection

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
    StatusPrefix = Get_(STATUS_PREFIX)
End Property
Public Property Let StatusPrefix(ByVal arg As String)
    If Len(arg) > 256 Then
        arg = Left(arg, 256)
    End If
    Call Set_(STATUS_PREFIX, arg)
End Property
Public Property Get VimToggleKey() As String
    VimToggleKey = Get_(VIM_TOGGLE_KEY)
End Property
Public Property Let VimToggleKey(ByVal arg As String)
    ' Todo: validate
    Call Set_(VIM_TOGGLE_KEY, arg)
End Property
Public Property Get ScrollOffset() As Double
    ScrollOffset = Get_(SCROLL_OFFSET)
End Property
Public Property Let ScrollOffset(ByVal arg As Double)
    If arg < 0 Then
        arg = 0
    ElseIf arg > 500 Then
        arg = 500
    End If
    Call Set_(SCROLL_OFFSET, arg)
End Property
Public Property Get MaxHistories() As Long
    MaxHistories = Get_(MAX_HISTORIES)
End Property
Public Property Let MaxHistories(ByVal arg As Long)
    If arg < 0 Then
        arg = 0
    ElseIf arg > 1000 Then
        arg = 1000
    End If
    Call Set_(MAX_HISTORIES, arg)
End Property
Public Property Get LangJapanese() As Boolean
    LangJapanese = Get_(LANG_JAPANESE)
End Property
Public Property Let LangJapanese(ByVal arg As Boolean)
    Call Set_(LANG_JAPANESE, arg)
End Property
Public Property Get IsJisKeyboard() As Boolean
    IsJisKeyboard = Get_(JIS_KEYBOARD)
End Property
Public Property Let IsJisKeyboard(ByVal arg As Boolean)
    Call Set_(JIS_KEYBOARD, arg)
End Property
Public Property Get DebugMode() As Boolean
    DebugMode = Get_(DEBUG_MODE)
End Property
Public Property Let DebugMode(ByVal arg As Boolean)
    Call Set_(DEBUG_MODE, arg)
End Property

Private Sub Class_Initialize()
    'Debug.Print Timer, "Config init - start"
    Set cConfig = New Dictionary
    Set cConfKeys = New Collection

    'defined keys
    With cConfig
        .Add STATUS_PREFIX, "vim.xlam: "
        .Add VIM_TOGGLE_KEY, "^m"
        .Add SCROLL_OFFSET, 54#
        .Add MAX_HISTORIES, 100
        .Add LANG_JAPANESE, True
        .Add JIS_KEYBOARD, True
        .Add DEBUG_MODE, Not ThisWorkbook.IsAddin
    End With

    'generate key colletion
    Dim k As Variant
    For Each k In cConfig.Keys
        cConfKeys.Add k
    Next
    'Debug.Print Timer, "Config init - stop"
End Sub

Private Sub Class_Terminate()
    Set cConfig = Nothing
    Set cConfKeys = Nothing
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
    'Todo: validate config value

    If cConfig.Exists(keyName) Then
        cConfig(keyName) = arg
    Else
        cConfig.Add keyName, arg
    End If
End Sub
