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

'custom config filename
Private Const CUSTOM_CONFIG = "_vimxlamrc"

'keyname
Private Const STATUS_PREFIX = "statusprefix"
Private Const VIM_TOGGLE_KEY = "togglekey"
Private Const SCROLL_OFFSET = "scrolloff"
Private Const MAX_HISTORIES = "jumplisthistory"
Private Const LANG_JAPANESE = "japanese"
Private Const JIS_KEYBOARD = "jiskeyboard"
Private Const QUIT_APP_ALL_WB_CLOSED = "quitapp"
Private Const NUMPAD_COUNT = "numpadcount"
Private Const SUGGEST_WAIT = "suggestwait"
Private Const SUGGEST_LABELS = "suggestlabels"
Private Const COLOR_PICKER_SIZE = "colorpickersize"
Private Const CUSTOM_COLOR_1 = "customcolor1"
Private Const CUSTOM_COLOR_2 = "customcolor2"
Private Const CUSTOM_COLOR_3 = "customcolor3"
Private Const CUSTOM_COLOR_4 = "customcolor4"
Private Const CUSTOM_COLOR_5 = "customcolor5"
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
Public Property Get QuitApp() As Boolean
    QuitApp = cConfig(QUIT_APP_ALL_WB_CLOSED)
End Property
Public Property Let QuitApp(ByVal arg As Boolean)
    cConfig(QUIT_APP_ALL_WB_CLOSED) = arg
End Property
Public Property Get NumpadCount() As Boolean
    NumpadCount = cConfig(NUMPAD_COUNT)
End Property
Public Property Let NumpadCount(ByVal arg As Boolean)
    cConfig(NUMPAD_COUNT) = arg
    Call ChangeNumpadKeymap(arg)
End Property
Public Property Get SuggestWait() As Long
    SuggestWait = cConfig(SUGGEST_WAIT)
End Property
Public Property Let SuggestWait(ByVal arg As Long)
    cConfig(SUGGEST_WAIT) = arg
End Property
Public Property Get SuggestLabels() As String
    SuggestLabels = cConfig(SUGGEST_LABELS)
End Property
Public Property Let SuggestLabels(ByVal arg As String)
    cConfig(SUGGEST_LABELS) = arg
End Property
Public Property Get ColorPickerSize() As Double
    ColorPickerSize = cConfig(COLOR_PICKER_SIZE)
End Property
Public Property Let ColorPickerSize(ByVal arg As Double)
    If arg < 6 Then
        arg = 6
    ElseIf arg > 48 Then
        arg = 48
    End If
    cConfig(COLOR_PICKER_SIZE) = arg
End Property
Public Property Get CustomColor1() As Variant
    CustomColor1 = cConfig(CUSTOM_COLOR_1)
End Property
Public Property Let CustomColor1(ByVal arg As Variant)
    Dim colorCode As Long: colorCode = ConvertCustomColorToLong(arg)
    If colorCode > -1 Then cConfig(CUSTOM_COLOR_1) = colorCode
End Property
Public Property Get CustomColor2() As Variant
    CustomColor2 = cConfig(CUSTOM_COLOR_2)
End Property
Public Property Let CustomColor2(ByVal arg As Variant)
    Dim colorCode As Long: colorCode = ConvertCustomColorToLong(arg)
    If colorCode > -1 Then cConfig(CUSTOM_COLOR_2) = colorCode
End Property
Public Property Get CustomColor3() As Variant
    CustomColor3 = cConfig(CUSTOM_COLOR_3)
End Property
Public Property Let CustomColor3(ByVal arg As Variant)
    Dim colorCode As Long: colorCode = ConvertCustomColorToLong(arg)
    If colorCode > -1 Then cConfig(CUSTOM_COLOR_3) = colorCode
End Property
Public Property Get CustomColor4() As Variant
    CustomColor4 = cConfig(CUSTOM_COLOR_4)
End Property
Public Property Let CustomColor4(ByVal arg As Variant)
    Dim colorCode As Long: colorCode = ConvertCustomColorToLong(arg)
    If colorCode > -1 Then cConfig(CUSTOM_COLOR_4) = colorCode
End Property
Public Property Get CustomColor5() As Variant
    CustomColor5 = cConfig(CUSTOM_COLOR_5)
End Property
Public Property Let CustomColor5(ByVal arg As Variant)
    Dim colorCode As Long: colorCode = ConvertCustomColorToLong(arg)
    If colorCode > -1 Then cConfig(CUSTOM_COLOR_5) = colorCode
End Property
Public Property Get DebugMode() As Boolean
    DebugMode = cConfig(DEBUG_MODE)
End Property
Public Property Let DebugMode(ByVal arg As Boolean)
    cConfig(DEBUG_MODE) = arg
End Property

Private Function ConvertCustomColorToLong(ByVal arg As Variant) As Long
    Select Case VarType(arg)
        Case vbString
            ConvertCustomColorToLong = HexColorCodeToLong(Replace(arg, "#", ""))
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte, vbCurrency, vbDecimal
            arg = CLng(arg)
            If 0 < arg And arg < 16777216 Then
                ConvertCustomColorToLong = arg
            Else
                ConvertCustomColorToLong = -1
            End If
    End Select
End Function

Private Sub ChangeNumpadKeymap(ByVal isEnabled As Boolean)
    Dim i As Long
    For i = 1 To 9
        If isEnabled Then
            Call Parent.KeyMap.Map("nmap <k" & CStr(i) & "> " & SHOW_CMD_PROCEDURE & " """ & CStr(i) & """")
        Else
            Call Parent.KeyMap.Map("nunmap <k" & CStr(i) & ">")
        End If
    Next
End Sub


Private Sub Class_Initialize()
    Set cConfig = New Dictionary

    'defined keys
    With cConfig
        .Add STATUS_PREFIX, "vim.xlam: "
        .Add VIM_TOGGLE_KEY, "^{m}"
        .Add SCROLL_OFFSET, 54#
        .Add MAX_HISTORIES, 100
        .Add LANG_JAPANESE, Application.LanguageSettings.LanguageID(msoLanguageIDUI) = msoLanguageIDJapanese
        .Add JIS_KEYBOARD, IsJISKeyboardLayout()
        .Add QUIT_APP_ALL_WB_CLOSED, True
        .Add NUMPAD_COUNT, False
        .Add SUGGEST_WAIT, 1000
        .Add SUGGEST_LABELS, "asdfghjkl;qwertyuiopzxcvbnm,./1234567890"
        .Add COLOR_PICKER_SIZE, 12#
        .Add CUSTOM_COLOR_1, &H66FF&
        .Add CUSTOM_COLOR_2, &H6699FF
        .Add CUSTOM_COLOR_3, &HFF00FF
        .Add CUSTOM_COLOR_4, &H8000&
        .Add CUSTOM_COLOR_5, &HFF0000
        .Add DEBUG_MODE, Not ThisWorkbook.IsAddin
    End With

End Sub

Private Sub Class_Terminate()
    Set cConfig = Nothing
    Set Parent = Nothing
End Sub

'/**
' * Get config value
' *
' * @param {String} keyName
' * @param {Variant} defaultValue - use this value when keyName does not exists
' * @return {Variant} - config value or default value
' */
Public Function Get_(ByRef keyName As String, Optional ByVal defaultValue As Variant) As Variant
    On Error GoTo Catch

    If cConfig.Exists(keyName) Then
        Get_ = cConfig(keyName)
    ElseIf Not defaultValue = Empty Then
        Get_ = defaultValue
    Else
        Err.Raise 50000, Description:="Config key """ & keyName & """ does not exist"
    End If
    Exit Function

Catch:
    Call ErrorHandler("Config.Get_")
End Function

'/**
' * Set config value
' *
' * @param {String} keyName
' * @param {Variant} arg - config value
' */
Public Sub Set_(ByRef keyName As String, ByVal arg As Variant)
    On Error GoTo Catch

    Select Case LCase(keyName)
        Case STATUS_PREFIX: StatusPrefix = arg
        Case VIM_TOGGLE_KEY: VimToggleKey = arg
        Case SCROLL_OFFSET: ScrollOffset = arg
        Case MAX_HISTORIES: MaxHistories = arg
        Case LANG_JAPANESE: LangJapanese = arg
        Case JIS_KEYBOARD: IsJisKeyboard = arg
        Case QUIT_APP_ALL_WB_CLOSED: QuitApp = arg
        Case NUMPAD_COUNT: NumpadCount = arg
        Case SUGGEST_WAIT: SuggestWait = arg
        Case SUGGEST_LABELS: SuggestLabels = arg
        Case COLOR_PICKER_SIZE: ColorPickerSize = arg
        Case CUSTOM_COLOR_1: CustomColor1 = arg
        Case CUSTOM_COLOR_2: CustomColor2 = arg
        Case CUSTOM_COLOR_3: CustomColor3 = arg
        Case CUSTOM_COLOR_4: CustomColor4 = arg
        Case CUSTOM_COLOR_5: CustomColor5 = arg
        Case DEBUG_MODE: DebugMode = arg

        Case Else
            Err.Raise 50000, Description:="Config key """ & keyName & """ does not exist"
    End Select
    Exit Sub

Catch:
    Call ErrorHandler("Config.Set_")
End Sub

Public Sub LoadCustomConfig(Optional ByVal configFilePath As String)
    If configFilePath = "" Then
        configFilePath = ThisWorkbook.Path

        If configFilePath Like "https://*" Then
            Dim sepIndex As Long
            If configFilePath Like "*my.sharepoint.com/*" Then
                sepIndex = InStr(configFilePath, "/Documents/")
                configFilePath = Mid(configFilePath, sepIndex + 10)
            Else
                sepIndex = InStr(10, configFilePath, "/")
                sepIndex = InStr(sepIndex + 1, configFilePath, "/")
                configFilePath = Mid(configFilePath, sepIndex)
            End If
            configFilePath = Environ("OneDrive") & Replace(configFilePath, "/", "\")
        End If

        configFilePath = configFilePath & "\" & CUSTOM_CONFIG
    End If

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject

    ' Check file exists
    If Not fso.FileExists(configFilePath) Then
        Exit Sub
    End If

    ' Load file
    Dim fs As TextStream
    Set fs = fso.OpenTextFile(configFilePath, ForReading, False, TristateFalse)

    ' Process each line
    While Not fs.AtEndOfStream
        Call ParseConfig(fs.ReadLine())
    Wend

    ' Close the file
    fs.Close
End Sub

Private Sub ParseConfig(ByRef configStr As String)
    On Error GoTo Catch

    configStr = LTrim(configStr)

    ' Ignore lines beginning with "#"
    If configStr Like "[#]*" Then
        Exit Sub
    End If

    ' Parse "set" config
    If configStr Like "set *" Then
        Dim configKey As String
        Dim arg As Variant

        ' set noXXXX
        If configStr Like "set no*" Then
            configKey = RTrim(Mid(configStr, 7))
            If cConfig.Exists(configKey) Then
                Call Set_(configKey, False)
                Exit Sub
            End If
        End If

        ' set XXXX=YYYY
        If configStr Like "set *=*" Then
            Dim equalIndex As Long: equalIndex = InStr(configStr, "=")
            configKey = Mid(configStr, 5, equalIndex - 5)
            arg = Mid(configStr, equalIndex + 1)
            If cConfig.Exists(configKey) Then
                Call Set_(configKey, arg)
                Exit Sub
            End If

        ' set XXXX
        Else
            configKey = RTrim(Mid(configStr, 5))
            If cConfig.Exists(configKey) Then
                Call Set_(configKey, True)
                Exit Sub
            End If
        End If
    End If

    ' Parse "map" config
    If configStr Like "*map *" Then
        Call Parent.KeyMap.Map(configStr)
    End If

    Exit Sub
Catch:
End Sub
