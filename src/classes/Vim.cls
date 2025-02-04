VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Vim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cMsg As cls_Message                 ' Message text
Private cHelp As cls_Help                   ' Help text
Private cVars As cls_Vars                   ' Global variables

Private cIsEnabled As Boolean               ' Vim mode (enabled / disabled)
Private cCount As Long                      ' Counter
Private cModes As cls_Modes                 ' Defined modes
Private cConfig As cls_Config               ' Loaded configs
Private cKeyMap As cls_KeymapManager        ' Keymap manager

Private cJumpList As cls_DblCircList        ' Hold jumplist

'/*
' * Initializes the Vim class with default configurations and components.
' */
Private Sub Class_Initialize()
    ' Initialize Message
    Set cMsg = New cls_Message
    Set cMsg.Parent = Me

    ' Initialize Help
    Set cHelp = New cls_Help
    Set cHelp.Parent = Me

    ' Initialize Vars
    Set cVars = New cls_Vars
    Set cVars.Parent = Me

    ' Set default config
    Set cConfig = New cls_Config
    Set cConfig.Parent = Me

    ' Initialize Keymap manager
    Set cKeyMap = New cls_KeymapManager
    Set cKeyMap.Parent = Me

    ' Initialize Modes
    Set cModes = New cls_Modes
    Set cModes.Parent = Me
End Sub

'/*
' * Cleans up resources and disables Vim addin when the class is terminated.
' */
Public Sub Quit()
    ' Flag to quit
    cCount = -1

    Me.Enabled = False
    Application.OnKey cConfig.VimToggleKey

    Set cMsg = Nothing
    Set cVars = Nothing
    Set cModes = Nothing
    Set cConfig = Nothing
    Set cKeyMap = Nothing
    Set cJumpList = Nothing

End Sub

'/*
' * Gets or sets the enable/disable status of the Vim addin.
' *
' * @returns {Boolean} - True if Vim is enabled, False otherwise.
' */
Public Property Get Enabled() As Boolean
    Enabled = cIsEnabled
End Property
Public Property Let Enabled(ByVal arg As Boolean)
    Dim statusMsg As String

    cIsEnabled = arg

    If cIsEnabled Then
        'Make enable
        cCount = 0

        If cJumpList Is Nothing Then
            Set cJumpList = New cls_DblCircList
            cJumpList.Max = Me.Config.MaxHistories
        End If

        Call DisableIME
        Call cModes.Change(MODE_NORMAL)

    Else
        'Make disable
        Call cModes.Disable
        Call SetStatusBar
        Call cKeyMap.UnbindAll
    End If

    Application.OnKey Me.Config.VimToggleKey, "ToggleVim"

    If Not gVim Is Nothing And cCount <> -1 Then
        If cIsEnabled Then
            statusMsg = cMsg.VimStarted
        Else
            statusMsg = cMsg.VimStopped
        End If
        Call SetStatusBarTemporarily(statusMsg, 3000)
    End If
End Property

'/*
' * Gets or sets the Japanese language setting.
' *
' * @returns {Boolean} - True if the language is set to Japanese, False otherwise.
' */
Public Property Get IsJapanese() As Boolean
    IsJapanese = Me.Config.LangJapanese
End Property
Public Property Let IsJapanese(ByVal arg As Boolean)
    Me.Config.LangJapanese = arg
    Call SetStatusBarTemporarily(gVim.Msg.LangChanged, 2000)
End Property

'/*
' * Gets or sets the debug mode setting.
' *
' * @returns {Boolean} - True if debug mode is enabled, False otherwise.
' */
Public Property Get DebugMode() As Boolean
    DebugMode = Me.Config.DebugMode
End Property
Public Property Let DebugMode(ByVal arg As Boolean)
    Me.Config.DebugMode = arg
    If arg Then
        Call SetStatusBarTemporarily(gVim.Msg.DebugModeEnabled, 2000)
    Else
        Call SetStatusBarTemporarily(gVim.Msg.DebugModeDisabled, 2000)
    End If
End Property

'/*
' * Gets or sets the count property.
' *
' * @returns {Long} - The current count value.
' */
Public Property Get Count1() As Long
    If cCount = 0 Then
        Count1 = 1
    Else
        Count1 = cCount
    End If
End Property
Public Property Let Count1(ByVal arg As Long)
    cCount = arg
End Property
Public Property Get Count() As Long
    Count = cCount
End Property
Public Property Let Count(ByVal arg As Long)
    cCount = arg
End Property

'/*
' * Gets the jumplist manager.
' *
' * @returns {cls_MainArena} - The jumplist manager.
' */
Public Property Get JumpList() As cls_DblCircList
    Set JumpList = cJumpList
End Property

'/*
' * Gets the message properties for displaying messages.
' *
' * @returns {cls_Message} - The message manager.
' */
Public Property Get Msg() As cls_Message
    Set Msg = cMsg
End Property

'/*
' * Gets the help properties for getting help text.
' *
' * @returns {cls_Help} - The help text manager.
' */
Public Property Get Help() As cls_Help
    Set Help = cHelp
End Property

'/*
' * Gets or sets the global variables.
' *
' * @returns {cls_Vars} - The global variables manager.
' */
Public Property Get Vars() As cls_Vars
    Set Vars = cVars
End Property

'/*
' * Gets or sets the configurations.
' *
' * @returns {cls_Config} - The configuration manager.
' */
Public Property Get Config() As cls_Config
    Set Config = cConfig
End Property

'/*
' * Gets the keymap manager.
' *
' * @returns {cls_KeymapManager} - The keymap manager.
' */
Public Property Get KeyMap() As cls_KeymapManager
    Set KeyMap = cKeyMap
End Property

'/*
' * Gets the modes manager.
' *
' * @returns {cls_Modes} - The modes manager.
' */
Public Property Get Mode() As cls_Modes
    Set Mode = cModes
End Property
