Attribute VB_Name = "C_Constants"
Option Explicit

' Definition for Keyboard event (keybd_event in user32)
Public Const KEYUP = &H2          'Key up
Public Const EXTENDED_KEY = &H1   'For using extended keys

Public Const LONG_MAX As Long = 2147483647
Public Const DUMMY_PROCEDURE As String = "Dummy"
Public Const SHOW_CMD_PROCEDURE As String = "ShowCmdForm"

' KeyMap
Public Const KEY_SEPARATOR As String = " - "
Public Const KEY_SUMMARY As String = " [SUMMARY]"
Public Const KEY_TEMP As String = " [TEMP]"
Public Const KEY_TERM_SYMBOL As String = vbLf

Public Const KEY_CMD As String = "<cmd>"    ' Must lower case
Public Const KEY_REMAP As String = "<key>"  ' Must lower case

' Modes
Public Const MODE_DUMMY As String = "_"
Public Const MODE_NORMAL As String = "n"
Public Const MODE_VISUAL As String = "v"
Public Const MODE_CMDLINE As String = "c"
Public Const MODE_SHAPEINSERT As String = "i"

' Cmdline
Public Const CMDLINE_CANCELED As String = vbLf

Public Enum eResultType
    VBASendkeys = 1
    VBAKeyCodes
End Enum

' ref: https://learn.microsoft.com/windows/win32/inputdev/virtual-key-codes
Public Enum eKey
    BackSpace_ = 8
    Tab_ = 9
    Enter_ = 13
    Pause_ = 19
    CapsLock_ = 20
    IME_On_ = 25
    IME_Off_
    Escape_
    Henkan_         ' JIS only  変換
    Muhenkan_       ' JIS only  無変換
    Space_ = 32
    PageUp_
    PageDown_
    End_
    Home_
    Left_
    Up_
    Right_
    Down_
    Select_
    Print_
    Execute_
    PrintScreen_
    Insert_
    Delete_
    Help_

    k0_
    k1_
    k2_
    k3_
    k4_
    k5_
    k6_
    k7_
    k8_
    k9_

    A_ = 65
    B_
    C_
    D_
    E_
    F_
    G_
    H_
    I_
    J_
    K_
    L_
    M_
    N_
    O_
    P_
    Q_
    R_
    S_
    T_
    U_
    V_
    W_
    X_
    Y_
    Z_
    WinLeft_
    WinRight_
    Application_
    Numpad0_ = 96
    Numpad1_
    Numpad2_
    Numpad3_
    Numpad4_
    Numpad5_
    Numpad6_
    Numpad7_
    Numpad8_
    Numpad9_
    NumpadMultiply_     '  *   in numeric keypad
    NumpadAdd_          '  +   in numeric keypad
    NumpadEnter_        'Enter in numeric keypad
    NumpadSubtract_     '  -   in numeric keypad
    NumpadDecimal_      '  .   in numeric keypad
    NumpadDivide_       '  /   in numeric keypad
    F1_
    F2_
    F3_
    F4_
    F5_
    F6_
    F7_
    F8_
    F9_
    F10_
    F11_
    F12_
    F13_
    F14_
    F15_
    F16_
    NumLock_ = 144
    ScrollLock_

    ShiftLeft_ = 160            'Left Shift
    ShiftRight_                 'Right Shift
    CtrlLeft_                   'Left Ctrl
    CtrlRight_                  'Right Ctrl
    AltLeft_                    'Left Alt
    AltRight_                   'Right Alt

    Semicoron_US_ = 186         ' US  ; (Shift : )      JIS  : (Shift * )
    Coron_JIS_ = 186            ' US  ; (Shift : )      JIS  : (Shift * )
    Equal_US_ = 187             ' US  = (Shift + )      JIS  ; (Shift + )
    Semicoron_JIS_ = 187        ' US  = (Shift + )      JIS  ; (Shift + )
    Comma_ = 188                ' US  , (Shift < )      JIS  , (Shift < )
    Minus_                      ' US  - (Shift _ )      JIS  - (Shift = )
    Period_                     ' US  . (Shift > )      JIS  . (Shift > )
    Slash_                      ' US  / (Shift ? )      JIS  / (Shift ? )
    BackQuote_US_ = 192         ' US  ` (Shift ‾ )      JIS  @ (Shift ` )
    AtMark_JIS_ = 192           ' US  ` (Shift ‾ )      JIS  @ (Shift ` )
    OpeningSquareBracket_ = 219 ' US  [ (Shift { )      JIS  [ (Shift { )
    Backslash_                  ' US  ¥ (Shift | )      JIS  ¥ (Shift | )  [upperside ¥ ]
    ClosingSquareBracket_       ' US  ] (Shift } )      JIS  ] (Shift } )
    SingleQuote_US_ = 222       ' US  ' (Shift " )      JIS  ^ (Shift ‾ )
    Caret_JIS_ = 222            ' US  ' (Shift " )      JIS  ^ (Shift ‾ )
    Underscore_ = 226           ' US  None              JIS  ¥ (Shift _ )  [downside ¥ ]

    'JIS keyboard only
    Eisu_ = 240                 ' 英数
    Katakana_ = 242             ' カタカナ ひらがな
    HankakuZenkaku_             ' 半角/全角

    'Customize
    Ctrl_ = 512
    Shift_ = 1024
    Alt_ = 2048
End Enum
