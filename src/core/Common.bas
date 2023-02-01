Attribute VB_Name = "C_Common"
Option Explicit
Option Private Module

#If Win64 Then
    Public Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Function GetKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long
    Public Declare Function GetKeyState Lib "user32.dll" (ByVal vKey As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


'Definition for Keyboard event (keybd_event in user32)
Public Const KEYUP = &H2          'Key up
Public Const EXTENDED_KEY = &H1   'For using extended keys
'
'Left Ctrl, Left Alt, Right Shift, Insert, delete,
'Home, End, Page Up, Page Down
'Arrow Keys (Up, Down, Left, Right)
'Num Lock, Break (Ctrl + Pause)
'Print Screen, Enter key on 10 keys

'Virtual Key codes
Public Const LSHIFT = &HA0 'Left Shift
Public Const RSHIFT = &HA1 'Right Shift
Public Const LCTRL = &HA2  'Left Ctrl
Public Const RCTRL = &HA3  'Right Ctrl
Public Const LMENU = &HA4  'Left Alt
Public Const RMENU = &HA5  'Right Alt
Public Const KANJI = &H19  'Kanji
Public Const APPS = &H5D   'Application Keys

Public Enum vKey
    BackSpace_ = 8
    Tab_ = 9
    Enter_ = 13
    Pause_ = 19
    CapsLock_ = 20
    Kanji_ = 25
    Escape_ = 27
    Henkan_ = 28
    Muhenkan_ = 29
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
    NumpadMultiply_     'テンキーの *
    NumpadAdd_          'テンキーの +
    NumpadEnter_        'テンキーの Enter
    NumpadSubstract_    'テンキーの -
    NumpadDecimal_      'テンキーの .
    NumpadDivide_       'テンキーの /
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

    ShiftLeft_ = 160       'Left Shift
    ShiftRight_            'Right Shift
    CtrlLeft_              'Left Ctrl
    CtrlRight_             'Right Ctrl
    AltLeft_               'Left Alt
    AltRight_              'Right Alt

    'Shift_JIS 配列
    Coron_ = 186                ' :  (Shift: *)
    Semicoron_                  ' ;  (Shift: +)
    Comma_                      ' ,  (Shift: <)
    Minus_                      ' -  (Shift: =)
    Period_                     ' .  (Shift: >)
    Slash_                      ' /  (Shift: ?)
    AtMark_                     ' @  (Shift: `)
    OpeningSquareBracket_ = 219 ' [  (Shift: {)
    Backslash_                  ' ¥  (Shift: |)  上側の ¥
    ClosingSquareBracket_       ' ]  (Shift: })
    Caret_                      ' ^  (Shift: ‾)
    Underscore_ = 226           ' ¥  (Shift: _)  下側の ¥
    Eisu_ = 240                 ' Caps Lock
    Katakana_ = 242             ' カタカナ ひらがな
    HankakuZenkaku_             ' 半角/全角

    Ctrl_ = 512
    Shift_ = 1024
    Alt_ = 2048
End Enum

Private Function keystrokeAPI(ByVal key As Integer)
    Dim Ctrl As Boolean
    Dim Shift As Boolean
    Dim Alt As Boolean

    Ctrl = ((key ¥ Ctrl_) And 1) = 1
    Shift = ((key ¥ Shift_) And 1) = 1
    Alt = ((key ¥ Alt_) And 1) = 1

    key = key And &HFF

    If Ctrl Then keybd_event vbKeyControl, 0, 0, 0
    If Shift Then keybd_event vbKeyShift, 0, 0, 0
    If Alt Then keybd_event vbKeyMenu, 0, 0, 0

    keybd_event key, 0, 0, 0
    keybd_event key, 0, KEYUP, 0

    If Alt Then keybd_event vbKeyMenu, 0, KEYUP, 0
    If Shift Then keybd_event vbKeyShift, 0, KEYUP, 0
    If Ctrl Then keybd_event vbKeyControl, 0, KEYUP, 0
End Function

' /**
'  * 指定されたキーを順番に押す
'  * (Ctrlキーなどの解放はしない。手動でやったときに使う)
'  *
'  * Alt + H を押した後 I を押す
'  *     Call keystrokeWithoutKeyup(Alt_ + H_, I_)
'  */
Function keystrokeWithoutKeyup(ParamArray keys() As Variant)
    Dim i As Integer
    Dim u As Integer

    u = UBound(keys)

    For i = LBound(keys) To u
        Call keystrokeAPI(keys(i))
    Next i
End Function

' /**
'  * 指定されたキーを順番に押す。なお遅くなるので長押しするようなキーには不向き
'  * (Ctrlキーなどを解放する。)
'  *
'  * @param releaseShiftKey: Shiftキーの解放をする(True)/しない(False)
'  *
'  * Alt + H を押した後 I を押す
'  *     Call keystroke(True, Alt_ + H_, I_)
'  */
Function keystroke(ByVal releaseShiftKey As Boolean, ParamArray keys() As Variant)
    Dim i As Integer
    Dim u As Integer

    Call keyupControlKeys

    If releaseShiftKey Then
        Call releaseShiftKeys
    End If

    u = UBound(keys)

    For i = LBound(keys) To u
        Call keystrokeAPI(keys(i))
    Next i

    Call unkeyupControlKeys
End Function

Function releaseShiftKeys()
    If GetKeyState(LSHIFT) > 0 Then
        keybd_event LSHIFT, 0, KEYUP, 0
    ElseIf GetKeyState(RSHIFT) > 0 Then
        keybd_event RSHIFT, 0, KEYUP, 0
    Else
        keybd_event vbKeyShift, 0, KEYUP, 0
    End If
End Function

Function keyupControlKeys()
    keybd_event LSHIFT, 0, KEYUP, 0
    keybd_event RSHIFT, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event LCTRL, 0, KEYUP, 0
    keybd_event RCTRL, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event LMENU, 0, KEYUP, 0
    keybd_event RMENU, 0, EXTENDED_KEY Or KEYUP, 0
End Function

Function unkeyupControlKeys()
    If (GetKeyState(LSHIFT) And &H8000) <> 0 Then
        keybd_event LSHIFT, 0, 0, 0
    ElseIf (GetKeyState(RSHIFT) And &H8000) <> 0 Then
        keybd_event RSHIFT, 0, EXTENDED_KEY, 0
    ElseIf (GetKeyState(vbKeyShift) And &H8000) <> 0 Then
        keybd_event vbKeyShift, 0, 0, 0
    End If

    If (GetKeyState(LCTRL) And &H8000) <> 0 Then
        keybd_event LCTRL, 0, 0, 0
    ElseIf (GetKeyState(RCTRL) And &H8000) <> 0 Then
        keybd_event RCTRL, 0, EXTENDED_KEY, 0
    ElseIf (GetKeyState(vbKeyControl) And &H8000) <> 0 Then
        keybd_event vbKeyControl, 0, 0, 0
    End If

    If (GetKeyState(LMENU) And &H8000) <> 0 Then
        keybd_event LMENU, 0, 0, 0
    ElseIf (GetKeyState(RMENU) And &H8000) <> 0 Then
        keybd_event RMENU, 0, EXTENDED_KEY, 0
    ElseIf (GetKeyState(vbKeyMenu) And &H8000) <> 0 Then
        keybd_event vbKeyMenu, 0, 0, 0
    End If
End Function

Function disableIME()
    If IMEStatus <> vbIMEOff Then
        Call keystrokeWithoutKeyup(Kanji_)
    End If
End Function

Function repeatRegister(ByVal funcName As String, ParamArray args() As Variant)
    If Repeater Is Nothing Then
        Set Repeater = New cls_Repeater
    End If

    Call Repeater.Register(funcName, gCount, args)
End Function

Function repeatAction()
    If Not Repeater Is Nothing Then
        Call Repeater.Run
    End If
End Function
