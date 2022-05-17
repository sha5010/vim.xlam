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
    If GetKeyState(LSHIFT) < 0 Then
        keybd_event LSHIFT, 0, 0, 0
    ElseIf GetKeyState(RSHIFT) < 0 Then
        keybd_event RSHIFT, 0, EXTENDED_KEY, 0
    ElseIf GetKeyState(vbKeyShift) < 0 Then
        keybd_event vbKeyShift, 0, 0, 0
    Else
        keybd_event vbKeyShift, 0, KEYUP, 0
    End If

    If GetKeyState(LCTRL) < 0 Then
        keybd_event LCTRL, 0, 0, 0
    ElseIf GetKeyState(RCTRL) < 0 Then
        keybd_event RCTRL, 0, EXTENDED_KEY, 0
    Else
        keybd_event vbKeyControl, 0, KEYUP, 0
    End If

    If GetKeyState(LMENU) < 0 Then
        keybd_event LMENU, 0, 0, 0
    ElseIf GetKeyState(RMENU) < 0 Then
        keybd_event RMENU, 0, EXTENDED_KEY, 0
    Else
        keybd_event vbKeyMenu, 0, KEYUP, 0
    End If
End Function

