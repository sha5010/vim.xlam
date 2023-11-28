Attribute VB_Name = "C_KeyStroke"
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

'/*
' * Simulates a keystroke for a single key.
' *
' * @param {Long} key - The key code for the desired key.
' */
Private Sub StrokeSingleKey(ByVal key As Long)
    Dim Ctrl As Boolean
    Dim Shift As Boolean
    Dim Alt As Boolean

    ' Extracting modifier key states
    Ctrl = ((key ¥ Ctrl_) And 1) = 1
    Shift = ((key ¥ Shift_) And 1) = 1
    Alt = ((key ¥ Alt_) And 1) = 1

    ' Extracting the actual key code
    key = key And &HFF

    ' Simulating keydown for modifier keys
    If Ctrl Then keybd_event vbKeyControl, 0, 0, 0
    If Shift Then keybd_event vbKeyShift, 0, 0, 0
    If Alt Then keybd_event vbKeyMenu, 0, 0, 0

    ' Simulating key stroke for the specified key
    keybd_event key, 0, 0, 0
    keybd_event key, 0, KEYUP, 0

    ' Simulating keyup for modifier keys
    If Alt Then keybd_event vbKeyMenu, 0, KEYUP, 0
    If Shift Then keybd_event vbKeyShift, 0, KEYUP, 0
    If Ctrl Then keybd_event vbKeyControl, 0, KEYUP, 0
End Sub

'/*
' * Press the specified keys in order (Do not release Ctrl/Shift/Alt keys)
' *
' * @param {eKey, ...} strokeKeys - An array of key codes for the desired keys.
' * @example: To press Alt + H and then I
' *     Call KeyStrokeWithoutKeyup(Alt_ + H_, I_)
' */
Sub KeyStrokeWithoutKeyup(ParamArray strokeKeys() As Variant)
    Dim i As Long

    For i = LBound(strokeKeys) To UBound(strokeKeys)
        Call StrokeSingleKey(strokeKeys(i))
    Next i
End Sub

'/*
' * Press the specified keys in order. (Release Ctrl/Alt keys)
' * Note: Not suitable for keys that need to be held down
' *
' * @param {Boolean} releaseShiftKey - Flag to release the Shift key before simulating other keys.
' * @param {eKey, ...} strokeKeys - An array of key codes for the desired keys.
' * @example: To press Alt + H and then I
' *     Call KeyStroke(True, Alt_ + H_, I_)
' */
Sub KeyStroke(ByVal releaseShiftKey As Boolean, ParamArray strokeKeys() As Variant)
    Dim i As Long

    Call KeyUpControlKeys

    If releaseShiftKey Then
        Call ReleaseShiftKeys
    End If

    For i = LBound(strokeKeys) To UBound(strokeKeys)
        Call StrokeSingleKey(strokeKeys(i))
    Next i

    Call UnkeyUpControlKeys
End Sub

'/*
' * Releases the Shift key based on its state.
' */
Sub ReleaseShiftKeys()
    If GetKeyState(ShiftLeft_) > 0 Then
        keybd_event ShiftLeft_, 0, KEYUP, 0
    ElseIf GetKeyState(ShiftRight_) > 0 Then
        keybd_event ShiftRight_, 0, KEYUP, 0
    Else
        keybd_event vbKeyShift, 0, KEYUP, 0
    End If
End Sub

'/*
' * Simulates keyup events for various control keys.
' */
Sub KeyUpControlKeys()
    keybd_event ShiftLeft_, 0, KEYUP, 0
    keybd_event ShiftRight_, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event CtrlLeft_, 0, KEYUP, 0
    keybd_event CtrlRight_, 0, EXTENDED_KEY Or KEYUP, 0
    keybd_event AltLeft_, 0, KEYUP, 0
    keybd_event AltRight_, 0, EXTENDED_KEY Or KEYUP, 0
End Sub

'/*
' * Simulates keyup events for control keys based on their current state.
' */
Sub UnkeyUpControlKeys()
    If (GetKeyState(ShiftLeft_) And &H8000) <> 0 Then
        keybd_event ShiftLeft_, 0, 0, 0
    ElseIf (GetKeyState(ShiftRight_) And &H8000) <> 0 Then
        keybd_event ShiftRight_, 0, EXTENDED_KEY, 0
    ElseIf (GetKeyState(vbKeyShift) And &H8000) <> 0 Then
        keybd_event vbKeyShift, 0, 0, 0
    End If

    If (GetKeyState(CtrlLeft_) And &H8000) <> 0 Then
        keybd_event CtrlLeft_, 0, 0, 0
    ElseIf (GetKeyState(CtrlRight_) And &H8000) <> 0 Then
        keybd_event CtrlRight_, 0, EXTENDED_KEY, 0
    ElseIf (GetKeyState(vbKeyControl) And &H8000) <> 0 Then
        keybd_event vbKeyControl, 0, 0, 0
    End If

    If (GetKeyState(AltLeft_) And &H8000) <> 0 Then
        keybd_event AltLeft_, 0, 0, 0
    ElseIf (GetKeyState(AltRight_) And &H8000) <> 0 Then
        keybd_event AltRight_, 0, EXTENDED_KEY, 0
    ElseIf (GetKeyState(vbKeyMenu) And &H8000) <> 0 Then
        keybd_event vbKeyMenu, 0, 0, 0
    End If
End Sub
