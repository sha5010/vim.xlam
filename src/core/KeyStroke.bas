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
    Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
    Public Declare Function GetKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private pHoldCtrlLeft As Boolean
Private pHoldCtrlRight As Boolean
Private pHoldShiftLeft As Boolean
Private pHoldShiftRight As Boolean
Private pHoldAltLeft As Boolean
Private pHoldAltRight As Boolean

'/*
' * Simulates a keystroke for a single key with optional modifier keys.
' *
' * @param {Long} key - The key code for the desired key.
' * @param {Boolean} [ignoreKeyUp=False] - Flag to ignore keyup simulation.
' */
Private Sub StrokeSingleKey(ByVal key As Long, Optional ByVal ignoreKeyUp As Boolean = False)
    Dim Ctrl As Boolean
    Dim Shift As Boolean
    Dim Alt As Boolean

    ' Extracting modifier key states
    Ctrl = ((key ¥ Ctrl_) And 1) = 1
    Shift = ((key ¥ Shift_) And 1) = 1
    Alt = ((key ¥ Alt_) And 1) = 1

    ' Extracting the actual key code
    key = key And &HFF

    ' Get current key status
    pHoldShiftLeft = ((GetKeyState(ShiftLeft_) And &H8000) <> 0) And Not ignoreKeyUp
    pHoldShiftRight = ((GetKeyState(ShiftRight_) And &H8000) <> 0) And Not ignoreKeyUp
    pHoldCtrlLeft = ((GetKeyState(CtrlLeft_) And &H8000) <> 0) And Not ignoreKeyUp
    pHoldCtrlRight = ((GetKeyState(CtrlRight_) And &H8000) <> 0) And Not ignoreKeyUp
    pHoldAltLeft = ((GetKeyState(AltLeft_) And &H8000) <> 0) And Not ignoreKeyUp
    pHoldAltRight = ((GetKeyState(AltRight_) And &H8000) <> 0) And Not ignoreKeyUp

    ' Simulating keydown for modifier keys
    If Alt Then
        keybd_event AltLeft_, 0, 0, 0
    ElseIf Not ignoreKeyUp Then
        If pHoldAltRight Then keybd_event AltRight_, 0, EXTENDED_KEY Or KEYUP, 0
        If pHoldAltLeft Then keybd_event AltLeft_, 0, KEYUP, 0
    End If

    If Ctrl Then
        keybd_event CtrlLeft_, 0, 0, 0
    ElseIf Not ignoreKeyUp Then
        If pHoldCtrlRight Then keybd_event CtrlRight_, 0, EXTENDED_KEY Or KEYUP, 0
        If pHoldCtrlLeft Then keybd_event CtrlLeft_, 0, KEYUP, 0
    End If

    If Shift Then
        keybd_event vbKeyShift, 0, EXTENDED_KEY, 0
    ElseIf Not ignoreKeyUp Then
        If pHoldShiftRight Then keybd_event ShiftRight_, 0, EXTENDED_KEY Or KEYUP, 0
        If pHoldShiftLeft Then keybd_event ShiftLeft_, 0, KEYUP, 0
    End If

    ' Simulating key stroke for the specified key
    keybd_event key, 0, 0, 0
    keybd_event key, 0, KEYUP, 0

    ' Simulating keyup for modifier keys
    If pHoldAltLeft Then keybd_event AltLeft_, 0, 0, 0
    If pHoldAltRight Then keybd_event AltRight_, 0, 0, 0
    If Not pHoldAltLeft And Alt Then keybd_event AltLeft_, 0, KEYUP, 0

    If pHoldShiftLeft Then keybd_event ShiftLeft_, 0, 0, 0
    If pHoldShiftRight Then keybd_event ShiftRight_, 0, 0, 0
    If Not pHoldShiftLeft And Shift Then keybd_event ShiftLeft_, 0, KEYUP, 0
    If Not pHoldShiftRight And Shift Then keybd_event ShiftRight_, 0, KEYUP, 0

    If pHoldCtrlLeft Then keybd_event CtrlLeft_, 0, 0, 0
    If pHoldCtrlRight Then keybd_event CtrlRight_, 0, 0, 0
    If Not pHoldCtrlLeft And Ctrl Then keybd_event CtrlLeft_, 0, KEYUP, 0

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
        Call StrokeSingleKey(strokeKeys(i), ignoreKeyUp:=True)
    Next i
End Sub

'/*
' * Press the specified keys in order. (Release Ctrl/Alt keys)
' * Note: Not suitable for keys that need to be held down
' *
' * @param {eKey, ...} strokeKeys - An array of key codes for the desired keys.
' * @example: To press Alt + H and then I
' *     Call KeyStroke(Alt_ + H_, I_)
' */
Sub KeyStroke(ParamArray strokeKeys() As Variant)
    Dim i As Long

    For i = LBound(strokeKeys) To UBound(strokeKeys)
        Call StrokeSingleKey(strokeKeys(i))
    Next i
End Sub

'/*
' * Simulates keyup events for various control keys.
' */
Sub KeyUpControlKeys()
    pHoldCtrlLeft = False
    pHoldCtrlRight = False
    pHoldShiftLeft = False
    pHoldShiftRight = False
    pHoldAltLeft = False
    pHoldAltRight = False

    If (GetKeyState(ShiftLeft_) And &H8000) <> 0 Then
        keybd_event ShiftLeft_, 0, KEYUP, 0
        pHoldShiftLeft = True
    End If
    If (GetKeyState(ShiftRight_) And &H8000) <> 0 Then
        keybd_event ShiftRight_, 0, EXTENDED_KEY Or KEYUP, 0
        pHoldShiftRight = True
    End If
    If (GetKeyState(CtrlLeft_) And &H8000) <> 0 Then
        keybd_event CtrlLeft_, 0, KEYUP, 0
        pHoldCtrlLeft = True
    End If
    If (GetKeyState(CtrlRight_) And &H8000) <> 0 Then
        keybd_event CtrlRight_, 0, EXTENDED_KEY Or KEYUP, 0
        pHoldCtrlRight = True
    End If
    If (GetKeyState(AltLeft_) And &H8000) <> 0 Then
        keybd_event AltLeft_, 0, KEYUP, 0
        pHoldAltLeft = True
    End If
    If (GetKeyState(AltRight_) And &H8000) <> 0 Then
        keybd_event AltRight_, 0, EXTENDED_KEY Or KEYUP, 0
        pHoldAltRight = True
    End If
End Sub

'/*
' * Simulates keyup events for control keys based on their current state.
' */
Sub UnkeyUpControlKeys()
    If pHoldShiftLeft Then keybd_event ShiftLeft_, 0, 0, 0
    If pHoldShiftRight Then keybd_event ShiftRight_, 0, EXTENDED_KEY, 0
    If pHoldCtrlLeft Then keybd_event CtrlLeft_, 0, 0, 0
    If pHoldCtrlRight Then keybd_event CtrlRight_, 0, EXTENDED_KEY, 0
    If pHoldAltLeft Then keybd_event AltLeft_, 0, 0, 0
    If pHoldAltRight Then keybd_event AltRight_, 0, EXTENDED_KEY, 0

    pHoldCtrlLeft = False
    pHoldCtrlRight = False
    pHoldShiftLeft = False
    pHoldShiftRight = False
    pHoldAltLeft = False
    pHoldAltRight = False
End Sub
