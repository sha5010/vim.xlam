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
    Select Case IMEStatus
        Case vbIMEModeOn, Is > 3
            Call keystrokeWithoutKeyup(Kanji_)
    End Select
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
