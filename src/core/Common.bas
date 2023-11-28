Attribute VB_Name = "C_Common"
Option Explicit
Option Private Module


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
