Attribute VB_Name = "C_Common"
Option Explicit
Option Private Module

' Repeater
Private pSavedFuncName As String
Private pSavedCount As Long
Private pSavedArgs As Variant

Sub RepeatRegister(ByVal funcName As String, ParamArray args() As Variant)
    ' Store values in module variables
    pSavedFuncName = funcName
    pSavedCount = gVim.Count
    pSavedArgs = args
End Sub

Function RepeatAction(Optional ByVal g As String) As Boolean
    ' Restore g:count
    gVim.Count = pSavedCount

    Select Case UBound(pSavedArgs)
        Case -1
            RepeatAction = Application.Run(pSavedFuncName)
        Case 0
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0))
        Case 1
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1))
        Case 2
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2))
        Case 3
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3))
        Case 4
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3), pSavedArgs(4))
        Case 5
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3), pSavedArgs(4), pSavedArgs(5))
        Case 6
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3), pSavedArgs(4), pSavedArgs(5), pSavedArgs(6))
        Case 7
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3), pSavedArgs(4), pSavedArgs(5), pSavedArgs(6), pSavedArgs(7))
        Case 8
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3), pSavedArgs(4), pSavedArgs(5), pSavedArgs(6), pSavedArgs(7), pSavedArgs(8))
        Case 9
            RepeatAction = Application.Run(pSavedFuncName, pSavedArgs(0), pSavedArgs(1), pSavedArgs(2), pSavedArgs(3), pSavedArgs(4), pSavedArgs(5), pSavedArgs(6), pSavedArgs(7), pSavedArgs(8), pSavedArgs(9))
        Case Else
            ' Error if argument is more than 10
            Call DebugPrint("Too many arguments", pSavedFuncName & " in RepeatAction")
    End Select

    ' Reset g:count after execution
    gVim.Count = 0
End Function

Function disableIME()
    Select Case IMEStatus
        Case vbIMEModeOn, Is > 3
            Call keystrokeWithoutKeyup(Kanji_)
    End Select
End Function
