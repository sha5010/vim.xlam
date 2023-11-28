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

'/*
' * Jumps to the next or previous position in the jump list.
' *
' * @param {Boolean} isNext - True to jump to the next position, False to jump to the previous position.
' */
Private Sub JumpInner(ByVal isNext As Boolean)
    On Error GoTo Catch

    ' Check if the jump list is available
    If Not gVim.JumpList Is Nothing Then
        Dim isAfterJumped As Boolean
        isAfterJumped = (TypeOf gVim.JumpList.Current Is Range And TypeOf Selection Is Range)

        If isAfterJumped Then
            Dim i As Long

            For i = 1 To 3
                Select Case i
                Case 1
                    isAfterJumped = (gVim.JumpList.Current.Parent.Parent Is Selection.Parent.Parent)
                Case 2
                    isAfterJumped = (gVim.JumpList.Current.Parent Is Selection.Parent)
                Case 3
                    isAfterJumped = (gVim.JumpList.Current.Address = Selection.Address)
                End Select
                If Not isAfterJumped Then
                    Exit For
                End If
            Next i
        End If

        Dim targetRange As Range
        ' Get the next or previous target range from the jump list
        If isNext Then
            Set targetRange = gVim.JumpList.Forward
        Else
            Set targetRange = gVim.JumpList.Back
        End If

        ' Check if the target range is not empty
        If Not targetRange Is Nothing Then
            ' Stop visual mode if active
            Call StopVisualMode

            ' Record the current position to the jump list if it's the latest position
            If Not isAfterJumped Then
                Call RecordToJumpList(CurrentToLatest:=False)
            End If

            Dim targetWorkbook As Workbook
            Dim targetWorksheet As Worksheet
            ' Get the workbook and worksheet from the target range
            Set targetWorkbook = targetRange.Parent.Parent
            Set targetWorksheet = targetRange.Parent

            ' Activate the target workbook and worksheet, and select the target range
            targetWorkbook.Activate
            targetWorksheet.Activate
            targetRange.Select
        Else
            ' Display a status message for reaching the latest or oldest position
            Dim statusMessage As String
            If isNext Then
                statusMessage = gVim.Msg.LatestJumplist
            Else
                statusMessage = gVim.Msg.OldestJumplist
            End If
            Call SetStatusBarTemporarily(statusMessage, 1000)
        End If
    End If
    Exit Sub

Catch:
    ' Handle errors and call the error handler
    Call ErrorHandler("JumpInner")
End Sub

Function JumpPrev(Optional ByVal g As String) As Boolean
    Call JumpInner(isNext:=False)
End Function

Function JumpNext(Optional ByVal g As String) As Boolean
    Call JumpInner(isNext:=True)
End Function

Function ClearJumps(Optional ByVal g As String) As Boolean
    If Not gVim.JumpList Is Nothing Then
        Call gVim.JumpList.ClearAll
        Call SetStatusBarTemporarily(gVim.Msg.ClearedJumplist, 2000)
    End If
End Function

'/*
' * Records the current or specified target range to the jump list.
' *
' * @param {Range} [Target] - The target range to add to the jump list. If not specified, uses the current selection or active cell.
' * @param {Boolean} [CurrentToLatest=True] - True to update the jump list from the current position to the latest position.
' * @returns {Boolean} - Always returns True.
' */
Function RecordToJumpList(Optional Target As Range, Optional ByVal CurrentToLatest As Boolean = True) As Boolean
    On Error GoTo Catch

    ' Verify if the jump list is available
    If gVim.JumpList Is Nothing Then
        Exit Function
    End If

    ' If Target is not specified, use the current selection or active cell
    If Target Is Nothing Then
        If TypeName(Selection) = "Range" Then
            Set Target = Selection
        ElseIf Not ActiveCell Is Nothing Then
            Set Target = ActiveCell
        Else
            Exit Function
        End If
    End If

    ' Add the target range to the jump list
    Call gVim.JumpList.Add(Target, CurrentToLatest)
    RecordToJumpList = True
    Exit Function

Catch:
    ' Handle errors and call the error handler
    Call ErrorHandler("RecordToJumpList")
End Function

Sub DisableIME()
    Select Case IMEStatus
        Case Is > 3, vbIMEHiragana
            Call KeyStrokeWithoutKeyup(IME_On_)
    End Select
End Sub
