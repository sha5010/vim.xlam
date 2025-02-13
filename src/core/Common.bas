Attribute VB_Name = "C_Common"
Option Explicit
Option Private Module

#If Win64 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function MonitorFromRect Lib "user32" (ByRef lpRect As RECT, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpmi As monitorInfo) As Long

#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Private Declare Function GetDpiForWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function MonitorFromRect Lib "user32" (ByRef lpRect As RECT, ByVal dwFlags As Long) As LongPtr
    Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As LongPtr, ByRef lpmi As monitorInfo) As Long
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type monitorInfo
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

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
    If pSavedFuncName = "" Then
        Exit Function
    End If

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
        Dim i As Long
        Dim isAfterJumped As Boolean
        isAfterJumped = (TypeOf gVim.JumpList.Current Is Range And TypeOf Selection Is Range)

        If isAfterJumped Then

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

        For i = 1 To gVim.Count1
            ' Get the next or previous target range from the jump list
            If isNext Then
                Set targetRange = gVim.JumpList.Forward
            Else
                Set targetRange = gVim.JumpList.Back
            End If

            If targetRange Is Nothing Then
                If i > 1 Then
                    Set targetRange = gVim.JumpList.Current
                End If
                Exit For
            End If
        Next i

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

Function SearchHelp(Optional ByVal key As String = "") As Boolean
    On Error GoTo Catch

    If key = "" Then
        Call SetStatusBarTemporarily(gVim.Msg.ArgumentsRequired, 3000)
        Exit Function
    End If

    Dim cmd As String
    If InStr(key, ":") = 1 Then
        Dim suggestsList() As String
        suggestsList = gVim.KeyMap.Suggest(Mid(key, 2), True)
        If UBound(suggestsList) = 0 Then
            cmd = gVim.KeyMap.Get_(suggestsList(0), True)
            key = ":" & suggestsList(0)
        End If
    Else
        cmd = gVim.KeyMap.Get_(gVim.KeyMap.VimToVBA(key, KEY_SEPARATOR))
    End If

    Dim helpText As String
    If cmd = DUMMY_PROCEDURE Then
        cmd = ""
        helpText = gVim.Msg.NoKeyAllocation & key
    ElseIf cmd = "" Then
        helpText = gVim.Msg.NoCommandAvailable & Mid(key, 2)
    Else
        helpText = gVim.Help.GetText(cmd)

        If helpText = cmd Then
            helpText = gVim.Msg.NoCommandHelp & key
        Else
            helpText = key & "    " & cmd & "    " & helpText
        End If
    End If

    Call SetStatusBarTemporarily(helpText, 3000)
    Exit Function
Catch:
    ' Handle errors and call the error handler
    Call ErrorHandler("SearchHelp")
End Function

Private Function CmdSuggest(ByVal key As String) As CommandBar
    If key = "" Then
        Exit Function
    End If

    Dim suggestsList() As String
    suggestsList = gVim.KeyMap.Suggest(key)

    If UBound(suggestsList) < 0 Then
        Exit Function
    End If

    Dim menuDict As Dictionary
    Set menuDict = New Dictionary

    Dim i As Long
    For i = LBound(suggestsList) To UBound(suggestsList)
        Dim nextChar As String
        nextChar = Replace(suggestsList(i), key & KEY_SEPARATOR, "", Count:=1)

        If InStr(nextChar, KEY_SEPARATOR) > 0 Then
            nextChar = Split(nextChar, KEY_SEPARATOR, 2)(0)
            If Not menuDict.Exists(nextChar) Then
                menuDict.Add nextChar, "  + more"
            End If
        Else
            If Not menuDict.Exists(nextChar) Then
                menuDict.Add nextChar, gVim.Help.GetText(gVim.KeyMap.Get_(suggestsList(i)))
            Else
                menuDict(nextChar) = gVim.Help.GetText(gVim.KeyMap.Get_(suggestsList(i)))
            End If
        End If
    Next i

    Set CmdSuggest = Application.CommandBars.Add(position:=msoBarPopup, Temporary:=True)
    Dim k
    For Each k In menuDict.Keys()
        With CmdSuggest.Controls.Add(Type:=msoControlButton)
            .Caption = UF_Cmd.Label_Text.Caption & "&" & gVim.KeyMap.SendKeysToDisplayText(k) & "    " & menuDict(k)
            .OnAction = "'CompleteSuggest """ & k & """'"
        End With
    Next k
End Function

Private Function CmdlineSuggest(ByVal key As String) As CommandBar
    Dim suggestsList() As String
    suggestsList = gVim.KeyMap.Suggest(key, True)

    If UBound(suggestsList) < 0 Then
        Exit Function
    ElseIf UBound(suggestsList) = 0 Then
        UF_CmdLine.TextBox.Text = suggestsList(0)
        Exit Function
    End If

    Set CmdlineSuggest = Application.CommandBars.Add(position:=msoBarPopup, Temporary:=True)

    Dim i As Long
    For i = LBound(suggestsList) To UBound(suggestsList)
        With CmdlineSuggest.Controls.Add(Type:=msoControlButton)
            Dim labelCaption As String
            If i < Len(gVim.Config.SuggestLabels) Then
                labelCaption = "(&" & Mid(gVim.Config.SuggestLabels, i + 1, 1) & ")  "
            Else
                labelCaption = "      "
            End If
            .Caption = labelCaption & suggestsList(i) & String(Int(32 - Len(suggestsList(i)) * 2), ChrW(&H2005)) & _
                gVim.Help.GetText(gVim.KeyMap.Get_(suggestsList(i), True))
            .OnAction = "'CompleteSuggest """ & suggestsList(i) & """'"
        End With
    Next i
End Function

Function PathSuggest(ByVal cmd As String, ByVal basePath As String) As CommandBar
    basePath = ResolvePath(basePath)

    Dim childItems As Collection
    Set childItems = DirGrob(basePath)

    If childItems.Count = 0 Then
        Exit Function
    End If

    cmd = Left(cmd, InStrRev(Replace(cmd, "/", "¥"), "¥"))
    If childItems.Count = 1 Then
        Call CompleteSuggest(cmd & childItems(1))
        Exit Function
    End If

    Set PathSuggest = Application.CommandBars.Add(position:=msoBarPopup, Temporary:=True)
    With PathSuggest.Controls.Add(Type:=msoControlButton)
        .Enabled = False
        .Caption = basePath
    End With

    Dim childItem
    Dim i As Long: i = 0
    Dim isDirEnded As Boolean
    For Each childItem In childItems
        With PathSuggest.Controls.Add(Type:=msoControlButton)
            Dim labelCaption As String
            If i < Len(gVim.Config.SuggestLabels) Then
                labelCaption = "(&" & Mid(gVim.Config.SuggestLabels, i + 1, 1) & ")  "
            Else
                labelCaption = "      "
            End If
            .Caption = labelCaption & childItem
            .OnAction = "'CompleteSuggest """ & cmd & childItem & """'"
            If (i = 0) Or (Not isDirEnded And Right(childItem, 1) <> "/") Then
                .BeginGroup = True
            End If
        End With
        isDirEnded = (Right(childItem, 1) <> "/")
        i = i + 1

        If i >= 50 And childItems.Count > 51 Then
            With PathSuggest.Controls.Add(Type:=msoControlButton)
                .Enabled = False
                .Caption = CStr(childItems.Count - 50) & gVim.Msg.RemainingResults
                .BeginGroup = True
            End With
            Exit For
        End If
    Next
End Function

Function ShowSuggest(Optional ByVal key As String = "") As Boolean
    Dim formCaption As String
    Dim tmpMenu As CommandBar
    Dim i As Long

    If UF_Cmd.Visible Then
        formCaption = UF_Cmd.Caption
        Set tmpMenu = CmdSuggest(key)

    ElseIf UF_CmdLine.Visible And UF_CmdLine.Label_Prefix.Caption = ":" Then
        key = UF_CmdLine.TextBox.Text
        formCaption = UF_CmdLine.Caption
        Set tmpMenu = CmdlineSuggest(key)

        If tmpMenu Is Nothing Then
            If InStr(key, " ") = 0 Then
                Exit Function
            End If

            Dim secondPart As String
            secondPart = Replace(Split(key, " ", 2)(1), "/", "¥")
            If Not StartsWith(secondPart, Array(".¥", "..¥", "‾¥", "¥")) Then
                Exit Function
            End If

            Set tmpMenu = PathSuggest(key, secondPart)
        End If
    End If

    If tmpMenu Is Nothing Then
        Exit Function
    End If

    Call ShowPopupMenu(tmpMenu, formCaption)
End Function

'/*
' * Displays a popup menu near the specified window.
' *
' * @param {CommandBar} popupMenu - The popup menu to be displayed.
' * @param {String} winCaption - The caption of the window to find.
' */
Sub ShowPopupMenu(ByRef popupMenu As CommandBar, ByVal winCaption As String)
    ' Declare variable for window handle
    Dim hWnd As LongPtr
    hWnd = FindWindowA(vbNullString, winCaption)

    ' If the window is not found, exit the subroutine
    If hWnd = 0 Then
        Exit Sub
    End If

    ' Declare variable for window rectangle
    Dim formRect As RECT
    GetWindowRect hWnd, formRect

    ' Obtain the DPI scale factor for the window
    Dim dpi As Long
    dpi = GetDpiForWindow(hWnd)
    Dim scaleFactor As Single
    scaleFactor = dpi / 96  ' Calculate scale factor based on 96 DPI

    ' Declare variable for monitor handle
    Dim monitorHandle As LongPtr
    monitorHandle = MonitorFromRect(formRect, 2)  ' 2: MONITOR_DEFAULTTONEAREST

    ' Declare variable for monitor information
    Dim monitorInfo As monitorInfo
    monitorInfo.cbSize = Len(monitorInfo)
    GetMonitorInfo monitorHandle, monitorInfo

    ' Calculate the adjusted position of the popup menu
    Dim adjustedLeft As Long
    Dim adjustedTop As Long
    adjustedLeft = formRect.Left * scaleFactor
    adjustedTop = (formRect.Top + (formRect.Bottom - formRect.Top) / 2) * scaleFactor - popupMenu.Height

    ' Adjust top position if the popup is out of the monitor's bounds
    If adjustedTop < monitorInfo.rcMonitor.Top Then
        adjustedTop = monitorInfo.rcMonitor.Top  ' Adjust to the bottom if too high
    End If

    ' Adjust left position if the popup is out of the monitor's bounds
    If adjustedLeft < monitorInfo.rcMonitor.Left Then
        adjustedLeft = monitorInfo.rcMonitor.Left  ' Adjust to the right if too far left
    End If

    ' Show the popup menu at the adjusted position
    popupMenu.ShowPopup adjustedLeft, adjustedTop
End Sub

'/*
' * Completes the suggestion by passing the key to the relevant control.
' *
' * @param {String} key - The key to be passed to the suggestion control.
' */
Sub CompleteSuggest(ByVal key As String)
    ' Check if the command interface (UF_Cmd) is visible
    If UF_Cmd.Visible Then
        ' Pass the key to UF_Cmd control
        Call UF_Cmd.ReceiveKey(key)

    ' Check if the command line interface (UF_CmdLine) is visible
    ElseIf UF_CmdLine.Visible Then
        ' Set the key text in the UF_CmdLine TextBox
        UF_CmdLine.TextBox.Text = key
    End If
End Sub
