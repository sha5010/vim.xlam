VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Cmd 
   Caption         =   "Command"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1590
   OleObjectBlob   =   "Cmd.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cUserForm As cls_UFKeyReceiver
Attribute cUserForm.VB_VarHelpID = -1

Private cCmdBuffer As String
Private cCmdBufferWithoutNumber As String
Private cCount As Long
Private cNumSeqFlag As Boolean

'/*
' * Handles key press events with SendKeys and performs associated actions.
' *
' * @param {String} key - The pressed key.
' */
Private Sub cUserForm_KeyPressWithSendKeys(ByVal key As String)
    ' Process when a number is entered on the Numpad, if NumpadCount option is enabled
    If gVim.Config.NumpadCount And InStr(key, "{") = 1 Then
        Dim keyValue As String
        keyValue = Left(Mid(key, 2), Len(key) - 2)
        If Not keyValue Like "*[!0-9]*" Then
            Select Case CLng(keyValue)
                Case 96 To 105
                    key = CLng(keyValue) - 96
            End Select
        End If
    End If

    ' Check if the pressed key is a single digit and a numeric sequence is ongoing
    If Len(key) = 1 And (cCount = 0 Or cNumSeqFlag) Then
        cNumSeqFlag = True
        If cCount < LONG_MAX / 10 - 1 Then
            cCount = cCount * 10 + CLng(key)
        Else
            cCount = LONG_MAX
        End If
    Else
        cNumSeqFlag = False
    End If

    ' Handle the key based on the current command buffer
    If Len(cCmdBuffer) > 0 Then
        cCmdBuffer = cCmdBuffer & KEY_SEPARATOR & key
        If Not cNumSeqFlag Then
            cCmdBufferWithoutNumber = cCmdBufferWithoutNumber & KEY_SEPARATOR & key
        End If
    ElseIf Not cNumSeqFlag And Len(cCmdBuffer) = 0 Then
        cCmdBuffer = key
        cCmdBufferWithoutNumber = key
    End If

    ' Exit if the command buffer is empty
    If Len(cCmdBuffer) = 0 Then
        Exit Sub
    End If

    ' Try to run the command
    Dim anyCommandExecuted As Boolean
    anyCommandExecuted = Try()

    ' Quit if no command available
    If Not anyCommandExecuted Then
        Dim valid As Boolean
        valid = gVim.KeyMap.IsStillValid(cCmdBuffer)
        If Not valid And cCmdBuffer <> cCmdBufferWithoutNumber Then
            valid = gVim.KeyMap.IsStillValid(cCmdBufferWithoutNumber)
        End If

    ' Quit if command was executed successfully
    ElseIf cCmdBuffer = "" And cCount = 0 Then
        Exit Sub
    Else
        valid = True
    End If

    ' Handle pre-defined cancel keys
    Select Case key
    Case "{ESC}", "^{[}", "^{c}"
        Call QuitForm
        Exit Sub
    End Select

    ' There is no available commands
    If Not valid Then
        Call QuitForm

        ' Display a status message for debugging if enabled
        If gVim.DebugMode Then
            Call SetStatusBarTemporarily(gVim.Msg.NoKeyAllocation, 2000)
        End If

        Exit Sub
    End If

    Call LazySuggest

    ' Redisplay if my form is invisible
    If Not Me.Visible Then
        Me.Show
    End If
End Sub

'/*
' * Tries to run a command and handles the result.
' *
' * @returns {Boolean} - True if any command was successfully executed, False otherwise.
' */
Private Function Try() As Boolean
    Dim isSucceeded As Boolean
    Dim cmd As String

    ' Try to run the command from the key map based on the current command buffer
    cmd = gVim.KeyMap.Get_(cCmdBuffer)
    If IsCommandAvailable(cmd) Then
        Try = True
        isSucceeded = Run(cmd, ignoreCount:=(cCmdBuffer <> cCmdBufferWithoutNumber))
        If isSucceeded Then
            Call QuitForm
            Exit Function
        End If
    End If

    ' If the current command buffer is different from the one without the number,
    ' try to run the command without the number from the key map
    cmd = gVim.KeyMap.Get_(cCmdBufferWithoutNumber)
    If cCmdBuffer <> cCmdBufferWithoutNumber And IsCommandAvailable(cmd) Then
        Try = True
        isSucceeded = Run(cmd)
        If isSucceeded Then
            Call QuitForm
            Exit Function
        End If
    End If

    ' If the current command buffer contains multiple keys separated by KEY_SEPARATOR,
    ' try to run the command by gradually removing the last key and checking at each step
    Dim checkCmd As String: checkCmd = cCmdBuffer
    Dim argStr As String: argStr = ""
    Dim sepIndex As Long: sepIndex = InStrRev(checkCmd, KEY_SEPARATOR)

    Do While sepIndex > 1
        ' Extract the argument string from the remaining part of the command buffer
        argStr = gVim.KeyMap.SendKeysToDisplayText(Mid(checkCmd, sepIndex + Len(KEY_SEPARATOR))) & argStr
        ' Remove the last key from the command buffer
        checkCmd = Left(checkCmd, sepIndex - 1)

        ' Try to run the command from the key map
        cmd = gVim.KeyMap.Get_(checkCmd)
        If IsCommandAvailable(cmd) Then
            Try = True
            ' Run the command with the accumulated argument string
            isSucceeded = Run(cmd, argStr, ignoreCount:=(Len(checkCmd) > Len(cCmdBufferWithoutNumber)))
            If isSucceeded Then
                Call QuitForm
                Exit Function
            End If
        End If

        ' Find the index of the previous KEY_SEPARATOR in the remaining command buffer
        sepIndex = InStrRev(checkCmd, KEY_SEPARATOR)
    Loop
End Function

'/*
' * Runs a command with optional arguments and handles the result.
' *
' * @param {String} cmd - The command to run.
' * @param {String} [arg=""] - The optional arguments for the command.
' * @param {Boolean} [ignoreCount=False] - Flag to ignore the current count.
' * @returns {Boolean} - True if the command was successfully run, False otherwise.
' */
Private Function Run(ByVal cmd As String, _
            Optional ByVal arg As String = "", _
            Optional ByVal ignoreCount As Boolean = False) As Boolean

    Dim result As Variant: result = True
    If Not ignoreCount Then
        gVim.Count = cCount
    Else
        gVim.Count = 0
    End If

    ' Hide the form and run the command
    Me.Hide
    On Error GoTo Catch
    If Len(arg) > 0 Then
        result = Application.Run(cmd, arg)
    Else
        result = Application.Run(cmd)
    End If

    ' Close form if command was successful
    If Not result Then
        Run = True

        ' Reset vars manually because form has already closed
        cCount = 0
        cCmdBuffer = ""
        cCmdBufferWithoutNumber = ""
        cNumSeqFlag = False
    End If

    gVim.Count = 0
    Exit Function

Catch:
    ' Handle the case where the macro associated with the command is missing
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.MissingMacro & cmd, 3000)
    Else
        ' Clear the error and resume execution
        Err.Clear
        Resume Next
    End If
End Function

'/*
' * Checks if a command is available for execution.
' *
' * @param {String} cmd - The command to check.
' * @returns {Boolean} - True if the command is available, False otherwise.
' */
Private Function IsCommandAvailable(ByVal cmd As String) As Boolean
    IsCommandAvailable = (cmd <> DUMMY_PROCEDURE And Not cmd Like "'" & SHOW_CMD_PROCEDURE & " ""*")
End Function

'/*
' * Handles key press events with a string and appends the string to the label text.
' *
' * @param {String} str - The pressed key.
' */
Private Sub cUserForm_KeyPressWithString(ByVal str As String)
    If gVim.Config.NumpadCount And str Like "<k[0-9]>" Then
        Me.Label_Text = Me.Label_Text & Mid(str, 3, 1)
    Else
        Me.Label_Text = Me.Label_Text & str
    End If
    Me.Width = Me.Label_Text.Left + Me.Label_Text.Width + 12
End Sub

'/*
' * Initializes the user form and sets up its properties.
' */
Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Top = Application.Top + Application.Height - Me.Height - 4
        .Left = Application.Left + 4
    End With

    With Me.Label_Text
        .WordWrap = False
        .AutoSize = True
    End With

    Set cUserForm = New cls_UFKeyReceiver
    Set cUserForm.Form = Me
End Sub

'/*
' * Activates the user form and sets its position.
' */
Private Sub UserForm_Activate()
    Me.Move Application.Left + 4, Application.Top + Application.Height - Me.Height - 4
    Call LazySuggest
End Sub

'/*
' * Resets the form-related variables.
' */
Private Sub ResetVars()
    cCount = 0
    gVim.Count = 0
    cCmdBuffer = ""
    cCmdBufferWithoutNumber = ""
    cNumSeqFlag = False
End Sub

'/*
' * Quits the form and resets related variables.
' */
Private Sub QuitForm()
    Call ResetVars
    Me.Hide
End Sub

'/*
' * Launches the user form with the provided prefix key.
' *
' * @param {String} prefixKey - The prefix key to display on the form.
' * @returns {Boolean} - True if the form was launched, False otherwise.
' */
Public Function Launch(ByVal prefixKey As String) As Boolean
    If Not (cCount = 0 And cCmdBuffer = "") Then
        Launch = True
        Exit Function
    End If

    ' Set the label text with the prefix key
    Me.Label_Text = gVim.KeyMap.SendKeysToDisplayText(prefixKey)
    Me.Width = 84

    ' Determine if the prefix key is a single character or a numeric sequence
    If Len(prefixKey) > 1 Then   ' [^0-9]
        cCmdBuffer = prefixKey
        cCmdBufferWithoutNumber = prefixKey
    ElseIf Asc(prefixKey) > 48 Then  ' [1-9]
        cCount = CLng(prefixKey)
        cNumSeqFlag = True
    End If

    ' Show the form
    Me.Show
End Function

Private Sub LazySuggest()
    Static lastRegisterTime
    Static lastRegisterProc

    ' Try to cancel the previous OnTime event
    On Error Resume Next
    Call Application.OnTime(lastRegisterTime, lastRegisterProc, , False)
    On Error GoTo 0

    ' Disabled if SuggestWait is less than 0
    If gVim.Config.SuggestWait <= 0 Then
        Exit Sub
    End If

    ' Calculate the time for the next OnTime event
    lastRegisterTime = Date + CDec(Timer + gVim.Config.SuggestWait / 1000) / 86400

    ' Set procedure name
    If gVim.KeyMap.IsStillValid(cCmdBuffer) Then
        lastRegisterProc = "'ShowSuggest """ & cCmdBuffer & """'"
    ElseIf gVim.KeyMap.IsStillValid(cCmdBufferWithoutNumber) Then
        lastRegisterProc = "'ShowSuggest """ & cCmdBufferWithoutNumber & """'"
    Else
        ' TODO: ArgumentSuggest
        Exit Sub
    End If

    ' Register the next OnTime event
    Call Application.OnTime(lastRegisterTime, lastRegisterProc)
End Sub

Public Sub ReceiveKey(ByVal key As String)
    Call cUserForm_KeyPressWithString(gVim.KeyMap.SendKeysToDisplayText(key))
    Call cUserForm_KeyPressWithSendKeys(key)
End Sub
