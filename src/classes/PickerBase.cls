VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_PickerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*
' * The PickerBase class provides fundamental functionalities for a selection UI (Picker).
' * It acts as a virtual base class for concrete Picker classes implementing the IPicker interface,
' * encapsulating common logic such as keybindings, list operations, and help display.
' */

Private WithEvents mUserForm As cls_UFKeyReceiver   ' Instance of cls_UFKeyReceiver, handles key events from the user form.
Attribute mUserForm.VB_VarHelpID = -1
Private mStatusLabel As MSForms.Label               ' Label control for displaying status messages.
Private mKeybindings As Dictionary                  ' Dictionary storing keybindings (mapping of key inputs to PickerAction objects).
Private mTextboxMap As Dictionary                   ' Dictionary storing keybindings specific to the textbox.
Private mOwner As cls_IPicker                       ' Reference to the Picker (a class implementing IPicker interface) that uses this base class.

Private Const AMOUNT As Long = 3            ' Number of rows to move for jump actions.
Private Const VISIBLE_ROWS As Long = 22     ' Number of visible rows in the picker's listbox.

'/*
' * Defines core actions that can be performed in the Picker.
' * These actions include list navigation, item selection, displaying help, and closing the Picker.
' */
Private Enum CoreAction
    MoveUp = -1337 ' Move selection up
    MoveDown       ' Move selection down
    JumpUp         ' Jump selection up by a specified amount
    JumpDown       ' Jump selection down by a specified amount
    ScrollUp       ' Scroll up one page
    ScrollDown     ' Scroll down one page
    ScrollUpHalf   ' Scroll up half a page
    ScrollDownHalf ' Scroll down half a page
    MoveToTop      ' Move selection to the top
    MoveToBottom   ' Move selection to the bottom
    FocusSearch    ' Focus the search box

    SelectItem     ' Select the current item
    ShowHelp       ' Display help message
    ClosePicker    ' Close the picker
End Enum

'/*
' * Defines actions specific to interaction within the textbox.
' * These actions include navigation, deletion, item selection, etc., within the textbox.
' */
Private Enum TextboxAction
    MoveUp         ' Move cursor up
    MoveDown       ' Move cursor down
    ScrollUp       ' Scroll up
    ScrollDown     ' Scroll down
    DeleteWord     ' Delete previous word
    DeleteBackward ' Delete backward to beginning of line

    SelectItem     ' Select item
    ShowHelp       ' Show help
    ClosePicker    ' Close picker
End Enum

'/*
' * Initializes the PickerBase object.
' * Sets up the user form event receiver and declares base keybindings.
' *
' * @param {cls_IPicker} owner - The instance of the IPicker implementation class using this base class.
' */
Public Sub Init(ByRef owner As cls_IPicker)
    Set mUserForm = New cls_UFKeyReceiver
    Set mKeybindings = New Dictionary
    Set mTextboxMap = New Dictionary
    Set mOwner = owner

    '/* Declare base actions */
    Call DeclareAction(Move, "j", "カーソルを下へ", "Move down", CoreAction.MoveDown)
    Call DeclareAction(Move, "k", "カーソルを上へ", "Move up", CoreAction.MoveUp)
    Call DeclareAction(Move, "<C-j>", "カーソルを" & AMOUNT & "行下へ", "Move down " & AMOUNT & " rows", CoreAction.JumpDown)
    Call DeclareAction(Move, "<C-k>", "カーソルを" & AMOUNT & "行上へ", "Move up " & AMOUNT & " rows", CoreAction.JumpUp)
    Call DeclareAction(Move, "g", "カーソルを一番上へ", "Move to top", CoreAction.MoveToTop)
    Call DeclareAction(Move, "G", "カーソルを一番下へ", "Move to bottom", CoreAction.MoveToBottom)
    Call DeclareAction(Move, "<C-f>", "1ページ下へ", "Scroll down a page", CoreAction.ScrollDown)
    Call DeclareAction(Move, "<C-b>", "1ページ上へ", "Scroll up a page", CoreAction.ScrollUp)
    Call DeclareAction(Move, "<C-d>", "半ページ下へ", "Scroll down a half page", CoreAction.ScrollDownHalf)
    Call DeclareAction(Move, "<C-u>", "半ページ上へ", "Scroll up a half page", CoreAction.ScrollUpHalf)
    Call DeclareAction(Move, "/", "検索ボックスに移動", "Focus the search box", CoreAction.FocusSearch)

    Call DeclareAction(Enter, "<CR>", "確定", "Select", CoreAction.SelectItem)
    Call DeclareAction(Enter, "[a-z0-9]", "プレフィックスから選択", "Select with prefix", xlNone, helpOnly:=True)

    Call DeclareAction(NotDisplayed, "?", "ヘルプを表示", "Show help", CoreAction.ShowHelp)
    Call DeclareAction(NotDisplayed, "<Esc>", "Pickerを閉じる", "Close the picker", CoreAction.ClosePicker)
    Call DeclareAction(NotDisplayed, "<C-[>", "Pickerを閉じる", "Close the picker", CoreAction.ClosePicker)
    Call DeclareAction(NotDisplayed, "<Tab>", "検索ボックスに移動", "Focus the search box", CoreAction.FocusSearch)
    Call DeclareAction(NotDisplayed, "<S-Tab>", "検索ボックスに移動", "Focus the search box", CoreAction.FocusSearch)

    '/* Textbox keymapping */
    Call KeymapTextbox("<C-j>", TextboxAction.MoveDown)
    Call KeymapTextbox("<C-k>", TextboxAction.MoveUp)
    Call KeymapTextbox("<C-f>", TextboxAction.ScrollDown)
    Call KeymapTextbox("<C-b>", TextboxAction.ScrollUp)
    Call KeymapTextbox("<C-w>", TextboxAction.DeleteWord)
    Call KeymapTextbox("<C-u>", TextboxAction.DeleteBackward)
    Call KeymapTextbox("<CR>", TextboxAction.SelectItem)
    Call KeymapTextbox("<C-/>", TextboxAction.ShowHelp)
    Call KeymapTextbox("<Esc>", TextboxAction.ClosePicker)
    Call KeymapTextbox("<C-[>", TextboxAction.ClosePicker)
End Sub

'/*
' * Declares a new action and associates it with a keybinding.
' *
' * @param {eActionType} ActionType - The type of action (e.g., Move, Func, Enter).
' * @param {String} lhs - The left-hand side keybinding (Vim-style, e.g., "s" for the 's' key).
' * @param {String} helpTextJa - Japanese help text.
' * @param {String} helpTextEn - English help text.
' * @param {Variant} registerAction - The numeric code for the action to register.
' * @param {Boolean} [helpOnly=False] - If True, the action is only displayed in help messages, and no keybinding is registered.
' */
Public Sub DeclareAction(ByVal ActionType As eActionType, ByVal lhs As String, ByVal helpTextJa As String, ByVal helpTextEn As String, ByVal registerAction As Variant, Optional ByVal helpOnly As Boolean = False)
    Dim pickerAction As cls_PickerAction
    Set pickerAction = New cls_PickerAction

    If helpOnly Then
        Set mKeybindings(lhs) = pickerAction.Setup(ActionType, lhs, helpTextJa, helpTextEn, xlNone)
        Exit Sub
    End If

    Dim sendKeysStyle As String
    sendKeysStyle = gVim.KeyMap.VimToVBA(lhs, KEY_SEPARATOR, firstKeyOnly:=True)

    Set mKeybindings(sendKeysStyle) = pickerAction.Setup(ActionType, sendKeysStyle, helpTextJa, helpTextEn, CLng(registerAction))
End Sub

'/*
' * Sets up a textbox-specific keybinding.
' *
' * @param {String} lhs - The left-hand side keybinding.
' * @param {TextboxAction} registerAction - The TextboxAction enum value to register.
' */
Private Sub KeymapTextbox(ByVal lhs As String, ByVal registerAction As TextboxAction)
    Dim sendKeysStyle As String
    sendKeysStyle = gVim.KeyMap.VimToVBA(lhs, KEY_SEPARATOR, firstKeyOnly:=True)

    mTextboxMap(sendKeysStyle) = registerAction
End Sub

'/*
' * Notifies the base class that the associated UserForm has been closed.
' * This method is called from the UserForm_Terminate event of UF_Picker.
' * It is used to release the strong reference to mOwner, breaking the circular dependency.
' */
Public Sub NotifyFormClosed()
    If Not mOwner Is Nothing Then
        Set mOwner = Nothing ' Release the reference to the owner to break the circular dependency.
    End If
End Sub

'/*
' * Releases object references when the class is terminated.
' */
Private Sub Class_Terminate()
    Set mUserForm = Nothing
    Set mStatusLabel = Nothing
    Set mKeybindings = Nothing
    Set mTextboxMap = Nothing
    Set mOwner = Nothing
End Sub

'/*
' * Returns the key of the currently selected list item.
' *
' * @returns {String} - The key of the currently selected item.
' */
Public Property Get CurrentKey() As String
    With mUserForm.ListBox
        If .ListIndex < 0 Or .ListCount = 0 Then
            Exit Property
        End If

        CurrentKey = .List(.ListIndex, 1)
    End With
End Property

'/*
' * Returns a concatenated string of all single-character keys that are bound.
' * This is used for displaying a list of single-character keys in the Picker's UI.
' *
' * @returns {String} - A string of bound single-character keys.
' */
Public Property Get BoundKeyChars() As String
    Dim chars As String
    Dim boundKey As Variant

    For Each boundKey In mKeybindings.Keys()
        Dim displayKey As String
        displayKey = gVim.KeyMap.SendKeysToDisplayText(boundKey)

        If Len(displayKey) = 1 Then
            chars = chars & displayKey
        End If
    Next
    BoundKeyChars = chars
End Property

'/*
' * Attaches the Picker to specific UI elements (UserForm, ListBox, TextBox, Label).
' *
' * @param {MSForms.UserForm} objUserForm - The UserForm to attach to.
' * @param {MSForms.ListBox} objListBox - The ListBox control within the UserForm.
' * @param {MSForms.TextBox} objTextBox - The TextBox control within the UserForm.
' * @param {MSForms.Label} objLabel - The Label control for displaying status messages.
' */
Public Sub Attach(ByRef objUserForm As MSForms.UserForm, ByRef objListBox As MSForms.ListBox, ByRef objTextBox As MSForms.TextBox, ByRef objLabel As MSForms.Label)
    Set mUserForm.Form = objUserForm
    Set mUserForm.ListBox = objListBox
    Set mUserForm.TextBox = objTextBox
    Set mStatusLabel = objLabel

    Call UpdateCaption("?: Show help")
End Sub

'/*
' * Updates the caption of the Picker.
' *
' * @param {String} Message - Additional message to display in the caption.
' */
Public Sub UpdateCaption(ByVal Message As String)
    If Message <> "" Then
        UF_Picker.Caption = mOwner.Caption & " (" & Message & ")"
    Else
        UF_Picker.Caption = mOwner.Caption
    End If
End Sub

'/*
' * Regenerates the Picker's list.
' */
Public Sub RegenerateList()
    Call UF_Picker.UpdateCache
End Sub

'/*
' * Handles Picker actions based on the specified action type.
' *
' * @param {String} selectedKey - The key of the currently selected item.
' * @param {Long} ActionType - The type of action to handle (from CoreAction enum).
' * @returns {Boolean} - True if the action was handled, False otherwise.
' */
Public Function HandleAction(ByVal selectedKey As String, ByVal ActionType As Long) As Boolean
    HandleAction = True
    With mUserForm.ListBox
        Select Case ActionType
            Case CoreAction.MoveDown
                Call ListDown
            Case CoreAction.MoveUp
                Call ListUp
            Case CoreAction.JumpDown
                Call ListDown(AMOUNT)
            Case CoreAction.JumpUp
                Call ListUp(AMOUNT)
            Case CoreAction.ScrollDown
                Call ListDown(VISIBLE_ROWS, isLoop:=False)
            Case CoreAction.ScrollUp
                Call ListUp(VISIBLE_ROWS, isLoop:=False)
            Case CoreAction.ScrollDownHalf
                Call ListDown(VISIBLE_ROWS \ 2, isLoop:=False)
            Case CoreAction.ScrollUpHalf
                Call ListUp(VISIBLE_ROWS \ 2, isLoop:=False)

            Case CoreAction.MoveToTop
                .ListIndex = 0
            Case CoreAction.MoveToBottom
                .ListIndex = .ListCount - 1

            Case CoreAction.FocusSearch
                UF_Picker.TextBox_Search.SetFocus

            Case CoreAction.SelectItem
                Call SelectNthItem(.ListIndex)

            Case CoreAction.ShowHelp
                Call ShowPickerHelp

            Case CoreAction.ClosePicker
                Unload UF_Picker

            Case Else
                mStatusLabel.Caption = "Undefined action was triggered"
                Call SetStatusBarTemporarily("Undefined action was triggered", 3000)
                HandleAction = False

        End Select
    End With
End Function

'/*
' * Handles key press events from the user form.
' * It checks for keybindings in either the listbox or textbox context
' * and executes the corresponding action.
' *
' * @param {String} key - The string representing the pressed key.
' * @param {Object} src - The control that raised the event (mUserForm.ListBox or mUserForm.TextBox).
' */
Private Sub mUserForm_KeyPressWithSendKeys(ByVal key As String, src As Object)
    If Not src Is mUserForm.TextBox Then
        ' If typed in the Listbox, disable other default behaviors
        mUserForm.IsIgnoreKeydown = True

        Dim selectedKey As String
        With mUserForm.ListBox
            If .ListIndex >= 0 And .ListCount > 0 Then
                selectedKey = .List(.ListIndex, 1)
            End If
        End With

        ' Check if an action is defined for the pressed key
        If mKeybindings.Exists(key) Then
            If mOwner.HandleAction(selectedKey, mKeybindings(key).ActionCode) Then
                Exit Sub
            ElseIf Me.HandleAction(selectedKey, mKeybindings(key).ActionCode) Then
                Exit Sub
            End If
        End If

        ' If the pressed key corresponds to a displayed index, confirm it.
        Dim displayKey As String
        displayKey = gVim.KeyMap.SendKeysToDisplayText(key)

        If Len(displayKey) = 1 Then
            Dim i As Long
            For i = 0 To 100
                With mUserForm.ListBox
                    If i >= .ListCount Then
                        Exit For
                    ElseIf .List(i, 0) = "" Then
                        Exit For
                    End If

                    If .List(i, 0) = displayKey Then
                        Call SelectNthItem(i)
                        Exit Sub
                    End If
                End With
            Next i
        End If
    Else
        If Not mTextboxMap.Exists(key) Then
            Exit Sub
        End If

        mUserForm.IsIgnoreKeydown = True
        Select Case mTextboxMap(key)
            Case TextboxAction.MoveUp
                Call ListUp
            Case TextboxAction.MoveDown
                Call ListDown
            Case TextboxAction.ScrollUp
                Call ListUp(VISIBLE_ROWS, isLoop:=False)
            Case TextboxAction.ScrollDown
                Call ListDown(VISIBLE_ROWS, isLoop:=False)
            Case TextboxAction.DeleteWord
                Call KeyStroke(Ctrl_ + BackSpace_)
            Case TextboxAction.DeleteBackward
                Call KeyStroke(Shift_ + Home_, BackSpace_)
            Case TextboxAction.SelectItem
                Call SelectNthItem(mUserForm.ListBox.ListIndex)
            Case TextboxAction.ShowHelp
                Call ShowPickerHelp
            Case TextboxAction.ClosePicker
                Unload UF_Picker
            Case Else
                mUserForm.IsIgnoreKeydown = False
        End Select
    End If
End Sub

'--- List manipulation helper functions ---

'/*
' * Moves the listbox selection up.
' *
' * @param {Long} [cnt=1] - The number of rows to move.
' * @param {Boolean} [isLoop=True] - Whether to loop around when reaching the end of the list.
' */
Private Sub ListUp(Optional ByVal cnt As Long = 1, Optional ByVal isLoop As Boolean = True)
    With mUserForm.ListBox
        If .ListCount = 0 Then
            Exit Sub
        ElseIf .ListIndex = 0 And isLoop Then
            .ListIndex = .ListCount - 1
        ElseIf .ListIndex - cnt < 0 Then
            .ListIndex = 0
        Else
            .ListIndex = .ListIndex - cnt
        End If
    End With
End Sub

'/*
' * Moves the listbox selection down.
' *
' * @param {Long} [cnt=1] - The number of rows to move.
' * @param {Boolean} [isLoop=True] - Whether to loop around when reaching the end of the list.
' */
Private Sub ListDown(Optional ByVal cnt As Long = 1, Optional ByVal isLoop As Boolean = True)
    With mUserForm.ListBox
        If .ListCount = 0 Then
            Exit Sub
        ElseIf .ListIndex = .ListCount - 1 And isLoop Then
            .ListIndex = 0
        ElseIf .ListIndex + cnt >= .ListCount Then
            .ListIndex = .ListCount - 1
        Else
            .ListIndex = .ListIndex + cnt
        End If
    End With
End Sub

'/*
' * Selects the item at the specified index and closes the Picker.
' *
' * @param {Long} n - The index of the item to select.
' */
Private Sub SelectNthItem(ByVal n As Long)
    With mUserForm.ListBox
        If .ListIndex < -1 Or .ListCount = 0 Then
            Exit Sub
        End If

        Call mOwner.OnEnter(.List(n, 1))
        Unload UF_Picker
    End With
End Sub

'/*
' * Displays the Picker's help message.
' * Organizes and presents all declared actions and their help text by category.
' */
Private Sub ShowPickerHelp()
    Dim pickerAction As Variant
    Dim strHelp As String
    Dim helpCategory As Dictionary

    Set helpCategory = New Dictionary
    For Each pickerAction In mKeybindings.Items()
        If pickerAction.ActionTypeName <> "" Then
            Dim tName As String
            tName = pickerAction.ActionTypeName
            If Not helpCategory.Exists(tName) Then
                helpCategory.Add tName, New Collection
            End If

            helpCategory(tName).Add pickerAction.HelpMessage
        End If
    Next

    Dim cat As Variant
    Dim helpContent As Variant
    For Each cat In helpCategory.Keys
        strHelp = strHelp & "[" & cat & "]" & vbCrLf
        For Each helpContent In helpCategory(cat)
            strHelp = strHelp & "  " & helpContent & vbCrLf
        Next
        strHelp = strHelp & vbCrLf
    Next
    strHelp = RTrim(strHelp)

    MsgBox strHelp, Title:="Picker help"
End Sub
