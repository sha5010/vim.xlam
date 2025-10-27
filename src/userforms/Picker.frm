VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Picker 
   Caption         =   "Picker"
   ClientHeight    =   6120
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4050
   OleObjectBlob   =   "Picker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/**
' * This UserForm serves as a generic picker interface, allowing users to select items
' * from a list that can be filtered via a search box. It dynamically adjusts its content
' * and behavior based on an attached IPicker implementation.
' */

' Master list of characters available for quick indexing of list items.
Private Const KEYLIST_MASTER As String = "abcdefghijklmnopqrstuvwxyz1234567890"
' Maximum number of items to display in the listbox to prevent performance issues.
Private Const MAX_SHOW_ITEMS As Long = 2000

Private mPickerSource       As cls_IPicker    ' Reference to the IPicker implementation that provides data and handles actions.
Private mItems              As Dictionary     ' Dictionary holding the items to be displayed in the picker.
Private mAvailableIndexKeys As String         ' Not currently used for indexing, but could be for optimized key assignment.
Private mUsedKeyChars       As String         ' Characters already bound to specific actions by the picker source.
Private mCacheUpdated       As Boolean        ' Flag to indicate if the item cache needs to be refreshed.
Private mSearchTerm         As String         ' The current search term entered by the user.
Private mResultCount        As Long           ' The number of results found after filtering.
Private mLastSelected       As String         ' The key of the last selected item, used for restoring selection.
Private mlastSelectedIndex  As Long           ' The index of the last selected item, used for restoring selection.
Private mIsGiveUp           As Boolean        ' Flag to indicate if the search yielded too many results and was truncated.
Private mIsUpdating         As Boolean        ' Flag to prevent re-entrancy during list updates.
Private mIsSearching        As Boolean        ' Flag to indicate if a search operation is currently in progress.

'/**
' * @brief Refreshes the internal item cache from the picker source.
' *
' * This method retrieves the latest items from the attached IPicker source
' * and marks the cache as updated, triggering a list refresh.
' */
Public Sub UpdateCache()
    Set mItems = mPickerSource.Item
    mCacheUpdated = True
End Sub

'/**
' * @brief Handles the Change event of the result listbox.
' *
' * This event is triggered when the selected item in the listbox changes.
' * It updates the status, notifies the `IPicker` source of the change,
' * and stores the last selected item's key and index.
' */
Private Sub ListBox_Result_Change()
    With Me.ListBox_Result
        If .ListIndex < 0 Or .ListCount = 0 Then
            ' No item is selected or list is empty, exit to prevent errors.
            Exit Sub
        End If

        Call UpdateStatus
        Call mPickerSource.OnChange(.List(.ListIndex, 1))
        mLastSelected = .List(.ListIndex, 1)
        mlastSelectedIndex = .ListIndex
    End With
End Sub

'/**
' * @brief Handles the Double-Click event of the result listbox.
' *
' * When an item in the listbox is double-clicked, this event notifies
' * the `IPicker` source that an item has been confirmed/selected.
' *
' * @param {MSForms.ReturnBoolean} Cancel - A boolean value that you can set to True
' *                                         to cancel the default action.
' */
Private Sub ListBox_Result_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.ListBox_Result
        If .ListIndex < 0 Or .ListCount = 0 Then
            ' No item is selected or list is empty, exit to prevent errors.
            Exit Sub
        End If

        Call mPickerSource.OnEnter(.List(.ListIndex, 1))
    End With
End Sub

'/**
' * @brief Handles the Change event of the search textbox.
' *
' * This event updates the internal search term whenever the user types
' * in the search box.
' */
Private Sub TextBox_Search_Change()
    mSearchTerm = TextBox_Search.Text
End Sub

'/**
' * @brief Updates and filters the listbox items based on the current search term.
' *
' * This procedure iterates through all available items, filters them based on
' * the search term, and populates the listbox. It handles performance by
' * yielding control (`DoEvents`) and allows the search to be interrupted
' * if the search term changes or the cache is updated.
' */
Private Sub UpdateList()
    Dim i                   As Long      ' Loop counter for items
    Dim t                   As Long      ' Tick count for performance monitoring
    Dim keyCharIndex        As Long      ' Index for assigning quick selection keys
    Dim listKey             As Variant   ' Key of the current item
    Dim listItem            As String    ' Display text of the current item
    Dim lastUpdated         As Long      ' Last tick count when status was updated
    Dim searchTerm          As String    ' Formatted search term for 'Like' operator
    Dim currentTerm         As String    ' Snapshot of the search term at the start of the loop
    Dim isKeyVisible        As Boolean   ' Indicates if the key column is visible
    Dim isFoundLastSelected As Boolean   ' Flag to track if the previously selected item is found

    ' Prevent re-entrant calls if an update is already in progress.
    If mIsUpdating Then
        Exit Sub
    End If
    mIsUpdating = True
    isKeyVisible = mPickerSource.KeyWidth > 0

    On Error GoTo Catch ' Set up error handling for the loop.

    ' Main loop for list updates, continues if search term changes or cache updates.
    Do
        i = 0
        mResultCount = 0
        keyCharIndex = 0
        mIsGiveUp = False
        ListBox_Result.Clear ' Clear the listbox before populating.

        currentTerm = mSearchTerm               ' Capture current search term for consistency within this iteration.
        searchTerm = "*" & mSearchTerm & "*"    ' Prepare search term for 'Like' operator.
        lastUpdated = GetTickCount()            ' Initialize performance timer.
        isFoundLastSelected = False             ' Reset flag for last selected item.

        If mItems Is Nothing Then
            Exit Sub
        End If

        mIsSearching = True ' Indicate that a search operation is active.

        ' Iterate through each item provided by the picker source.
        For Each listKey In mItems.Keys
            i = i + 1
            listItem = mItems.Item(listKey)
            t = GetTickCount()
            ' Update status and yield control periodically to keep the UI responsive.
            If t - lastUpdated > 16 Then ' Approximately 60 FPS (1000ms / 60 = 16.6ms)
                Label_Status.Caption = "Searching... " & i & " / " & mItems.Count
                DoEvents
                Sleep 1
                lastUpdated = t
            End If

            ' Check for interruption conditions:
            If mItems Is Nothing Then
                Exit Sub
            ElseIf mSearchTerm <> currentTerm Then
                Exit For ' Search term changed, restart search with new term.
            ElseIf mCacheUpdated Then
                mCacheUpdated = False ' Cache updated, restart with fresh items.
                Exit For
            ElseIf currentTerm <> "" Then ' Only filter if a search term is present.
                ' Apply search filtering logic.
                If listItem Like searchTerm Then
                ElseIf LCase(listItem) Like searchTerm Then
                ElseIf Not isKeyVisible Then
                    GoTo Continue ' If key column is not visible, don't search by key.
                ElseIf listKey Like searchTerm Then
                ElseIf LCase(listKey) Like searchTerm Then
                Else
                    GoTo Continue ' Item does not match search term, skip to next.
                End If
            End If

            mResultCount = mResultCount + 1
            ' If too many results, truncate the list and set the "give up" flag.
            If ListBox_Result.ListCount >= MAX_SHOW_ITEMS Then
                mIsGiveUp = True
            Else
                ListBox_Result.AddItem "" ' Add a new row to the listbox.

                ' Assign a quick selection key if the key column is visible and available.
                Dim keyChar As String
                Do While keyCharIndex < Len(KEYLIST_MASTER)
                    keyCharIndex = keyCharIndex + 1
                    keyChar = Mid(KEYLIST_MASTER, keyCharIndex, 1)
                    If InStr(mUsedKeyChars, keyChar) <= 0 Then
                        ListBox_Result.List(mResultCount - 1, 0) = keyChar
                        Exit Do ' Key assigned, exit inner loop.
                    End If
                Loop
                ListBox_Result.List(mResultCount - 1, 1) = listKey    ' Store item key in column 1.
                ListBox_Result.List(mResultCount - 1, 2) = listItem ' Store item display text in column 2.

                ' Restore previous selection if this item matches.
                If listKey = mLastSelected Then
                    ListBox_Result.ListIndex = mResultCount - 1
                    isFoundLastSelected = True
                End If
            End If
Continue: ' Label for skipping items that don't match the search criteria.
        Next

        ' Adjust selection after the list is populated.
        If Not isFoundLastSelected And mlastSelectedIndex < ListBox_Result.ListCount Then
            ListBox_Result.ListIndex = mlastSelectedIndex
        ElseIf ListBox_Result.ListCount > 0 And ListBox_Result.ListIndex < 0 Then
            ListBox_Result.ListIndex = 0 ' Select the first item if nothing else is selected.
        End If
        Call ListBox_Result_Change ' Trigger change event to update status and picker source.

        mIsSearching = False ' Search operation completed.
        Call UpdateStatus     ' Update the status label.

        ' Wait for search term to change or cache to update, then restart the loop.
        Do
            If currentTerm <> mSearchTerm Then
                Exit Do ' Search term changed, restart update.
            ElseIf mCacheUpdated Then
                mCacheUpdated = False ' Cache updated, restart update.
                Exit Do
            ElseIf mItems Is Nothing Then
                Exit Sub ' Items were cleared, exit.
            End If
            Sleep 1
            DoEvents
        Loop
    Loop
    Exit Sub

Catch:
    If Err.Number = 93 Then ' Invalid pattern string (e.g., malformed regex in search term)
        Err.Clear       ' Clear the error.
        Resume Next     ' Continue to the next statement (skips the problematic search filter).
    Else
        Call ErrorHandler("UpdateList")
    End If
End Sub

'/**
' * @brief Updates the status label displaying search results and item count.
' *
' * This method shows the current selection index, total result count,
' * and indicates if the results were truncated due to `MAX_SHOW_ITEMS`.
' */
Private Sub UpdateStatus()
    If mIsSearching Then
        Exit Sub
    End If

    Dim statusMsg As String
    With Me.ListBox_Result
        If .ListCount = 0 Then
            statusMsg = "0 / 0"
        Else
            statusMsg = .ListIndex + 1 & " / " & mResultCount ' Current selection / total results.
        End If

        ' Add a truncation warning if applicable.
        If mIsGiveUp Then
            statusMsg = statusMsg & " (only top " & MAX_SHOW_ITEMS & " items)"
        End If
    End With

    Me.Label_Status.Caption = statusMsg ' Display the status message.
End Sub

'/**
' * @brief Initializes the UserForm and its controls before it is displayed.
' *
' * This sets the initial positions and sizes of the search box, status label,
' * and result listbox, as well as the initial position of the UserForm itself.
' */
Private Sub UserForm_Initialize()
    ' Initialize position and size of the search textbox.
    With Me.TextBox_Search
        .Top = 3
        .Left = 3
        .Height = 18
        .TabIndex = 2
    End With

    ' Initialize position and size of the status label.
    With Me.Label_Status
        .Top = 30
        .Left = 3
        .Height = 12
    End With

    ' Initialize position, column count, and default column widths for the result listbox.
    With Me.ListBox_Result
        .Top = 48
        .Left = 3
        .ColumnCount = 3
        .ColumnWidths = "18; 48" ' Default column widths: Assigned character (18pt), Key (48pt)
        .TabIndex = 1
    End With

    ' Set initial UserForm position (bottom-right of the application window).
    With Me
        .StartUpPosition = 0 ' Manual positioning.
    End With
End Sub

'/**
' * @brief Handles the Activate event of the UserForm.
' *
' * This event is triggered when the UserForm becomes the active window.
' * It resizes the internal controls to fit the UserForm's current dimensions
' * and initiates the initial list population.
' */
Private Sub UserForm_Activate()
    ' Adjust width of search textbox and status label to fit the form's inner width.
    Me.TextBox_Search.Width = Me.InsideWidth - 6
    Me.Label_Status.Width = Me.InsideWidth - 6
    With Me.ListBox_Result
        ' Adjust height and width of the result listbox to fill remaining space.
        .Height = Me.InsideHeight - .Top - 6
        .Width = Me.InsideWidth - 6
    End With

    Call UpdateList ' Populate the listbox with items.
End Sub

'/**
' * @brief Handles the Terminate event of the UserForm.
' *
' * This event is triggered when the UserForm is unloaded from memory.
' * It ensures that all object references are released to prevent memory leaks.
' */
Private Sub UserForm_Terminate()
    Set mPickerSource = Nothing
    Set mItems = Nothing
End Sub

'/**
' * @brief Launches the picker UserForm with a specified `IPicker` implementation.
' *
' * This is the primary method to display the picker. It sets up the picker's data source,
' * initializes key bindings, attaches UI controls, and adjusts the form's size and position.
' *
' * @param {cls_IPicker} pickerSource - An object implementing the `IPicker` interface,
' *                                     providing data and handling actions for the picker.
' * @param {Boolean} [ShowLeft=False] - Optional. If True, the picker will be shown on the left
' *                                     side of the application window; otherwise, on the right.
' */
Public Sub Launch(ByRef pickerSource As cls_IPicker, Optional ByVal ShowLeft As Boolean = False)
    Dim pickerBase As cls_PickerBase
    Set pickerBase = New cls_PickerBase

    Set mPickerSource = pickerSource             ' Assign the IPicker implementation.
    Set mItems = mPickerSource.Item              ' Get initial items from the source.
    mUsedKeyChars = mPickerSource.BoundKeyChars  ' Get keys already bound by the picker source.
    mLastSelected = mPickerSource.Default        ' Get the default selected item.
    ' Attach UserForm controls to the picker source for interaction.
    mPickerSource.Attach Me, Me.ListBox_Result, Me.TextBox_Search, Me.Label_Status

    With Me
        ' Adjust listbox column widths to accommodate the picker source's key width.
        .ListBox_Result.ColumnWidths = "18; " & mPickerSource.KeyWidth
        ' Adjust the form's width based on the picker source's explicit form width or key width.
        If mPickerSource.PickerFormWidth > 0 Then
            .Width = mPickerSource.PickerFormWidth
        Else
            .Width = .Width + mPickerSource.KeyWidth    ' Original logic if no explicit form width or 0 is returned
            .Height = .Width * 1.612                    ' Set a proportional height for the form.
        End If
         ' Position at the bottom of the Excel window.
        .Top = Application.Top + Application.Height - .Height - 45

        ' Position the form either on the left or right side of the application window.
        If ShowLeft Then
            .Left = Application.Left + 6
        Else
            .Left = Application.Left + Application.Width - Me.Width - 21
        End If
        .Show ' Display the UserForm.
    End With
End Sub
