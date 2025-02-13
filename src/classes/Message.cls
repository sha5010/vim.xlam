VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cls_Vim

' --- TEMPLATE ---
'Public Property Get EventNameOrMessageTitle() As String
'    EventNameOrMessageTitle = IIf(Parent.IsJapanese, _
'        "日本語のメッセージ (Japanese message)", _
'        "English message")
'End Property

Public Property Get VimStarted() As String
    VimStarted = IIf(Parent.IsJapanese, _
        "Vim モードを開始しました。", _
        "Vim mode started.")
End Property

Public Property Get VimStopped() As String
    VimStopped = IIf(Parent.IsJapanese, _
        "Vim モードを停止しました。", _
        "Vim mode stopped.")
End Property

Public Property Get LangChanged() As String
    LangChanged = IIf(Parent.IsJapanese, _
        "日本語モードに切り替えました。", _
        "Switched to English mode.")
End Property

Public Property Get DebugModeEnabled() As String
    DebugModeEnabled = IIf(Parent.IsJapanese, _
        "デバッグモードを有効化しました。", _
        "Debug mode enabled.")
End Property

Public Property Get DebugModeDisabled() As String
    DebugModeDisabled = IIf(Parent.IsJapanese, _
        "デバッグモードを無効化しました。", _
        "Debug mode disabled.")
End Property

Public Property Get MissingMacro() As String
    MissingMacro = IIf(Parent.IsJapanese, _
        "登録されているマクロが存在しません。マクロ名: ", _
        "The registered macro does not exist. Macro name: ")
End Property

Public Property Get NoKeyAllocation() As String
    NoKeyAllocation = IIf(Parent.IsJapanese, _
        "キー割り当てがありません。", _
        "No allocation.")
End Property

Public Property Get AmbiguousCommand() As String
    AmbiguousCommand = IIf(Parent.IsJapanese, _
        "コマンドが曖昧です: ", _
        "Command is ambiguous: ")
End Property

Public Property Get NoCommandAvailable() As String
    NoCommandAvailable = IIf(Parent.IsJapanese, _
        "コマンドがありません: ", _
        "No command is available: ")
End Property

Public Property Get NoCommandHelp() As String
    NoCommandHelp = IIf(Parent.IsJapanese, _
        "コマンドヘルプがありません: ", _
        "No command help is available: ")
End Property

Public Property Get ArgumentsRequired() As String
    ArgumentsRequired = IIf(Parent.IsJapanese, _
        "引数が必要です。", _
        "Argument is required.")
End Property

Public Property Get LatestJumplist() As String
    LatestJumplist = IIf(Parent.IsJapanese, _
        "一番新しいジャンプリスト履歴です。", _
        "Latest jump list entry.")
End Property

Public Property Get OldestJumplist() As String
    OldestJumplist = IIf(Parent.IsJapanese, _
        "一番古いジャンプリスト履歴です。", _
        "Oldest jump list entry.")
End Property

Public Property Get ClearedJumplist() As String
    ClearedJumplist = IIf(Parent.IsJapanese, _
        "ジャンプリストをクリアしました。", _
        "Jumplist cleared.")
End Property

Public Property Get YankInProgress() As String
    YankInProgress = IIf(Parent.IsJapanese, _
        "テキストをコピーしています...", _
        "Copying text...")
End Property

Public Property Get YankDone() As String
    YankDone = IIf(Parent.IsJapanese, _
        "クリップボードにコピーしました｡", _
        "Copied to clipboard. ")
End Property

Public Property Get ProcessingNumber() As String
    ProcessingNumber = IIf(Parent.IsJapanese, _
        "数値の増減を処理中です...", _
        "Processing numeric increase/decrease...")
End Property

Public Property Get TooManyCells() As String
    TooManyCells = IIf(Parent.IsJapanese, _
        "選択セル数が多すぎます。", _
        "Too many cells selected.")
End Property

Public Property Get InitializedExtendedSelection() As String
    InitializedExtendedSelection = IIf(Parent.IsJapanese, _
        "異なるシートで拡張選択はできないため、選択範囲は初期化されました。", _
        "The selection was initialized because extended selections cannot be made on different sheets.")
End Property

Public Property Get ClearedExtendedSelection() As String
    ClearedExtendedSelection = IIf(Parent.IsJapanese, _
        "保存されている拡張選択範囲をクリアしました。", _
        "Saved extended selection has been cleared.")
End Property

Public Property Get UnableIdentifySource() As String
    UnableIdentifySource = IIf(Parent.IsJapanese, _
        "元となるデータを特定できません。", _
        "Unable to identify the source data.")
End Property

Public Property Get NoDataInSelectedCells() As String
    NoDataInSelectedCells = IIf(Parent.IsJapanese, _
        "選択セルの先頭、または末尾にデータがありません。", _
        "There is no data at the beginning or end of the selected cell.")
End Property

Public Property Get ConfirmToDeleteAllComments() As String
    ConfirmToDeleteAllComments = IIf(Parent.IsJapanese, _
        "アクティブシート上のすべてのコメントを削除します。よろしいですか?" & vbLf & "　※この操作は取り消せません。", _
        "Delete all comments on the active sheet. Are you sure?" & vbLf & "  * THIS ACTION CANNOT BE UNDONE.")
End Property

Public Property Get NoMatchingCell() As String
    NoMatchingCell = IIf(Parent.IsJapanese, _
        "該当するセルが見つかりません。", _
        "No matching cell found.")
End Property

Public Property Get EmptyClipboard() As String
    EmptyClipboard = IIf(Parent.IsJapanese, _
        "クリップボードが空です。", _
        "Clipboard is empty.")
End Property

Public Property Get ConfirmToSaveBeforeReopening() As String
    ConfirmToSaveBeforeReopening = IIf(Parent.IsJapanese, _
        "ファイルを開き直す前に、編集内容を保存しますか?", _
        "Do you want to save your edits before reopening the file?")
End Property

Public Property Get ConfirmToSaveBeforeSwitchReadonly() As String
    ConfirmToSaveBeforeSwitchReadonly = IIf(Parent.IsJapanese, _
        "読み取り専用の切り替えを行う前に、編集内容を保存しますか?", _
        "Do you want to save your edits before switching to read-only?")
End Property

Public Property Get EnterNewSheetName() As String
    EnterNewSheetName = IIf(Parent.IsJapanese, _
        "新しいシート名を入力してください。", _
        "Please enter a new sheet name.")
End Property

Public Property Get RenameSheetTitle() As String
    RenameSheetTitle = IIf(Parent.IsJapanese, _
        "シート名の変更", _
        "Rename sheet")
End Property

Public Function SheetAlreadyExists(ByVal sheetName As String) As String
    SheetAlreadyExists = IIf(Parent.IsJapanese, _
        "すでに """ & sheetName & """ シートが存在します。", _
        "There is already a """ & sheetName & """ sheet.")
End Function

Public Property Get ChangedSheetName() As String
    ChangedSheetName = IIf(Parent.IsJapanese, _
        "シート名を変更しました", _
        "Sheet name has been changed")
End Property

Public Property Get DeleteOrHideAllSheets() As String
    DeleteOrHideAllSheets = IIf(Parent.IsJapanese, _
        "シートをすべて削除、または非表示にすることはできません。", _
        "You cannot delete or hide all sheets.")
End Property

Public Property Get HideAllSheets() As String
    HideAllSheets = IIf(Parent.IsJapanese, _
        "すべてのシートを非表示にすることはできません。", _
        "You cannnot hide all sheets.")
End Property

Public Property Get CannotDeleteVeryHiddenSheet() As String
    CannotDeleteVeryHiddenSheet = IIf(Parent.IsJapanese, _
        "VeryHidden のシートは消すことができません。" & vbLf & "非表示を解除してから再実行してください｡ ", _
        "The 'VeryHidden' sheet cannot be deleted." & vbLf & "Please unhide it and try again.")
End Property

Public Property Get SheetRenameError() As String
    SheetRenameError = IIf(Parent.IsJapanese, _
        "入力されたシートまたはグラフの名前が正しくありません。次の点を確認して修正してください。" & vbLf & vbLf & _
        "- 入力文字が 31 文字以内であること" & vbLf & _
        "- 次の文字を含まないこと:" & vbLf & "    :  ¥ / ? * [ ]", _
        "The sheet or chart name entered is not valid. Please ensure that:" & vbLf & vbLf & _
        "- The name is 31 characters or less" & vbLf & _
        "- It does not contain any of the following characters:" & vbLf & "    :  ¥ / ? * [ ]")
End Property

Public Property Get LockingRange() As String
    LockingRange = IIf(Parent.IsJapanese, _
        "セル選択ロック範囲: ", _
        "Cell selection lock range: ")
End Property

Public Property Get ClearedSelectionLock() As String
    ClearedSelectionLock = IIf(Parent.IsJapanese, _
        "セル選択ロックを解除しました", _
        "Cleared cell selection lock.")
End Property

Public Property Get RemainingResults() As String
    RemainingResults = IIf(Parent.IsJapanese, _
        " 個の結果は省略", _
        " results remaining")
End Property
