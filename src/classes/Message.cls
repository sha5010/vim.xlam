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
'Public Property Get EVENT_NAME_OR_MESSAGE_TITLE() As String
'    EVENT_NAME_OR_MESSAGE_TITLE = IIf(Parent.IsJapanese, _
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
