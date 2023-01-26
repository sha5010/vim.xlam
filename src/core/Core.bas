Attribute VB_Name = "C_Core"
Option Explicit
Option Private Module

Public X As cls_EventHook               'Workbook 系のイベントを捕捉するために使用

Public gVimMode As Boolean              'Vim モードが有効/無効
Public gCount As Long                   'カウントを使用できる機能で使用
Public gLangJa As Boolean               '日本語モードが有効/無効
Public gKeyMap As Dictionary            'マッピングされたキーと機能の辞書
Public gCmdBuf As String                'コマンドフォームでのコマンド保持
Public gDebugMode As Boolean            'デバッグモードが有効/無効
Public gLastYanked As Range             'Copy,Cutコマンドの対象セルを保持
Public gExtendRange As Range            '複数選択機能で選択したセルを保持
Public gRegisteredKeys As Dictionary    '登録済みのキーリストを保持

Public JumpList As cls_MainArena        '編集したセルを格納するジャンプリスト
Public Repeater As cls_Repeater         'コマンドの繰り返しを保持するためのクラス

Public Const KEY_NORMAL As String = "N"
Public Const KEY_RETURNONLY As String = "RO"
Public Const KEY_NORMAL_ARG As String = "NRQ"
Public Const KEY_RETURNONLY_ARG As String = "RORQ"

Public Const STATUS_PREFIX As String = "vim.xlam: "  'ステータスバーの一時メッセージ表示の prefix

Sub startVim()
    Dim startTimer As Double
    Dim message As String

    startTimer = Timer

    gVimMode = True
    gCount = 1
    gLangJa = DEFAULT_LANG_JA

    If X Is Nothing Then
        Set X = New cls_EventHook
        Set X.App = Application
    Else
        Set X.TempApp = Nothing
    End If

    If JumpList Is Nothing Then
        Set JumpList = New cls_MainArena
        Call JumpList.SetMax(MAX_HISTORIES)
    End If

    If Repeater Is Nothing Then
        Set Repeater = New cls_Repeater
    End If

    Call disableIME

    If gRegisteredKeys Is Nothing Then
        Call initMapping
        gDebugMode = Not ThisWorkbook.IsAddin
    Else
        Call enableKeys
    End If

    Application.OnKey VIM_TOOGLE_KEY, "toggleVim"

    message = "Vim モードを開始しました。"
    If gDebugMode Then
        message = message & "読み込み時間: " & Format(Timer - startTimer, "0.000") & "s"
    End If

    Call setStatusBarTemporarily(message, 3)
End Sub

Sub stopVim()
    If gRegisteredKeys Is Nothing Then
        Call mapResetAll
    Else
        Call disableKeys
    End If

    Call setStatusBarTemporarily("Vim モードを停止しました。", 3)

    Set X.App = Nothing
    Set X = Nothing

    Application.OnKey VIM_TOOGLE_KEY, "toggleVim"

    gVimMode = False
End Sub

Sub reloadVim(Optional isForce As Boolean = False)
    gVimMode = True
    gCount = 1
    gLangJa = DEFAULT_LANG_JA

    If isForce Then
        Set JumpList = New cls_MainArena
        Call JumpList.SetMax(MAX_HISTORIES)
    End If

    Call initMapping

    Application.OnKey VIM_TOOGLE_KEY, "toggleVim"
    Call setStatusBarTemporarily("Vim モードをリロードしました。", 2)
End Sub

Sub toggleVim()
    If gVimMode Then
        Call stopVim
    Else
        Call startVim
    End If
End Sub

Sub temporarilyDisableVim()
    Set X.TempApp = Application

    If gRegisteredKeys Is Nothing Then
        Call mapResetAll
    Else
        Call disableKeys
    End If

    Call setStatusBarTemporarily("図形の挿入モードに移行したため、vim.xlam を一時的に無効化しました。ESC (または Ctrl + [) で復帰します。", 3)

    Application.OnKey "{ESC}", "resumeVim"
    Application.OnKey "^{[}", "resumeVim"
    Application.OnKey VIM_TOOGLE_KEY, "toggleVim"

    gVimMode = False
End Sub

Sub resumeVim()
    Application.OnKey "{ESC}"
    Application.OnKey "^{[}"

    If gRegisteredKeys Is Nothing Then
        Call initMapping
    Else
        Call enableKeys
    End If

    Application.OnKey VIM_TOOGLE_KEY, "toggleVim"
    Call setStatusBarTemporarily("再開しました。", 1)

    Set X.TempApp = Nothing
    gVimMode = True

    Call disableIME
End Sub

Sub showCmdForm(ByVal prefix As String)
    If Not gVimMode Then
        'gVimMode = False なのにキーマップが有効な場合は
        'エラー等でインスタンスが落ちたと考えられるので再起動
        Call startVim
    End If

    gCmdBuf = prefix
    UF_Cmd.Show
End Sub

Sub toggleLang()
    gLangJa = Not gLangJa
    If gLangJa Then
        Call setStatusBarTemporarily("日本語モードに切り替えました。", 2)
    Else
        Call setStatusBarTemporarily("Switched to English mode.", 2)
    End If
End Sub

Sub toggleDebugMode()
    gDebugMode = Not gDebugMode
    If gDebugMode Then
        Call setStatusBarTemporarily("デバッグモードを有効化しました。", 2)
    Else
        Call setStatusBarTemporarily("デバッグモードを無効化しました。", 2)
    End If
End Sub
