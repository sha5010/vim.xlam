VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_SheetPicker 
   Caption         =   "SheetPicker"
   ClientHeight    =   5526
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   4032
   OleObjectBlob   =   "SheetPicker.frx":0000
End
Attribute VB_Name = "UF_SheetPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'キーリストを定義
Private Const KEYLIST As String = "1234567890abcdefimnopqrstuvwxyz"
Private Const INVISIBLE As String = "(hidden) "
Private Const VERY_HIDDEN As String = "(HIDDEN) "
Private Const AMOUNT As Byte = 3   'Ctrl で一気に移動する量
Private Const FORM_CAPTION As String = "SheetPicker"

'プレビューモード
Private previewMode As Boolean

Private Function Activate_Nth_sheet(ByVal n As Integer) As Boolean
    'N番目のシートをアクティベート
    If ActiveWorkbook.Sheets.Count < n Or n < 1 Then
        Exit Function
    End If

    If Not ActiveWorkbook.Sheets(n).Visible Then
        ActiveWorkbook.Sheets(n).Visible = True
    End If

    ActiveWorkbook.Sheets(n).Activate
    Activate_Nth_sheet = True
End Function

Private Sub Toggle_Sheet_Visible(ByVal n As Integer, _
                                 Optional ByVal VeryHidden As Boolean = False)
    '変数宣言
    Dim idx As Integer
    Dim sheetVisibility As Integer
    Dim hiddenText As String
    Dim sheetName As String
    Dim i As Integer
    Dim cnt As Integer

    'N番目のシートの可視/不可視状態をトグル
    If ActiveWorkbook.Sheets.Count < n Then
        Exit Sub
    End If

    idx = List_Sheets.ListIndex
    sheetName = List_Sheets.List(idx, 1)

    sheetVisibility = IIf(VeryHidden, xlVeryHidden, xlSheetHidden)
    hiddenText = IIf(VeryHidden, VERY_HIDDEN, INVISIBLE)

    With ActiveWorkbook.Sheets(n)
        If .Visible <> sheetVisibility Then
            'check the number of visible sheets
            For i = 1 To ActiveWorkbook.Sheets.Count
                If ActiveWorkbook.Sheets(i).Visible = xlSheetVisible Then
                    cnt = cnt + 1
                    If cnt > 1 Then
                        Exit For
                    End If
                End If
            Next i

            'not all sheets can be hidden
            If cnt = 1 And .Visible = xlSheetVisible Then
                MsgBox gVim.Msg.HideAllSheets, vbExclamation
                Exit Sub
            End If

            .Visible = sheetVisibility
            sheetName = hiddenText & .Name
        Else
            .Visible = xlSheetVisible
            sheetName = .Name
        End If
    End With

    List_Sheets.List(idx, 1) = sheetName
End Sub

Private Sub Rename_Sheet(ByVal n As Integer)
    '変数宣言
    Dim ret As Variant
    Dim cur As String

    'N番目のシートが存在しなければ終了
    If ActiveWorkbook.Sheets.Count < n Then
        Exit Sub
    End If

    On Error GoTo Catch

    'N番目のシートをリネームするためのダイアログを表示
    With ActiveWorkbook.Sheets(n)
        cur = .Name
        ret = InputBox(gVim.Msg.EnterNewSheetName, gVim.Msg.RenameSheetTitle, cur)
        Call DisableIME

        If ret <> "" Then
            '同名だったら何もしない
            If ret = cur Then
                Exit Sub

            '新しい名前のシートがすでに存在する場合はエラー
            ElseIf IsSheetExists(ret) Then
                MsgBox gVim.Msg.SheetAlreadyExists(ret), vbExclamation
                Exit Sub
            End If

            'リネーム
            .Name = ret

            'リストボックス更新
            If .Visible <> xlSheetVisible Then
                ret = INVISIBLE & ret
            End If
            List_Sheets.List(List_Sheets.ListIndex, 1) = ret
        End If
    End With
    Exit Sub

Catch:
    MsgBox gVim.Msg.SheetRenameError, vbExclamation
End Sub

Private Sub Delete_Sheet(ByVal n As Integer)
    '変数宣言
    Dim cur As Integer

    'N番目のシートが存在しなければ終了
    If ActiveWorkbook.Sheets.Count < n Then
        Exit Sub
    End If

    '対象シートがVeryHiddenの場合は消せないので警告表示
    If ActiveWorkbook.Sheets(n).Visible = xlVeryHidden Then
        MsgBox gVim.Msg.CannotDeleteVeryHiddenSheet, vbExclamation
        Exit Sub
    End If

    '対象シートが最後の可視シートの場合はエラー
    If ActiveSheet.Visible = xlSheetVisible And GetVisibleSheetsCount() = 1 Then
        MsgBox gVim.Msg.DeleteOrHideAllSheets, vbExclamation
        Exit Sub
    End If

    '削除前のシート数を保持
    cur = ActiveWorkbook.Sheets.Count

    'N番目のシートを削除 (デフォルトでダイアログが表示される)
    ActiveWorkbook.Sheets(n).Delete

    '削除されたか確認
    If ActiveWorkbook.Sheets.Count < cur Then
        '削除された場合はリスト再生成
        List_Sheets.Clear
        Call MakeList
    End If
End Sub

Private Sub Move_Sheet(ByVal n As Long, ByVal moveDirection As XlSearchDirection)
    With ActiveWorkbook
        ' Check n-th sheet is exists
        If n < 1 Or .Sheets.Count < n Then
            Exit Sub
        End If

        ' Exit if number of sheets = 1
        If .Sheets.Count = 1 Then
            Exit Sub
        End If

        ' Calculate destination index
        Dim destIndex As Long
        Dim isWrap As Boolean: isWrap = False

        If moveDirection = xlNext Then
            If n = .Sheets.Count Then
                destIndex = 1
                isWrap = True
            Else
                destIndex = n + 1
            End If
        ElseIf moveDirection = xlPrevious Then
            If n = 1 Then
                destIndex = .Sheets.Count
                isWrap = True
            Else
                destIndex = n - 1
            End If
        End If

        ' Move Sheet
        Dim hidState As XlSheetVisibility
        Dim targetSheet As Object

        Set targetSheet = .Sheets(destIndex)
        hidState = targetSheet.Visible

        If Not hidState = xlSheetVisible Then
            Application.ScreenUpdating = False
            targetSheet.Visible = xlSheetVisible
        End If

        If n < destIndex Then
            .Sheets(n).Move After:=targetSheet
        Else
            .Sheets(n).Move Before:=targetSheet
        End If

        targetSheet.Visible = hidState

        Application.ScreenUpdating = True

        ' Remake list
        If isWrap Then
            Me.List_Sheets.Clear
            Call MakeList
        Else
            Dim buf As String
            buf = List_Sheets.List(n - 1, 1)
            List_Sheets.List(n - 1, 1) = List_Sheets.List(destIndex - 1, 1)
            List_Sheets.List(destIndex - 1, 1) = buf
        End If

        ' Reselect
        List_Sheets.ListIndex = destIndex - 1
    End With
End Sub

Private Sub Show_Help()
    'ヘルプ文字列
    Dim HELP_MSG As String
    HELP_MSG = "[Move Cursor]¥n" & _
        "  j/k¥tMove down/up¥n" & _
        "  C-j/C-k¥tMove down/up (" & AMOUNT & " rows)¥n" & _
        "  g/G¥tMove to top/bottom¥n" & _
        "¥n" & _
        "[Sheet Action]¥n" & _
        "  J/K¥tSwap sheet with lower/upper¥n" & _
        "  h/H¥tToogle sheet visible/(Very hidden)¥n" & _
        "  l¥tPreview the sheet for current row¥n" & _
        "  R¥tRename sheet¥n" & _
        "  D/X¥tDelete sheet¥n" & _
        "¥n" & _
        "[Change sheet]¥n" & _
        "  Enter¥tActivate the sheet for current row¥n" & _
        "  [0-9a-z]¥tActivate specify sheet¥n" & _
        "¥n" & _
        "[Preview mode]¥n" & _
        "  P¥tToggle preview mode"

    HELP_MSG = Replace(HELP_MSG, "¥n", vbLf)
    HELP_MSG = Replace(HELP_MSG, "¥t", vbTab)
    Call MsgBox(HELP_MSG)

    Me.Caption = Replace(Me.Caption, " (?: Show help)", "")
End Sub

Private Sub List_Sheets_Change()
    Dim idx As Integer

    idx = List_Sheets.ListIndex + 1

    If previewMode And idx > 0 Then
        If ActiveWorkbook.Sheets(idx).Visible And idx <> ActiveWorkbook.ActiveSheet.Index Then
            Call Activate_Nth_sheet(idx)
        End If
    End If
End Sub

Private Sub List_Sheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Activate_Nth_sheet(List_Sheets.ListIndex + 1) Then
        Unload Me
    End If
End Sub

Private Sub List_Sheets_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    '変数宣言
    Const CTRL_OFFSET As Integer = -96
    Dim idx As Byte

    'Escキーを押されたらアンロード
    If KeyAscii = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If

    'Enterキーを押されたらアクティブなものに切り替え
    If KeyAscii = 13 Then
        Activate_Nth_sheet (List_Sheets.ListIndex + 1)
        Unload Me
        Exit Sub
    End If

    'vim 風の上下移動とか
    With List_Sheets
        Select Case KeyAscii
            Case Asc("j")
                If .ListIndex = .ListCount - 1 Then
                    .ListIndex = 0
                Else
                    .ListIndex = .ListIndex + 1
                End If

            Case Asc("k")
                If .ListIndex = 0 Then
                    .ListIndex = .ListCount - 1
                Else
                    .ListIndex = .ListIndex - 1
                End If

            Case CTRL_OFFSET + Asc("j")
                If .ListIndex = .ListCount - 1 Then
                    .ListIndex = 0
                ElseIf .ListIndex + AMOUNT >= .ListCount Then
                    .ListIndex = .ListCount - 1
                Else
                    .ListIndex = .ListIndex + AMOUNT
                End If

            Case CTRL_OFFSET + Asc("k")
                If .ListIndex = 0 Then
                    .ListIndex = .ListCount - 1
                ElseIf .ListIndex - AMOUNT < 0 Then
                    .ListIndex = 0
                Else
                    .ListIndex = .ListIndex - AMOUNT
                End If

            Case Asc("J")
                Call Move_Sheet(.ListIndex + 1, xlNext)

            Case Asc("K")
                Call Move_Sheet(.ListIndex + 1, xlPrevious)

            Case Asc("g")
                .ListIndex = 0

            Case CTRL_OFFSET + Asc("g")
                .ListIndex = .ListCount - 1

            Case Asc("G")
                .ListIndex = .ListCount - 1

            Case Asc("h")
                Call Toggle_Sheet_Visible(.ListIndex + 1)

            Case Asc("H")
                Call Toggle_Sheet_Visible(.ListIndex + 1, VeryHidden:=True)

            Case Asc("l")
                Call Activate_Nth_sheet(.ListIndex + 1)

            Case Asc("P")
                previewMode = Not previewMode
                If previewMode Then
                    Me.Caption = FORM_CAPTION & " (Preview mode)"
                    Call List_Sheets_Change
                Else
                    Me.Caption = FORM_CAPTION
                End If

            Case Asc("R")
                Call Rename_Sheet(.ListIndex + 1)

            Case Asc("D"), Asc("X")
                Call Delete_Sheet(.ListIndex + 1)

            Case Asc("?")
                Call Show_Help
        End Select
    End With

    'それ以外でインデックスが指定された場合
    idx = InStr(KEYLIST, Chr(KeyAscii))
    If idx > 0 Then

        '表示されていないインデックスの場合は無効
        If idx > List_Sheets.ListCount Then
            Exit Sub
        End If

        'アクティブシートに設定
        If Activate_Nth_sheet(idx) Then
            Unload Me
        End If

    ElseIf Asc("A") <= KeyAscii And KeyAscii <= Asc("Z") Then
        '大文字入力時、選択が移動してしまうのを防ぐ
        KeyAscii = -1
    End If
End Sub

Private Sub UserForm_Activate()
    '表示位置
    With Me
        .Top = Application.Top + Application.Height - .Height - 36 + (Application.WindowState = xlMaximized) * 6
        .Left = Application.Left - (Application.WindowState = xlMaximized) * 6
    End With
End Sub

Private Sub UserForm_Initialize()
    'フォームのキャプションを設定
    Me.Caption = FORM_CAPTION & " (?: Show help)"

    'デフォルトではプレビューモードを無効化
    previewMode = False

    'リストボックスのサイズをUserFormに合わせる
    With List_Sheets
        .Top = 3
        .Left = 3
        .Height = Me.InsideHeight - 3
        .Width = Me.InsideWidth - 6

        'キー列の表示幅を設定
        .ColumnWidths = "18 pt"
    End With

    'シート一覧を表示
    Call MakeList
End Sub

Private Sub MakeList()
    'エラーハンドリング
    On Error GoTo Catch

    '変数宣言
    Dim i As Integer
    Dim keyLength As Integer
    Dim sheetName As String

    '使用できるキーの数を取得
    keyLength = Len(KEYLIST)

    'アクティブブックのシート一覧をリストに表示
    With List_Sheets
        For i = 1 To ActiveWorkbook.Sheets.Count
            .AddItem ""

            'キーが使えれば割当
            If i <= keyLength Then
                .List(i - 1, 0) = Mid(KEYLIST, i, 1)
            End If

            'シート名を表示
            sheetName = ActiveWorkbook.Sheets(i).Name
            If Not ActiveWorkbook.Sheets(i).Visible Then
                sheetName = INVISIBLE & sheetName
            End If
            .List(i - 1, 1) = sheetName

            'アクティブシートならアクティブに
            If i = ActiveWorkbook.ActiveSheet.Index Then
                .ListIndex = i - 1
            End If
        Next
    End With
    Exit Sub

Catch:
    If ErrorHandler("MakeList in UF_SheetPicker") Then
        Unload Me
    End If
End Sub
