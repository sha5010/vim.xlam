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
Private Const INVISIBLE As String = "(非表示) "
Private Const AMOUNT As Byte = 3   'Ctrl で一気に移動する量

'プレビューモード
Private previewMode As Boolean

Private Function Activate_Nth_sheet(ByVal n As Integer) As Boolean
    'N番目のシートをアクティベート
    If ActiveWorkbook.Worksheets.Count < n Or n < 1 Then
        Exit Function
    End If
    
    If Not ActiveWorkbook.Worksheets(n).Visible Then
        ActiveWorkbook.Worksheets(n).Visible = True
    End If
    
    ActiveWorkbook.Worksheets(n).Activate
    Activate_Nth_sheet = True
End Function

Private Sub Toggle_Sheet_Visible(ByVal n As Integer, _
                                 Optional ByVal VeryHidden As Boolean = False)
    '変数宣言
    Dim idx As Integer
    Dim sheetName As String
                                 
    'N番目のシートの可視/不可視状態をトグル
    If ActiveWorkbook.Worksheets.Count < n Then
        Exit Sub
    End If
    
    idx = List_Sheets.ListIndex
    sheetName = List_Sheets.List(idx, 1)
    
    With ActiveWorkbook.Worksheets(n)
        If .Visible <> xlSheetVisible Then
            .Visible = xlSheetVisible
            sheetName = Replace(sheetName, INVISIBLE, "", Count:=1)
        Else
            If VeryHidden Then
                .Visible = xlVeryHidden
            Else
                .Visible = xlSheetHidden
            End If
            sheetName = INVISIBLE & sheetName
        End If
    End With
    
    List_Sheets.List(idx, 1) = sheetName
End Sub

Private Sub Rename_Sheet(ByVal n As Integer)
    '変数宣言
    Dim ret As Variant
    Dim cur As String
    
    'N番目のシートが存在しなければ終了
    If ActiveWorkbook.Worksheets.Count < n Then
        Exit Sub
    End If
    
    'N番目のシートをリネームするためのダイアログを表示
    With ActiveWorkbook.Worksheets(n)
        cur = .Name
        ret = InputBox("新しいシート名を入力してください。", "シートの名前変更", cur)
        
        If ret <> "" Then
            'リネーム
            .Name = ret
            
            'リストボックス更新
            If .Visible <> xlSheetVisible Then
                ret = INVISIBLE & ret
            End If
            List_Sheets.List(List_Sheets.ListIndex, 1) = ret
        End If
    End With
End Sub

Private Sub Show_Help()
    'ヘルプ文字列
    Dim HELP As String
    HELP = "[Move]" & vbLf & _
        "j/k" & Chr(9) & "Move down/up" & vbLf & _
        "Ctrl+(j/k)" & Chr(9) & "Move down/up " & AMOUNT & " rows" & vbLf & _
        "g/G" & Chr(9) & "Move to top/bottom" & vbLf & vbLf & _
        "[Sheet Action]" & vbLf & _
        "h/H" & Chr(9) & "Toogle sheet visible/(Very hidden)" & vbLf & _
        "l" & Chr(9) & "Preview the sheet for current row" & vbLf & _
        "R" & Chr(9) & "Rename sheet" & vbLf & vbLf & _
        "[Change sheet]" & vbLf & _
        "Enter" & Chr(9) & "Activate the sheet for current row" & vbLf & _
        "[0-9a-z]" & Chr(9) & "Activate specify sheet" & vbLf & vbLf & _
        "[Preview mode]" & vbLf & _
        "P" & Chr(9) & "Toggle preview mode"
        
    Call MsgBox(HELP)
End Sub

Private Sub List_Sheets_Change()
    Dim idx As Integer
    
    idx = List_Sheets.ListIndex + 1
    
    If previewMode Then
        If ActiveWorkbook.Worksheets(idx).Visible And idx <> ActiveWorkbook.ActiveSheet.Index Then
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
                    Me.Caption = "シートを選択 (プレビューモードON)"
                    Call List_Sheets_Change
                Else
                    Me.Caption = "シートを選択"
                End If
            
            Case Asc("R")
                Call Rename_Sheet(.ListIndex + 1)
            
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
    
    '変数宣言
    Dim i As Integer
    Dim keyLength As Integer
    Dim sheetName As String
    
    '使用できるキーの数を取得
    keyLength = Len(KEYLIST)
    
    'デフォルトではプレビューモードを無効化
    previewMode = False
    
    'リストボックスのサイズをUserFormに合わせる
    With List_Sheets
        .Top = 3
        .Left = 3
        .Height = Me.InsideHeight - 3
        .Width = Me.InsideWidth - 6
    End With
    
    'アクティブブックのシート一覧をリストに表示
    With List_Sheets
        For i = 1 To ActiveWorkbook.Worksheets.Count
            .AddItem ""
        
            'キーが使えれば割当
            If i <= keyLength Then
                .List(i - 1, 0) = Mid(KEYLIST, i, 1)
            End If
        
            'シート名を表示
            sheetName = ActiveWorkbook.Worksheets(i).Name
            If Not ActiveWorkbook.Worksheets(i).Visible Then
                sheetName = INVISIBLE & sheetName
            End If
            .List(i - 1, 1) = sheetName
            
            'アクティブシートならアクティブに
            If i = ActiveWorkbook.ActiveSheet.Index Then
                .ListIndex = i - 1
            End If
        Next
    End With
End Sub
