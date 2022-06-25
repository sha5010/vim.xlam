VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Repeater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private savedFuncName As String
Private savedCount As Long
Private savedArgs As Variant

Public Sub Register(ByVal funcName As String, ByVal Count As Long, ByVal args As Variant)
    'クラス変数に値を保存
    savedFuncName = funcName
    savedCount = Count
    savedArgs = args
End Sub

Public Sub Run()
    'gCount を復元
    gCount = savedCount

    '引数なしの場合
    If LBound(savedArgs) = -1 Then
        Application.Run savedFuncName
        GoTo Finally
    End If

    '引数ありの場合
    Select Case UBound(savedArgs)
        Case 0
            Application.Run savedFuncName, savedArgs(0)
        Case 1
            Application.Run savedFuncName, savedArgs(0), savedArgs(1)
        Case 2
            Application.Run savedFuncName, savedArgs(0), savedArgs(1), savedArgs(2)
        Case 3
            Application.Run savedFuncName, savedArgs(0), savedArgs(1), savedArgs(2), savedArgs(3)
        Case 4
            Application.Run savedFuncName, savedArgs(0), savedArgs(1), savedArgs(2), savedArgs(3), savedArgs(4)
        Case Else
            '5個以上の引数はエラー扱いとする
            Call debugPrint("Too many arguments", savedFuncName & " in Repeater")
    End Select

Finally:
    '実行後は gCount をリセット
    gCount = 1
End Sub
