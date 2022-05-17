VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_FindForm 
   Caption         =   "Find"
   ClientHeight    =   306
   ClientLeft      =   42
   ClientTop       =   432
   ClientWidth     =   4578
   OleObjectBlob   =   "FindForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_FindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_Find_Click()
    Dim inputText As String
    Dim findString As String
    Dim t As Range
    
    inputText = TextBox1.Value
    If inputText = "/" Then
        Call nextFoundCell
    Else
        If inputText Like "/*" Then
            findString = Mid(inputText, 2)
        Else
            findString = inputText
        End If
        
        Set t = ActiveSheet.Cells.Find(What:=findString, _
                                       LookIn:=xlValues, _
                                       LookAt:=xlPart, _
                                       SearchOrder:=xlByColumns, _
                                       MatchByte:=False)
        If Not t Is Nothing Then
            ActiveWorkbook.ActiveSheet.Activate
            t.Activate
        End If
    End If
    
    TextBox1.Value = "/"
    Me.Hide
End Sub

Private Sub TextBox1_Change()
    If Len(TextBox1.Value) < 1 Then
        TextBox1.Value = "/"
    End If
    
    If InStr(TextBox1.Value, Chr(9)) > 0 Then
        TextBox1.Value = Replace(TextBox1.Value, Chr(9), "")
    End If
    
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Or (Shift = 2 And KeyCode = 219) Then 'Escape
        Me.TextBox1.Value = "/"
        Me.Hide
    End If
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Find"
        .StartUpPosition = 0
        .Top = 0
        .Left = 0
    End With
      
    With TextBox1
        .Value = ""
        .Multiline = False
        .EnterKeyBehavior = True
    End With
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Top = 570
    Me.Left = 80
    
    With Me.TextBox1
        .Value = "/"
        If gLangJa Then
            .IMEMode = fmIMEModeHiragana
        Else
            .IMEMode = fmIMEModeOff
        End If
        .SetFocus
    End With
End Sub
