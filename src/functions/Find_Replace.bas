Attribute VB_Name = "F_Find_Replace"
Option Explicit
Option Private Module

Function showFindFollowLang()
    UF_FindForm.Show
End Function

Function showFindNotFollowLang()
    gLangJa = Not gLangJa
    UF_FindForm.Show
    gLangJa = Not gLangJa
End Function


Function nextFoundCell()
    Dim t As Range
    Dim i As Integer

    If gCount > 1 Then
        Application.ScreenUpdating = False
    End If

    For i = gCount To 1 Step -1
        If gCount = 1 Then
            Application.ScreenUpdating = True
        End If

        Set t = Cells.FindNext(After:=ActiveCell)
        If Not t Is Nothing Then
            t.Activate
        Else
            Application.ScreenUpdating = True
            Exit Function
        End If

    Next i
End Function

Function previousFoundCell()
    Dim t As Range
    Dim i As Integer

    If gCount > 1 Then
        Application.ScreenUpdating = False
    End If

    For i = gCount To 1 Step -1
        If i = 1 Then
            Application.ScreenUpdating = True
        End If

        Set t = Cells.FindPrevious(After:=ActiveCell)
        If Not t Is Nothing Then
            t.Activate
        Else
            Application.ScreenUpdating = True
            Exit Function
        End If

    Next i
End Function

Function showReplaceWindow()
    Call keystroke(True, Alt_ + E_, E_)
End Function

