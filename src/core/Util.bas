Attribute VB_Name = "C_Util"
Option Explicit
Option Private Module

'/**
'  * ステータスバーの表示を変更する。(プログレスバー表示可能)
'  * str ..................... ステータスバーに表示する文字列。空白の場合はデフォルトに戻す。
'  * Count ................... 処理完了数。Max も必須。
'  * Max ..................... 処理最大数。Cnt も必須。
'  * Percent ................. 処理の進捗度。[0.00 - 1.00] Count、Max がセットされている場合は無効。
'  * NumDigitsAfterDecimal ... パーセント表示の小数点以下の桁数。(default: 0)
'  * ProgressBar ............. True/*False*: プログレスバーを表示する/表示しない
'  * Cer_per_Max ............. True/*False*: "( Count / Max )" を表示する/表示しない
'  */
Sub setStatusBar(Optional ByVal str As String = "", _
                 Optional ByVal Count As Long = -1, _
                 Optional ByVal Max As Long = -1, _
                 Optional ByVal Percent As Single = -1, _
                 Optional ByVal NumDigitsAfterDecimal As Byte = 0, _
                 Optional ByVal ProgressBar As Boolean = False, _
                 Optional ByVal Cnt_per_Max As Boolean = False)

    Const MAX_LEN As Byte = 13

    Dim txt As String
    Dim int_l As Integer, l As Single, det As Byte
    Static last As Single

    If str = "" Then
        Application.StatusBar = False
        Exit Sub
    End If

    If ProgressBar Then
        If Count >= 0 And Max >= Count Then
            Percent = Count / Max
        End If

        Percent = Percent * 100

        If Percent < 0 Or 100 < Percent Then
            Application.StatusBar = False
            Exit Sub
        End If

        l = Percent * (MAX_LEN / 100)
        int_l = Int(l)
        det = Round((l - int_l) * 8)

        txt = ChrW(&H2595)
        txt = txt & String(int_l, ChrW(&H2588))

        If det = 0 And MAX_LEN > int_l Then
            txt = txt & ChrW(&H2003)
        ElseIf det > 0 Then
            txt = txt & ChrW(&H2590 - det)
        End If

        If MAX_LEN > int_l Then
            txt = txt & String(MAX_LEN - int_l - 1, ChrW(&H2003))
        End If
        txt = txt & ChrW(&H258F)

        txt = "        進捗:" & txt & " " & _
            Format(WorksheetFunction.RoundDown(Percent, NumDigitsAfterDecimal), _
            "0" & Choose((NumDigitsAfterDecimal > 0) + 2, "." & String(NumDigitsAfterDecimal, "0"), "")) & " %"

        If Cnt_per_Max And Count >= 0 And Max >= Count Then
            txt = txt & " ( " & Count & " / " & Max & " )"
        End If
    End If

    If Not ProgressBar Or Timer - last > 0.1 Then
        Application.StatusBar = str & txt
        last = Timer
    End If
End Sub

Sub setStatusBarTemporarily(ByVal str As String, ByVal seconds As Byte)
    Dim i As Integer
    Dim startDate As Date
    Static lastRegisterTime As Double

    startDate = Date

    On Error Resume Next
    Call Application.OnTime(lastRegisterTime, "setStatusBar", , False)
    On Error GoTo 0

    lastRegisterTime = CDbl(startDate) + (Timer + seconds) / 86400

    Call setStatusBar(STATUS_PREFIX & str)
    Call Application.OnTime(lastRegisterTime, "setStatusBar")
End Sub

Function reMatch(ByVal str As String, ByVal Pattern As String, _
                 Optional ByVal IgnoreCase As Boolean = False, _
                 Optional ByVal Global_ As Boolean = True, _
                 Optional ByVal Multiline As Boolean = False) As Boolean

    Dim re As RegExp

    Set re = New RegExp
    With re
        .Global = Global_
        .IgnoreCase = IgnoreCase
        .Multiline = Multiline
        .Pattern = Pattern
        reMatch = .test(str)
    End With

    Set re = Nothing
End Function

Function reSearch(ByVal str As String, ByVal Pattern As String, _
                 Optional ByVal IgnoreCase As Boolean = False, _
                 Optional ByVal Global_ As Boolean = True, _
                 Optional ByVal Multiline As Boolean = False) As String

    Dim re As RegExp
    Dim mc As MatchCollection

    Set re = New RegExp
    With re
        .Global = Global_
        .IgnoreCase = IgnoreCase
        .Multiline = Multiline
        .Pattern = Pattern

        Set mc = .Execute(str)

        If mc.Count = 0 Then
            reSearch = ""
        Else
            reSearch = mc(0).Value
        End If
    End With

    Set re = Nothing
    Set mc = Nothing
End Function

Function reReplace(ByVal str As String, ByVal Pattern As String, ByVal replaceStr As String, _
                 Optional ByVal IgnoreCase As Boolean = False, _
                 Optional ByVal Global_ As Boolean = True, _
                 Optional ByVal Multiline As Boolean = False) As String

    Dim re As RegExp

    Set re = New RegExp
    With re
        .Global = Global_
        .IgnoreCase = IgnoreCase
        .Multiline = Multiline
        .Pattern = Pattern

        reReplace = .Replace(str, replaceStr)
    End With

    Set re = Nothing
End Function

Function getWorkbookIndex(ByVal targetWorkbook As Workbook) As Integer
    Dim i As Integer

    If targetWorkbook Is Nothing Then
        getWorkbookIndex = 0
        Exit Function
    End If

    For i = 1 To Workbooks.Count
        If Workbooks(i).FullName = targetWorkbook.FullName Then
            getWorkbookIndex = i
            Exit For
        End If
    Next i
End Function

'#####################################################################################'
' Source: https://mohayonao.hatenadiary.org/entry/20080617/1213712469                 '
'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv'
'
' 和集合
' Union2(ParamArray ArgList() As Variant) As Range
'
' 積集合
' Intersect2(ParamArray ArgList() As Variant) As Range
'
' 差集合
' Except2(ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range
'
' セル範囲の反転
' Invert2(ByRef SourceRange As Variant) As Range
'
'
'# 複数のセル ArgList の和集合を返す
'# Application.Union の拡張版 Nothing でもOK
Public Function Union2(ParamArray ArgList() As Variant) As Range

    Dim buf As Range

    Dim i As Long
    For i = 0 To UBound(ArgList)
        If TypeName(ArgList(i)) = "Range" Then
            If buf Is Nothing Then
                Set buf = ArgList(i)
            Else
                Set buf = Application.Union(buf, ArgList(i))
            End If
        End If
    Next

    Set Union2 = buf

End Function


'# 複数のセル ArgList の積集合を返す
'# Application.Intersect の拡張版 Nothing でもOK
Public Function Intersect2(ParamArray ArgList() As Variant) As Range

    Dim buf As Range

    Dim i As Long

    For i = 0 To UBound(ArgList)
        If Not TypeName(ArgList(i)) = "Range" Then
            Exit Function
        ElseIf buf Is Nothing Then
            Set buf = ArgList(i)
        Else
            Set buf = Application.Intersect(buf, ArgList(i))
        End If

        If buf Is Nothing Then Exit Function
    Next

    Set Intersect2 = buf

End Function


'# SourceRange から ArgList を差し引いた差集合を返す
'# (SourceRange と 反転した ArgList との積集合を返す)
Public Function Except2 _
    (ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range

    If TypeName(SourceRange) = "Range" Then

        Dim buf As Range

        Set buf = SourceRange

        Dim i As Long

        For i = 0 To UBound(ArgList)
            If TypeName(ArgList(i)) = "Range" Then
                Set buf = Intersect2(buf, Invert2(ArgList(i)))
            End If
        Next

        Set Except2 = buf

    End If

End Function


'# SourceRange の選択範囲を反転する
Public Function Invert2(ByRef SourceRange As Variant) As Range

    If Not TypeName(SourceRange) = "Range" Then Exit Function

    Dim Sh As Worksheet
    Set Sh = SourceRange.Parent

    Dim buf As Range
    Set buf = SourceRange.Parent.Cells

    Dim a As Range
    For Each a In SourceRange.Areas

        Dim AreaTop    As Long
        Dim AreaBottom As Long
        Dim AreaLeft   As Long
        Dim AreaRight  As Long

        AreaTop = a.Row
        AreaBottom = AreaTop + a.Rows.Count - 1
        AreaLeft = a.Column
        AreaRight = AreaLeft + a.Columns.Count - 1


        '■□□
        '■×□
        '■□□  ■の部分
        Dim RangeLeft   As Range
        Set RangeLeft = GetRangeWithPosition(Sh, _
            Sh.Cells.Row, Sh.Cells.Column, Sh.Rows.Count, AreaLeft - 1)
        '   Top           Left             Bottom         Right

        '□□■
        '□×■
        '□□■  ■の部分
        Dim RangeRight  As Range
        Set RangeRight = GetRangeWithPosition(Sh, _
            Sh.Cells.Row, AreaRight + 1, Sh.Rows.Count, Sh.Columns.Count)
        '   Top           Left           Bottom         Right


        '□■□
        '□×□
        '□□□  ■の部分
        Dim RangeTop    As Range
        Set RangeTop = GetRangeWithPosition(Sh, _
            Sh.Cells.Row, AreaLeft, AreaTop - 1, AreaRight)
        '   Top           Left      Bottom       Right


        '□□□
        '□×□
        '□■□  ■の部分
        Dim RangeBottom As Range
        Set RangeBottom = GetRangeWithPosition(Sh, _
            AreaBottom + 1, AreaLeft, Sh.Rows.Count, AreaRight)
        '   Top              Left      Bottom         Right


        Set buf = Intersect2(buf, _
            Union2(RangeLeft, RangeRight, RangeTop, RangeBottom))

    Next

    Set Invert2 = buf

End Function


'# 四隅を指定して Range を得る
Private Function GetRangeWithPosition(ByRef Sh As Worksheet, _
    ByVal Top As Long, ByVal Left As Long, _
    ByVal Bottom As Long, ByVal Right As Long) As Range

    '# 無効条件
    If Top > Bottom Or Left > Right Then
        Exit Function
    ElseIf Top < 0 Or Left < 0 Then
        Exit Function
    ElseIf Bottom > Cells.Rows.Count Or Right > Cells.Columns.Count Then
        Exit Function
    End If

    Set GetRangeWithPosition _
        = Sh.Range(Sh.Cells(Top, Left), Sh.Cells(Bottom, Right))

End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
' Source: https://mohayonao.hatenadiary.org/entry/20080617/1213712469                 '
'#####################################################################################'


Sub debugPrint(ByVal str As String, Optional ByVal funcName As String = "")
    If Not gDebugMode Then
        Exit Sub
    End If

    If funcName <> "" Then
        funcName = "[" & funcName & "] "
    End If

    Call setStatusBarTemporarily("[DEBUG] " & funcName & str, 5)
    Debug.Print "[" & Now & "] [DEBUG] " & funcName & str
End Sub
