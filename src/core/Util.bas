Attribute VB_Name = "C_Util"
Option Explicit
Option Private Module

#If Win64 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (x As Currency) As Boolean
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (x As Currency) As Boolean
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (x As Currency) As Boolean
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (x As Currency) As Boolean
#End If

Private Ctr1 As Currency
Private Ctr2 As Currency
Private Freq As Currency
Private Overhead As Currency

Private savedStatusMsg As String

'/*
' * Initializes the timer by querying the performance frequency and counters.
' * This should be called before using the timer functions.
' */
Public Sub TimeClear()
    QueryPerformanceFrequency Freq
    QueryPerformanceCounter Ctr1
    QueryPerformanceCounter Ctr2

    Overhead = Ctr2 - Ctr1 ' Determine API overhead

    QueryPerformanceCounter Ctr1 ' Time loop
End Sub

'/*
' * Gets the elapsed time using the QueryPerformanceCounter function.
' *
' * @param {String} [vFormat="0.0000"] - The format for the returned time value.
' * @returns {Double} - The elapsed time in seconds.
' */
Public Function GetQueryPerformanceTime(Optional ByVal vFormat As String = "0.0000") As Double
    Dim result As Currency
    QueryPerformanceCounter Ctr2

    ' Remove API overhead and perform unit conversion if needed
    result = (Ctr2 - Ctr1 - Overhead) / Freq

    ' Return the elapsed time with the specified format
    GetQueryPerformanceTime = Format(CDbl(result), vFormat)

    ' Reset the timer
    Call TimeClear
End Function

'/*
' * Sets the status bar text with optional progress bar.
' *
' * @param {String} [str=""] - The text for the status bar.
' * @param {Long} [count=-1] - The current count for the progress bar.
' * @param {Long} [max=-1] - The maximum count for the progress bar.
' * @param {Double} [percent=-1] - The percentage (0.0-1.0) of completion for the progress bar.
' * @param {Byte} [numDigitsAfterDecimal=0] - Number of digits after the decimal point in the percentage.
' * @param {Boolean} [progressBar=False] - Flag to enable or disable the progress bar.
' * @param {Boolean} [countPerMax=False] - Flag to display count and max in the status bar.
' */
Sub SetStatusBar(Optional ByVal str As String = "", _
                 Optional ByVal currentCount As Long = -1, _
                 Optional ByVal maximumCount As Long = -1, _
                 Optional ByVal percent As Double = -1, _
                 Optional ByVal numDigitsAfterDecimal As Byte = 0, _
                 Optional ByVal progressBar As Boolean = False, _
                 Optional ByVal countPerMax As Boolean = False)

    On Error GoTo Catch

    Const PROG_BAR_LENGTH = 13

    Dim progBarText As String
    Dim progBarLength As Long
    Dim progBarValue As Double
    Dim decimalPart As Byte
    Static lastUpdateTime As Double

    ' If the text is empty, reset the status bar and exit the subroutine
    If str = "" Then
        savedStatusMsg = ""
        Application.StatusBar = False
        Exit Sub
    End If

    ' If the progress bar is enabled
    If progressBar Then
        ' Calculate the percentage of completion
        If currentCount >= 0 And maximumCount >= currentCount Then
            percent = currentCount / maximumCount
        End If

        ' Convert the percentage to a scale of 100
        percent = percent * 100

        ' If the percentage is less than 0 or greater than 100, reset the status bar and exit the subroutine
        If percent < 0 Or 100 < percent Then
            Application.StatusBar = False
            Exit Sub
        End If

        ' Calculate the length of the progress bar
        progBarValue = percent * (PROG_BAR_LENGTH / 100)
        progBarLength = CLng(progBarValue)
        decimalPart = Round((progBarValue - progBarLength) * 8)

        ' Create the progress bar string
        progBarText = ChrW(&H2595)
        progBarText = progBarText & String(progBarLength, ChrW(&H2588))

        ' Add spaces or partial blocks to the progress bar string as needed
        If decimalPart = 0 And PROG_BAR_LENGTH > progBarLength Then
            progBarText = progBarText & ChrW(&H2003)
        ElseIf decimalPart > 0 Then
            progBarText = progBarText & ChrW(&H2590 - decimalPart)
        End If

        ' Add spaces to the end of the progress bar string if needed
        If PROG_BAR_LENGTH > progBarLength Then
            progBarText = progBarText & String(PROG_BAR_LENGTH - progBarLength - 1, ChrW(&H2003))
        End If
        progBarText = progBarText & ChrW(&H258F)

        ' Add the percentage and progress bar to the status bar string
        progBarText = "        Progress:" & progBarText & " " & _
            Format(WorksheetFunction.RoundDown(percent, numDigitsAfterDecimal), _
            "0" & Choose((numDigitsAfterDecimal > 0) + 2, "." & String(numDigitsAfterDecimal, "0"), "")) & " %"

        ' If countPerMax is enabled and count is greater than or equal to 0 and max is greater than or equal to count
        If countPerMax And currentCount >= 0 And maximumCount >= currentCount Then
            ' Add the count and max to the status bar string
            progBarText = progBarText & " ( " & currentCount & " / " & maximumCount & " )"
        End If
    End If

    ' If the progress bar is not enabled or the timer has exceeded 0.1 seconds
    If Not progressBar Or Timer - lastUpdateTime > 0.1 Then
        ' Update the status bar
        savedStatusMsg = str & progBarText
        Application.StatusBar = str & progBarText
        lastUpdateTime = Timer
    End If

Catch:
End Sub

'/*
' * Sets the status bar text temporarily with optional delay.
' *
' * @param {String} str - The text for the status bar.
' * @param {Long} miliseconds - The delay in miliseconds.
' * @param {Boolean} [disablePrefix=False] - Flag to disable the prefix in the status bar.
' */
Sub SetStatusBarTemporarily(ByVal str As String, _
                            ByVal miliseconds As Long, _
                   Optional ByVal disablePrefix As Boolean = False)

    Dim startDate As Date
    Static lastRegisterTime As Double

    ' Get the current date
    startDate = Date

    ' Try to cancel the previous OnTime event
    On Error Resume Next
    Call Application.OnTime(lastRegisterTime, "'SetStatusBar """ & savedStatusMsg & """'", , False)
    On Error GoTo Catch

    ' Calculate the time for the next OnTime event
    lastRegisterTime = CDbl(startDate) + (Timer + miliseconds / 1000) / 86400

    ' If disablePrefix is enabled, set the status bar with the string
    If disablePrefix Then
        Application.StatusBar = str
    Else
        ' Otherwise, set the status bar with the prefix and the string
        Application.StatusBar = gVim.Config.StatusPrefix & str
    End If

    ' Register the next OnTime event
    Call Application.OnTime(lastRegisterTime, "'SetStatusBar """ & savedStatusMsg & """'")

Catch:
End Sub

'/*
' * Checks if there is a match for the regular expression pattern in the input string.
' *
' * @param {String} str - The input string.
' * @param {String} matchPattern - The regular expression pattern.
' * @param {Boolean} [isIgnoreCase=False] - Flag to enable case-insensitive matching.
' * @param {Boolean} [isGlobal=True] - Flag to enable global matching.
' * @param {Boolean} [isMultiline=False] - Flag to enable multiline mode.
' * @returns {Boolean} - True if there is a match, False otherwise.
' */
Function RegExpMatch(ByVal str As String, ByVal matchPattern As String, _
            Optional ByVal isIgnoreCase As Boolean = False, _
            Optional ByVal isGlobal As Boolean = True, _
            Optional ByVal isMultiline As Boolean = False) As Boolean

    Dim re As RegExp
    Set re = New RegExp

    With re
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .Multiline = isMultiline
        .Pattern = matchPattern

        RegExpMatch = .Test(str)
    End With

    Set re = Nothing
End Function

'/*
' * Searches for the first occurrence of the regular expression pattern in the input string.
' *
' * @param {String} str - The input string.
' * @param {String} matchPattern - The regular expression pattern.
' * @param {Boolean} [isIgnoreCase=False] - Flag to enable case-insensitive matching.
' * @param {Boolean} [isGlobal=True] - Flag to enable global searching.
' * @param {Boolean} [isMultiline=False] - Flag to enable multiline mode.
' * @returns {String} - The matched string or an empty string if no match.
' */
Function RegExpSearch(ByVal str As String, ByVal matchPattern As String, _
             Optional ByVal isIgnoreCase As Boolean = False, _
             Optional ByVal isGlobal As Boolean = True, _
             Optional ByVal isMultiline As Boolean = False) As String

    Dim re As RegExp
    Dim mc As MatchCollection

    Set re = New RegExp
    With re
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .Multiline = isMultiline
        .Pattern = matchPattern

        Set mc = .Execute(str)

        If mc.Count = 0 Then
            RegExpSearch = ""
        Else
            RegExpSearch = mc(0).Value
        End If
    End With

    Set re = Nothing
    Set mc = Nothing
End Function

'/*
' * Replaces occurrences of the regular expression pattern in the input string with the specified replacement.
' *
' * @param {String} str - The input string.
' * @param {String} matchPattern - The regular expression pattern.
' * @param {String} replaceStr - The replacement string.
' * @param {Boolean} [isIgnoreCase=False] - Flag to enable case-insensitive matching.
' * @param {Boolean} [isGlobal=True] - Flag to enable global replacing.
' * @param {Boolean} [isMultiline=False] - Flag to enable multiline mode.
' * @returns {String} - The input string with replacements applied.
' */
Function RegExpReplace(ByVal str As String, ByVal matchPattern As String, ByVal replaceStr As String, _
              Optional ByVal isIgnoreCase As Boolean = False, _
              Optional ByVal isGlobal As Boolean = True, _
              Optional ByVal isMultiline As Boolean = False) As String

    Dim re As RegExp
    Set re = New RegExp

    With re
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .Multiline = isMultiline
        .Pattern = matchPattern

        RegExpReplace = .Replace(str, replaceStr)
    End With

    Set re = Nothing
End Function

'/*
' * Returns the index of the target workbook in the Workbooks collection.
' *
' * @param {Workbook} targetWorkbook - The target workbook.
' * @returns {Long} - The index of the target workbook. Returns 0 if the target workbook is Nothing.
' */
Function GetWorkbookIndex(ByVal targetWorkbook As Workbook) As Long
    Dim i As Long

    If targetWorkbook Is Nothing Then
        GetWorkbookIndex = 0
        Exit Function
    End If

    For i = 1 To Workbooks.Count
        If Workbooks(i).FullName = targetWorkbook.FullName Then
            GetWorkbookIndex = i
            Exit Function
        End If
    Next i
End Function

'/*
' * Checks if a worksheet with the specified name exists in the active workbook.
' *
' * @param {String} targetSheetName - The name of the target worksheet.
' * @returns {Boolean} - True if the worksheet exists, False otherwise.
' */
Function IsSheetExists(ByVal targetSheetName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In Worksheets
        If ws.Name = targetSheetName Then
            IsSheetExists = True
            Exit Function
        End If
    Next
End Function

'/*
' * Returns the count of visible sheets in the active workbook.
' *
' * @returns {Long} - The count of visible sheets.
' */
Function GetVisibleSheetsCount() As Long
    Dim ws As Worksheet

    For Each ws In Worksheets
        If ws.Visible = xlSheetVisible Then
            GetVisibleSheetsCount = GetVisibleSheetsCount + 1
        End If
    Next
End Function

'/*
' * Converts a hex color code to a long integer.
' *
' * @param {String} colorCode - The hex color code to convert.
' * @returns {Long} - The corresponding long integer color value, or -1 if the conversion fails.
' */
Function HexColorCodeToLong(ByVal colorCode As String) As Long
    If colorCode Like "*[!0-9a-fA-F]*" Then
        ' Invalid characters in the color code
        HexColorCodeToLong = -1
    ElseIf Len(colorCode) = 3 Then
        ' Convert 3-digit hex color code to long
        HexColorCodeToLong = Val("&H" & Mid(colorCode, 3, 1) & Mid(colorCode, 3, 1) & _
            Mid(colorCode, 2, 1) & Mid(colorCode, 2, 1) & Mid(colorCode, 1, 1) & Mid(colorCode, 1, 1) & "&")
    ElseIf Len(colorCode) = 6 Then
        ' Convert 6-digit hex color code to long
        HexColorCodeToLong = Val("&H" & Mid(colorCode, 5, 2) & Mid(colorCode, 3, 2) & Mid(colorCode, 1, 2) & "&")
    Else
        ' Invalid length of the color code
        HexColorCodeToLong = -1
    End If
End Function

'/*
' * Converts a long integer color code to a 6-digit hex string.
' *
' * @param {Long} colorCode - The long integer color code to convert.
' * @returns {String} - The corresponding 6-digit hex color code string in lowercase.
' */
Function ColorCodeToHex(ByVal colorCode As Long) As String
    ' Convert long integer color code to 6-digit hex string
    ColorCodeToHex = Right("0" & Hex(colorCode Mod 256), 2) & _
                     Right("0" & Hex(colorCode ¥ 256 Mod 256), 2) & _
                     Right("0" & Hex(colorCode ¥ 256 ¥ 256), 2)
    ColorCodeToHex = LCase(ColorCodeToHex)
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

'文字列中の出現回数を返す
Function StrCount(baseStr As String, chkStr As String) As Long
    Dim n As Long: n = 0
    Dim ret As Long: ret = 0
    Do
        n = InStr(n + 1, baseStr, chkStr)
        If n = 0 Then
            Exit Do
        Else
            ret = ret + 1
        End If
    Loop
    StrCount = ret
End Function

'文字列中のN回目に現れた位置を返す
Function StrNPos(baseStr As String, chkStr As String, ByVal n As Long) As Long
    Dim i As Long: i = 0
    Dim l As Long: l = 0
    For i = 1 To n
        l = InStr(l + 1, baseStr, chkStr)
        If l = 0 Then
            Exit For
        End If
        StrNPos = l
    Next i
End Function

Sub DebugPrint(ByVal str As String, Optional ByVal funcName As String = "")
    If Not gVim.Config.DebugMode Then
        Exit Sub
    End If

    If funcName <> "" Then
        funcName = "[" & funcName & "] "
    End If

    Call SetStatusBarTemporarily("[DEBUG] " & funcName & str, 5000)
    Debug.Print "[" & Now & "] [DEBUG] " & funcName & str
End Sub

Function ErrorHandler(Optional ByVal funcName As String = "") As Boolean
    Dim strMsg As String

    If Err.Number <> 0 Then
        strMsg = "[ERROR] "
        If funcName <> "" Then
            strMsg = strMsg & funcName & ": "
        End If
        strMsg = strMsg & Err.Description & " (" & Err.Number & ")"

        Call SetStatusBarTemporarily(strMsg, 5000)
        Debug.Print "[" & Now & "] " & strMsg

        Err.Clear
        ErrorHandler = True
    End If
End Function
