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
Private isVisibleTempMsg As Boolean
Private Const APPLY_SAVED_MSG = "## APPLY SAVED MESSAGE ##"

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

    ' Restore status bar message if the input string matches the "APPLY_SAVED_MSG"
    If str = APPLY_SAVED_MSG Then
        isVisibleTempMsg = False
        Call SetStatusBar(savedStatusMsg)
        Exit Sub

    ' If a temporary message is visible, save the current string as the new saved message
    ElseIf isVisibleTempMsg Then
        savedStatusMsg = str
        Exit Sub

    ' If the input string is empty, clear the status bar
    ElseIf str = "" Then
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
    Static lastRegisterTime

    ' Get the current date
    startDate = Date

    ' Try to cancel the previous OnTime event
    On Error Resume Next
    Call Application.OnTime(lastRegisterTime, "'SetStatusBar """ & APPLY_SAVED_MSG & """'", , False)
    On Error GoTo Catch

    ' Set the visibility flag for temporary messages to True
    isVisibleTempMsg = True

    ' Calculate the time for the next OnTime event
    lastRegisterTime = startDate + CDec(Timer + miliseconds / 1000) / 86400

    ' If disablePrefix is enabled, set the status bar with the string
    If disablePrefix Then
        Application.StatusBar = str
    Else
        ' Otherwise, set the status bar with the prefix and the string
        Application.StatusBar = gVim.Config.StatusPrefix & str
    End If

    ' Register the next OnTime event
    Call Application.OnTime(lastRegisterTime, "'SetStatusBar """ & APPLY_SAVED_MSG & """'")

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
' * Checks if the input string starts with any of the prefixes in the provided parameter.
' * The parameter can be either a single string or an array of string.
' * If the parameter type is not String or String(), an error is raised.
' *
' * @param {String} str - The string to check.
' * @param {Variant} prefixes - A single prefix string or an array of prefix strings to check for.
' * @returns {Boolean} - True if the string starts with any of the prefixes, False otherwise.
' */
Function StartsWith(ByRef str As String, ByVal prefixes As Variant) As Boolean
    Dim i As Long

    ' Check the type of prefixes
    If Not (VarType(prefixes) = vbString Or IsArray(prefixes)) Then
        Err.Raise 5, , "Type mismatch: 'prefixes' must be either a String or String()"
    End If

    ' Check if prefixes is an array
    If IsArray(prefixes) Then
        ' Loop through the array of prefixes and check if any match the beginning of the text
        For i = LBound(prefixes) To UBound(prefixes)
            ' Check the type of prefix
            If VarType(prefixes(i)) = vbString Then
                If InStr(str, prefixes(i)) = 1 Then ' InStr starts searching from position 1
                    StartsWith = True
                    Exit Function
                End If
            End If
        Next i
    Else
        ' If a single string is provided, check if it matches the beginning of the text
        If InStr(str, prefixes) = 1 Then
            StartsWith = True
        Else
            StartsWith = False
        End If
    End If
End Function

'/*
' * Checks if the input string ends with any of the suffixes in the provided parameter.
' * The parameter can be either a single string or an array of string.
' * If the parameter type is not String or String(), an error is raised.
' *
' * @param {String} str - The string to check.
' * @param {Variant} suffixes - A single suffix string or an array of suffix strings to check for.
' * @returns {Boolean} - True if the string ends with any of the suffixes, False otherwise.
' */
Function EndsWith(ByRef str As String, ByVal suffixes As Variant) As Boolean
    Dim i As Long
    Dim textLen As Long
    Dim suffixLen As Long

    ' Check the type of suffixes
    If Not (VarType(suffixes) = vbString Or IsArray(suffixes)) Then
        Err.Raise 5, , "Type mismatch: 'suffixes' must be either a String or String()"
    End If

    ' Check if suffixes is an array
    If IsArray(suffixes) Then
        ' Loop through the array of suffixes and check if any match the end of the text
        For i = LBound(suffixes) To UBound(suffixes)
            suffixLen = Len(suffixes(i))
            textLen = Len(str)

            ' Check the type of suffix
            If VarType(suffixes(i)) = vbString Then
                ' Using InStrRev to search from the end of the string
                If InStrRev(str, suffixes(i), textLen) = textLen - suffixLen + 1 Then
                    EndsWith = True
                    Exit Function
                End If
            End If
        Next i
    Else
        ' If a single string is provided, check if it matches the end of the text
        suffixLen = Len(suffixes)
        textLen = Len(str)

        If InStrRev(str, suffixes, textLen) = textLen - suffixLen + 1 Then
            EndsWith = True
        Else
            EndsWith = False
        End If
    End If
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
    Dim ws As Object

    For Each ws In Sheets
        If ws.Visible = xlSheetVisible Then
            GetVisibleSheetsCount = GetVisibleSheetsCount + 1
        End If
    Next
End Function

'/*
' * Retrieves the list of files and subfolders from the specified folder.
' *
' * @param {String} folderPath - The path of the folder to list files and subfolders from.
' * @returns {Collection} - A collection of file and subfolder names.
' */
Function DirGrob(ByVal folderPath As String) As Collection
    Dim fso As FileSystemObject
    Dim objFolder As folder
    Dim objFile As file
    Dim objSubFolder As folder

    ' Setup Collection
    Set DirGrob = New Collection

    ' Create FileSystemObject
    Set fso = New FileSystemObject

    Dim sepIndex As Long
    Dim lastPart As String
    folderPath = Replace(folderPath, "/", "¥")
    sepIndex = InStrRev(folderPath, "¥")

    lastPart = LCase(Mid(folderPath, sepIndex + 1))
    folderPath = Left(folderPath, sepIndex)

    ' Ignore errors
    On Error Resume Next

    ' Ensure the folder exists before proceeding
    If fso.FolderExists(folderPath) Then
        Set objFolder = fso.GetFolder(folderPath)

        ' List subfolders in the folder
        For Each objSubFolder In objFolder.SubFolders
            If StartsWith(LCase(fso.GetFileName(objSubFolder.Path)), lastPart) Then
                DirGrob.Add fso.GetFileName(objSubFolder.Path) & "/" ' Append "/" to indicate it's a folder
            End If
        Next objSubFolder

        ' List files in the folder
        For Each objFile In objFolder.Files
            If StartsWith(LCase(fso.GetFileName(objFile.Path)), lastPart) Then
                DirGrob.Add fso.GetFileName(objFile.Path) ' Add the file name
            End If
        Next objFile
    End If

    ' Release the FileSystemObject
    Set fso = Nothing
End Function

'/*
' * Converts a relative path to an absolute path.
' *
' * @param {String} cwd - The current working directory.
' * @param {String} relativePath - The relative path to be converted.
' * @returns {String} - The corresponding absolute path.
' */
Function GetAbsolutePath(ByRef cwd As String, ByRef relativePath As String) As String
    ' Declare variables
    Dim fso As FileSystemObject
    Dim fullPath As String

    ' Create FileSystemObject
    Set fso = New FileSystemObject

    ' Combine the current workbook path with the relative path and get the absolute path
    fullPath = cwd & "¥" & relativePath
    GetAbsolutePath = fso.GetAbsolutePathName(fullPath)

    ' Release the FileSystemObject
    Set fso = Nothing
End Function

'/*
' * Resolves the provided relative or special path into an absolute path.
' *
' * @param {String} strPath - The relative or special path that needs to be resolved.
' * @returns {String} - The resolved absolute path.
' */
Function ResolvePath(ByVal strPath As String) As String
    ' Replace path to windows style
    strPath = Replace(strPath, "/", "¥")

    ' Declare variable to store the resolved absolute path
    Dim absPath As String

    ' Check if the relative path starts with a backslash
    If StartsWith(strPath, "¥") Then
        ' Resolve the absolute path
        absPath = GetAbsolutePath("", Mid(strPath, 2))

    ' Check if the relative path starts with a tilde (‾), indicating the user profile directory
    ElseIf StartsWith(strPath, "‾¥") Then
        ' Resolve the absolute path using the user's profile directory
        absPath = GetAbsolutePath(Environ$("USERPROFILE"), Mid(strPath, 3))

    ' Otherwise, resolve the relative path using the active workbook's directory
    Else
        absPath = GetAbsolutePath(ActiveWorkbook.Path, strPath)
    End If

    ' If the relative path ends with a backslash, append it to the absolute path
    If Right(strPath, 1) = "¥" And Right(absPath, 1) <> "¥" Then
        absPath = absPath & "¥"
    End If

    ' Return the resolved absolute path
    ResolvePath = absPath
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
