VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_KeymapManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cls_Vim
Private cKeyMap As Dictionary
Private cOnKeyMap As Dictionary

Private REDIRECT_TABLE_US  As Dictionary
Private REDIRECT_TABLE_JIS As Dictionary

' cKeyMap structure
' {
'   "n - {h}": "MoveLeft",                                  # Normal mode & "h" key  => Map to "MoveLeft" procedure
'   "n - {d} [SUMMARY]": "¥n{d} - {d}¥n{d} - {c} ..."       # Normal mode & "d" key summary  => List of available subkeys
'   "n - {r} {TEMP]": "rr SelectRow¥nrd DeleteRow¥n ..."    # Normal mode & "r" key temp  => Unresolved key list
' }

'/*
' * Sets key mappings for the specified mode.
' *
' * @param {String} mode_ - The mode for which the key mapping is set.
' * @param {String} lhs - The left-hand side key in Vim style.
' * @param {String} rhs - The right-hand side key in Vim like style.
' * @param {Boolean} reMap - Indicates whether to remap the key.
' * @param {Boolean} cmdLine - Indicates whether the key mapping is for the command line.
' * @param {Variant} args() - Optional arguments.
' */
Public Sub Set_(ByRef mode_ As String, ByRef lhs As String, ByRef rhs As String, _
                ByVal reMap As Boolean, ByVal cmdLine As Boolean, ParamArray args())

    On Error GoTo Catch

    ' Check the mode
    If Not Parent.Mode.Exists(mode_) And Not cmdLine Then
        Err.Raise 50000, Description:="Unsupported mode: " & mode_
    End If

    ' Make dict key
    Dim lhsKey As String
    Dim strKey As String
    Dim keyPrefix As String: keyPrefix = mode_ & KEY_SEPARATOR
    Dim firstKey As String: firstKey = ""
    Dim sepIndex As Long: sepIndex = 0

    If cmdLine Then
        lhsKey = lhs

    Else
        ' Convert Vim style key to SendKeys style
        lhsKey = VimToVBA(lhs, KEY_SEPARATOR)

        sepIndex = InStr(lhsKey, KEY_SEPARATOR)
        firstKey = lhsKey
        If sepIndex > 0 Then
            firstKey = Left(lhsKey, sepIndex - 1)
        End If

        ' Register the first key for non-command line modes to ensure Ctrl series OnKey execution
        Call RegisterOnKey(mode_, firstKey)
    End If
    strKey = keyPrefix & lhsKey

    ' Make dict value
    Dim strValue As String
    If reMap Then
        strValue = "'KeyStroke " & VimToVBA(rhs, ",", resultType:=VBAKeyCodes) & "'"
    ElseIf rhs = "" Then
        ' Unmap
        strValue = ""
    Else
        strValue = FormatArgs(rhs, args)
    End If

    ' Set the key mapping in the keymap table
    If cKeyMap.Exists(strKey) Then
        cKeyMap(strKey) = strValue
    Else
        cKeyMap.Add strKey, strValue
    End If

    ' For mappings with two or more keys, add to SUMMARY
    If sepIndex > 0 Or cmdLine Then
        strKey = keyPrefix
        If Not cmdLine Then
            strKey = strKey & firstKey
        End If
        strKey = strKey & KEY_SUMMARY

        If cKeyMap.Exists(strKey) Then
            cKeyMap(strKey) = cKeyMap(strKey) & KEY_TERM_SYMBOL & lhsKey
        Else
            cKeyMap.Add strKey, KEY_TERM_SYMBOL & lhsKey
        End If
    End If
    Exit Sub

Catch:
    Call ErrorHandler("KeymapManager.Set_")
End Sub

'/*
' * Formats arguments for Application.Run based on the function name and arguments array.
' *
' * @param {String} funcName - The name of the function.
' * @param {Variant} args - The array of arguments.
' * @return {String} - The formatted argument string for Application.Run.
' */
Private Function FormatArgs(ByVal funcName As String, ByVal args As Variant) As String
    If IsMissing(args) Or Not IsArray(args) Then
        FormatArgs = funcName
        Exit Function
    End If

    If UBound(args) = -1 Then
        FormatArgs = funcName
        Exit Function
    ElseIf Right(TypeName(args(0)), 2) = "()" Then
        FormatArgs = FormatArgs(funcName, args(0))
        Exit Function
    End If

    Dim i As Long
    For i = LBound(args) To UBound(args)
        Select Case TypeName(args(i))
        Case "String"
            FormatArgs = FormatArgs & """" & args(i) & """" & ","
        Case "Integer", "Long", "LongLong", "Double", "Single", "Byte", "Boolean"
            FormatArgs = FormatArgs & args(i) & ","
        Case Else
            Call DebugPrint("Unsupported argument type: " & TypeName(args(i)), "FormatArgs")
        End Select
    Next i
    FormatArgs = "'" & funcName & " " & Left(FormatArgs, Len(FormatArgs) - 1) & "'"
End Function

'/*
' * Parses a Vim-style map string and adds it to the TEMP keymap.
' *
' * @param {String} mapStr - The Vim-style map string to parse.
' */
Public Sub Map(ByVal mapStr As String)
    ' <mode>map [<cmd>]<lhs> [<key>]<rhs>

    On Error GoTo Catch

    ' Split the string into parts based on spaces (using InStr for optimization).
    Dim buf(0 To 2) As String
    Dim i As Long
    Dim spaceIndex As Long: spaceIndex = 1
    Dim lastSpaceIndex As Long: lastSpaceIndex = 1

    ' Trim and replace tabs with spaces for consistency
    mapStr = Trim(Replace(mapStr, vbTab, " "))

    For i = LBound(buf) To UBound(buf)
        ' Determine the next space index
        If i = UBound(buf) Then
            spaceIndex = Len(mapStr) + 1
        Else
            spaceIndex = InStr(lastSpaceIndex, mapStr, " ")
        End If

        ' Exit for loop if space missing
        If spaceIndex = 0 Then
            buf(i) = Mid(mapStr, lastSpaceIndex)
            Exit For
        End If

        ' Extract the current part of the string
        buf(i) = Mid(mapStr, lastSpaceIndex, spaceIndex - lastSpaceIndex)

        ' Update the last space index, skipping consecutive spaces
        lastSpaceIndex = spaceIndex + 1
        While i < UBound(buf) And Mid(mapStr, lastSpaceIndex, 1) = " "
            lastSpaceIndex = lastSpaceIndex + 1
        Wend
    Next i

    ' First element (mode)
    buf(0) = LCase(buf(0))

    ' Check for map style
    If InStr(buf(0), "map") = 0 Then
        Err.Raise 50000, Description:="Syntax error: does not match map style"
    End If

    ' Extract mode (remove "map" and use MODE_NORMAL if unspecified)
    Dim mode_ As String: mode_ = Replace(buf(0), "map", "")
    Dim unmap As Boolean: unmap = False

    If mode_ Like "*un" Then
        ' Unmap
        unmap = True
        mode_ = Replace(mode_, "un", "")
    ElseIf buf(2) = "" Then
        ' Check for syntax error (missing parameters)
        Err.Raise 50000, Description:="Syntax error: parameters are missing"
    End If

    If Len(mode_) > 1 Then
        Err.Raise 50000, Description:="Syntax error: mode indicator must be 1 character"
    ElseIf mode_ = "" Then
        mode_ = MODE_NORMAL
    End If

    ' Validate mode
    If Not Parent.Mode.Exists(mode_) Then
        Err.Raise 50000, Description:="Mode error: mode """ & mode_ & """ is invalid"
    End If

    ' Second element (lhs)
    Dim lhs As String: lhs = buf(1)
    Dim cmdLine As Boolean: cmdLine = False

    ' Check if lhs starts with KEY_CMD
    Dim firstKey As String
    If InStr(LCase(lhs), KEY_CMD) = 1 Then
        ' Remove KEY_CMD prefix and set cmdLine flag
        lhs = Mid(lhs, Len(KEY_CMD) + 1)
        firstKey = ""
        mode_ = KEY_CMD
        cmdLine = True

    Else
        ' Convert lhs to VBA format (firstKey)
        firstKey = VimToVBA(buf(1), KEY_SEPARATOR, True)

        ' This is necessary for Ctrl-series OnKey to work (not mapped by default)
        Call RegisterOnKey(mode_, firstKey, unmap:=unmap)
    End If

    ' Unmap
    If unmap Then
        Call Del_(mode_, lhs, cmdLine)
        Exit Sub
    End If

    ' Third element (rhs)
    Dim rhs As String: rhs = buf(2)

    ' Append to TEMP keymap
    Dim strKey As String: strKey = mode_ & KEY_SEPARATOR & firstKey & KEY_TEMP
    Dim strValue As String: strValue = lhs & " " & rhs & KEY_TERM_SYMBOL
    If cKeyMap.Exists(strKey) Then
        cKeyMap(strKey) = cKeyMap(strKey) & strValue
    Else
        cKeyMap.Add strKey, strValue
    End If
    Exit Sub

Catch:
    Call ErrorHandler("KeymapManager.Map")
End Sub

'/*
' * Removes a key mapping from the keymap table.
' *
' * @param {String} mode_ - The mode of the key mapping.
' * @param {String} lhs - The left-hand side key of the mapping.
' * @param {Boolean} [cmdLine=False] - Flag to indicate if the mapping is for command line mode.
' */
Public Sub Del_(ByVal mode_ As String, ByVal lhs As String, _
       Optional ByVal cmdLine As Boolean = False)
    ' Make dictionary key
    Dim lhsKey As String
    Dim strKey As String
    Dim keyPrefix As String: keyPrefix = mode_ & KEY_SEPARATOR
    Dim firstKey As String: firstKey = ""
    Dim sepIndex As Long: sepIndex = 0

    If cmdLine Then
        lhsKey = lhs
    Else
        ' Convert Vim style key to SendKeys style
        lhsKey = VimToVBA(lhs, KEY_SEPARATOR)

        sepIndex = InStr(lhsKey, KEY_SEPARATOR)
        firstKey = lhsKey
        If sepIndex > 0 Then
            firstKey = Left(lhsKey, sepIndex - 1)
        End If
    End If
    strKey = keyPrefix & lhsKey

    ' Delete the key mapping from the keymap table
    If cKeyMap.Exists(strKey) Then
        cKeyMap.Remove strKey
    End If

    ' For mappings with two or more keys, delete from SUMMARY
    If sepIndex > 0 Or cmdLine Then
        strKey = keyPrefix
        If Not cmdLine Then
            strKey = strKey & firstKey
        End If
        strKey = strKey & KEY_SUMMARY

        If cKeyMap.Exists(strKey) Then
            cKeyMap(strKey) = Replace(cKeyMap(strKey), KEY_TERM_SYMBOL & lhsKey, "")
        End If
    End If

    ' Delete from TEMP keymap
    strKey = mode_ & KEY_SEPARATOR & firstKey & KEY_TEMP

    If cKeyMap.Exists(strKey) Then
        Dim searchStr As String: searchStr = lhs & " "
        Dim searchIdx As Long: searchIdx = InStr(cKeyMap(strKey), searchStr)

        If searchIdx > 0 Then
            Dim termIdx As Long: termIdx = InStr(searchIdx, cKeyMap(strKey), KEY_TERM_SYMBOL)
            If termIdx = 0 Then
                cKeyMap(strKey) = ""
            Else
                cKeyMap(strKey) = Replace(cKeyMap(strKey), Mid(cKeyMap(strKey), searchIdx, termIdx - searchIdx + 1), "")
            End If

            If Len(cKeyMap(strKey)) = 0 Then
                cKeyMap.Remove strKey
            End If
        End If
    End If
End Sub

'/*
' * Parses the TEMP key and calls Set_ for each parsed mapping.
' *
' * @param {String} mode_ - The mode to be parsed from the TEMP keymap.
' * @param {String} key - The key to be parsed from the TEMP keymap.
' */
Private Sub LazyParseKeys(ByVal mode_ As String, ByVal key As String)
    ' Get the TEMP key for the provided key
    Dim tempKey As String: tempKey = mode_ & KEY_SEPARATOR
    If mode_ <> KEY_CMD Then
        tempKey = tempKey & key
    End If
    tempKey = tempKey & KEY_TEMP

    ' Check if the TEMP key exists
    If Not cKeyMap.Exists(tempKey) Then
        Dim tmpLen As Long
        tmpLen = InStr(key, KEY_SEPARATOR)
        If tmpLen > 0 Then
            Call LazyParseKeys(mode_, Left(key, tmpLen - 1))
            If Not cKeyMap.Exists(tempKey) Then
                Exit Sub
            End If
        End If
    End If

    ' TEMP key contains mappings separated by KEY_TERM_SYMBOL
    Dim tempStr As String: tempStr = cKeyMap(tempKey)
    Dim strLen As Long: strLen = Len(tempStr)
    Dim seekStart As Long: seekStart = 1
    Dim seekEnd As Long: seekEnd = 0

    ' Iterate through mappings in TEMP key
    While seekStart <= strLen
        ' Find the next occurrence of KEY_TERM_SYMBOL
        seekEnd = InStr(seekStart, tempStr, KEY_TERM_SYMBOL)

        ' Extract the mapping substring
        Dim targetStr As String
        targetStr = Mid(tempStr, seekStart, seekEnd - seekStart)

        ' Split the mapping into lhs and rhs based on the first space
        Dim spaceIndex As Long: spaceIndex = InStr(targetStr, " ")
        Dim lhs As String: lhs = Left(targetStr, spaceIndex - 1)
        Dim rhs As String: rhs = Mid(targetStr, spaceIndex + 1)

        ' Check if lhs starts with KEY_CMD
        Dim cmdLine As Boolean: cmdLine = (mode_ = KEY_CMD)

        ' Check if rhs starts with KEY_REMAP
        Dim reMap As Boolean: reMap = False
        If InStr(LCase(rhs), KEY_REMAP) = 1 Then
            reMap = True
            rhs = Mid(rhs, Len(KEY_REMAP) + 1)
        End If

        ' Parse arguments from rhs
        Dim args As Variant: args = ParseArgsString(rhs)

        ' Call Set_ to apply the mapping
        Call Set_(mode_, lhs, rhs, reMap, cmdLine, args)

        ' Move to the next mapping
        seekStart = seekEnd + 1
    Wend

    ' Remove the TEMP key
    cKeyMap.Remove tempKey
End Sub

'/*
' * Parses the arguments string from rhs and returns an array of arguments.
' *
' * @param {String} rhs - The input string containing arguments.
' * @returns {Variant()} - An array of parsed arguments.
' */
Private Function ParseArgsString(ByRef rhs As String) As Variant()
    On Error GoTo Catch

    ' Initialize variables
    Dim spaceIndex As Long: spaceIndex = InStr(rhs, " ")
    Dim quoteIndex As Long: quoteIndex = InStr(rhs, """")
    Dim ret() As Variant

    ' Check if there are no spaces and quotes in the input string
    If spaceIndex = 0 And quoteIndex = 0 Then
        ParseArgsString = Array()
        Exit Function
    End If

    Dim targetStr As String: targetStr = rhs
    rhs = ""

    Dim inQuotes As Boolean: inQuotes = False
    Dim seekPos As Long: seekPos = 1
    Dim strLen As Long: strLen = Len(targetStr)

    ' Loop through the input string
    Do While seekPos <= strLen
        Dim currentToken As String

        If (quoteIndex = 0 Or spaceIndex < quoteIndex) And spaceIndex > 0 And Not inQuotes Then
            ' Extract token until the next space
            currentToken = Mid(targetStr, seekPos, spaceIndex - seekPos)
            seekPos = spaceIndex + 1

            ' Skip consecutive spaces
            While Mid(targetStr, seekPos, 1) = " "
                seekPos = seekPos + 1
            Wend

            ' Update spaceIndex for the next iteration
            spaceIndex = InStr(seekPos, targetStr, " ")

        ElseIf quoteIndex > 0 Then
            Dim afterQuoteChar As String: afterQuoteChar = Mid(targetStr, quoteIndex + 1, 1)
            Dim beforeQuoteChar As String: beforeQuoteChar = ""

            ' Check characters before and after the quote
            If quoteIndex > 1 Then
                beforeQuoteChar = Mid(targetStr, quoteIndex - 1, 1)
            End If

            ' Handle quote positions
            If afterQuoteChar = """" Then
                ' Two consecutive quotes, include in the token
                currentToken = currentToken & Mid(targetStr, seekPos, quoteIndex - seekPos + 1)
                seekPos = quoteIndex + 2
            ElseIf Not inQuotes And (beforeQuoteChar = " " Or beforeQuoteChar = "") Then
                ' Start of quoted string
                seekPos = quoteIndex + 1
                inQuotes = True
            ElseIf inQuotes And (afterQuoteChar = " " Or afterQuoteChar = "") Then
                ' End of quoted string
                currentToken = currentToken & Mid(targetStr, seekPos, quoteIndex - seekPos)
                seekPos = quoteIndex + 1

                ' Skip consecutive spaces after the quote
                While Mid(targetStr, seekPos, 1) = " "
                    seekPos = seekPos + 1
                Wend

                ' Reset inQuotes flag
                inQuotes = False
            Else
                ' Invalid quote position
                Err.Raise 50000, Description:="Syntax error: invalid quote("") position"
            End If

            ' Update spaceIndex and quoteIndex for the next iteration
            spaceIndex = InStr(seekPos, targetStr, " ")
            quoteIndex = InStr(seekPos, targetStr, """")

        Else
            ' Extract the remaining token if no spaces or quotes are found
            currentToken = Mid(targetStr, seekPos, strLen - seekPos + 1)
            seekPos = strLen + 1
        End If

        ' Check if not inQuotes to add the token to the result array
        If Not inQuotes Then
            If rhs = "" Then
                rhs = currentToken
            Else
                If (Not Not ret) = 0 Then
                    ' Uninitialized array
                    ReDim ret(0 To 0)
                Else
                    ReDim Preserve ret(LBound(ret) To UBound(ret) + 1)
                End If
                ret(UBound(ret)) = currentToken
            End If
            currentToken = ""
        End If
    Loop

    ' Check for an uneven number of quotes
    If inQuotes Then
         Err.Raise 50000, Description:="Syntax error: invalid number of quotes("")"
    End If

    ' Return the parsed arguments array
    ParseArgsString = ret

    Exit Function

Catch:
    Call ErrorHandler("KeymapManager.ParseArgsString")
End Function

'/*
' * Gets the value associated with the specified key from the keymap table.
' *
' * @param {String} key - The key for which to retrieve the value.
' * @param {Boolean} [cmdLine=False] - A flag indicating whether the key is associated with command-line mode.
' * @returns {String} - The value associated with the key or DUMMY_PROCEDURE if not found.
' */
Public Function Get_(ByVal key As String, Optional cmdLine As Boolean = False) As String
    ' Fix the key according to the keyboard type
    key = FixWithKeyboardType(key)

    ' Determine the mode based on the current mode or command-line mode
    Dim mode_ As String

    If cmdLine Then
        mode_ = KEY_CMD
    Else
        mode_ = Parent.Mode.Current
    End If

    ' Search for the key in the keymap table
    Get_ = SearchKey(key, mode_)
End Function

'/*
' * Searches for the specified key in the keymap table.
' *
' * @param {String} key - The key to search for.
' * @param {String} mode_ - The mode in which to search for the key.
' * @returns {String} - The value associated with the key or DUMMY_PROCEDURE if not found.
' */
Private Function SearchKey(ByVal key As String, ByVal mode_ As String) As String
    ' Check if the mode is empty and return DUMMY_PROCEDURE if true
    If mode_ = "" Then
        SearchKey = ""
        Exit Function
    End If

    ' Construct the key to search in the keymap table
    Dim strKey As String: strKey = mode_ & KEY_SEPARATOR & key

    ' Lazily parse keys to handle TEMP bindings
    Call LazyParseKeys(mode_, key)

    ' Check if the exact key exists in the keymap table
    If cKeyMap.Exists(strKey) Then
        SearchKey = cKeyMap(strKey)
        Exit Function
    ElseIf cKeyMap.Exists(strKey & KEY_SUMMARY) Then
        ' If SUMMARY exists, call SHOW_CMD_PROCEDURE for subkeys
        key = Replace(key, "{""}", "{""""}")
        key = Replace(key, "{'}", "{''}")
        SearchKey = "'" & SHOW_CMD_PROCEDURE & " """ & key & """'"
        Exit Function
    End If

    ' If the key is not found, recursively search in the fallback mode
    If mode_ = MODE_NORMAL Then
        SearchKey = DUMMY_PROCEDURE
    Else
        SearchKey = SearchKey(key, Parent.Mode.KeyMapFallback(mode_))
    End If
End Function

'/*
' * Splits the result of suggested keys based on the termination symbol.
' *
' * @param {String} key - The key for which to retrieve suggested keys.
' * @returns {String()} - An array of suggested keys.
' */
Public Function Suggest(ByVal key As String, Optional ByVal cmdLine As Boolean = False) As String()
    Dim mode_ As String
    If cmdLine Then
        mode_ = KEY_CMD
    Else
        mode_ = Parent.Mode.Current
    End If

    Dim suggestResult As String: suggestResult = SearchSuggestKeys(key, mode_)
    Suggest = Split(suggestResult, KEY_TERM_SYMBOL)
End Function

'/*
' * Checks if the specified key is still valid based on suggested keys.
' *
' * @param {String} key - The key to check for validity.
' * @returns {Boolean} - True if the key is still valid, False otherwise.
' */
Public Function IsStillValid(ByVal key As String) As Boolean
    IsStillValid = Len(SearchSuggestKeys(key, Parent.Mode.Current, True)) > 0
End Function

'/*
' * Searches for suggested keys based on the specified key and mode.
' *
' * @param {String} key - The key for which to retrieve suggested keys.
' * @param {String} mode_ - The mode in which to search for suggested keys.
' * @param {Boolean} [checkOnly=False] - Flag to indicate whether to perform only a validity check.
' * @returns {String} - The result of suggested keys.
' */
Private Function SearchSuggestKeys(ByVal key As String, ByVal mode_ As String, _
                          Optional ByVal checkOnly As Boolean = False) As String
    If mode_ = "" Then
        SearchSuggestKeys = ""
        Exit Function
    End If

    ' Extract the first key for search
    Dim strKey As String
    If mode_ = KEY_CMD Then
        strKey = mode_ & KEY_SEPARATOR
    Else
        Dim sepIndex As Long: sepIndex = InStr(key, KEY_SEPARATOR)
        Dim firstKey As String: firstKey = key
        If sepIndex > 0 Then
            firstKey = Left(key, sepIndex - 1)
        End If
        strKey = mode_ & KEY_SEPARATOR & firstKey
    End If

    ' Lazily parse keys to handle TEMP bindings
    Call LazyParseKeys(mode_, strKey)

    ' Construct the summary key
    Dim summaryKey As String: summaryKey = strKey & KEY_SUMMARY

    ' Check if SUMMARY exists and search for the key
    If cKeyMap.Exists(summaryKey) Then
        Dim summaryStr As String: summaryStr = cKeyMap(summaryKey)
        Dim searchStr As String:  searchStr = KEY_TERM_SYMBOL & key
        Dim seekStart As Long:    seekStart = 1
        Dim searchIndex As Long:  searchIndex = InStr(seekStart, summaryStr, searchStr)

        Do While searchIndex > 0
            If checkOnly Then
                SearchSuggestKeys = "True"
                Exit Function
            End If

            seekStart = searchIndex + 1
            searchIndex = InStr(seekStart, summaryStr, KEY_TERM_SYMBOL)
            If searchIndex = 0 Then
                SearchSuggestKeys = SearchSuggestKeys & Mid(summaryStr, seekStart)
                Exit Do
            Else
                SearchSuggestKeys = SearchSuggestKeys & Mid(summaryStr, seekStart, searchIndex - seekStart + 1)
            End If
            seekStart = searchIndex
            searchIndex = InStr(seekStart, summaryStr, searchStr)
        Loop
    End If

    ' Recursive search in the fallback mode
    SearchSuggestKeys = SearchSuggestKeys & SearchSuggestKeys(key, Parent.Mode.KeyMapFallback(mode_), checkOnly)

    ' Remove the trailing termination symbol if present
    If SearchSuggestKeys Like "*" & KEY_TERM_SYMBOL Then
        SearchSuggestKeys = Left(SearchSuggestKeys, Len(SearchSuggestKeys) - 1)
    End If
End Function

'/*
' * Registers or remove the specified key in the OnKeyMap table.
' *
' * @param {String} mode_ - The mode to register.
' * @param {String} firstKey - The first key to register.
' * @param {Boolean} [unmap=False] - Flag to unmap
' */
Private Sub RegisterOnKey(ByVal mode_ As String, ByVal firstKey As String, Optional ByVal unmap As Boolean = False)
    If cOnKeyMap.Count = 0 Then
        Call InitDummyKey
    End If

    Dim targetModeBit As Long: targetModeBit = Parent.Mode.ModeBit(mode_)

    If unmap Then
        If cOnKeyMap.Exists(firstKey) Then
            cOnKeyMap(firstKey) = cOnKeyMap(firstKey) And (Not targetModeBit)
        End If
    Else
        If cOnKeyMap.Exists(firstKey) Then
            cOnKeyMap(firstKey) = cOnKeyMap(firstKey) Or targetModeBit
        Else
            cOnKeyMap.Add firstKey, targetModeBit
        End If
    End If
End Sub

Private Sub InitDummyKey()
    Dim dummyModeBit As Long: dummyModeBit = Parent.Mode.ModeBit(MODE_DUMMY)

    With cOnKeyMap
        .Add "{a}", dummyModeBit
        .Add "{b}", dummyModeBit
        .Add "{c}", dummyModeBit
        .Add "{d}", dummyModeBit
        .Add "{e}", dummyModeBit
        .Add "{f}", dummyModeBit
        .Add "{g}", dummyModeBit
        .Add "{h}", dummyModeBit
        .Add "{i}", dummyModeBit
        .Add "{j}", dummyModeBit
        .Add "{k}", dummyModeBit
        .Add "{l}", dummyModeBit
        .Add "{m}", dummyModeBit
        .Add "{n}", dummyModeBit
        .Add "{o}", dummyModeBit
        .Add "{p}", dummyModeBit
        .Add "{q}", dummyModeBit
        .Add "{r}", dummyModeBit
        .Add "{s}", dummyModeBit
        .Add "{t}", dummyModeBit
        .Add "{u}", dummyModeBit
        .Add "{v}", dummyModeBit
        .Add "{w}", dummyModeBit
        .Add "{x}", dummyModeBit
        .Add "{y}", dummyModeBit
        .Add "{z}", dummyModeBit

        .Add "0", dummyModeBit
        .Add "1", dummyModeBit
        .Add "2", dummyModeBit
        .Add "3", dummyModeBit
        .Add "4", dummyModeBit
        .Add "5", dummyModeBit
        .Add "6", dummyModeBit
        .Add "7", dummyModeBit
        .Add "8", dummyModeBit
        .Add "9", dummyModeBit

        'Symbols
        .Add "{ }", dummyModeBit
        .Add "{!}", dummyModeBit
        .Add "{""}", dummyModeBit
        .Add "{#}", dummyModeBit
        .Add "{$}", dummyModeBit
        .Add "{%}", dummyModeBit
        .Add "{&}", dummyModeBit
        .Add "{'}", dummyModeBit
        .Add "{(}", dummyModeBit
        .Add "{)}", dummyModeBit
        .Add "{*}", dummyModeBit
        .Add "{+}", dummyModeBit
        .Add "{,}", dummyModeBit
        .Add "{-}", dummyModeBit
        .Add "{.}", dummyModeBit
        .Add "{/}", dummyModeBit
        .Add "{:}", dummyModeBit
        .Add "{;}", dummyModeBit
        .Add "{<}", dummyModeBit
        .Add "{=}", dummyModeBit
        .Add "{>}", dummyModeBit
        .Add "{?}", dummyModeBit
        .Add "{@}", dummyModeBit
        .Add "{[}", dummyModeBit
        .Add "{¥}", dummyModeBit
        .Add "{]}", dummyModeBit
        .Add "{^}", dummyModeBit
        .Add "{_}", dummyModeBit
        .Add "{`}", dummyModeBit
        .Add "{{}", dummyModeBit
        .Add "{|}", dummyModeBit
        .Add "{}}", dummyModeBit
        .Add "{‾}", dummyModeBit
        .Add "{226}", dummyModeBit
        .Add "+0", dummyModeBit

        'Shift
        .Add "+{a}", dummyModeBit
        .Add "+{b}", dummyModeBit
        .Add "+{c}", dummyModeBit
        .Add "+{d}", dummyModeBit
        .Add "+{e}", dummyModeBit
        .Add "+{f}", dummyModeBit
        .Add "+{g}", dummyModeBit
        .Add "+{h}", dummyModeBit
        .Add "+{i}", dummyModeBit
        .Add "+{j}", dummyModeBit
        .Add "+{k}", dummyModeBit
        .Add "+{l}", dummyModeBit
        .Add "+{m}", dummyModeBit
        .Add "+{n}", dummyModeBit
        .Add "+{o}", dummyModeBit
        .Add "+{p}", dummyModeBit
        .Add "+{q}", dummyModeBit
        .Add "+{r}", dummyModeBit
        .Add "+{s}", dummyModeBit
        .Add "+{t}", dummyModeBit
        .Add "+{u}", dummyModeBit
        .Add "+{v}", dummyModeBit
        .Add "+{w}", dummyModeBit
        .Add "+{x}", dummyModeBit
        .Add "+{y}", dummyModeBit
        .Add "+{z}", dummyModeBit
    End With

End Sub

Private Sub Class_Initialize()
    Set cKeyMap = New Dictionary
    Set cOnKeyMap = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set cKeyMap = Nothing
    Set cOnKeyMap = Nothing
    Set REDIRECT_TABLE_US = Nothing
    Set REDIRECT_TABLE_JIS = Nothing
    Set Parent = Nothing
End Sub


'/*
' * Converts Vim-style key sequences to VBA-compatible SendKeys format.
' *
' * @param {String} lhs - The input string in Vim-style key notation.
' * @param {String} sep - The separator used between keys in the output.
' * @param {Boolean} [firstKeyOnly] - Flag to include only the first key in the output.
' * @returns {String} - The converted key sequence in VBA-compatible format.
' */
Public Function VimToVBA(ByVal lhs As String, ByVal sep As String, _
                Optional ByVal firstKeyOnly As Boolean = False, _
                Optional ByVal resultType As eResultType = VBASendkeys) As String

    ' Parameter validation
    If Len(lhs) = 0 Then
        Err.Raise 50000, Description:="lhs must be at least 1 char"
    ElseIf Len(sep) < 1 Then
        Err.Raise 50000, Description:="sep must be at least 1 chars"
    End If

    Dim char As String
    Dim converted As String

    Dim i As Long
    For i = 1 To Len(lhs)
        char = Mid(lhs, i, 1)

        ' Normal char
        If char <> "<" Then
            If resultType = VBASendkeys Then
                converted = ConvertSendKeysStyle(char)
            ElseIf resultType = VBAKeyCodes Then
                converted = ConvertKeyCodes(char)
            End If
            GoTo Continue
        End If

        ' Check end of special char
        Dim endOfGt As Long
        endOfGt = InStr(i, lhs, ">")

        If endOfGt = 0 Then
            If resultType = VBASendkeys Then
                converted = ConvertSendKeysStyle(char)
            ElseIf resultType = VBAKeyCodes Then
                converted = ConvertKeyCodes(char)
            End If
            GoTo Continue
        End If

        ' Resolve special char
        Dim components As Variant
        components = Split(Mid(lhs, i + 1, endOfGt - i - 1), "-")

        Dim ub As Long
        ub = UBound(components)

        Dim isCtrl As Boolean: isCtrl = False
        Dim isShift As Boolean: isShift = False
        Dim isAlt As Boolean: isAlt = False

        If ub > 0 Then
            Dim j As Long

            For j = 0 To ub - 1
                Select Case LCase(components(j))
                Case "c"
                    isCtrl = True
                Case "s"
                    isShift = True
                Case "m", "a"
                    isAlt = True
                Case Else
                    Err.Raise 50000, Description:="Unsupported key prefix: " & components(j)
                    Exit Function
                End Select
            Next j
        End If

        If resultType = VBASendkeys Then
            converted = ConvertSendKeysStyle(LCase(components(ub)), isCtrl, isShift, isAlt)
        ElseIf resultType = VBAKeyCodes Then
            converted = ConvertKeyCodes(LCase(components(ub)), isCtrl, isShift, isAlt)
        End If
        i = endOfGt
Continue:

        If VimToVBA = "" Then
            VimToVBA = converted
        Else
            VimToVBA = VimToVBA & sep & converted
        End If

        If firstKeyOnly Then
            Exit For
        End If
    Next i
End Function

'/*
' * Converts key codes to the corresponding SendKeys string.
' *
' * @param {String} key - The key code to be converted.
' * @param {Boolean} isCtrl - True if the Ctrl key is pressed, False otherwise.
' * @param {Boolean} isShift - True if the Shift key is pressed, False otherwise.
' * @param {Boolean} isAlt - True if the Alt key is pressed, False otherwise.
' * @returns {String} - The SendKeys string representing the key code.
' */
Private Function ConvertKeyCodes(ByVal key As String, _
                        Optional ByVal isCtrl As Boolean, _
                        Optional ByVal isShift As Boolean, _
                        Optional ByVal isAlt As Boolean) As String

    ' Convert key to lowercase
    Dim lowerKey As String: lowerKey = LCase(key)

    ' ASCII code of the key
    Dim charCode As Long: charCode = Asc(lowerKey)

    ' Flag indicating if the keyboard is JIS/US
    Dim isJis As Boolean: isJis = Parent.Config.IsJisKeyboard
    Dim isUs As Boolean: isUs = Not isJis

    ' Resulting key code
    Dim result As Long: result = 0

    ' Shift enable if uppercase character
    If Len(key) = 1 And key <> lowerKey Then
        isShift = True
    End If

    ' Handle special keys
    Dim i As Long
    Select Case lowerKey
        Case "bs": result = BackSpace_
        Case "tab": result = Tab_
        Case "esc": result = Escape_
        Case "del": result = Delete_
        Case "up": result = Up_
        Case "down": result = Down_
        Case "left": result = Left_
        Case "right": result = Right_
        Case "insert": result = Insert_
        Case "home": result = Home_
        Case "end": result = End_
        Case "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f12", "f13", "f14", "f15", "f16"
            i = CLng(Mid(lowerKey, 2))
            result = F1_ - 1 + i
        Case "nl", "cr", "return", "enter": result = Enter_
        Case "space": result = Space_
        Case "lt", "<"
            result = Comma_
            isShift = True
        Case "bslash", "¥": result = Backslash_
        Case "bar", "|"
            result = Backslash_
            isShift = True
        Case "pageup": result = PageUp_
        Case "pagedown": result = PageDown_
        Case "kplus": result = NumpadAdd_
        Case "kminus": result = NumpadSubtract_
        Case "kmultiply": result = NumpadMultiply_
        Case "kdivide": result = NumpadDivide_
        Case "kenter": result = NumpadEnter_
        Case "kpoint": result = NumpadDecimal_
        Case "k0", "k1", "k2", "k3", "k4", "k5", "k6", "k7", "k8", "k9"
            i = CLng(Mid(lowerKey, 2, 1))
            result = Numpad0_ + i

        Case "+", ",": result = Equal_US_
        Case "-": result = Minus_
        Case ".", ">"
            result = Period_
            isShift = (lowerKey = ">")
        Case "/", "?"
            result = Slash_
            isShift = (lowerKey = "?")
        Case "[", "{"
            result = OpeningSquareBracket_
            isShift = (lowerKey = "{")
        Case "]", "}"
            result = ClosingSquareBracket_
            isShift = (lowerKey = "}")

        ' Handle keys specific to US keyboard
        Case isUs
            Select Case lowerKey
                Case ";", ":"
                    result = Semicoron_US_
                    isShift = (lowerKey = ":")
                Case "=": result = Equal_US_
                Case "_"
                    result = Minus_
                    isShift = True
                Case "`", "‾"
                    result = BackQuote_US_
                    isShift = (lowerKey = "‾")
                Case "'", """"
                    result = SingleQuote_US_
                    isShift = (lowerKey = """")
                Case "@"
                    result = k2_
                    isShift = True
                Case "^"
                    result = k6_
                    isShift = True
                Case "*"
                    result = k8_
                    isShift = True
                Case "&", "(", ")"
                    result = charCode + 17
                    isShift = True
            End Select

        ' Handle keys specific to JIS keyboard
        Case isJis
            Select Case lowerKey
                Case ";": result = Semicoron_JIS_
                Case ":", "*"
                    result = Coron_JIS_
                    isShift = (lowerKey = "*")
                Case "=":
                    result = Minus_
                    isShift = True
                Case "@", "`"
                    result = AtMark_JIS_
                    isShift = (lowerKey = "`")
                Case "^", "‾"
                    result = Caret_JIS_
                    isShift = (lowerKey = "‾")
                Case "_"
                    result = Underscore_
                    isShift = True
            End Select

        ' Handle keys specific to US and JIS keyboards
        Case "!", """", "#", "$", "%", "&", "'", "(", ")"
            result = charCode + 16
            isShift = True

    End Select

    ' If result is still 0, it means the key is not a special key
    If result = 0 Then
        ' Handle alphanumeric keys
        If 48 <= charCode And charCode <= 57 Then  '0-9
            result = charCode
        ElseIf 97 <= charCode And charCode <= 122 Then  'a-z
            result = charCode - 32
        Else
            ' Unexpected key
            Err.Raise 50000, Description:="Unexpected key: " & key
        End If
    End If

    ' Apply modifiers
    If isCtrl Then result = result + Ctrl_
    If isShift Then result = result + Shift_
    If isAlt Then result = result + Alt_

    ' Convert the result to a string and return
    ConvertKeyCodes = CStr(result)
End Function


'/*
' * Converts a single key to VBA-compatible SendKeys format.
' *
' * @param {String} key - The input key to be converted.
' * @param {Boolean} isCtrl - Flag indicating if Ctrl is pressed.
' * @param {Boolean} isShift - Flag indicating if Shift is pressed.
' * @param {Boolean} isAlt - Flag indicating if Alt is pressed.
' * @returns {String} - The converted key in VBA-compatible format.
' */
Public Function ConvertSendKeysStyle(ByVal key As String, _
                            Optional ByVal isCtrl As Boolean, _
                            Optional ByVal isShift As Boolean, _
                            Optional ByVal isAlt As Boolean) As String

    Dim lowerKey As String: lowerKey = LCase(key)

    If Len(key) = 1 Then
        If key <> lowerKey Then
            isShift = True
        End If

        If 48 <= Asc(key) And Asc(key) <= 57 Then  'is number
            ConvertSendKeysStyle = key
        Else
            ConvertSendKeysStyle = "{" & lowerKey & "}"
        End If
    Else
        Select Case lowerKey
        Case "bs", "tab", "esc", "del", "up", "down", "left", "right", "insert", "home", "end", _
             "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f12", "f13", "f14", "f15", "f16"
            ConvertSendKeysStyle = "{" & UCase(key) & "}"
        Case "nl", "cr", "return", "enter"
            ConvertSendKeysStyle = "{ENTER}"
        Case "space":     ConvertSendKeysStyle = "{ }"
        Case "lt":        ConvertSendKeysStyle = "{<}"
        Case "bslash":    ConvertSendKeysStyle = "{¥}"
        Case "bar":       ConvertSendKeysStyle = "{|}"
        Case "pageup":    ConvertSendKeysStyle = "{PGUP}"
        Case "pagedown":  ConvertSendKeysStyle = "{PGDN}"
        Case "kplus":     ConvertSendKeysStyle = "{107}"
        Case "kminus":    ConvertSendKeysStyle = "{109}"
        Case "kmultiply": ConvertSendKeysStyle = "{106}"
        Case "kdivide":   ConvertSendKeysStyle = "{111}"
        Case "kenter":    ConvertSendKeysStyle = "{108}"
        Case "kpoint":    ConvertSendKeysStyle = "{110}"
        Case "k0", "k1", "k2", "k3", "k4", "k5", "k6", "k7", "k8", "k9"
            ConvertSendKeysStyle = "{" & CStr(Numpad0_ + CLng(Mid(lowerKey, 2, 1))) & "}"
        Case Else
            Err.Raise 50000, Description:="Unsupported key name: " & key
        End Select
    End If

    If isAlt Then ConvertSendKeysStyle = "%" & ConvertSendKeysStyle
    If isShift Then ConvertSendKeysStyle = "+" & ConvertSendKeysStyle
    If isCtrl Then ConvertSendKeysStyle = "^" & ConvertSendKeysStyle

    ConvertSendKeysStyle = FixWithKeyboardType(ConvertSendKeysStyle)
End Function

'/*
' * Converts a key code and decorated key information to VBA-compatible SendKeys format.
' *
' * @param {Long} code - The key code.
' * @param {Integer} decoratedKey - The decorated key information.
' * @returns {String} - The converted key in VBA-compatible format.
' */
Public Function KeycodeToSendKeys(ByVal code As Long, ByVal decoratedKey As Integer) As String
    Dim isCtrl As Boolean: isCtrl = decoratedKey And 2
    Dim isShift As Boolean: isShift = decoratedKey And 1
    Dim isAlt As Boolean: isAlt = decoratedKey And 4
    Dim convertedKey As String

    convertedKey = GetKeyCharacter(code, isShift)
    If convertedKey = "" Then
        Exit Function
    End If

    isShift = isShift And Not ((96 <= code And code <= 111) Or code = 226)

    convertedKey = ConvertSendKeysStyle(convertedKey, isCtrl, isShift, isAlt)
    KeycodeToSendKeys = FixWithKeyboardType(convertedKey)
End Function

'/*
' * Retrieves the character corresponding to the key code.
' *
' * @param {Long} code - The key code.
' * @param {Boolean} isShift - Flag indicating if Shift is pressed.
' * @returns {String} - The corresponding character.
' */
Private Function GetKeyCharacter(ByVal code As Long, ByVal isShift As Boolean) As String
    Dim isJis As Boolean
    Dim isUs As Boolean
    isJis = Parent.Config.IsJisKeyboard
    isUs = Not isJis

    Select Case code
        ' ignore key codes
        Case 0 To 7, 10 To 12, 14 To 26, 28 To 31, 41 To 44, 47, 91 To 93, 144, 145, 160 To 185, 229:

        Case 8: GetKeyCharacter = "bs"
        Case 9: GetKeyCharacter = "tab"
        Case 13: GetKeyCharacter = "enter"
        Case 27: GetKeyCharacter = "esc"
        Case 33: GetKeyCharacter = "pageup"
        Case 34: GetKeyCharacter = "pagedown"
        Case 35: GetKeyCharacter = "end"
        Case 36: GetKeyCharacter = "home"
        Case 37: GetKeyCharacter = "left"
        Case 38: GetKeyCharacter = "up"
        Case 39: GetKeyCharacter = "right"
        Case 40: GetKeyCharacter = "down"
        Case 45: GetKeyCharacter = "insert"
        Case 46: GetKeyCharacter = "del"
        Case 106: GetKeyCharacter = "kmultiply"
        Case 107: GetKeyCharacter = "kplus"
        Case 108: GetKeyCharacter = "kenter"
        Case 109: GetKeyCharacter = "kminus"
        Case 110: GetKeyCharacter = "kpoint"
        Case 111: GetKeyCharacter = "kdivide"
        Case 186 And isUs: GetKeyCharacter = ";"
        Case 186 And isJis: GetKeyCharacter = ":"
        Case 187 And isUs: GetKeyCharacter = "="
        Case 187 And isJis: GetKeyCharacter = ";"
        Case 188: GetKeyCharacter = ","
        Case 189: GetKeyCharacter = "-"
        Case 190: GetKeyCharacter = "."
        Case 191: GetKeyCharacter = "/"
        Case 192 And isUs: GetKeyCharacter = "`"
        Case 192 And isJis: GetKeyCharacter = "@"
        Case 219: GetKeyCharacter = "["
        Case 220: GetKeyCharacter = "¥"
        Case 221: GetKeyCharacter = "]"
        Case 222 And isUs: GetKeyCharacter = "'"
        Case 222 And isJis: GetKeyCharacter = "^"
        Case 226 And isShift: GetKeyCharacter = "_"
        Case 226: GetKeyCharacter = "¥"
        Case 32, 48 To 57, 65 To 90
            GetKeyCharacter = LCase(Chr(code))
        Case 96 To 105
            GetKeyCharacter = "k" & Chr(code - 48)
        Case 112 To 127
            GetKeyCharacter = "f" & code - 111

        Case Else
            Call DebugPrint("Unsupported key code: " & code, "GetKeyCharacter")
    End Select
End Function

'/*
' * Applies keyboard type-specific corrections to the given key.
' *
' * @param {String} key - The key to be corrected.
' * @returns {String} - The corrected key.
' */
Private Function FixWithKeyboardType(ByVal key As String) As String
    ' Initialize the redirect tables if not already done
    Call InitializeRedirectTables

    ' Set the initial value to the input key
    FixWithKeyboardType = key

    ' Determine which redirect table to use based on the current keyboard type
    Dim redirectTable As Dictionary
    If Parent.Config.IsJisKeyboard Then
        Set redirectTable = REDIRECT_TABLE_JIS
    Else
        Set redirectTable = REDIRECT_TABLE_US
    End If

    ' Check if the key exists in the redirect table and update the result if necessary
    If Not redirectTable Is Nothing And redirectTable.Exists(key) Then
        FixWithKeyboardType = redirectTable(key)
    End If
End Function

'/*
' * Transforms the SendKeys-style key into display text.
' *
' * @param {String} key - The SendKeys-style key to be transformed.
' * @returns {String} - The display text representing the key.
' */
Public Function SendKeysToDisplayText(ByVal key As String) As String
    Dim char As String

    ' Only shift key pressed
    If Left(key, 2) = "+{" And Len(key) = 4 Then
        char = Mid(key, 3, 1)
        If 97 <= Asc(char) And Asc(char) <= 122 Then  'a-z
            SendKeysToDisplayText = UCase(char)
        Else
            Debug.Print "Unexpected case:", char
        End If

    ' Special chars
    ElseIf Right(key, 1) = "}" Then
        Dim braceIndex As Long
        braceIndex = InStr(key, "{")
        If Len(key) - braceIndex > 2 Then
            Dim keyName As String
            keyName = Mid(key, braceIndex + 1, Len(key) - braceIndex - 1)
            Select Case keyName
                Case "226": SendKeysToDisplayText = "^_"
                Case "96", "97", "98", "99", "100", "101", "102", "103", "104", "105"
                    SendKeysToDisplayText = "<k" & Chr(CLng(keyName) - 48) & ">"
                Case "106": SendKeysToDisplayText = "<kmultiply>"
                Case "107": SendKeysToDisplayText = "<kplus>"
                Case "108": SendKeysToDisplayText = "<kenter>"
                Case "109": SendKeysToDisplayText = "<kminus>"
                Case "110": SendKeysToDisplayText = "<kpoint>"
                Case "111": SendKeysToDisplayText = "<kdivide>"

                Case Else
                    SendKeysToDisplayText = Left(key, braceIndex - 1) & "<" & keyName & ">"
            End Select
        Else
            SendKeysToDisplayText = Left(key, braceIndex - 1)
            If braceIndex > 1 Then
                SendKeysToDisplayText = SendKeysToDisplayText & UCase(Mid(key, braceIndex + 1, 1))
            Else
                SendKeysToDisplayText = SendKeysToDisplayText & LCase(Mid(key, braceIndex + 1, 1))
            End If
        End If

    ' Normal chars
    Else
        SendKeysToDisplayText = key
    End If
End Function

'/*
' * Initializes the redirect tables if not already done.
' */
Private Sub InitializeRedirectTables()
    If REDIRECT_TABLE_US Is Nothing Then
        Set REDIRECT_TABLE_US = New Dictionary
        With REDIRECT_TABLE_US
            .Add "+{`}", "{‾}"
            .Add "+1", "{!}"
            .Add "+2", "{@}"
            .Add "+3", "{#}"
            .Add "+4", "{$}"
            .Add "+5", "{%}"
            .Add "+6", "{^}"
            .Add "+7", "{&}"
            .Add "+8", "{*}"
            .Add "+9", "{(}"
            .Add "+0", "{)}"
            .Add "+{-}", "{_}"
            .Add "+{=}", "{+}"
            .Add "+{[}", "{{}"
            .Add "+{]}", "{}}"
            .Add "+{¥}", "{|}"
            .Add "+{;}", "{:}"
            .Add "+{'}", "{""}"
            .Add "+{,}", "{<}"
            .Add "+{.}", "{>}"
            .Add "+{/}", "{?}"

            .Add "^{‾}", "^+{`}"
            .Add "^{!}", "^+1"
            .Add "^{@}", "^+2"
            .Add "^{#}", "^+3"
            .Add "^{$}", "^+4"
            .Add "^{%}", "^+5"
            .Add "^{^}", "^+6"
            .Add "^{&}", "^+7"
            .Add "^{*}", "^+8"
            .Add "^{(}", "^+9"
            .Add "^{)}", "^+0"
            .Add "^{_}", "^+{-}"
            .Add "^{+}", "^+{=}"
            .Add "^{{}", "^+{[}"
            .Add "^{}}", "^+{]}"
            .Add "^{|}", "^+{¥}"
            .Add "^{:}", "^+{;}"
            .Add "^{""}", "^+{'}"
            .Add "^{<}", "^+{,}"
            .Add "^{>}", "^+{.}"
            .Add "^{?}", "^+{/}"
        End With
    End If

    If REDIRECT_TABLE_JIS Is Nothing Then
        Set REDIRECT_TABLE_JIS = New Dictionary
        With REDIRECT_TABLE_JIS
            .Add "+1", "{!}"
            .Add "+2", "{""}"
            .Add "+3", "{#}"
            .Add "+4", "{$}"
            .Add "+5", "{%}"
            .Add "+6", "{&}"
            .Add "+7", "{'}"
            .Add "+8", "{(}"
            .Add "+9", "{)}"
            .Add "+{-}", "{=}"
            .Add "+{^}", "{‾}"
            .Add "+{¥}", "{|}"
            .Add "+{@}", "{`}"
            .Add "+{[}", "{{}"
            .Add "+{;}", "{+}"
            .Add "+{:}", "{*}"
            .Add "+{]}", "{}}"
            .Add "+{,}", "{<}"
            .Add "+{.}", "{>}"
            .Add "+{/}", "{?}"
            .Add "+{226}", "{_}"

            .Add "^{!}", "^+1"
            .Add "^{""}", "^+2"
            .Add "^{#}", "^+3"
            .Add "^{$}", "^+4"
            .Add "^{%}", "^+5"
            .Add "^{&}", "^+6"
            .Add "^{'}", "^+7"
            .Add "^{(}", "^+8"
            .Add "^{)}", "^+9"
            .Add "^{=}", "^+{-}"
            .Add "^{‾}", "^+{^}"
            .Add "^{|}", "^+{¥}"
            .Add "^{`}", "^+{@}"
            .Add "^{{}", "^+{[}"
            .Add "^{+}", "^+{;}"
            .Add "^{*}", "^+{:}"
            .Add "^{}}", "^+{]}"
            .Add "^{<}", "^+{,}"
            .Add "^{>}", "^+{.}"
            .Add "^{?}", "^+{/}"
            .Add "^{_}", "^+{226}"
        End With
    End If
End Sub

'/*
' * Unbinds all the previously bound keys using the OnKey method.
' */
Public Sub UnbindAll()
    Dim key As Variant

    For Each key In cOnKeyMap.Keys()
        Application.OnKey key
    Next key
End Sub

'/*
' * Binds all the keys from the OnKeyMap using the OnKey method.
' */
Public Sub BindAll()
    If cOnKeyMap.Count = 0 Then
        Call InitDummyKey
    End If

    Dim currentModeBit As Long
    currentModeBit = Parent.Mode.ModeBit(Parent.Mode.Current, isFallback:=True)

    Dim key As Variant
    For Each key In cOnKeyMap.Keys()
        If Parent.Mode.CheckBits(cOnKeyMap(key), currentModeBit) Then
            ' Create the lazy load procedure for the key
            Dim lazyLoadProcedure As String
            If key = "{226}" Then
                lazyLoadProcedure = "{¥}"
            Else
                lazyLoadProcedure = Replace(key, "{'}", "{''}")
                lazyLoadProcedure = Replace(lazyLoadProcedure, "{""}", "{""""}")
            End If

            lazyLoadProcedure = "'LazyLoad """ & lazyLoadProcedure & """'"

            Application.OnKey key, lazyLoadProcedure
        Else
            Application.OnKey key
        End If
    Next key

End Sub
