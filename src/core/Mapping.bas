Attribute VB_Name = "C_Mapping"
Option Explicit
Option Private Module

Function map(ByVal key As String, ByVal subKey As String, ByVal funcName As String, _
             Optional ByVal arg1 As Variant, _
             Optional ByVal arg2 As Variant, _
             Optional ByVal arg3 As Variant, _
             Optional ByVal arg4 As Variant, _
             Optional ByVal arg5 As Variant, _
             Optional ByVal returnOnly As Boolean = False, _
             Optional ByVal requireArguments As Boolean = False)

    Dim firstKey As String
    Dim funcNameWithArg As String

    'key なしの場合
    If key = "" Then
        Exit Function
    End If

    'argX の指定と returnArguments は同時指定不可
    If Not (IsMissing(arg1) And IsMissing(arg2) And IsMissing(arg3) And IsMissing(arg4) And IsMissing(arg5)) _
        And requireArguments Then

        Err.Raise 50000, Description:="argX と requireArguments は同時指定できません。"
        Exit Function
    End If

    '引数付きの名前を算出
    funcNameWithArg = "'" & funcName
    funcNameWithArg = funcNameWithArg & parseArg(arg1)
    funcNameWithArg = funcNameWithArg & parseArg(arg2)
    funcNameWithArg = funcNameWithArg & parseArg(arg3)
    funcNameWithArg = funcNameWithArg & parseArg(arg4)
    funcNameWithArg = funcNameWithArg & parseArg(arg5)
    If Right(funcNameWithArg, 1) = "," Then
        funcNameWithArg = Left(funcNameWithArg, Len(funcNameWithArg) - 1)
    End If
    funcNameWithArg = funcNameWithArg & "'"

    'subKey がない場合
    If subKey = "" And Not requireArguments Then
        Call registerOnKey(key, funcNameWithArg)
        Call registerKeyMap(key, funcNameWithArg, returnOnly, requireArguments)
        Exit Function
    End If

    '2文字以上のマッピング
    Call registerOnKey(key, "'showCmdForm """ & key & """'")
    Call registerKeyMap(key & subKey, funcNameWithArg, returnOnly, requireArguments)
End Function

Private Function parseArg(ByVal arg As Variant) As String
    If IsMissing(arg) Then
        Exit Function
    End If

    Select Case TypeName(arg)
        Case "String"
            parseArg = " """ & arg & ""","
        Case "Integer", "Long", "LongLong", "Double", "Single", "Byte"
            parseArg = " " & arg & ","
        Case "Boolean"
            parseArg = " " & CStr(arg) & ","
        Case Else
            Call debugPrint("Unsupport argument type: " & TypeName(arg), "parseArg")
    End Select
End Function

Private Sub registerOnKey(ByVal key As String, Optional funcName As String = "")
    Dim lowerKey As String

    lowerKey = LCase(key)
    If lowerKey <> key Then
        key = "+" & lowerKey
    End If

    If gRegisteredKeys.Exists(key) Then
        gRegisteredKeys(key) = funcName
    Else
        Call gRegisteredKeys.Add(key, funcName)
    End If

    If funcName = "" Then
        Call Application.OnKey(key)
    Else
        Call Application.OnKey(key, funcName)
    End If
End Sub

Sub disableKeys()
    Dim key As Variant

    If gRegisteredKeys Is Nothing Then
        Err.Raise 50000, Description:="キーの初回セットアップが済んでいません。"
        Exit Sub
    End If

    For Each key In gRegisteredKeys
        Call Application.OnKey(key)
    Next key
End Sub

Sub enableKeys()
    Dim key As Variant

    If gRegisteredKeys Is Nothing Then
        Err.Raise 50000, Description:="キーの初回セットアップが済んでいません。"
        Exit Sub
    End If

    For Each key In gRegisteredKeys
        Call Application.OnKey(key, gRegisteredKeys(key))
    Next key
End Sub

Private Sub unregisterKeyMap(ByVal key As String)
    If gKeyMap(KEY_NORMAL).Exists(key) Then Call gKeyMap(KEY_NORMAL).Remove(key)
    If gKeyMap(KEY_NORMAL_ARG).Exists(key) Then Call gKeyMap(KEY_NORMAL_ARG).Remove(key)
    If gKeyMap(KEY_RETURNONLY).Exists(key) Then Call gKeyMap(KEY_RETURNONLY).Remove(key)
    If gKeyMap(KEY_RETURNONLY_ARG).Exists(key) Then Call gKeyMap(KEY_RETURNONLY_ARG).Remove(key)
End Sub

Private Sub registerKeyMap(ByVal key As String, ByVal funcName As String, _
                           ByVal returnOnly As Boolean, _
                           ByVal requireArguments As Boolean)

    'Remove {}
    key = reReplace(key, "¥{(.+)¥}", "$1")

    Call unregisterKeyMap(key)
    If returnOnly Then
        If requireArguments Then
            Call gKeyMap(KEY_RETURNONLY_ARG).Add(key, funcName)
        Else
            Call gKeyMap(KEY_RETURNONLY).Add(key, funcName)
        End If

    Else
        If requireArguments Then
            Call gKeyMap(KEY_NORMAL_ARG).Add(key, funcName)
        Else
            Call gKeyMap(KEY_NORMAL).Add(key, funcName)
        End If
    End If
End Sub

Function primitiveKeyMapping(ByVal KeyCode As Byte)
    Call keyupControlKeys
    Call releaseShiftKeys

    keybd_event KeyCode, 0, 0, 0
    keybd_event KeyCode, 0, KEYUP, 0

    Call unkeyupControlKeys
End Function

Function prepareMapping()
    Call mapResetAll

    Set gKeyMap = New Dictionary
    Set gRegisteredKeys = New Dictionary

    gKeyMap.Add KEY_NORMAL, New Dictionary
    gKeyMap.Add KEY_NORMAL_ARG, New Dictionary
    gKeyMap.Add KEY_RETURNONLY, New Dictionary
    gKeyMap.Add KEY_RETURNONLY_ARG, New Dictionary

    '主要なキーを無効化
    Call mapToAllDummy
End Function

Function mapToAllDummy()
    registerOnKey "a", "dummy"
    registerOnKey "b", "dummy"
    registerOnKey "c", "dummy"
    registerOnKey "d", "dummy"
    registerOnKey "e", "dummy"
    registerOnKey "f", "dummy"
    registerOnKey "g", "dummy"
    registerOnKey "h", "dummy"
    registerOnKey "i", "dummy"
    registerOnKey "j", "dummy"
    registerOnKey "k", "dummy"
    registerOnKey "l", "dummy"
    registerOnKey "m", "dummy"
    registerOnKey "n", "dummy"
    registerOnKey "o", "dummy"
    registerOnKey "p", "dummy"
    registerOnKey "q", "dummy"
    registerOnKey "r", "dummy"
    registerOnKey "s", "dummy"
    registerOnKey "t", "dummy"
    registerOnKey "u", "dummy"
    registerOnKey "v", "dummy"
    registerOnKey "w", "dummy"
    registerOnKey "x", "dummy"
    registerOnKey "y", "dummy"
    registerOnKey "z", "dummy"

    registerOnKey "0", "dummy"
    registerOnKey "1", "dummy"
    registerOnKey "2", "dummy"
    registerOnKey "3", "dummy"
    registerOnKey "4", "dummy"
    registerOnKey "5", "dummy"
    registerOnKey "6", "dummy"
    registerOnKey "7", "dummy"
    registerOnKey "8", "dummy"
    registerOnKey "9", "dummy"

    registerOnKey "=", "dummy"
    registerOnKey "-", "dummy"
    registerOnKey "{^}", "dummy"
    registerOnKey "¥", "dummy"
    registerOnKey "@", "dummy"
    registerOnKey "{[}", "dummy"
    registerOnKey ";", "dummy"
    registerOnKey ":", "dummy"
    registerOnKey "{]}", "dummy"
    registerOnKey ",", "dummy"
    registerOnKey ".", "dummy"
    registerOnKey "/", "dummy"
    registerOnKey "¥", "dummy"
    registerOnKey " ", "dummy"
    registerOnKey "{226}", "dummy"

    registerOnKey "+a", "dummy"
    registerOnKey "+b", "dummy"
    registerOnKey "+c", "dummy"
    registerOnKey "+d", "dummy"
    registerOnKey "+e", "dummy"
    registerOnKey "+f", "dummy"
    registerOnKey "+g", "dummy"
    registerOnKey "+h", "dummy"
    registerOnKey "+i", "dummy"
    registerOnKey "+j", "dummy"
    registerOnKey "+k", "dummy"
    registerOnKey "+l", "dummy"
    registerOnKey "+m", "dummy"
    registerOnKey "+n", "dummy"
    registerOnKey "+o", "dummy"
    registerOnKey "+p", "dummy"
    registerOnKey "+q", "dummy"
    registerOnKey "+r", "dummy"
    registerOnKey "+s", "dummy"
    registerOnKey "+t", "dummy"
    registerOnKey "+u", "dummy"
    registerOnKey "+v", "dummy"
    registerOnKey "+w", "dummy"
    registerOnKey "+x", "dummy"
    registerOnKey "+y", "dummy"
    registerOnKey "+z", "dummy"

    registerOnKey "+0", "dummy"
    registerOnKey "+1", "dummy"
    registerOnKey "+2", "dummy"
    registerOnKey "+3", "dummy"
    registerOnKey "+4", "dummy"
    registerOnKey "+5", "dummy"
    registerOnKey "+6", "dummy"
    registerOnKey "+7", "dummy"
    registerOnKey "+8", "dummy"
    registerOnKey "+9", "dummy"

    registerOnKey "+-", "dummy"
    registerOnKey "+{^}", "dummy"
    registerOnKey "+¥", "dummy"
    registerOnKey "+@", "dummy"
    registerOnKey "+{[}", "dummy"
    registerOnKey "+;", "dummy"
    registerOnKey "+:", "dummy"
    registerOnKey "+{]}", "dummy"
    registerOnKey "<", "dummy"
    registerOnKey "+.", "dummy"
    registerOnKey "+/", "dummy"
    registerOnKey "_", "dummy"
    registerOnKey "+ ", "dummy"
End Function

Function mapResetAll()
    Set gKeyMap = Nothing

    With Application
        .OnKey "a"
        .OnKey "b"
        .OnKey "c"
        .OnKey "d"
        .OnKey "e"
        .OnKey "f"
        .OnKey "g"
        .OnKey "h"
        .OnKey "i"
        .OnKey "j"
        .OnKey "k"
        .OnKey "l"
        .OnKey "m"
        .OnKey "n"
        .OnKey "o"
        .OnKey "p"
        .OnKey "q"
        .OnKey "r"
        .OnKey "s"
        .OnKey "t"
        .OnKey "u"
        .OnKey "v"
        .OnKey "w"
        .OnKey "x"
        .OnKey "y"
        .OnKey "z"

        .OnKey "0"
        .OnKey "1"
        .OnKey "2"
        .OnKey "3"
        .OnKey "4"
        .OnKey "5"
        .OnKey "6"
        .OnKey "7"
        .OnKey "8"
        .OnKey "9"

        .OnKey "="
        .OnKey "-"
        .OnKey "{^}"
        .OnKey "¥"
        .OnKey "@"
        .OnKey "{[}"
        .OnKey ";"
        .OnKey ":"
        .OnKey "{]}"
        .OnKey ","
        .OnKey "."
        .OnKey "/"
        .OnKey "¥"
        .OnKey " "
        .OnKey "{226}"

        'Shift
        .OnKey "+a"
        .OnKey "+b"
        .OnKey "+c"
        .OnKey "+d"
        .OnKey "+e"
        .OnKey "+f"
        .OnKey "+g"
        .OnKey "+h"
        .OnKey "+i"
        .OnKey "+j"
        .OnKey "+k"
        .OnKey "+l"
        .OnKey "+m"
        .OnKey "+n"
        .OnKey "+o"
        .OnKey "+p"
        .OnKey "+q"
        .OnKey "+r"
        .OnKey "+s"
        .OnKey "+t"
        .OnKey "+u"
        .OnKey "+v"
        .OnKey "+w"
        .OnKey "+x"
        .OnKey "+y"
        .OnKey "+z"

        .OnKey "+0"
        .OnKey "+1"
        .OnKey "+2"
        .OnKey "+3"
        .OnKey "+4"
        .OnKey "+5"
        .OnKey "+6"
        .OnKey "+7"
        .OnKey "+8"
        .OnKey "+9"

        .OnKey "+-"
        .OnKey "+{^}"
        .OnKey "+¥"
        .OnKey "+@"
        .OnKey "+{[}"
        .OnKey "+;"
        .OnKey "+:"
        .OnKey "+{]}"
        .OnKey "<"
        .OnKey "+."
        .OnKey "+/"
        .OnKey "_"
        .OnKey "+ "

        'Ctrl
        .OnKey "^a"
        .OnKey "^b"
        .OnKey "^c"
        .OnKey "^d"
        .OnKey "^e"
        .OnKey "^f"
        .OnKey "^g"
        .OnKey "^h"
        .OnKey "^i"
        .OnKey "^j"
        .OnKey "^k"
        .OnKey "^l"
        .OnKey "^m"
        .OnKey "^n"
        .OnKey "^o"
        .OnKey "^p"
        .OnKey "^q"
        .OnKey "^r"
        .OnKey "^s"
        .OnKey "^t"
        .OnKey "^u"
        .OnKey "^v"
        .OnKey "^w"
        .OnKey "^x"
        .OnKey "^y"
        .OnKey "^z"

        .OnKey "^0"
        .OnKey "^1"
        .OnKey "^2"
        .OnKey "^3"
        .OnKey "^4"
        .OnKey "^5"
        .OnKey "^6"
        .OnKey "^7"
        .OnKey "^8"
        .OnKey "^9"

        .OnKey "^-"
        .OnKey "^{^}"
        .OnKey "^¥"
        .OnKey "^@"
        .OnKey "^{[}"
        .OnKey "^;"
        .OnKey "^:"
        .OnKey "^{]}"
        .OnKey "^,"
        .OnKey "^."
        .OnKey "^/"
        .OnKey "^¥"

        'Ctrl + Shift
        .OnKey "^+a"
        .OnKey "^+b"
        .OnKey "^+c"
        .OnKey "^+d"
        .OnKey "^+e"
        .OnKey "^+f"
        .OnKey "^+g"
        .OnKey "^+h"
        .OnKey "^+i"
        .OnKey "^+j"
        .OnKey "^+k"
        .OnKey "^+l"
        .OnKey "^+m"
        .OnKey "^+n"
        .OnKey "^+o"
        .OnKey "^+p"
        .OnKey "^+q"
        .OnKey "^+r"
        .OnKey "^+s"
        .OnKey "^+t"
        .OnKey "^+u"
        .OnKey "^+v"
        .OnKey "^+w"
        .OnKey "^+x"
        .OnKey "^+y"
        .OnKey "^+z"

        .OnKey "^+0"
        .OnKey "^+1"
        .OnKey "^+2"
        .OnKey "^+3"
        .OnKey "^+4"
        .OnKey "^+5"
        .OnKey "^+6"
        .OnKey "^+7"
        .OnKey "^+8"
        .OnKey "^+9"

        .OnKey "^+-"
        .OnKey "^+{^}"
        .OnKey "^+¥"
        .OnKey "^+@"
        .OnKey "^+{[}"
        .OnKey "^+;"
        .OnKey "^+:"
        .OnKey "^+{]}"
        .OnKey "^<"
        .OnKey "^+."
        .OnKey "^+/"
        .OnKey "^_"
    End With
End Function

Function dummy()
    If gDebugMode Then
        Call setStatusBarTemporarily("No allocation", 1)
    End If
End Function
