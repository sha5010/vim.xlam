Attribute VB_Name = "A_UserConfig"
' ==================================================   vim.xlam   ==================================================
' Author: SHA-5010 (@sha5010)
'
' [Description]
' vim.xlam は Excel上でVimエディタの操作感を実現するための Excelアドインソフトウェアです。
' 藤原崇 氏 (@VimExcel) が開発された VimExcel* のコードを一部使わせていただいております。
'   * https://www.vector.co.jp/soft/winnt/business/se494158.html
'
' vim.xlam では、自由にキーマッピングを変更できることを念頭に開発しています。
' VBA を記述できる方であれば、誰でも機能を開発し、好きなキーマップで登録することができます。
'
' LICENCE: MIT
' ==================================================================================================================
'
' *** 簡単な使い方 ***
' 1. Function または Sub で使いたい機能を作る。(すでに用意されていれば不要)
'    - F_ から始まるモジュール内に追加することを推奨。
' 2. initMapping 内で map メソッドを使い、キーを登録する。
'
' ==================================================================================================================
'
' *** map の使い方 ***
' Call map(key, subKey, funcName, [arg1, arg2, ..., arg5], [returnOnly], [requireArguments])
'
'   key .................. 登録するキーの1文字目を指定。Application.OnKey で指定する Key に準拠。
'   subKey ............... key を押した後、さらにキーを押す場合に指定。ここは普通の文字を入れる。
'   funcName ............. 登録したい機能のメソッド名を String型 で指定。
'   [argX] ............... 登録するメソッドで引数がある場合、それを指定。最大5つまで。
'   [returnOnly] ......... Enterキーを押すまでは実行されないようにする。:コマンドなど。
'   [requireArguments] ... コマンドの文字が入力された後は引数として扱うようにする。
'
' *** 例 ***
' - "j" で下入力したい
'     Call map("j", "", "moveDown")
'
' - "gg" で一番上に移動したい
'     Call map("g", "g", "moveToTopRow")
'
' - ":w" で Enterキーが押されたら保存したい
'     Call map(":", "w", "saveWorkbook", returnOnly:=True)
'
' - "w1" 〜 "w9" のように "wN" を受け取って、N番目のシートに移動したい
'     Call map("w", "", "activateWorksheet", requireArguments:=True)
'
' - "W" の後に引数を受け取って、それを基にセルを移動したい
'     Call map("W", "", "moveToSpecifiedCell", returnOnly:=True, requireArguments:=True)
'
' - 引数を受け取るメソッドを作って、map するときに幅を持たせたい
'     Call map("b", "b", "toggleBorderAll", xlContinuous, xlThin)
'
' ==================================================================================================================

Option Explicit
Option Private Module

Public Const VIM_TOOGLE_KEY As String = "^m"    'Vimモードを切り替えるショートカット (default: Ctrl + m)

Public Const SCROLL_OFFSET As Byte = 3          'scrollCurrentTop, scrollCurrentBottom N行分空ける
Public Const MAX_HISTORIES As Integer = 100     'ジャンプリストに格納する最大数


Sub initMapping()
    'マッピングの準備
    Call prepareMapping

    'Core
    Call map("^p", "", "toggleLang")


    'InsertMode
    Call map("a", "", "appendFollowLangMode")
    Call map("A", "", "appendNotFollowLangMode")
    Call map("i", "", "insertFollowLangMode")
    Call map("I", "", "insertNotFollowLangMode")
    Call map("s", "", "substituteFollowLangMode")
    Call map("S", "", "substituteNotFollowLangMode")


    'Moving
    Call map("h", "", "moveLeft")
    Call map("j", "", "moveDown")
    Call map("k", "", "moveUp")
    Call map("l", "", "moveRight")
    Call map("H", "", "moveLeft")
    Call map("J", "", "moveDown")
    Call map("K", "", "moveUp")
    Call map("L", "", "moveRight")
    Call map("^h", "", "moveLeft")
    Call map("^j", "", "moveDown")
    Call map("^k", "", "moveUp")
    Call map("^l", "", "moveRight")
    Call map("^H", "", "moveLeft")
    Call map("^J", "", "moveDown")
    Call map("^K", "", "moveUp")
    Call map("^L", "", "moveRight")

    Call map("g", "g", "moveToTopRow")
    Call map("G", "", "moveToLastRow")
    Call map("0", "", "moveToFirstColumn")
    Call map("{^}", "", "moveToLeftEnd")
    Call map("$", "", "moveToRightEnd")
    Call map("g", "0", "moveToA1")

    Call map("+{[}", "", "moveToTopOfCurrentRegion")
    Call map("+{]}", "", "moveToBottomOfCurrentRegion")

    Call map("W", "", "moveToSpecifiedCell", returnOnly:=True, requireArguments:=True)
    Call map(":", "", "moveToSpecifiedRow", returnOnly:=True, requireArguments:=True)


    'Cell
    Call map("d", "d", "cutCell")
    Call map("y", "y", "yankCell")
    Call map("o", "", "insertCellsDown")
    Call map("O", "", "insertCellsUp")
    Call map("t", "", "insertCellsRight")
    Call map("T", "", "insertCellsLeft")
    Call map("{+}", "", "incrementText")
    Call map("{-}", "", "decrementText")
    Call map("{[}", "", "increaseDecimal")
    Call map("{]}", "", "decreaseDecimal")
    Call map(" ", "", "unionSelectCells")
    Call map("+ ", "", "exceptSelectCells")
    Call map("@", "f", "followHyperlinkOfActiveCell")


    'Border
    Call map("b", "b", "toggleBorderAll", xlContinuous, xlThin)
    Call map("b", "a", "toggleBorderAround", xlContinuous, xlThin)
    Call map("b", "h", "toggleBorderLeft", xlContinuous, xlThin)
    Call map("b", "j", "toggleBorderBottom", xlContinuous, xlThin)
    Call map("b", "k", "toggleBorderTop", xlContinuous, xlThin)
    Call map("b", "l", "toggleBorderRight", xlContinuous, xlThin)
    Call map("b", "ia", "toggleBorderInner", xlContinuous, xlThin)
    Call map("b", "is", "toggleBorderInnerHorizontal", xlContinuous, xlThin)
    Call map("b", "iv", "toggleBorderInnerVertical", xlContinuous, xlThin)
    Call map("b", "/", "toggleBorderDiagonalUp", xlContinuous, xlThin)
    Call map("b", "¥", "toggleBorderDiagonalDown", xlContinuous, xlThin)

    Call map("b", "B", "toggleBorderAll", xlContinuous, xlThick)
    Call map("b", "A", "toggleBorderAround", xlContinuous, xlThick)
    Call map("b", "H", "toggleBorderLeft", xlContinuous, xlThick)
    Call map("b", "J", "toggleBorderBottom", xlContinuous, xlThick)
    Call map("b", "K", "toggleBorderTop", xlContinuous, xlThick)
    Call map("b", "L", "toggleBorderRight", xlContinuous, xlThick)
    Call map("B", "b", "toggleBorderAll", xlContinuous, xlThick)
    Call map("B", "a", "toggleBorderAround", xlContinuous, xlThick)
    Call map("B", "h", "toggleBorderLeft", xlContinuous, xlThick)
    Call map("B", "j", "toggleBorderBottom", xlContinuous, xlThick)
    Call map("B", "k", "toggleBorderTop", xlContinuous, xlThick)
    Call map("B", "l", "toggleBorderRight", xlContinuous, xlThick)
    Call map("B", "ia", "toggleBorderInner", xlContinuous, xlThick)
    Call map("B", "is", "toggleBorderInnerHorizontal", xlContinuous, xlThick)
    Call map("B", "iv", "toggleBorderInnerVertical", xlContinuous, xlThick)
    Call map("B", "/", "toggleBorderDiagonalUp", xlContinuous, xlThick)
    Call map("B", "¥", "toggleBorderDiagonalDown", xlContinuous, xlThick)

    Call map("b", "ob", "toggleBorderAll", xlContinuous, xlHairline)
    Call map("b", "oa", "toggleBorderAround", xlContinuous, xlHairline)
    Call map("b", "oh", "toggleBorderLeft", xlContinuous, xlHairline)
    Call map("b", "oj", "toggleBorderBottom", xlContinuous, xlHairline)
    Call map("b", "ok", "toggleBorderTop", xlContinuous, xlHairline)
    Call map("b", "ol", "toggleBorderRight", xlContinuous, xlHairline)
    Call map("b", "oia", "toggleBorderInner", xlContinuous, xlHairline)
    Call map("b", "ois", "toggleBorderInnerHorizontal", xlContinuous, xlHairline)
    Call map("b", "oiv", "toggleBorderInnerVertical", xlContinuous, xlHairline)
    Call map("b", "o/", "toggleBorderDiagonalUp", xlContinuous, xlHairline)
    Call map("b", "o¥", "toggleBorderDiagonalDown", xlContinuous, xlHairline)

    Call map("b", "mb", "toggleBorderAll", xlContinuous, xlMedium)
    Call map("b", "ma", "toggleBorderAround", xlContinuous, xlMedium)
    Call map("b", "mh", "toggleBorderLeft", xlContinuous, xlMedium)
    Call map("b", "mj", "toggleBorderBottom", xlContinuous, xlMedium)
    Call map("b", "mk", "toggleBorderTop", xlContinuous, xlMedium)
    Call map("b", "ml", "toggleBorderRight", xlContinuous, xlMedium)
    Call map("b", "mia", "toggleBorderInner", xlContinuous, xlMedium)
    Call map("b", "mis", "toggleBorderInnerHorizontal", xlContinuous, xlMedium)
    Call map("b", "miv", "toggleBorderInnerVertical", xlContinuous, xlMedium)
    Call map("b", "m/", "toggleBorderDiagonalUp", xlContinuous, xlMedium)
    Call map("b", "m¥", "toggleBorderDiagonalDown", xlContinuous, xlMedium)

    Call map("b", "tb", "toggleBorderAll", xlDouble, xlThick)
    Call map("b", "ta", "toggleBorderAround", xlDouble, xlThick)
    Call map("b", "th", "toggleBorderLeft", xlDouble, xlThick)
    Call map("b", "tj", "toggleBorderBottom", xlDouble, xlThick)
    Call map("b", "tk", "toggleBorderTop", xlDouble, xlThick)
    Call map("b", "tl", "toggleBorderRight", xlDouble, xlThick)
    Call map("b", "tia", "toggleBorderInner", xlDouble, xlThick)
    Call map("b", "tis", "toggleBorderInnerHorizontal", xlDouble, xlThick)
    Call map("b", "tiv", "toggleBorderInnerVertical", xlDouble, xlThick)
    Call map("b", "t/", "toggleBorderDiagonalUp", xlDouble, xlThick)
    Call map("b", "t¥", "toggleBorderDiagonalDown", xlDouble, xlThick)

    Call map("b", "dd", "deleteBorderAll")
    Call map("b", "da", "deleteBorderAround")
    Call map("b", "dh", "deleteBorderLeft")
    Call map("b", "dj", "deleteBorderBottom")
    Call map("b", "dk", "deleteBorderTop")
    Call map("b", "dl", "deleteBorderRight")
    Call map("b", "dia", "deleteBorderInner")
    Call map("b", "dis", "deleteBorderInnerHorizontal")
    Call map("b", "div", "deleteBorderInnerVertical")
    Call map("b", "d/", "deleteBorderDiagonalUp")
    Call map("b", "d¥", "deleteBorderDiagonalDown")

    Call map("b", "cc", "setBorderColorAll")
    Call map("b", "ca", "setBorderColorAround")
    Call map("b", "ch", "setBorderColorLeft")
    Call map("b", "cj", "setBorderColorBottom")
    Call map("b", "ck", "setBorderColorTop")
    Call map("b", "cl", "setBorderColorRight")
    Call map("b", "cia", "setBorderColorInner")
    Call map("b", "cis", "setBorderColorInnerHorizontal")
    Call map("b", "civ", "setBorderColorInnerVertical")
    Call map("b", "c/", "setBorderColorDiagonalUp")
    Call map("b", "c¥", "setBorderColorDiagonalDown")


    'Row
    Call map("e", "", "narrowRowsHeight")
    Call map("E", "", "wideRowsHeight")
    Call map("r", "r", "selectRows")
    Call map("r", "a", "appendRows")
    Call map("r", "i", "insertRows")
    Call map("r", "x", "deleteRows")
    Call map("r", "y", "yankRows")
    Call map("r", "d", "cutRows")
    Call map("r", "h", "hideRows")
    Call map("r", "H", "unhideRows")
    Call map("r", "g", "groupRows")
    Call map("r", "u", "ungroupRows")
    Call map("r", "f", "foldRowsGroup")
    Call map("r", "s", "spreadRowsGroup")
    Call map("r", "j", "adjustRowsHeight")
    Call map("r", "w", "setRowsHeight")


    'Column
    Call map("q", "", "narrowColumnsWidth")
    Call map("Q", "", "wideColumnsWidth")
    Call map("c", "c", "selectColumns")
    Call map("c", "a", "appendColumns")
    Call map("c", "i", "insertColumns")
    Call map("c", "x", "deleteColumns")
    Call map("c", "y", "yankColumns")
    Call map("c", "d", "cutColumns")
    Call map("c", "h", "hideColumns")
    Call map("c", "H", "unhideColumns")
    Call map("c", "g", "groupColumns")
    Call map("c", "u", "ungroupColumns")
    Call map("c", "f", "foldColumnsGroup")
    Call map("c", "s", "spreadColumnsGroup")
    Call map("c", "j", "adjustColumnsWidth")
    Call map("c", "w", "setColumnsWidth")


    'Yank
    Call map("y", "r", "yankRows")
    Call map("y", "c", "yankColumns")
    Call map("y", "gg", "yankToTopRows")
    Call map("y", "G", "yankToBottomRows")
    Call map("y", "{", "yankToTopOfCurrentRegionRows")
    Call map("y", "}", "yankToBottomOfCurrentRegionRows")
    Call map("y", "0", "yankToLeftEndColumns")
    Call map("y", "$", "yankToRightEndColumns")
    Call map("y", "^", "yankToLeftOfCurrentRegionColumns")
    Call map("y", "g$", "yankToRightOfCurrentRegionColumns")

    Call map("y", "h", "yankFromLeftCell")
    Call map("y", "j", "yankFromDownCell")
    Call map("y", "k", "yankFromUpCell")
    Call map("y", "l", "yankFromRightCell")


    'Delete
    Call map("X", "", "deleteValue")
    Call map("D", "", "deleteValue")
    Call map("x", "x", "deleteRows")
    Call map("x", "r", "deleteRows")
    Call map("x", "c", "deleteColumns")
    Call map("x", "gg", "deleteToTopRows")
    Call map("x", "G", "deleteToBottomRows")
    Call map("x", "{", "deleteToTopOfCurrentRegionRows")
    Call map("x", "}", "deleteToBottomOfCurrentRegionRows")
    Call map("x", "0", "deleteToLeftEndColumns")
    Call map("x", "$", "deleteToRightEndColumns")
    Call map("x", "^", "deleteToLeftOfCurrentRegionColumns")
    Call map("x", "g$", "deleteToRightOfCurrentRegionColumns")

    Call map("x", "h", "deleteToLeft")
    Call map("x", "j", "deleteToUp")
    Call map("x", "k", "deleteToUp")
    Call map("x", "l", "deleteToLeft")


    'Cut
    Call map("d", "r", "cutRows")
    Call map("d", "c", "cutColumns")
    Call map("d", "gg", "cutToTopRows")
    Call map("d", "G", "cutToBottomRows")
    Call map("d", "{", "cutToTopOfCurrentRegionRows")
    Call map("d", "}", "cutToBottomOfCurrentRegionRows")
    Call map("d", "0", "cutToLeftEndColumns")
    Call map("d", "$", "cutToRightEndColumns")
    Call map("d", "^", "cutToLeftOfCurrentRegionColumns")
    Call map("d", "g$", "cutToRightOfCurrentRegionColumns")


    'Paste
    Call map("p", "", "pasteSmart")
    Call map("P", "", "pasteSpecial")
    Call map("U", "", "pasteValue")


    'Font
    Call map("<", "", "decreaseFontSize")
    Call map(">", "", "increaseFontSize")
    Call map("f", "n", "changeFontName")
    Call map("f", "s", "changeFontSize")
    Call map("f", "h", "alignLeft")
    Call map("f", "j", "alignBottom")
    Call map("f", "k", "alignTop")
    Call map("f", "l", "alignRight")
    Call map("f", "o", "alignCenter")
    Call map("f", "m", "alignMiddle")
    Call map("f", "b", "toggleBold")
    Call map("f", "i", "toggleItalic")
    Call map("f", "u", "toggleUnderline")
    Call map("f", "f", "showFontDialog")


    'Color
    Call map("f", "c", "smartFontColor")
    Call map("F", "c", "smartFillColor")
    Call map("b", "c", "changeShapeBorderColor", requireArguments:=True)


    'Find & Replace
    Call map("/", "", "showFindFollowLang")
    Call map("?", "", "showFindNotFollowLang")
    Call map("n", "", "nextFoundCell")
    Call map("N", "", "previousFoundCell")
    Call map("R", "", "showReplaceWindow")


    'Scrolling
    Call map("^u", "", "scrollUpHalf")
    Call map("^d", "", "scrollDownHalf")
    Call map("^b", "", "scrollUp")
    Call map("^f", "", "scrollDown")
    Call map("^y", "", "scrollUp1Row")
    Call map("^e", "", "scrollDown1Row")
    Call map("z", "t", "scrollCurrentTop")
    Call map("z", "z", "scrollCurrentMiddle")
    Call map("z", "b", "scrollCurrentBottom")
    Call map("z", "s", "scrollCurrentLeft")
    Call map("z", "m", "scrollCurrentCenter")
    Call map("z", "e", "scrollCurrentRight")


    'Worksheet Function
    Call map("v", "", "nextWorksheet")
    Call map("V", "", "previousWorksheet")

    Call map("w", "w", "showSheetPicker")
    Call map("w", "s", "showSheetPicker")
    Call map("w", "n", "nextWorksheet")
    Call map("w", "p", "previousWorksheet")
    Call map("w", "r", "renameWorksheet")
    Call map("w", "h", "moveWorksheetBack")
    Call map("w", "l", "moveWorksheetForward")
    Call map("w", "i", "insertWorksheet")
    Call map("w", "a", "appendWorksheet")
    Call map("w", "d", "deleteWorksheet")
    'Call map("w", "1", "activateFirstWorksheet")
    Call map("w", "0", "activateLastWorksheet")
    Call map("w", "$", "activateLastWorksheet")
    Call map("w", "c", "cloneWorksheet")
    Call map("w", "y", "cloneWorksheet")
    Call map("w", "", "activateWorksheet", requireArguments:=True)


    'Useful Command
    Call map("u", "", "undo_CtrlZ")
    'Call map("^r", "", "repeat_F4")  'Currently Broken
    Call map("m", "", "zoomIn")
    Call map("M", "", "zoomOut")
    Call map(":", "e", "openWorkbook", returnOnly:=True)
    Call map(":", "e!", "reopenActiveWorkbook", returnOnly:=True)
    Call map(":", "w", "saveWorkbook", returnOnly:=True)
    Call map(":", "q", "closeAskSaving", returnOnly:=True)
    Call map(":", "q!", "closeWithoutSaving", returnOnly:=True)
    Call map(":", "wq", "closeWithSaving", returnOnly:=True)
    Call map(":", "x", "closeWithSaving", returnOnly:=True)
    Call map(":", "r", "reloadVim", returnOnly:=True)
    Call map(":", "r!", "reloadVim", True, returnOnly:=True)

    Call map("^i", "", "jumpNext")
    Call map("^o", "", "jumpPrev")
    Call map(":", "cle", "clearJumps", returnOnly:=True)

    Call map(":", "b", "activateWorkbook", returnOnly:=True, requireArguments:=True)
    Call map(":", "bn", "nextWorkbook", returnOnly:=True)
    Call map(":", "bp", "previousWorkbook", returnOnly:=True)


    'Atmark Command
    Call map("@", "w", "toggleFreezePanes")
    Call map("@", "r", "toggleWrapText")
    Call map("@", "m", "toggleMergeCells")
    Call map("@", "x", "toggleFormulaBar")
    Call map("@", "s", "showSummaryInfo")
    Call map("@", "a", "toggleReadOnly")
    Call map("@", "n", "nextWorkbook")
    Call map("@", "N", "previousWorkbook")


    'Count
    Call map("1", "", "showCmdForm", "1")
    Call map("2", "", "showCmdForm", "2")
    Call map("3", "", "showCmdForm", "3")
    Call map("4", "", "showCmdForm", "4")
    Call map("5", "", "showCmdForm", "5")
    Call map("6", "", "showCmdForm", "6")
    Call map("7", "", "showCmdForm", "7")
    Call map("8", "", "showCmdForm", "8")
    Call map("9", "", "showCmdForm", "9")


    'KeyMapping
    Call map("^{[}", "", "primitiveKeyMapping", vbKeyEscape)

End Sub
