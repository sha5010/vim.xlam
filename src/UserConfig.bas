Attribute VB_Name = "A_UserConfig"
' ==================================================   vim.xlam   ==================================================
' Author: SHA-5010 (@sha_5010)
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

Sub DefaultMapping()

    With gVim.KeyMap
        'Core
        Call .Map("nmap <C-p> ToggleLang")
        Call .Map("nmap : EnterCmdlineMode")
        Call .Map("nmap <cmd>debug ToggleDebugMode")


        'InsertMode
        Call .Map("nmap a AppendFollowLangMode")
        Call .Map("nmap A AppendNotFollowLangMode")
        Call .Map("nmap i InsertFollowLangMode")
        Call .Map("nmap I InsertNotFollowLangMode")
        Call .Map("nmap s SubstituteFollowLangMode")
        Call .Map("nmap S SubstituteNotFollowLangMode")


        'Moving
        Call .Map("nmap h MoveLeft")
        Call .Map("nmap j MoveDown")
        Call .Map("nmap k MoveUp")
        Call .Map("nmap l MoveRight")
        Call .Map("nmap H MoveLeftWithShift")
        Call .Map("nmap J MoveDownWithShift")
        Call .Map("nmap K MoveUpWithShift")
        Call .Map("nmap L MoveRightWithShift")
        Call .Map("nmap <C-h> MoveLeft")
        Call .Map("nmap <C-j> MoveDown")
        Call .Map("nmap <C-k> MoveUp")
        Call .Map("nmap <C-l> MoveRight")
        Call .Map("nmap <C-S-H> MoveLeft")
        Call .Map("nmap <C-S-J> MoveDown")
        Call .Map("nmap <C-S-K> MoveUp")
        Call .Map("nmap <C-S-L> MoveRight")

        Call .Map("nmap gg MoveToTopRow")
        Call .Map("nmap G MoveToLastRow")
        Call .Map("nmap <bar> MoveToNthColumn")
        Call .Map("nmap 0 MoveToFirstColumn")
        Call .Map("nmap ^ MoveToLeftEnd")
        Call .Map("nmap $ MoveToRightEnd")
        Call .Map("nmap g0 MoveToA1")

        Call .Map("nmap { MoveToTopOfCurrentRegion")
        Call .Map("nmap } MoveToBottomOfCurrentRegion")

        Call .Map("nmap W MoveToSpecifiedCell")


        'Cell
        Call .Map("nmap xx CutCell")
        Call .Map("nmap yy YankCell")
        Call .Map("nmap o InsertCellsDown")
        Call .Map("nmap O InsertCellsUp")
        Call .Map("nmap t InsertCellsRight")
        Call .Map("nmap T InsertCellsLeft")
        Call .Map("nmap > IncrementText")
        Call .Map("nmap <lt> DecrementText")
        Call .Map("nmap ( IncreaseDecimal")
        Call .Map("nmap ) DecreaseDecimal")
        Call .Map("nmap zw ToggleWrapText")
        Call .Map("nmap & ToggleMergeCells")
        Call .Map("nmap f, ApplyCommaStyle")
        Call .Map("nmap <Space> UnionSelectCells")
        Call .Map("nmap <S-Space> ExceptSelectCells")
        Call .Map("nmap <S-BS> ClearSelectCells")
        Call .Map("nmap gf FollowHyperlinkOfActiveCell")
        Call .Map("nmap Ff ApplyFlashFill")
        Call .Map("nmap FF ApplyFlashFill")
        Call .Map("nmap Fa ApplyAutoFill")
        Call .Map("nmap FA ApplyAutoFill")

        Call .Map("nmap =s AutoSum")
        Call .Map("nmap =a AutoAverage")
        Call .Map("nmap =c AutoCount")
        Call .Map("nmap =m AutoMax")
        Call .Map("nmap =i AutoMin")
        Call .Map("nmap == InsertFunction")

        Call .Map("nmap v ToggleVisualMode")
        Call .Map("nmap V ToggleVisualLine")


        ' Border
        Call .Map("nmap bb ToggleBorderAll " & xlContinuous & " " & xlThin)
        Call .Map("nmap ba ToggleBorderAround " & xlContinuous & " " & xlThin)
        Call .Map("nmap bh ToggleBorderLeft " & xlContinuous & " " & xlThin)
        Call .Map("nmap bj ToggleBorderBottom " & xlContinuous & " " & xlThin)
        Call .Map("nmap bk ToggleBorderTop " & xlContinuous & " " & xlThin)
        Call .Map("nmap bl ToggleBorderRight " & xlContinuous & " " & xlThin)
        Call .Map("nmap bia ToggleBorderInner " & xlContinuous & " " & xlThin)
        Call .Map("nmap bis ToggleBorderInnerHorizontal " & xlContinuous & " " & xlThin)
        Call .Map("nmap biv ToggleBorderInnerVertical " & xlContinuous & " " & xlThin)
        Call .Map("nmap b/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlThin)
        Call .Map("nmap b<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlThin)

        Call .Map("nmap bB ToggleBorderAll " & xlContinuous & " " & xlThick)
        Call .Map("nmap bA ToggleBorderAround " & xlContinuous & " " & xlThick)
        Call .Map("nmap bH ToggleBorderLeft " & xlContinuous & " " & xlThick)
        Call .Map("nmap bJ ToggleBorderBottom " & xlContinuous & " " & xlThick)
        Call .Map("nmap bK ToggleBorderTop " & xlContinuous & " " & xlThick)
        Call .Map("nmap bL ToggleBorderRight " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bb ToggleBorderAll " & xlContinuous & " " & xlThick)
        Call .Map("nmap Ba ToggleBorderAround " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bh ToggleBorderLeft " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bj ToggleBorderBottom " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bk ToggleBorderTop " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bl ToggleBorderRight " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bia ToggleBorderInner " & xlContinuous & " " & xlThick)
        Call .Map("nmap Bis ToggleBorderInnerHorizontal " & xlContinuous & " " & xlThick)
        Call .Map("nmap Biv ToggleBorderInnerVertical " & xlContinuous & " " & xlThick)
        Call .Map("nmap B/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlThick)
        Call .Map("nmap B<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlThick)

        Call .Map("nmap bob ToggleBorderAll " & xlContinuous & " " & xlHairline)
        Call .Map("nmap boa ToggleBorderAround " & xlContinuous & " " & xlHairline)
        Call .Map("nmap boh ToggleBorderLeft " & xlContinuous & " " & xlHairline)
        Call .Map("nmap boj ToggleBorderBottom " & xlContinuous & " " & xlHairline)
        Call .Map("nmap bok ToggleBorderTop " & xlContinuous & " " & xlHairline)
        Call .Map("nmap bol ToggleBorderRight " & xlContinuous & " " & xlHairline)
        Call .Map("nmap boia ToggleBorderInner " & xlContinuous & " " & xlHairline)
        Call .Map("nmap bois ToggleBorderInnerHorizontal " & xlContinuous & " " & xlHairline)
        Call .Map("nmap boiv ToggleBorderInnerVertical " & xlContinuous & " " & xlHairline)
        Call .Map("nmap bo/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlHairline)
        Call .Map("nmap bo<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlHairline)

        Call .Map("nmap bmb ToggleBorderAll " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bma ToggleBorderAround " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bmh ToggleBorderLeft " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bmj ToggleBorderBottom " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bmk ToggleBorderTop " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bml ToggleBorderRight " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bmia ToggleBorderInner " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bmis ToggleBorderInnerHorizontal " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bmiv ToggleBorderInnerVertical " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bm/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlMedium)
        Call .Map("nmap bm<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlMedium)

        Call .Map("nmap btb ToggleBorderAll " & xlDouble & " " & xlThick)
        Call .Map("nmap bta ToggleBorderAround " & xlDouble & " " & xlThick)
        Call .Map("nmap bth ToggleBorderLeft " & xlDouble & " " & xlThick)
        Call .Map("nmap btj ToggleBorderBottom " & xlDouble & " " & xlThick)
        Call .Map("nmap btk ToggleBorderTop " & xlDouble & " " & xlThick)
        Call .Map("nmap btl ToggleBorderRight " & xlDouble & " " & xlThick)
        Call .Map("nmap btia ToggleBorderInner " & xlDouble & " " & xlThick)
        Call .Map("nmap btis ToggleBorderInnerHorizontal " & xlDouble & " " & xlThick)
        Call .Map("nmap btiv ToggleBorderInnerVertical " & xlDouble & " " & xlThick)
        Call .Map("nmap bt/ ToggleBorderDiagonalUp " & xlDouble & " " & xlThick)
        Call .Map("nmap bt<bslash> ToggleBorderDiagonalDown " & xlDouble & " " & xlThick)

        Call .Map("nmap bdd DeleteBorderAll")
        Call .Map("nmap bda DeleteBorderAround")
        Call .Map("nmap bdh DeleteBorderLeft")
        Call .Map("nmap bdj DeleteBorderBottom")
        Call .Map("nmap bdk DeleteBorderTop")
        Call .Map("nmap bdl DeleteBorderRight")
        Call .Map("nmap bdia DeleteBorderInner")
        Call .Map("nmap bdis DeleteBorderInnerHorizontal")
        Call .Map("nmap bdiv DeleteBorderInnerVertical")
        Call .Map("nmap bd/ DeleteBorderDiagonalUp")
        Call .Map("nmap bd<bslash> DeleteBorderDiagonalDown")

        Call .Map("nmap bcc SetBorderColorAll")
        Call .Map("nmap bca SetBorderColorAround")
        Call .Map("nmap bch SetBorderColorLeft")
        Call .Map("nmap bcj SetBorderColorBottom")
        Call .Map("nmap bck SetBorderColorTop")
        Call .Map("nmap bcl SetBorderColorRight")
        Call .Map("nmap bcia SetBorderColorInner")
        Call .Map("nmap bcis SetBorderColorInnerHorizontal")
        Call .Map("nmap bciv SetBorderColorInnerVertical")
        Call .Map("nmap bc/ SetBorderColorDiagonalUp")
        Call .Map("nmap bc<bslash> SetBorderColorDiagonalDown")


        'Row
        Call .Map("nmap r- NarrowRowsHeight")
        Call .Map("nmap r+ WideRowsHeight")
        Call .Map("nmap rr SelectRows")
        Call .Map("nmap ra AppendRows")
        Call .Map("nmap ri InsertRows")
        Call .Map("nmap rd DeleteRows")
        Call .Map("nmap ry YankRows")
        Call .Map("nmap rx CutRows")
        Call .Map("nmap rh HideRows")
        Call .Map("nmap rH UnhideRows")
        Call .Map("nmap rg GroupRows")
        Call .Map("nmap ru UngroupRows")
        Call .Map("nmap rf FoldRowsGroup")
        Call .Map("nmap rs SpreadRowsGroup")
        Call .Map("nmap rj AdjustRowsHeight")
        Call .Map("nmap rw SetRowsHeight")


        'Column
        Call .Map("nmap c- NarrowColumnsWidth")
        Call .Map("nmap c+ WideColumnsWidth")
        Call .Map("nmap cc SelectColumns")
        Call .Map("nmap ca AppendColumns")
        Call .Map("nmap ci InsertColumns")
        Call .Map("nmap cd DeleteColumns")
        Call .Map("nmap cy YankColumns")
        Call .Map("nmap cx CutColumns")
        Call .Map("nmap ch HideColumns")
        Call .Map("nmap cH UnhideColumns")
        Call .Map("nmap cg GroupColumns")
        Call .Map("nmap cu UngroupColumns")
        Call .Map("nmap cf FoldColumnsGroup")
        Call .Map("nmap cs SpreadColumnsGroup")
        Call .Map("nmap cj AdjustColumnsWidth")
        Call .Map("nmap cw SetColumnsWidth")


        'Yank
        Call .Map("nmap yr YankRows")
        Call .Map("nmap yc YankColumns")
        Call .Map("nmap ygg YankRows " & eTargetRowType.ToTopRows)
        Call .Map("nmap yG YankRows " & eTargetRowType.ToBottomRows)
        Call .Map("nmap y{ YankRows " & eTargetRowType.ToTopOfCurrentRegionRows)
        Call .Map("nmap y} YankRows " & eTargetRowType.ToBottomOfCurrentRegionRows)
        Call .Map("nmap y0 YankColumns " & eTargetColumnType.ToLeftEndColumns)
        Call .Map("nmap y$ YankColumns " & eTargetColumnType.ToRightEndColumns)
        Call .Map("nmap y^ YankColumns " & eTargetColumnType.ToLeftOfCurrentRegionColumns)
        Call .Map("nmap yg$ YankColumns " & eTargetColumnType.ToRightOfCurrentRegionColumns)

        Call .Map("nmap yh YankFromLeftCell")
        Call .Map("nmap yj YankFromDownCell")
        Call .Map("nmap yk YankFromUpCell")
        Call .Map("nmap yl YankFromRightCell")
        Call .Map("nmap Y YankAsPlaintext")


        'Delete
        Call .Map("nmap X DeleteValue")
        Call .Map("nmap D DeleteValue")
        Call .Map("nmap dd DeleteRows")
        Call .Map("nmap dr DeleteRows")
        Call .Map("nmap dc DeleteColumns")
        Call .Map("nmap dgg DeleteRows " & eTargetRowType.ToTopRows)
        Call .Map("nmap dG DeleteRows " & eTargetRowType.ToBottomRows)
        Call .Map("nmap d{ DeleteRows " & eTargetRowType.ToTopOfCurrentRegionRows)
        Call .Map("nmap d} DeleteRows " & eTargetRowType.ToBottomOfCurrentRegionRows)
        Call .Map("nmap d0 DeleteColumns " & eTargetColumnType.ToLeftEndColumns)
        Call .Map("nmap d$ DeleteColumns " & eTargetColumnType.ToRightEndColumns)
        Call .Map("nmap d^ DeleteColumns " & eTargetColumnType.ToLeftOfCurrentRegionColumns)
        Call .Map("nmap dg$ DeleteColumns " & eTargetColumnType.ToRightOfCurrentRegionColumns)

        Call .Map("nmap dh DeleteToLeft")
        Call .Map("nmap dj DeleteToUp")
        Call .Map("nmap dk DeleteToUp")
        Call .Map("nmap dl DeleteToLeft")


        'Cut
        Call .Map("nmap xr CutRows")
        Call .Map("nmap xc CutColumns")
        Call .Map("nmap xgg CutRows " & eTargetRowType.ToTopRows)
        Call .Map("nmap xG CutRows " & eTargetRowType.ToBottomRows)
        Call .Map("nmap x{ CutRows " & eTargetRowType.ToTopOfCurrentRegionRows)
        Call .Map("nmap x} CutRows " & eTargetRowType.ToBottomOfCurrentRegionRows)
        Call .Map("nmap x0 CutColumns " & eTargetColumnType.ToLeftEndColumns)
        Call .Map("nmap x$ CutColumns " & eTargetColumnType.ToRightEndColumns)
        Call .Map("nmap x^ CutColumns " & eTargetColumnType.ToLeftOfCurrentRegionColumns)
        Call .Map("nmap xg$ CutColumns " & eTargetColumnType.ToRightOfCurrentRegionColumns)


        'Paste
        Call .Map("nmap p PasteSmart " & xlNext)
        Call .Map("nmap P PasteSmart " & xlPrevious)
        Call .Map("nmap gp PasteSpecial")
        Call .Map("nmap U PasteValue")


        'Font
        Call .Map("nmap - DecreaseFontSize")
        Call .Map("nmap + IncreaseFontSize")
        Call .Map("nmap fn ChangeFontName")
        Call .Map("nmap fs ChangeFontSize")
        Call .Map("nmap fh AlignLeft")
        Call .Map("nmap fj AlignBottom")
        Call .Map("nmap fk AlignTop")
        Call .Map("nmap fl AlignRight")
        Call .Map("nmap fo AlignCenter")
        Call .Map("nmap fm AlignMiddle")
        Call .Map("nmap fb ToggleBold")
        Call .Map("nmap fi ToggleItalic")
        Call .Map("nmap fu ToggleUnderline")
        Call .Map("nmap f- ToggleStrikethrough")
        Call .Map("nmap ft ChangeFormat")
        Call .Map("nmap ff ShowFontDialog")


        'Color
        Call .Map("nmap fc SmartFontColor")
        Call .Map("nmap FC SmartFillColor")
        Call .Map("nmap Fc SmartFillColor")
        Call .Map("nmap bc ChangeShapeBorderColor")


        'Comment
        Call .Map("nmap Ci EditCellComment")
        Call .Map("nmap Cc EditCellComment")
        Call .Map("nmap Ce DeleteCellComment")
        Call .Map("nmap Cx DeleteCellComment")
        Call .Map("nmap Cd DeleteCellComment")
        Call .Map("nmap CE DeleteCellCommentAll")
        Call .Map("nmap CD DeleteCellCommentAll")
        Call .Map("nmap Ca ToggleCellComment")
        Call .Map("nmap Cr ShowCellComment")
        Call .Map("nmap Cm HideCellComment")
        Call .Map("nmap CA ToggleCellCommentAll")
        Call .Map("nmap CR ShowCellCommentAll")
        Call .Map("nmap CM HideCellCommentAll")
        Call .Map("nmap CH HideCellCommentIndicator")
        Call .Map("nmap Cn NextCommentedCell")
        Call .Map("nmap Cp PrevCommentedCell")


        'Find & Replace
        Call .Map("nmap / ShowFindFollowLang")
        Call .Map("nmap ? ShowFindNotFollowLang")
        Call .Map("nmap n NextFoundCell")
        Call .Map("nmap N PreviousFoundCell")
        Call .Map("nmap R ShowReplaceWindow")
        Call .Map("nmap * FindActiveValueNext")
        Call .Map("nmap # FindActiveValuePrev")
        Call .Map("nmap ]c NextSpecialCells " & xlCellTypeComments)
        Call .Map("nmap [c PrevSpecialCells " & xlCellTypeComments)
        Call .Map("nmap ]o NextSpecialCells " & xlCellTypeConstants)
        Call .Map("nmap [o PrevSpecialCells " & xlCellTypeConstants)
        Call .Map("nmap ]f NextSpecialCells " & xlCellTypeFormulas)
        Call .Map("nmap [f PrevSpecialCells " & xlCellTypeFormulas)
        Call .Map("nmap ]k NextSpecialCells " & xlCellTypeBlanks)
        Call .Map("nmap [k PrevSpecialCells " & xlCellTypeBlanks)
        Call .Map("nmap ]t NextSpecialCells " & xlCellTypeSameFormatConditions)
        Call .Map("nmap [t PrevSpecialCells " & xlCellTypeSameFormatConditions)
        Call .Map("nmap ]v NextSpecialCells " & xlCellTypeSameValidation)
        Call .Map("nmap [v PrevSpecialCells " & xlCellTypeSameValidation)

        Call .Map("nmap ]s NextShape")
        Call .Map("nmap [s PrevShape")


        'Scrolling
        Call .Map("nmap <C-u> ScrollUpHalf")
        Call .Map("nmap <C-d> ScrollDownHalf")
        Call .Map("nmap <C-b> ScrollUp")
        Call .Map("nmap <C-f> ScrollDown")
        Call .Map("nmap <C-y> ScrollUp1Row")
        Call .Map("nmap <C-e> ScrollDown1Row")
        Call .Map("nmap zh ScrollLeft1Column")
        Call .Map("nmap zl ScrollRight1Column")
        Call .Map("nmap zH ScrollLeft")
        Call .Map("nmap zL ScrollRight")
        Call .Map("nmap zt ScrollCurrentTop")
        Call .Map("nmap zz ScrollCurrentMiddle")
        Call .Map("nmap zb ScrollCurrentBottom")
        Call .Map("nmap zs ScrollCurrentLeft")
        Call .Map("nmap zm ScrollCurrentCenter")
        Call .Map("nmap ze ScrollCurrentRight")


        'Worksheet Function
        Call .Map("nmap e NextWorksheet")
        Call .Map("nmap E PreviousWorksheet")

        Call .Map("nmap ww ShowSheetPicker")
        Call .Map("nmap ws ShowSheetPicker")
        Call .Map("nmap wn NextWorksheet")
        Call .Map("nmap wp PreviousWorksheet")
        Call .Map("nmap wr RenameWorksheet")
        Call .Map("nmap wh MoveWorksheetBack")
        Call .Map("nmap wl MoveWorksheetForward")
        Call .Map("nmap wi InsertWorksheet")
        Call .Map("nmap wa AppendWorksheet")
        Call .Map("nmap wd DeleteWorksheet")
        Call .Map("nmap w0 ActivateLastWorksheet")
        Call .Map("nmap w$ ActivateLastWorksheet")
        Call .Map("nmap wc ChangeWorksheetTabColor")
        Call .Map("nmap wy CloneWorksheet")
        Call .Map("nmap we ExportWorksheet")
        Call .Map("nmap w ActivateWorksheet")

        Call .Map("nmap <cmd>printpreview PrintPreviewOfActiveSheet")


        'Workbook Function
        Call .Map("nmap <cmd>e OpenWorkbook")
        Call .Map("nmap <cmd>e! ReopenActiveWorkbook")
        Call .Map("nmap <cmd>w SaveWorkbook")
        Call .Map("nmap <cmd>q CloseAskSaving")
        Call .Map("nmap <cmd>q! CloseWithoutSaving")
        Call .Map("nmap <cmd>wq CloseWithSaving")
        Call .Map("nmap <cmd>x CloseWithSaving")
        Call .Map("nmap <cmd>sav SaveAsNewWorkbook")
        Call .Map("nmap <cmd>b ActivateWorkbook")
        Call .Map("nmap <cmd>bn NextWorkbook")
        Call .Map("nmap <cmd>bp PreviousWorkbook")

        Call .Map("nmap ZZ CloseWithSaving")
        Call .Map("nmap ZQ CloseWithoutSaving")
        Call .Map("nmap ‾ ToggleReadOnly")
        Call .Map("nmap ]b NextWorkbook")
        Call .Map("nmap [b PreviousWorkbook")


        'Useful Command
        Call .Map("nmap u Undo_CtrlZ")
        Call .Map("nmap <C-r> RedoExecute")
        Call .Map("nmap . RepeatAction")
        Call .Map("nmap m ZoomIn")
        Call .Map("nmap M ZoomOut")
        Call .Map("nmap % ZoomSpecifiedScale")
        Call .Map("nmap <bslash> ShowContextMenu")

        Call .Map("nmap <C-i> JumpNext")
        Call .Map("nmap <C-o> JumpPrev")
        Call .Map("nmap <cmd>cle ClearJumps")


        'Other Commands
        Call .Map("nmap zf ToggleFreezePanes")
        Call .Map("nmap =v ToggleFormulaBar")
        Call .Map("nmap gs ShowSummaryInfo")
        Call .Map("nmap zp SetPrintArea")
        Call .Map("nmap zP ClearPrintArea")
        Call .Map("nmap @@ ShowMacroDialog")


        'Count
        Call .Map("nmap 1 ShowCmdForm " & "1")
        Call .Map("nmap 2 ShowCmdForm " & "2")
        Call .Map("nmap 3 ShowCmdForm " & "3")
        Call .Map("nmap 4 ShowCmdForm " & "4")
        Call .Map("nmap 5 ShowCmdForm " & "5")
        Call .Map("nmap 6 ShowCmdForm " & "6")
        Call .Map("nmap 7 ShowCmdForm " & "7")
        Call .Map("nmap 8 ShowCmdForm " & "8")
        Call .Map("nmap 9 ShowCmdForm " & "9")


        'KeyMapping
        Call .Map("nmap <C-[> <key><esc>")


        'Visual Mode
        Call .Map("vmap <Esc> StopVisualMode")
        Call .Map("vmap <C-[> StopVisualMode")
        Call .Map("vmap <S-0> SwapVisualBase")


        'Shape Insert Mode
        Call .Map("imap <Esc> ChangeToNormalMode")
        Call .Map("imap <C-[> ChangeToNormalMode")
    End With
End Sub
