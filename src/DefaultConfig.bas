Attribute VB_Name = "A_DefaultConfig"
' ==================================================   vim.xlam   ==================================================
' Author: SHA-5010 (@sha_5010)
'
' [Description]
' vim.xlam is an Excel add-in designed to provide a Vim-like experience,
' allowing you to navigate and operate within Excel using keyboard shortcuts.
'
' Designed with extensibility in mind, you can create your own methods and
' easily customize keybindings using the `Map` method. It's also designed
' to allow you to change keybindings easily from the default ones, so you can
' configure it to suit your preferences.
'
' LICENCE: MIT
' ==================================================================================================================

Option Explicit
Option Private Module

Sub DefaultConfig()
    With gVim.Config
        ' Note: Default value is defined in cls_Config.Class_Initialize
        ' --- Example ---
        ' .MaxHistories = 50

        .VimToggleKey = "<C-m>"
    End With

    With gVim.KeyMap
        ' --- Syntax ---
        ' [n|v|c|i]map [<cmd>]{lhs} [<key>]{rhs} [arg1] [arg2] [...]
        '  ^^^^^^^      ^^^^^ ^^^^^  ^^^^^ ^^^^^  ^^^^^^^^^^^^^^^^^
        '    |            |     |      |     |            `---- argX: Arguments of the function specified by {rhs}
        '    |            |     |      |     `----------------- rhs : Function name to be execute
        '    |            |     |      `----------------------- key : Flag to simulate keys (specify keys with {rhs})
        '    |            |     `------------------------------ lhs : Key sequence (vim style)
        '    |            `------------------------------------ cmd : Flag to enable in command mode (plain text)
        '    `------------------------------------------------- mode: Specify pre-defined mode ("n" if omitted)

        'Core
        .Map "nmap <C-p> ToggleLang"
        .Map "nmap : EnterCmdlineMode"
        .Map "nmap <cmd>debug ToggleDebugMode"


        'InsertMode
        .Map "nmap a AppendFollowLangMode"
        .Map "nmap A AppendNotFollowLangMode"
        .Map "nmap i InsertFollowLangMode"
        .Map "nmap I InsertNotFollowLangMode"
        .Map "nmap s SubstituteFollowLangMode"
        .Map "nmap S SubstituteNotFollowLangMode"


        'Moving
        .Map "nmap h MoveLeft"
        .Map "nmap j MoveDown"
        .Map "nmap k MoveUp"
        .Map "nmap l MoveRight"
        .Map "nmap H MoveLeftWithShift"
        .Map "nmap J MoveDownWithShift"
        .Map "nmap K MoveUpWithShift"
        .Map "nmap L MoveRightWithShift"
        .Map "nmap <C-h> MoveLeft"
        .Map "nmap <C-j> MoveDown"
        .Map "nmap <C-k> MoveUp"
        .Map "nmap <C-l> MoveRight"
        .Map "nmap <C-S-H> MoveLeft"
        .Map "nmap <C-S-J> MoveDown"
        .Map "nmap <C-S-K> MoveUp"
        .Map "nmap <C-S-L> MoveRight"

        .Map "nmap gg MoveToTopRow"
        .Map "nmap G MoveToLastRow"
        .Map "nmap <bar> MoveToNthColumn"
        .Map "nmap 0 MoveToFirstColumn"
        .Map "nmap ^ MoveToLeftEnd"
        .Map "nmap $ MoveToRightEnd"
        .Map "nmap g0 MoveToA1"

        .Map "nmap { MoveToTopOfCurrentRegion"
        .Map "nmap } MoveToBottomOfCurrentRegion"

        .Map "nmap W MoveToSpecifiedCell"


        'Cell
        .Map "nmap xx CutCell"
        .Map "nmap yy YankCell"
        .Map "nmap o InsertCellsDown"
        .Map "nmap O InsertCellsUp"
        .Map "nmap t InsertCellsRight"
        .Map "nmap T InsertCellsLeft"
        .Map "nmap > IncrementText"
        .Map "nmap <lt> DecrementText"
        .Map "nmap ( IncreaseDecimal"
        .Map "nmap ) DecreaseDecimal"
        .Map "nmap zw ToggleWrapText"
        .Map "nmap & ToggleMergeCells"
        .Map "nmap f, ApplyCommaStyle"
        .Map "nmap <Space> UnionSelectCells"
        .Map "nmap <S-Space> ExceptSelectCells"
        .Map "nmap <S-BS> ClearSelectCells"
        .Map "nmap gf FollowHyperlinkOfActiveCell"
        .Map "nmap Ff ApplyFlashFill"
        .Map "nmap FF ApplyFlashFill"
        .Map "nmap Fa ApplyAutoFill"
        .Map "nmap FA ApplyAutoFill"

        .Map "nmap =s AutoSum"
        .Map "nmap =a AutoAverage"
        .Map "nmap =c AutoCount"
        .Map "nmap =m AutoMax"
        .Map "nmap =i AutoMin"
        .Map "nmap == InsertFunction"

        .Map "nmap v ToggleVisualMode"
        .Map "nmap V ToggleVisualLine"


        ' Border
        .Map "nmap bb ToggleBorderAll " & xlContinuous & " " & xlThin
        .Map "nmap ba ToggleBorderAround " & xlContinuous & " " & xlThin
        .Map "nmap bh ToggleBorderLeft " & xlContinuous & " " & xlThin
        .Map "nmap bj ToggleBorderBottom " & xlContinuous & " " & xlThin
        .Map "nmap bk ToggleBorderTop " & xlContinuous & " " & xlThin
        .Map "nmap bl ToggleBorderRight " & xlContinuous & " " & xlThin
        .Map "nmap bia ToggleBorderInner " & xlContinuous & " " & xlThin
        .Map "nmap bis ToggleBorderInnerHorizontal " & xlContinuous & " " & xlThin
        .Map "nmap biv ToggleBorderInnerVertical " & xlContinuous & " " & xlThin
        .Map "nmap b/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlThin
        .Map "nmap b<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlThin

        .Map "nmap bB ToggleBorderAll " & xlContinuous & " " & xlThick
        .Map "nmap bA ToggleBorderAround " & xlContinuous & " " & xlThick
        .Map "nmap bH ToggleBorderLeft " & xlContinuous & " " & xlThick
        .Map "nmap bJ ToggleBorderBottom " & xlContinuous & " " & xlThick
        .Map "nmap bK ToggleBorderTop " & xlContinuous & " " & xlThick
        .Map "nmap bL ToggleBorderRight " & xlContinuous & " " & xlThick
        .Map "nmap Bb ToggleBorderAll " & xlContinuous & " " & xlThick
        .Map "nmap Ba ToggleBorderAround " & xlContinuous & " " & xlThick
        .Map "nmap Bh ToggleBorderLeft " & xlContinuous & " " & xlThick
        .Map "nmap Bj ToggleBorderBottom " & xlContinuous & " " & xlThick
        .Map "nmap Bk ToggleBorderTop " & xlContinuous & " " & xlThick
        .Map "nmap Bl ToggleBorderRight " & xlContinuous & " " & xlThick
        .Map "nmap Bia ToggleBorderInner " & xlContinuous & " " & xlThick
        .Map "nmap Bis ToggleBorderInnerHorizontal " & xlContinuous & " " & xlThick
        .Map "nmap Biv ToggleBorderInnerVertical " & xlContinuous & " " & xlThick
        .Map "nmap B/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlThick
        .Map "nmap B<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlThick

        .Map "nmap bob ToggleBorderAll " & xlContinuous & " " & xlHairline
        .Map "nmap boa ToggleBorderAround " & xlContinuous & " " & xlHairline
        .Map "nmap boh ToggleBorderLeft " & xlContinuous & " " & xlHairline
        .Map "nmap boj ToggleBorderBottom " & xlContinuous & " " & xlHairline
        .Map "nmap bok ToggleBorderTop " & xlContinuous & " " & xlHairline
        .Map "nmap bol ToggleBorderRight " & xlContinuous & " " & xlHairline
        .Map "nmap boia ToggleBorderInner " & xlContinuous & " " & xlHairline
        .Map "nmap bois ToggleBorderInnerHorizontal " & xlContinuous & " " & xlHairline
        .Map "nmap boiv ToggleBorderInnerVertical " & xlContinuous & " " & xlHairline
        .Map "nmap bo/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlHairline
        .Map "nmap bo<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlHairline

        .Map "nmap bmb ToggleBorderAll " & xlContinuous & " " & xlMedium
        .Map "nmap bma ToggleBorderAround " & xlContinuous & " " & xlMedium
        .Map "nmap bmh ToggleBorderLeft " & xlContinuous & " " & xlMedium
        .Map "nmap bmj ToggleBorderBottom " & xlContinuous & " " & xlMedium
        .Map "nmap bmk ToggleBorderTop " & xlContinuous & " " & xlMedium
        .Map "nmap bml ToggleBorderRight " & xlContinuous & " " & xlMedium
        .Map "nmap bmia ToggleBorderInner " & xlContinuous & " " & xlMedium
        .Map "nmap bmis ToggleBorderInnerHorizontal " & xlContinuous & " " & xlMedium
        .Map "nmap bmiv ToggleBorderInnerVertical " & xlContinuous & " " & xlMedium
        .Map "nmap bm/ ToggleBorderDiagonalUp " & xlContinuous & " " & xlMedium
        .Map "nmap bm<bslash> ToggleBorderDiagonalDown " & xlContinuous & " " & xlMedium

        .Map "nmap btb ToggleBorderAll " & xlDouble & " " & xlThick
        .Map "nmap bta ToggleBorderAround " & xlDouble & " " & xlThick
        .Map "nmap bth ToggleBorderLeft " & xlDouble & " " & xlThick
        .Map "nmap btj ToggleBorderBottom " & xlDouble & " " & xlThick
        .Map "nmap btk ToggleBorderTop " & xlDouble & " " & xlThick
        .Map "nmap btl ToggleBorderRight " & xlDouble & " " & xlThick
        .Map "nmap btia ToggleBorderInner " & xlDouble & " " & xlThick
        .Map "nmap btis ToggleBorderInnerHorizontal " & xlDouble & " " & xlThick
        .Map "nmap btiv ToggleBorderInnerVertical " & xlDouble & " " & xlThick
        .Map "nmap bt/ ToggleBorderDiagonalUp " & xlDouble & " " & xlThick
        .Map "nmap bt<bslash> ToggleBorderDiagonalDown " & xlDouble & " " & xlThick

        .Map "nmap bdd DeleteBorderAll"
        .Map "nmap bda DeleteBorderAround"
        .Map "nmap bdh DeleteBorderLeft"
        .Map "nmap bdj DeleteBorderBottom"
        .Map "nmap bdk DeleteBorderTop"
        .Map "nmap bdl DeleteBorderRight"
        .Map "nmap bdia DeleteBorderInner"
        .Map "nmap bdis DeleteBorderInnerHorizontal"
        .Map "nmap bdiv DeleteBorderInnerVertical"
        .Map "nmap bd/ DeleteBorderDiagonalUp"
        .Map "nmap bd<bslash> DeleteBorderDiagonalDown"

        .Map "nmap bcc SetBorderColorAll"
        .Map "nmap bca SetBorderColorAround"
        .Map "nmap bch SetBorderColorLeft"
        .Map "nmap bcj SetBorderColorBottom"
        .Map "nmap bck SetBorderColorTop"
        .Map "nmap bcl SetBorderColorRight"
        .Map "nmap bcia SetBorderColorInner"
        .Map "nmap bcis SetBorderColorInnerHorizontal"
        .Map "nmap bciv SetBorderColorInnerVertical"
        .Map "nmap bc/ SetBorderColorDiagonalUp"
        .Map "nmap bc<bslash> SetBorderColorDiagonalDown"


        'Row
        .Map "nmap r- NarrowRowsHeight"
        .Map "nmap r+ WideRowsHeight"
        .Map "nmap rr SelectRows"
        .Map "nmap ra AppendRows"
        .Map "nmap ri InsertRows"
        .Map "nmap rd DeleteRows"
        .Map "nmap ry YankRows"
        .Map "nmap rx CutRows"
        .Map "nmap rh HideRows"
        .Map "nmap rH UnhideRows"
        .Map "nmap rg GroupRows"
        .Map "nmap ru UngroupRows"
        .Map "nmap rf FoldRowsGroup"
        .Map "nmap rs SpreadRowsGroup"
        .Map "nmap rj AdjustRowsHeight"
        .Map "nmap rw SetRowsHeight"


        'Column
        .Map "nmap c- NarrowColumnsWidth"
        .Map "nmap c+ WideColumnsWidth"
        .Map "nmap cc SelectColumns"
        .Map "nmap ca AppendColumns"
        .Map "nmap ci InsertColumns"
        .Map "nmap cd DeleteColumns"
        .Map "nmap cy YankColumns"
        .Map "nmap cx CutColumns"
        .Map "nmap ch HideColumns"
        .Map "nmap cH UnhideColumns"
        .Map "nmap cg GroupColumns"
        .Map "nmap cu UngroupColumns"
        .Map "nmap cf FoldColumnsGroup"
        .Map "nmap cs SpreadColumnsGroup"
        .Map "nmap cj AdjustColumnsWidth"
        .Map "nmap cw SetColumnsWidth"


        'Yank
        .Map "nmap yr YankRows"
        .Map "nmap yc YankColumns"
        .Map "nmap ygg YankRows " & eTargetRowType.ToTopRows
        .Map "nmap yG YankRows " & eTargetRowType.ToBottomRows
        .Map "nmap y{ YankRows " & eTargetRowType.ToTopOfCurrentRegionRows
        .Map "nmap y} YankRows " & eTargetRowType.ToBottomOfCurrentRegionRows
        .Map "nmap y0 YankColumns " & eTargetColumnType.ToLeftEndColumns
        .Map "nmap y$ YankColumns " & eTargetColumnType.ToRightEndColumns
        .Map "nmap y^ YankColumns " & eTargetColumnType.ToLeftOfCurrentRegionColumns
        .Map "nmap yg$ YankColumns " & eTargetColumnType.ToRightOfCurrentRegionColumns

        .Map "nmap yh YankFromLeftCell"
        .Map "nmap yj YankFromDownCell"
        .Map "nmap yk YankFromUpCell"
        .Map "nmap yl YankFromRightCell"
        .Map "nmap Y YankAsPlaintext"


        'Delete
        .Map "nmap X DeleteValue"
        .Map "nmap D DeleteValue"
        .Map "nmap dd DeleteRows"
        .Map "nmap dr DeleteRows"
        .Map "nmap dc DeleteColumns"
        .Map "nmap dgg DeleteRows " & eTargetRowType.ToTopRows
        .Map "nmap dG DeleteRows " & eTargetRowType.ToBottomRows
        .Map "nmap d{ DeleteRows " & eTargetRowType.ToTopOfCurrentRegionRows
        .Map "nmap d} DeleteRows " & eTargetRowType.ToBottomOfCurrentRegionRows
        .Map "nmap d0 DeleteColumns " & eTargetColumnType.ToLeftEndColumns
        .Map "nmap d$ DeleteColumns " & eTargetColumnType.ToRightEndColumns
        .Map "nmap d^ DeleteColumns " & eTargetColumnType.ToLeftOfCurrentRegionColumns
        .Map "nmap dg$ DeleteColumns " & eTargetColumnType.ToRightOfCurrentRegionColumns

        .Map "nmap dh DeleteToLeft"
        .Map "nmap dj DeleteToUp"
        .Map "nmap dk DeleteToUp"
        .Map "nmap dl DeleteToLeft"


        'Cut
        .Map "nmap xr CutRows"
        .Map "nmap xc CutColumns"
        .Map "nmap xgg CutRows " & eTargetRowType.ToTopRows
        .Map "nmap xG CutRows " & eTargetRowType.ToBottomRows
        .Map "nmap x{ CutRows " & eTargetRowType.ToTopOfCurrentRegionRows
        .Map "nmap x} CutRows " & eTargetRowType.ToBottomOfCurrentRegionRows
        .Map "nmap x0 CutColumns " & eTargetColumnType.ToLeftEndColumns
        .Map "nmap x$ CutColumns " & eTargetColumnType.ToRightEndColumns
        .Map "nmap x^ CutColumns " & eTargetColumnType.ToLeftOfCurrentRegionColumns
        .Map "nmap xg$ CutColumns " & eTargetColumnType.ToRightOfCurrentRegionColumns


        'Paste
        .Map "nmap p PasteSmart " & xlNext
        .Map "nmap P PasteSmart " & xlPrevious
        .Map "nmap gp PasteSpecial"
        .Map "nmap U PasteValue"


        'Font
        .Map "nmap - DecreaseFontSize"
        .Map "nmap + IncreaseFontSize"
        .Map "nmap fn ChangeFontName"
        .Map "nmap fs ChangeFontSize"
        .Map "nmap fh AlignLeft"
        .Map "nmap fj AlignBottom"
        .Map "nmap fk AlignTop"
        .Map "nmap fl AlignRight"
        .Map "nmap fo AlignCenter"
        .Map "nmap fm AlignMiddle"
        .Map "nmap fb ToggleBold"
        .Map "nmap fi ToggleItalic"
        .Map "nmap fu ToggleUnderline"
        .Map "nmap f- ToggleStrikethrough"
        .Map "nmap ft ChangeFormat"
        .Map "nmap ff ShowFontDialog"


        'Color
        .Map "nmap fc SmartFontColor"
        .Map "nmap FC SmartFillColor"
        .Map "nmap Fc SmartFillColor"
        .Map "nmap bc ChangeShapeBorderColor"


        'Comment
        .Map "nmap Ci EditCellComment"
        .Map "nmap Cc EditCellComment"
        .Map "nmap Ce DeleteCellComment"
        .Map "nmap Cx DeleteCellComment"
        .Map "nmap Cd DeleteCellComment"
        .Map "nmap CE DeleteCellCommentAll"
        .Map "nmap CD DeleteCellCommentAll"
        .Map "nmap Ca ToggleCellComment"
        .Map "nmap Cr ShowCellComment"
        .Map "nmap Cm HideCellComment"
        .Map "nmap CA ToggleCellCommentAll"
        .Map "nmap CR ShowCellCommentAll"
        .Map "nmap CM HideCellCommentAll"
        .Map "nmap CH HideCellCommentIndicator"
        .Map "nmap Cn NextCommentedCell"
        .Map "nmap Cp PrevCommentedCell"


        'Find & Replace
        .Map "nmap / ShowFindFollowLang"
        .Map "nmap ? ShowFindNotFollowLang"
        .Map "nmap n NextFoundCell"
        .Map "nmap N PreviousFoundCell"
        .Map "nmap R ShowReplaceWindow"
        .Map "nmap * FindActiveValueNext"
        .Map "nmap # FindActiveValuePrev"
        .Map "nmap ]c NextSpecialCells " & xlCellTypeComments
        .Map "nmap [c PrevSpecialCells " & xlCellTypeComments
        .Map "nmap ]o NextSpecialCells " & xlCellTypeConstants
        .Map "nmap [o PrevSpecialCells " & xlCellTypeConstants
        .Map "nmap ]f NextSpecialCells " & xlCellTypeFormulas
        .Map "nmap [f PrevSpecialCells " & xlCellTypeFormulas
        .Map "nmap ]k NextSpecialCells " & xlCellTypeBlanks
        .Map "nmap [k PrevSpecialCells " & xlCellTypeBlanks
        .Map "nmap ]t NextSpecialCells " & xlCellTypeSameFormatConditions
        .Map "nmap [t PrevSpecialCells " & xlCellTypeSameFormatConditions
        .Map "nmap ]v NextSpecialCells " & xlCellTypeSameValidation
        .Map "nmap [v PrevSpecialCells " & xlCellTypeSameValidation

        .Map "nmap ]s NextShape"
        .Map "nmap [s PrevShape"


        'Scrolling
        .Map "nmap <C-u> ScrollUpHalf"
        .Map "nmap <C-d> ScrollDownHalf"
        .Map "nmap <C-b> ScrollUp"
        .Map "nmap <C-f> ScrollDown"
        .Map "nmap <C-y> ScrollUp1Row"
        .Map "nmap <C-e> ScrollDown1Row"
        .Map "nmap zh ScrollLeft1Column"
        .Map "nmap zl ScrollRight1Column"
        .Map "nmap zH ScrollLeft"
        .Map "nmap zL ScrollRight"
        .Map "nmap zt ScrollCurrentTop"
        .Map "nmap zz ScrollCurrentMiddle"
        .Map "nmap zb ScrollCurrentBottom"
        .Map "nmap zs ScrollCurrentLeft"
        .Map "nmap zm ScrollCurrentCenter"
        .Map "nmap ze ScrollCurrentRight"


        'Worksheet Function
        .Map "nmap e NextWorksheet"
        .Map "nmap E PreviousWorksheet"

        .Map "nmap ww ShowSheetPicker"
        .Map "nmap ws ShowSheetPicker"
        .Map "nmap wn NextWorksheet"
        .Map "nmap wp PreviousWorksheet"
        .Map "nmap wr RenameWorksheet"
        .Map "nmap wh MoveWorksheetBack"
        .Map "nmap wl MoveWorksheetForward"
        .Map "nmap wi InsertWorksheet"
        .Map "nmap wa AppendWorksheet"
        .Map "nmap wd DeleteWorksheet"
        .Map "nmap w0 ActivateLastWorksheet"
        .Map "nmap w$ ActivateLastWorksheet"
        .Map "nmap wc ChangeWorksheetTabColor"
        .Map "nmap wy CloneWorksheet"
        .Map "nmap we ExportWorksheet"
        .Map "nmap w ActivateWorksheet"

        .Map "nmap <cmd>printpreview PrintPreviewOfActiveSheet"


        'Workbook Function
        .Map "nmap <cmd>e OpenWorkbook"
        .Map "nmap <cmd>e! ReopenActiveWorkbook"
        .Map "nmap <cmd>w SaveWorkbook"
        .Map "nmap <cmd>q CloseAskSaving"
        .Map "nmap <cmd>q! CloseWithoutSaving"
        .Map "nmap <cmd>wq CloseWithSaving"
        .Map "nmap <cmd>x CloseWithSaving"
        .Map "nmap <cmd>sav SaveAsNewWorkbook"
        .Map "nmap <cmd>b ActivateWorkbook"
        .Map "nmap <cmd>bn NextWorkbook"
        .Map "nmap <cmd>bp PreviousWorkbook"

        .Map "nmap ZZ CloseWithSaving"
        .Map "nmap ZQ CloseWithoutSaving"
        .Map "nmap ‾ ToggleReadOnly"
        .Map "nmap ]b NextWorkbook"
        .Map "nmap [b PreviousWorkbook"


        'Useful Command
        .Map "nmap u Undo_CtrlZ"
        .Map "nmap <C-r> RedoExecute"
        .Map "nmap . RepeatAction"
        .Map "nmap m ZoomIn"
        .Map "nmap M ZoomOut"
        .Map "nmap % ZoomSpecifiedScale"
        .Map "nmap <bslash> ShowContextMenu"

        .Map "nmap <C-i> JumpNext"
        .Map "nmap <C-o> JumpPrev"
        .Map "nmap <cmd>cle ClearJumps"


        'Other Commands
        .Map "nmap zf ToggleFreezePanes"
        .Map "nmap =v ToggleFormulaBar"
        .Map "nmap gs ShowSummaryInfo"
        .Map "nmap zp SetPrintArea"
        .Map "nmap zP ClearPrintArea"
        .Map "nmap @@ ShowMacroDialog"


        'Count
        .Map "nmap 1 ShowCmdForm " & """1"""
        .Map "nmap 2 ShowCmdForm " & """2"""
        .Map "nmap 3 ShowCmdForm " & """3"""
        .Map "nmap 4 ShowCmdForm " & """4"""
        .Map "nmap 5 ShowCmdForm " & """5"""
        .Map "nmap 6 ShowCmdForm " & """6"""
        .Map "nmap 7 ShowCmdForm " & """7"""
        .Map "nmap 8 ShowCmdForm " & """8"""
        .Map "nmap 9 ShowCmdForm " & """9"""


        'KeyMapping
        .Map "nmap <C-[> <key><esc>"


        'Visual Mode
        .Map "vmap <Esc> StopVisualMode"
        .Map "vmap <C-[> StopVisualMode"
        .Map "vmap <S-0> SwapVisualBase"


        'Shape Insert Mode
        .Map "imap <Esc> ChangeToNormalMode"
        .Map "imap <C-[> ChangeToNormalMode"


        'Cmdline Mode
        .Map "cmap <C-w> <key><C-BS>"
        .Map "cmap <C-u> <key><S-Home><BS>"
        .Map "cmap <C-k> <key><S-End><Del>"
        .Map "cmap <C-h> <key><Left>"
        .Map "cmap <C-l> <key><Right>"
        .Map "cmap <C-a> <key><Home>"
        .Map "cmap <C-e> <key><End>"
    End With
End Sub
