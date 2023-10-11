<p align="center">
  <a href="https://github.com/sha5010/vim.xlam">
    <img alt="vim xlam_banner" src="https://user-images.githubusercontent.com/95682647/175554011-5f9b5a37-a08d-47f7-ac63-b162620cc99d.png" width="600">
  </a>
</p>

<p align="center">
  <a href="https://github.com/sha5010/vim.xlam/releases/latest">
    <img alt="release" src="https://img.shields.io/github/v/release/sha5010/vim.xlam">
  </a>
  <a href="./LICENSE">
    <img alt="license" src="https://img.shields.io/github/license/sha5010/vim.xlam">
  </a>
  <a href="https://twitter.com/sha_5010">
    <img alt="Twitter Follow" src="https://img.shields.io/twitter/follow/sha_5010?style=social">
  </a>
</p>

# vim.xlam

Vim experience in Excel. This is an Excel add-in that allows you to use Vim keybindings within Excel.

## Description

vim.xlam is an Excel add-in designed to provide a Vim-like experience, allowing you to navigate and operate within Excel using keyboard shortcuts.

Designed with extensibility in mind, you can create your own methods and easily customize keybindings using the `map` method. It's also designed to allow you to change keybindings easily from the default ones, so you can configure it to suit your preferences.

***Demo:***

![demo](https://user-images.githubusercontent.com/95682647/175773473-50376812-afcc-4ced-b436-7150d7b97872.gif)

\* Sample file courtesy of [https://atelierkobato.com](https://atelierkobato.com/download/).

## Features

- Supports not only basic cell navigation using `hjkl` but also various jump commands like `gg`, `G`, `^`, `$`.
- Efficiently perform tasks such as font, background color, and border settings without mouse interaction.
- Equipped with features for commenting, scrolling, and worksheet operations.
- Remembers the last edited cell and the cell before a jump, providing a jump list feature.
- Designed for easy customization, making it accessible to anyone.

## Installation

1. Download the latest vim.xlam from the [Release Page](https://github.com/sha5010/vim.xlam/releases/latest) (or [directly download the latest version](https://github.com/sha5010/vim.xlam/releases/latest/download/vim.xlam)).
2. Save the downloaded vim.xlam in `C:\Users\<USERNAME>\AppData\Roaming\Microsoft\AddIns`.
3. Launch Excel, go to File &gt; Options &gt; Add-Ins, and click the **Go...** button at the bottom of the screen.
4. Click the **Browse...** button, select the saved vim.xlam, and add the add-in.

## Usage

- By default, you can toggle Vim mode on/off with the `Ctrl + M` key combination.
- You can navigate cells using `hjkl`, and also perform cell editing with commands like `a` and `i`.
- Many other commands are available for use.


Please note that the default configuration is tailored for Japanese users. If you are not a Japanese user, you should modify the following section in [UserConfig.bas](./src/UserConfig.bas) to set it to `False`:

```vb
Public Const DEFAULT_LANG_JA As Boolean = True
```

### Default Keybindings

**Primary Commands**

| Type | Keystroke | Action | Description | Count |
| ---- | --------- | ------ | ----------- | ----- |
| Core | `<C-m>` | `toggleVim` | Toggle Vim mode | |
| InsertMode | `a` | `appendFollowLangMode` | Edit cell from the end, following language mode for IME | |
| InsertMode | `i` | `insertFollowLangMode` | Edit cell from the start, following language mode for IME | |
| InsertMode | `s` | `substituteFollowLangMode` | Clear and edit cell, following language mode for IME | |
| Moving | `h` | `moveLeft` | ← | ✓ |
| Moving | `j` | `moveDown` | ↓ | ✓ |
| Moving | `k` | `moveUp` | ↑ | ✓ |
| Moving | `l` | `moveRight` | → | ✓ |
| Moving | `gg` | `moveToTopRow` | Move to the 1st row or the `[count]`-th row | ✓ |
| Cell | `FF`/`Ff` | `applyFlashFill` | Flash Fill (fallback to Auto Fill if not applicable) | |
| Cell | `v` | `toggleVisualMode` | Toggle visual mode (extend selection) | |
| Border | `bb` | `toggleBorderAll` | Set borders around and inside the selected cells (solid, thin line) | |
| Border | `ba` | `toggleBorderAround` | Set borders around the selected cells (solid, thin line) | |
| Border | `bia` | `toggleBorderInner` | Set borders inside the selected cells (solid, thin line) | |
| Row | `ra` | `appendRows` | Insert rows below the current row | ✓ |
| Row | `ri` | `insertRows` | Insert rows above the current row | ✓ |
| Row | `rd` | `deleteRows` | Delete the current row | ✓ |
| Column | `ca` | `appendColumns` | Insert columns to the right of the current column | ✓ |
| Column | `ci` | `insertColumns` | Insert columns to the left of the current column | ✓ |
| Column | `cd` | `deleteColumns` | Delete the current column | ✓ |
| Delete | `D`/`X` | `deleteValue` | Delete the value in the cell | |
| Paste | `p` | `pasteSmart` | Paste after copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Paste | `P` | `pasteSmart` | Paste before copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Font | `-` | `decreaseFontSize` | Decrease font size | |
| Font | `+` | `increaseFontSize` | Increase font size | |
| Color | `fc` | `smartFontColor` | Show the font color selection dialog | |
| Find & Replace | `/` | `showFindFollowLang` | Show the find dialog, following the language mode of IME | |
| Find & Replace | `n` | `nextFoundCell` | Select the next found cell | ✓ |
| Find & Replace | `N` | `previousFoundCell` | Select the previous found cell | ✓ |
| Scrolling | `<C-u>` | `scrollUpHalf` | Scroll up by half a page | |
| Scrolling | `<C-d>` | `scrollDownHalf` | Scroll down by half a page | |
| Scrolling | `zt` | `scrollCurrentTop` | Scroll to make the current row at the top (`SCREEN_OFFSET` pts of padding) | |
| Scrolling | `zz` | `scrollCurrentMiddle` | Scroll to make the current row in the middle | |
| Scrolling | `zb` | `scrollCurrentBottom` | Scroll to make the current row at the bottom (`SCREEN_OFFSET` pts of padding) | |
| Worksheet | `e` | `nextWorksheet` | Select the next worksheet | |
| Worksheet | `E` | `previousWorksheet` | Select the previous worksheet | |
| Worksheet | `ww` | `showSheetPicker` | Launch the Sheet Picker | |
| Worksheet | `wr` | `renameWorksheet` | Change the name of the active worksheet | |
| Workbook | `:w` | `saveWorkbook` | Save the active workbook | |
| Workbook | `:q` | `closeAskSaving` | Close the active workbook (show a dialog if there are unsaved changes) | |
| Workbook | `:q!`/`ZQ` | `closeWithoutSaving` | Close the active workbook without saving | |
| Workbook | `:wq`/`:x`/`ZZ` | `closeWithSaving` | Save and close the active workbook | |
| Other | `u` | `undo_CtrlZ` | Undo (send `Ctrl + Z`) | |
| Other | `<C-r>` | `redoExecute` | Redo | |

<details><summary>Expand all commands</summary><div>

| Type | Keystroke | Action | Description | Count |
| ---- | --------- | ------ | ----------- | ----- |
| Core | `<C-m>` | `toggleVim` | Toggle Vim mode | |
| Core | `<C-p>` | `toggleLang` | Toggle language mode (Japanese/English) | |
| Core | `:r` | `reloadVim` | Reload vim.xlam | |
| Core | `:r!` | `reloadVim` | Reload vim.xlam (reapply keybindings) | |
| Core | `:debug` | `toggleDebugMode` | Toggle debug mode | |
| InsertMode | `a` | `appendFollowLangMode` | Edit cell from the end, following language mode for IME | |
| InsertMode | `A` | `appendNotFollowLangMode` | Edit cell from the end without following language mode for IME | |
| InsertMode | `i` | `insertFollowLangMode` | Edit cell from the start, following language mode for IME | |
| InsertMode | `I` | `insertNotFollowLangMode` | Edit cell from the start without following language mode for IME | |
| InsertMode | `s` | `substituteFollowLangMode` | Clear and edit cell, following language mode for IME | |
| InsertMode | `S` | `substituteNotFollowLangMode` | Clear and edit cell without following language mode for IME | |
| Moving | `h` | `moveLeft` | ← | ✓ |
| Moving | `j` | `moveDown` | ↓ | ✓ |
| Moving | `k` | `moveUp` | ↑ | ✓ |
| Moving | `l` | `moveRight` | → | ✓ |
| Moving | `H` | `moveLeft` | Shift + ← | ✓ |
| Moving | `J` | `moveDown` | Shift + ↓ | ✓ |
| Moving | `K` | `moveUp` | Shift + ↑ | ✓ |
| Moving | `L` | `moveRight` | Shift + → | ✓ |
| Moving | `<C-h>` | `moveLeft` | Ctrl + ← | |
| Moving | `<C-j>` | `moveDown` | Ctrl + ↓ | |
| Moving | `<C-k>` | `moveUp` | Ctrl + ↑ | |
| Moving | `<C-l>` | `moveRight` | Ctrl + → | |
| Moving | `<C-S-H>` | `moveLeft` | Ctrl + Shift + ← | |
| Moving | `<C-S-J>` | `moveDown` | Ctrl + Shift + ↓ | |
| Moving | `<C-S-K>` | `moveUp` | Ctrl + Shift + ↑ | |
| Moving | `<C-S-L>` | `moveRight` | Ctrl + Shift + → | |
| Moving | `gg` | `moveToTopRow` | Move to the 1st row or the `[count]`-th row | ✓ |
| Moving | `G` | `moveToLastRow` | Move to the last row of UsedRange or the `[count]`-th row | ✓ |
| Moving | `\|` | `moveToNthColumn` | Move to the `[count]`-th column | ✓ |
| Moving | `0` | `moveToFirstColumn` | Move to the 1st column | |
| Moving | `^` | `moveToLeftEnd` | Move to the first column of UsedRange | |
| Moving | `$` | `moveToRightEnd` | Move to the last column of UsedRange | |
| Moving | `g0` | `moveToA1` | Move to cell A1 | |
| Moving | `{` | `moveToTopOfCurrentRegion` | Move to the first row within the CurrentRegion | |
| Moving | `}` | `moveToBottomOfCurrentRegion` | Move to the last row within the CurrentRegion | |
| Moving | `W[cell]` | `moveToSpecifiedCell` | Move to the specified `[cell]` | |
| Moving | `:[num]` | `moveToSpecifiedRow` | Move to the specified `[num]`-th row | |
| Cell | `xx` | `cutCell` | Cut cell | |
| Cell | `yy` | `yankCell` | Copy cell | |
| Cell | `o` | `insertCellsDown` | Insert cells below the selected cells | ✓ |
| Cell | `O` | `insertCellsUp` | Insert cells above the selected cells | ✓ |
| Cell | `t` | `insertCellsRight` | Insert cells to the right of the selected cells | ✓ |
| Cell | `T` | `insertCellsLeft` | Insert cells to the left of the selected cells | ✓ |
| Cell | `>` | `incrementText` | Increase indentation | ✓ |
| Cell | `<` | `decrementText` | Decrease indentation | ✓ |
| Cell | `(` | `increaseDecimal` | Increase decimal places | ✓ |
| Cell | `)` | `decreaseDecimal` | Decrease decimal places | ✓ |
| Cell | `zw` | `toggleWrapText` | Toggle cell wrap text | |
| Cell | `&` | `toggleMergeCells` | Toggle cell merge | |
| Cell | `f,` | `applyCommaStyle` | Apply comma style | |
| Cell | `<Space>` | `unionSelectCells` | Add the current cell to the selection memory and select the remembered cells (allows selecting multiple cells) | |
| Cell | `<S-Space>` | `exceptSelectCells` | Remove the current cell from the remembered selected cells | |
| Cell | `<S-BS>` | `clearSelectCells` | Clear the remembered selected cells | |
| Cell | `gf` | `followHyperlinkOfActiveCell` | Open the hyperlink in the cell | |
| Cell | `FF`/`Ff` | `applyFlashFill` | Flash Fill (fallback to Auto Fill if not applicable) | |
| Cell | `FA`/`Fa` | `applyAutoFill` | Auto Fill | |
| Cell | `=s` | `autoSum` | Auto SUM | |
| Cell | `=a` | `autoAverage` | Auto SUM (average) | |
| Cell | `=c` | `autoCount` | Auto SUM (count) | |
| Cell | `=m` | `autoMax` | Auto SUM (maximum) | |
| Cell | `=i` | `autoMin` | Auto SUM (minimum) | |
| Cell | `==` | `insertFunction` | Insert function | |
| Cell | `v` | `toggleVisualMode` | Toggle visual mode (extend selection) | |
| Cell | `V` | `toggleVisualLine` | Toggle visual line mode (extend selection) | |
| Border | `bb` | `toggleBorderAll` | Set borders around and inside the selected cells (solid, thin line) | |
| Border | `ba` | `toggleBorderAround` | Set borders around the selected cells (solid, thin line) | |
| Border | `bh` | `toggleBorderLeft` | Set left borders of the selected cells (solid, thin line) | |
| Border | `bj` | `toggleBorderBottom` | Set bottom borders of the selected cells (solid, thin line) | |
| Border | `bk` | `toggleBorderTop` | Set top borders of the selected cells (solid, thin line) | |
| Border | `bl` | `toggleBorderRight` | Set right borders of the selected cells (solid, thin line) | |
| Border | `bia` | `toggleBorderInner` | Set borders inside the selected cells (solid, thin line) | |
| Border | `bis` | `toggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, thin line) | |
| Border | `biv` | `toggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, thin line) | |
| Border | `b/` | `toggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, thin line) | |
| Border | `b\` | `toggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, thin line) | |
| Border | `bB` | `toggleBorderAll` | Set borders around and inside the selected cells (solid, thick line) | |
| Border | `bA` | `toggleBorderAround` | Set borders around the selected cells (solid, thick line) | |
| Border | `bH` | `toggleBorderLeft` | Set left borders of the selected cells (solid, thick line) | |
| Border | `bJ` | `toggleBorderBottom` | Set bottom borders of the selected cells (solid, thick line) | |
| Border | `bK` | `toggleBorderTop` | Set top borders of the selected cells (solid, thick line) | |
| Border | `bL` | `toggleBorderRight` | Set right borders of the selected cells (solid, thick line) | |
| Border | `Bb` | `toggleBorderAll` | Set borders around and inside the selected cells (solid, thick line) | |
| Border | `Ba` | `toggleBorderAround` | Set borders around the selected cells (solid, thick line) | |
| Border | `Bh` | `toggleBorderLeft` | Set left borders of the selected cells (solid, thick line) | |
| Border | `Bj` | `toggleBorderBottom` | Set bottom borders of the selected cells (solid, thick line) | |
| Border | `Bk` | `toggleBorderTop` | Set top borders of the selected cells (solid, thick line) | |
| Border | `Bl` | `toggleBorderRight` | Set right borders of the selected cells (solid, thick line) | |
| Border | `Bia` | `toggleBorderInner` | Set borders inside the selected cells (solid, thick line) | |
| Border | `Bis` | `toggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, thick line) | |
| Border | `Biv` | `toggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, thick line) | |
| Border | `B/` | `toggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, thick line) | |
| Border | `B\` | `toggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, thick line) | |
| Border | `bob` | `toggleBorderAll` | Set borders around and inside the selected cells (solid, hairline) | |
| Border | `boa` | `toggleBorderAround` | Set borders around the selected cells (solid, hairline) | |
| Border | `boh` | `toggleBorderLeft` | Set left borders of the selected cells (solid, hairline) | |
| Border | `boj` | `toggleBorderBottom` | Set bottom borders of the selected cells (solid, hairline) | |
| Border | `bok` | `toggleBorderTop` | Set top borders of the selected cells (solid, hairline) | |
| Border | `bol` | `toggleBorderRight` | Set right borders of the selected cells (solid, hairline) | |
| Border | `boia` | `toggleBorderInner` | Set borders inside the selected cells (solid, hairline) | |
| Border | `bois` | `toggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, hairline) | |
| Border | `boiv` | `toggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, hairline) | |
| Border | `bo/` | `toggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, hairline) | |
| Border | `bo\` | `toggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, hairline) | |
| Border | `bmb` | `toggleBorderAll` | Set borders around and inside the selected cells (solid, medium line) | |
| Border | `bma` | `toggleBorderAround` | Set borders around the selected cells (solid, medium line) | |
| Border | `bmh` | `toggleBorderLeft` | Set left borders of the selected cells (solid, medium line) | |
| Border | `bmj` | `toggleBorderBottom` | Set bottom borders of the selected cells (solid, medium line) | |
| Border | `bmk` | `toggleBorderTop` | Set top borders of the selected cells (solid, medium line) | |
| Border | `bml` | `toggleBorderRight` | Set right borders of the selected cells (solid, medium line) | |
| Border | `bmia` | `toggleBorderInner` | Set borders inside the selected cells (solid, medium line) | |
| Border | `bmis` | `toggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, medium line) | |
| Border | `bmiv` | `toggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, medium line) | |
| Border | `bm/` | `toggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, medium line) | |
| Border | `bm\` | `toggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, medium line) | |
| Border | `btb` | `toggleBorderAll` | Set borders around and inside the selected cells (double line, thick line) | |
| Border | `bta` | `toggleBorderAround` | Set borders around the selected cells (double line, thick line) | |
| Border | `bth` | `toggleBorderLeft` | Set left borders of the selected cells (double line, thick line) | |
| Border | `btj` | `toggleBorderBottom` | Set bottom borders of the selected cells (double line, thick line) | |
| Border | `btk` | `toggleBorderTop` | Set top borders of the selected cells (double line, thick line) | |
| Border | `btl` | `toggleBorderRight` | Set right borders of the selected cells (double line, thick line) | |
| Border | `btia` | `toggleBorderInner` | Set borders inside the selected cells (double line, thick line) | |
| Border | `btis` | `toggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (double line, thick line) | |
| Border | `btiv` | `toggleBorderInnerVertical` | Set vertical borders inside the selected cells (double line, thick line) | |
| Border | `bt/` | `toggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (double line, thick line) | |
| Border | `bt\` | `toggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (double line, thick line) | |
| Border | `bdd` | `deleteBorderAll` | Delete all borders around and inside the selected cells | |
| Border | `bda` | `deleteBorderAround` | Delete borders around the selected cells | |
| Border | `bdh` | `deleteBorderLeft` | Delete left borders of the selected cells | |
| Border | `bdj` | `deleteBorderBottom` | Delete bottom borders of the selected cells | |
| Border | `bdk` | `deleteBorderTop` | Delete top borders of the selected cells | |
| Border | `bdl` | `deleteBorderRight` | Delete right borders of the selected cells | |
| Border | `bdia` | `deleteBorderInner` | Delete all inner borders of the selected cells | |
| Border | `bdis` | `deleteBorderInnerHorizontal` | Delete horizontal inner borders of the selected cells | |
| Border | `bdiv` | `deleteBorderInnerVertical` | Delete vertical inner borders of the selected cells | |
| Border | `bd/` | `deleteBorderDiagonalUp` | Delete diagonal up borders in the selected cells | |
| Border | `bd\` | `deleteBorderDiagonalDown` | Delete diagonal down borders in the selected cells | |
| Border | `bcc` | `setBorderColorAll` | Set the color of all borders around and inside the selected cells | |
| Border | `bca` | `setBorderColorAround` | Set the color of borders around the selected cells | |
| Border | `bch` | `setBorderColorLeft` | Set the color of left borders of the selected cells | |
| Border | `bcj` | `setBorderColorBottom` | Set the color of bottom borders of the selected cells | |
| Border | `bck` | `setBorderColorTop` | Set the color of top borders of the selected cells | |
| Border | `bcl` | `setBorderColorRight` | Set the color of right borders of the selected cells | |
| Border | `bcia` | `setBorderColorInner` | Set the color of all inner borders of the selected cells | |
| Border | `bcis` | `setBorderColorInnerHorizontal` | Set the color of horizontal inner borders of the selected cells | |
| Border | `bciv` | `setBorderColorInnerVertical` | Set the color of vertical inner borders of the selected cells | |
| Border | `bc/` | `setBorderColorDiagonalUp` | Set the color of diagonal up borders in the selected cells | |
| Border | `bc\` | `setBorderColorDiagonalDown` | Set the color of diagonal down borders in the selected cells | |
| Row | `r-` | `narrowRowsHeight` | Narrow the height of the row | ✓ |
| Row | `r+` | `wideRowsHeight` | Widen the height of the row | ✓ |
| Row | `rr` | `selectRows` | Select rows | ✓ |
| Row | `ra` | `appendRows` | Insert rows below the current row | ✓ |
| Row | `ri` | `insertRows` | Insert rows above the current row | ✓ |
| Row | `rd` | `deleteRows` | Delete the current row | ✓ |
| Row | `ry` | `yankRows` | Copy the current row | ✓ |
| Row | `rx` | `cutRows` | Cut the current row | ✓ |
| Row | `rh` | `hideRows` | Hide the current row | ✓ |
| Row | `rH` | `unhideRows` | Unhide the current row | ✓ |
| Row | `rg` | `groupRows` | Group the current row | ✓ |
| Row | `ru` | `ungroupRows` | Ungroup the current row | ✓ |
| Row | `rf` | `foldRowsGroup` | Fold the current row group | ✓ |
| Row | `rs` | `spreadRowsGroup` | Expand the folding of the current row group | ✓ |
| Row | `rj` | `adjustRowsHeight` | Automatically adjust the height of the current row | ✓ |
| Row | `rw` | `setRowsHeight` | Set the height of the current row to a custom value | ✓ |
| Column | `c-` | `narrowColumnsWidth` | Narrow the width of the column | ✓ |
| Column | `c+` | `wideColumnsWidth` | Widen the width of the column | ✓ |
| Column | `cc` | `selectColumns` | Select columns | ✓ |
| Column | `ca` | `appendColumns` | Insert columns to the right of the current column | ✓ |
| Column | `ci` | `insertColumns` | Insert columns to the left of the current column | ✓ |
| Column | `cd` | `deleteColumns` | Delete the current column | ✓ |
| Column | `cy` | `yankColumns` | Copy the current column | ✓ |
| Column | `cx` | `cutColumns` | Cut the current column | ✓ |
| Column | `ch` | `hideColumns` | Hide the current column | ✓ |
| Column | `cH` | `unhideColumns` | Unhide the current column | ✓ |
| Column | `cg` | `groupColumns` | Group the current column | ✓ |
| Column | `cu` | `ungroupColumns` | Ungroup the current column | ✓ |
| Column | `cf` | `foldColumnsGroup` | Fold the current column group | ✓ |
| Column | `cs` | `spreadColumnsGroup` | Expand the folding of the current column group | ✓ |
| Column | `cj` | `adjustColumnsWidth` | Automatically adjust the width of the current column | ✓ |
| Column | `cw` | `setColumnsWidth` | Set the width of the current column to a custom value | ✓ |
| Yank | `yr` | `yankRows` | Copy the current row | ✓ |
| Yank | `yc` | `yankColumns` | Copy the current column | ✓ |
| Yank | `ygg` | `yankToTopRows` | Copy from the current row to the first row | |
| Yank | `yG` | `yankToBottomRows` | Copy from the current row to the last row in UsedRange | |
| Yank | `y{` | `yankToTopOfCurrentRegionRows` | Copy from the current row to the first row of the CurrentRegion | |
| Yank | `y}` | `yankToBottomOfCurrentRegionRows` | Copy from the current row to the last row of the CurrentRegion | |
| Yank | `y0` | `yankToLeftEndColumns` | Copy from the current column to the first column in UsedRange | |
| Yank | `y$` | `yankToRightEndColumns` | Copy from the current column to the last column in UsedRange | |
| Yank | `y^` | `yankToLeftOfCurrentRegionColumns` | Copy from the current column to the first column of the CurrentRegion | |
| Yank | `yg$` | `yankToRightOfCurrentRegionColumns` | Copy from the current column to the last column of the CurrentRegion | |
| Yank | `yh` | `yankFromLeftCell` | Copy and paste the value from the cell to the left of the current cell | |
| Yank | `yj` | `yankFromDownCell` | Copy and paste the value from the cell below the current cell | |
| Yank | `yk` | `yankFromUpCell` | Copy and paste the value from the cell above the current cell | |
| Yank | `yl` | `yankFromRightCell` | Copy and paste the value from the cell to the right of the current cell | |
| Yank | `Y` | `yankAsPlaintext` | Copy the selected cells as plaintext | |
| Delete | `D`/`X` | `deleteValue` | Delete the value in the cell | |
| Delete | `dx` | `deleteRows` | Delete the current row | ✓ |
| Delete | `dr` | `deleteRows` | Delete the current row | ✓ |
| Delete | `dc` | `deleteColumns` | Delete the current column | ✓ |
| Delete | `dgg` | `deleteToTopRows` | Delete from the current row to the top row | |
| Delete | `dG` | `deleteToBottomRows` | Delete from the current row to the last row in UsedRange | |
| Delete | `d{` | `deleteToTopOfCurrentRegionRows` | Delete from the current row to the first row of the CurrentRegion | |
| Delete | `d}` | `deleteToBottomOfCurrentRegionRows` | Delete from the current row to the last row of the CurrentRegion | |
| Delete | `d0` | `deleteToLeftEndColumns` | Delete from the current column to the first column in UsedRange | |
| Delete | `d$` | `deleteToRightEndColumns` | Delete from the current column to the last column in UsedRange | |
| Delete | `d^` | `deleteToLeftOfCurrentRegionColumns` | Delete from the current column to the first column of the CurrentRegion | |
| Delete | `dg$` | `deleteToRightOfCurrentRegionColumns` | Delete from the current column to the last column of the CurrentRegion | |
| Delete | `dh` | `deleteToLeft` | Delete the current cell and shift left | ✓ |
| Delete | `dj` | `deleteToUp` | Delete the current cell and shift up | ✓ |
| Delete | `dk` | `deleteToUp` | Delete the current cell and shift up | ✓ |
| Delete | `dl` | `deleteToLeft` | Delete the current cell and shift left | ✓ |
| Cut | `xr` | `cutRows` | Cut the current row | ✓ |
| Cut | `xc` | `cutColumns` | Cut the current column | ✓ |
| Cut | `xgg` | `cutToTopRows` | Cut from the current row to the first row | |
| Cut | `xG` | `cutToBottomRows` | Cut from the current row to the last row in UsedRange | |
| Cut | `x{` | `cutToTopOfCurrentRegionRows` | Cut from the current row to the first row of the CurrentRegion | |
| Cut | `x}` | `cutToBottomOfCurrentRegionRows` | Cut from the current row to the last row of the CurrentRegion | |
| Cut | `x0` | `cutToLeftEndColumns` | Cut from the current column to the first column in UsedRange | |
| Cut | `x$` | `cutToRightEndColumns` | Cut from the current column to the last column in UsedRange | |
| Cut | `x^` | `cutToLeftOfCurrentRegionColumns` | Cut from the current column to the first column of the CurrentRegion | |
| Cut | `xg$` | `cutToRightOfCurrentRegionColumns` | Cut from the current column to the last column of the CurrentRegion | |
| Paste | `p` | `pasteSmart` | Paste after copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Paste | `P` | `pasteSmart` | Paste before copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Paste | `gp` | `pasteSpecial` | Show the paste special format dialog | |
| Paste | `U` | `pasteValue` | Paste values only | |
| Font | `-` | `decreaseFontSize` | Decrease font size | |
| Font | `+` | `increaseFontSize` | Increase font size | |
| Font | `fn` | `changeFontName` | Focus on font name | |
| Font | `fs` | `changeFontSize` | Focus on font size | |
| Font | `fh` | `alignLeft` | Align left | |
| Font | `fj` | `alignBottom` | Align bottom | |
| Font | `fk` | `alignTop` | Align top | |
| Font | `fl` | `alignRight` | Align right | |
| Font | `fo` | `alignCenter` | Align center horizontally | |
| Font | `fm` | `alignMiddle` | Align center vertically | |
| Font | `fb` | `toggleBold` | Toggle bold | |
| Font | `fi` | `toggleItalic` | Toggle italic | |
| Font | `fu` | `toggleUnderline` | Toggle underline | |
| Font | `f-` | `toggleStrikethrough` | Toggle strikethrough | |
| Font | `ft` | `changeFormat` | Focus on cell format | |
| Font | `ff` | `showFontDialog` | Show the cell format dialog | |
| Color | `fc` | `smartFontColor` | Show the font color selection dialog | |
| Color | `FC`/`Fc` | `smartFillColor` | Show the fill color selection dialog | |
| Color | `bc` | `changeShapeBorderColor` | (when a shape is selected) Show the border color selection dialog | |
| Comment | `Ci`/`Cc` | `editCellComment` | Edit the cell comment (add if none exists) | |
| Comment | `Ce`/`Cx`/`Cd` | `deleteCellComment` | Delete the current cell's comment | |
| Comment | `CE`/`CD` | `deleteCellCommentAll` | Delete all comments on the sheet | |
| Comment | `Ca` | `toggleCellComment` | Toggle the display of the current cell's comment | |
| Comment | `Cr` | `showCellComment` | Show the current cell's comment | |
| Comment | `Cm` | `hideCellComment` | Hide the current cell's comment | |
| Comment | `CA` | `toggleCellCommentAll` | Toggle the display of all comments | |
| Comment | `CR` | `showCellCommentAll` | Show all comments | |
| Comment | `CM` | `hideCellCommentAll` | Hide all comments | |
| Comment | `CH` | `hideCellCommentIndicator` | Hide the current cell's comment indicator | |
| Comment | `Cn` | `nextCommentedCell` | Select the next commented cell | |
| Comment | `Cp` | `prevCommentedCell` | Select the previous commented cell | |
| Find & Replace | `/` | `showFindFollowLang` | Show the find dialog, following the language mode of IME | |
| Find & Replace | `?` | `showFindNotFollowLang` | Show the find dialog without following the language mode of IME | |
| Find & Replace | `n` | `nextFoundCell` | Select the next found cell | ✓ |
| Find & Replace | `N` | `previousFoundCell` | Select the previous found cell | ✓ |
| Find & Replace | `R` | `showReplaceWindow` | Show the find and replace dialog | |
| Find & Replace | `*` | `findActiveValueNext` | Find the next cell with the active cell's value and select it | |
| Find & Replace | `#` | `findActiveValuePrev` | Find the previous cell with the active cell's value and select it | |
| Find & Replace | `]c` | `nextSpecialCells` | Select the next cell with a comment | ✓ |
| Find & Replace | `[c` | `prevSpecialCells` | Select the previous cell with a comment | ✓ |
| Find & Replace | `]o` | `nextSpecialCells` | Select the next cell with a constant value | ✓ |
| Find & Replace | `[o` | `prevSpecialCells` | Select the previous cell with a constant value | ✓ |
| Find & Replace | `]f` | `nextSpecialCells` | Select the next cell with a formula | ✓ |
| Find & Replace | `[f` | `prevSpecialCells` | Select the previous cell with a formula | ✓ |
| Find & Replace | `]k` | `nextSpecialCells` | Select the next empty cell | ✓ |
| Find & Replace | `[k` | `prevSpecialCells` | Select the previous empty cell | ✓ |
| Find & Replace | `]t` | `nextSpecialCells` | Select the next cell with conditional formatting | ✓ |
| Find & Replace | `[t` | `prevSpecialCells` | Select the previous cell with conditional formatting | ✓ |
| Find & Replace | `]v` | `nextSpecialCells` | Select the next cell with data validation | ✓ |
| Find & Replace | `[v` | `prevSpecialCells` | Select the previous cell with data validation | ✓ |
| Find & Replace | `]s` | `nextShape` | Select the next shape | ✓ |
| Find & Replace | `[s` | `prevShape` | Select the previous shape | ✓ |
| Scrolling | `<C-u>` | `scrollUpHalf` | Scroll up by half a page | |
| Scrolling | `<C-d>` | `scrollDownHalf` | Scroll down by half a page | |
| Scrolling | `<C-b>` | `scrollUp` | Scroll up by one page | |
| Scrolling | `<C-f>` | `scrollDown` | Scroll down by one page | |
| Scrolling | `<C-y>` | `scrollUp1Row` | Scroll up by one row | |
| Scrolling | `<C-e>` | `scrollDown1Row` | Scroll down by one row | |
| Scrolling | `zh` | `scrollLeft1Column` | Scroll left by one column | ✓ |
| Scrolling | `zl` | `scrollRight1Column` | Scroll right by one column | ✓ |
| Scrolling | `zH` | `scrollLeft` | Scroll left by one page | ✓ |
| Scrolling | `zL` | `scrollRight` | Scroll right by one page | ✓ |
| Scrolling | `zt` | `scrollCurrentTop` | Scroll to make the current row at the top (`SCREEN_OFFSET` pts of padding) | |
| Scrolling | `zz` | `scrollCurrentMiddle` | Scroll to make the current row in the middle | |
| Scrolling | `zb` | `scrollCurrentBottom` | Scroll to make the current row at the bottom (`SCREEN_OFFSET` pts of padding) | |
| Scrolling | `zs` | `scrollCurrentLeft` | Scroll to make the current column at the left | |
| Scrolling | `zm` | `scrollCurrentCenter` | Scroll to make the current column in the center | |
| Scrolling | `ze` | `scrollCurrentRight` | Scroll to make the current column at the right | |
| Worksheet | `e`/`wn` | `nextWorksheet` | Select the next worksheet | |
| Worksheet | `E`/`wp` | `previousWorksheet` | Select the previous worksheet | |
| Worksheet | `ww`/`ws` | `showSheetPicker` | Launch the Sheet Picker | |
| Worksheet | `wr` | `renameWorksheet` | Change the name of the active worksheet | |
| Worksheet | `wh` | `moveWorksheetBack` | Move the active worksheet one position to the front | ✓ |
| Worksheet | `wl` | `moveWorksheetForward` | Move the active worksheet one position to the back | ✓ |
| Worksheet | `wi` | `insertWorksheet` | Insert a new worksheet in front of the active worksheet | |
| Worksheet | `wa` | `appendWorksheet` | Insert a new worksheet after the active worksheet | |
| Worksheet | `wd` | `deleteWorksheet` | Delete the active worksheet | |
| Worksheet | `w0` | `activateLastWorksheet` | Select the last worksheet | |
| Worksheet | `w$` | `activateLastWorksheet` | Select the last worksheet | |
| Worksheet | `wc` | `changeWorksheetTabColor` | Change the color of the active worksheet tab | |
| Worksheet | `wy` | `cloneWorksheet` | Clone the active worksheet | |
| Worksheet | `we` | `exportWorksheet` | Show the move or copy sheet dialog | |
| Worksheet | `w[num]` | `activateWorksheet` | Select the worksheet at position `[num]` (only 1-9) | |
| Worksheet | `:p` | `printPreviewOfActiveSheet` | Show the print preview of the active sheet | |
| Workbook | `:e` | `openWorkbook` | Open a workbook | |
| Workbook | `:e!` | `reopenActiveWorkbook` | Discard changes to the active workbook and reopen it | |
| Workbook | `:w` | `saveWorkbook` | Save the active workbook | |
| Workbook | `:q` | `closeAskSaving` | Close the active workbook (show a dialog if there are unsaved changes) | |
| Workbook | `:q!`/`ZQ` | `closeWithoutSaving` | Close the active workbook without saving | |
| Workbook | `:wq`/`:x`/`ZZ` | `closeWithSaving` | Save and close the active workbook | |
| Workbook | `:b[num]` | `activateWorkbook` | Select the workbook at position `[num]` | |
| Workbook | `]b`/`:bn` | `nextWorkbook` | Select the next workbook | |
| Workbook | `[b`/`:bp` | `previousWorkbook` | Select the previous workbook | |
| Workbook | `~` | `toggleReadOnly` | Toggle read-only mode | |
| Other | `u` | `undo_CtrlZ` | Undo (send `Ctrl + Z`) | |
| Other | `<C-r>` | `redoExecute` | Redo | |
| Other | `.` | `repeatAction` | Repeat the previous action (limited to commands where `repeatRegister` is called) | |
| Other | `m` | `zoomIn` | Zoom in by 10% or `[count]`% | ✓ |
| Other | `M` | `zoomOut` | Zoom out by 10% or `[count]`% | ✓ |
| Other | `%` | `zoomSpecifiedScale` | Set the zoom level to `[count]`% (digits `1`-`9` correspond to predefined zoom levels) | ✓ |
| Other | `\` | `showContextMenu` | Show the context menu | |
| Other | `<C-i>` | `jumpNext` | Move to the next cell in the jump list | |
| Other | `<C-o>` | `jumpPrev` | Move to the previous cell in the jump list | |
| Other | `:cle` | `clearJumps` | Clear the jump list | |
| Other | `zf` | `toggleFreezePanes` | Toggle freeze panes on/off | |
| Other | `=v` | `toggleFormulaBar` | Toggle the visibility of the formula bar | |
| Other | `gs` | `showSummaryInfo` | Show the file properties | |
| Other | `zp` | `setPrintArea` | Set the selected cells as the print area | |
| Other | `zP` | `clearPrintArea` | Clear the print area | |
| Other | `@@` | `showMacroDialog` | Show the macro dialog | |
| Other | `1-9` | `showCmdForm` | Specify `[count]` (only works with features marked with ✓ in Count) | |

</div></details>

\* The keymaps are defined with `map` method in [UserConfig.bas](./src/UserConfig.bas).

### Custom Key Mapping

- `<C-[>` → `<Esc>`

## Customization

Under construction...

## Contributing

[Issues](https://github.com/sha5010/vim.xlam/issues) and [Pull Requests](https://github.com/sha5010/vim.xlam/pulls) are welcome. If you've developed your own features and would like to contribute, I'd appreciate your help.

English version of the README was generated by ChatGPT. If you come across any errors or have suggestions for improvements, please don't hesitate to let me know. Your feedback is highly appreciated.

## Author

[@sha_5010](https://twitter.com/sha_5010)

## License

[MIT](./LICENSE)
