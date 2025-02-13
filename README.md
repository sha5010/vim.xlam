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

<p align="center">
  [English]
  [<a href="README_ja.md">日本語</a>]
</p>


# vim.xlam

Vim experience in Excel. This is an Excel add-in that allows you to use Vim keybindings within Excel.

## 📝 Description

vim.xlam is an Excel add-in designed to provide a Vim-like experience, allowing you to navigate and operate within Excel using keyboard shortcuts.

Designed with extensibility in mind, you can create your own methods and easily customize keybindings using the `Map` method. It's also designed to allow you to change keybindings easily from the default ones, so you can configure it to suit your preferences.

***Demo:***

![demo](https://user-images.githubusercontent.com/95682647/175773473-50376812-afcc-4ced-b436-7150d7b97872.gif)

\* Sample file courtesy of [https://atelierkobato.com](https://atelierkobato.com/download/).

## ✨ Features

- 🚀 Over 350 commands, launching in less than 0.1 seconds, significantly improving Excel productivity.
- ⚡ Supports not only cell movement with `hjkl`, but also jump commands like `gg`, `G`, `^`, `$`, and more.
- 🎯 Vim features such as `count` specification, `.` repeat, jump lists, and visual mode, and more are recreated as faithfully as possible.
- 💡 Command suggestion feature, making it easy to use by simply remembering a minimal set of prefixes.
- 🎨 Custom ColorPicker, allowing you to select colors within 3 keystrokes.
- 🗂️ SheetPicker for easy navigation and management, even with a large number of sheets.
- 🛠️ Simple [customization](#%EF%B8%8F-customization) by writing a vimrc-like configuration file.

## 📦 Installation

1. Download the latest vim.xlam from the [Release Page](https://github.com/sha5010/vim.xlam/releases/latest) (or [directly download the latest version](https://github.com/sha5010/vim.xlam/releases/latest/download/vim.xlam)).
2. Save the downloaded vim.xlam in `C:\Users\<USERNAME>\AppData\Roaming\Microsoft\AddIns`.
3. Launch Excel, go to File &gt; Options &gt; Add-Ins, and click the **Go...** button at the bottom of the screen.
4. Click the **Browse...** button, select the saved vim.xlam, and add the add-in.

| :exclamation: **Attention** |
| ---- |
| **This project is currently under development**. There may be frequent breaking changes during releases. Please check the release notes before updating. |


### (Optional) Recommended Initial Setup

By default, the `/` key cannot be recognized by vim.xlam. You can enable it by following these steps:

1. Click **File** in the Excel menu, then select **Options**.
2. In the left panel, choose **Advanced** and scroll down to the **Lotus Compatibility** section near the bottom.
3. Clear the **Microsoft Excel menu key:** field and click **OK**.

## 📘 Usage

- As it automatically launches when added to the add-ins, enjoy the ultimate Vim experience in Excel!
- Operations such as cell navigation/editing, adding/deleting rows and columns, setting colors and borders, etc., can be performed using assigned keys.
  - For a list of all implemented commands, refer to the table below.
- Customization of settings and keymaps is also possible by placing a [configuration file](./config/_vimxlamrc). See [Customization](#%EF%B8%8F-customization) for details.

### ⌨️ Default Keybindings

**Primary Commands**

| Type | Keystroke | Action | Description | Count |
| ---- | --------- | ------ | ----------- | ----- |
| Core | `<C-m>` | `ToggleVim` | Toggle Vim mode | |
| InsertMode | `a` | `AppendFollowLangMode` | Edit cell from the end, following language mode for IME | |
| InsertMode | `i` | `InsertFollowLangMode` | Edit cell from the start, following language mode for IME | |
| InsertMode | `s` | `SubstituteFollowLangMode` | Clear and edit cell, following language mode for IME | |
| Moving | `h` | `MoveLeft` | ← | ✓ |
| Moving | `j` | `MoveDown` | ↓ | ✓ |
| Moving | `k` | `MoveUp` | ↑ | ✓ |
| Moving | `l` | `MoveRight` | → | ✓ |
| Moving | `gg` | `MoveToTopRow` | Move to the 1st row or the `[count]`-th row | ✓ |
| Cell | `FF`/`Ff` | `ApplyFlashFill` | Flash Fill (fallback to Auto Fill if not applicable) | |
| Mode | `v` | `ToggleVisualMode` | Toggle visual mode (extend selection) | |
| Border | `bb` | `ToggleBorderAll` | Set borders around and inside the selected cells (solid, thin line) | |
| Border | `ba` | `ToggleBorderAround` | Set borders around the selected cells (solid, thin line) | |
| Border | `bia` | `ToggleBorderInner` | Set borders inside the selected cells (solid, thin line) | |
| Row | `ra` | `AppendRows` | Insert rows below the current row | ✓ |
| Row | `ri` | `InsertRows` | Insert rows above the current row | ✓ |
| Row | `rd` | `DeleteRows` | Delete the current row | ✓ |
| Column | `ca` | `AppendColumns` | Insert columns to the right of the current column | ✓ |
| Column | `ci` | `InsertColumns` | Insert columns to the left of the current column | ✓ |
| Column | `cd` | `DeleteColumns` | Delete the current column | ✓ |
| Delete | `D`/`X` | `DeleteValue` | Delete the value in the cell | |
| Paste | `p` | `PasteSmart` | Paste after copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Paste | `P` | `PasteSmart` | Paste before copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Font | `-` | `DecreaseFontSize` | Decrease font size | ✓ |
| Font | `+` | `IncreaseFontSize` | Increase font size | ✓ |
| Color | `fc` | `SmartFontColor` | Show the font color selection dialog | |
| Find & Replace | `/` | `ShowFindFollowLang` | Show the find dialog, following the language mode of IME | |
| Find & Replace | `n` | `NextFoundCell` | Select the next found cell | ✓ |
| Find & Replace | `N` | `PreviousFoundCell` | Select the previous found cell | ✓ |
| Scrolling | `<C-u>` | `ScrollUpHalf` | Scroll up by half a page | ✓ |
| Scrolling | `<C-d>` | `ScrollDownHalf` | Scroll down by half a page | ✓ |
| Scrolling | `zt` | `ScrollCurrentTop` | Scroll to make the current row at the top (`SCREEN_OFFSET` pts of padding) | ✓ |
| Scrolling | `zz` | `ScrollCurrentMiddle` | Scroll to make the current row in the middle | ✓ |
| Scrolling | `zb` | `ScrollCurrentBottom` | Scroll to make the current row at the bottom (`SCREEN_OFFSET` pts of padding) | ✓ |
| Sheet | `e` | `NextSheet` | Select the next sheet | ✓ |
| Sheet | `E` | `PreviousSheet` | Select the previous sheet | ✓ |
| Sheet | `ww` | `ShowSheetPicker` | Launch the Sheet Picker | |
| Sheet | `wr` | `RenameSheet` | Change the name of the active sheet | |
| Workbook | `:w` | `SaveWorkbook` | Save the active workbook | |
| Workbook | `:q` | `CloseAskSaving` | Close the active workbook (show a dialog if there are unsaved changes) | |
| Workbook | `:q!`/`ZQ` | `CloseWithoutSaving` | Close the active workbook without saving | |
| Workbook | `:wq`/`:x`/`ZZ` | `CloseWithSaving` | Save and close the active workbook | |
| Other | `u` | `Undo_CtrlZ` | Undo (send `Ctrl + Z`) | |
| Other | `<C-r>` | `RedoExecute` | Redo | |

<details><summary>Expand all commands</summary><div>

| Type | Keystroke | Action | Description | Count |
| ---- | --------- | ------ | ----------- | ----- |
| Core | `<C-m>` | `ToggleVim` | Toggle Vim mode | |
| Core | `<C-p>` | `ToggleLang` | Toggle language mode (Japanese/English) | |
| Core | `:` | `EnterCmdlineMode` | Enter the Cmdline mode | |
| Core | `:reload` | `ReloadVim` | Reload vim.xlam | |
| Core | `:debug` | `ToggleDebugMode` | Toggle debug mode | |
| Core | `:version` | `ShowVersion` | Show vim.xlam version info | |
| InsertMode | `a` | `AppendFollowLangMode` | Edit cell from the end, following language mode for IME | |
| InsertMode | `A` | `AppendNotFollowLangMode` | Edit cell from the end without following language mode for IME | |
| InsertMode | `i` | `InsertFollowLangMode` | Edit cell from the start, following language mode for IME | |
| InsertMode | `I` | `InsertNotFollowLangMode` | Edit cell from the start without following language mode for IME | |
| InsertMode | `s` | `SubstituteFollowLangMode` | Clear and edit cell, following language mode for IME | |
| InsertMode | `S` | `SubstituteNotFollowLangMode` | Clear and edit cell without following language mode for IME | |
| Moving | `h` | `MoveLeft` | ← | ✓ |
| Moving | `j` | `MoveDown` | ↓ | ✓ |
| Moving | `k` | `MoveUp` | ↑ | ✓ |
| Moving | `l` | `MoveRight` | → | ✓ |
| Moving | `H` | `MoveLeftWithShift` | Shift + ← | ✓ |
| Moving | `J` | `MoveDownWithShift` | Shift + ↓ | ✓ |
| Moving | `K` | `MoveUpWithShift` | Shift + ↑ | ✓ |
| Moving | `L` | `MoveRightWithShift` | Shift + → | ✓ |
| Moving | `<C-h>` | `MoveLeft` | Ctrl + ← | |
| Moving | `<C-j>` | `MoveDown` | Ctrl + ↓ | |
| Moving | `<C-k>` | `MoveUp` | Ctrl + ↑ | |
| Moving | `<C-l>` | `MoveRight` | Ctrl + → | |
| Moving | `<C-S-H>` | `MoveLeft` | Ctrl + Shift + ← | |
| Moving | `<C-S-J>` | `MoveDown` | Ctrl + Shift + ↓ | |
| Moving | `<C-S-K>` | `MoveUp` | Ctrl + Shift + ↑ | |
| Moving | `<C-S-L>` | `MoveRight` | Ctrl + Shift + → | |
| Moving | `gg` | `MoveToTopRow` | Move to the 1st row or the `[count]`-th row | ✓ |
| Moving | `G` | `MoveToLastRow` | Move to the last row of UsedRange or the `[count]`-th row | ✓ |
| Moving | `\|` | `MoveToNthColumn` | Move to the `[count]`-th column | ✓ |
| Moving | `0` | `MoveToFirstColumn` | Move to the 1st column | |
| Moving | `^` | `MoveToLeftEnd` | Move to the first column of UsedRange | |
| Moving | `$` | `MoveToRightEnd` | Move to the last column of UsedRange | |
| Moving | `g0` | `MoveToA1` | Move to cell A1 | |
| Moving | `{` | `MoveToTopOfCurrentRegion` | Move to the first row within the CurrentRegion | |
| Moving | `}` | `MoveToBottomOfCurrentRegion` | Move to the last row within the CurrentRegion | |
| Moving | `W[cell]` | `MoveToSpecifiedCell` | Move to the specified `[cell]` | |
| Moving | `:[num]` | `MoveToSpecifiedRow` | Move to the specified `[num]`-th row | |
| Cell | `xx` | `CutCell` | Cut cell | |
| Cell | `yy` | `YankCell` | Copy cell | |
| Cell | `o` | `InsertCellsDown` | Insert cells below the selected cells | ✓ |
| Cell | `O` | `InsertCellsUp` | Insert cells above the selected cells | ✓ |
| Cell | `t` | `InsertCellsRight` | Insert cells to the right of the selected cells | ✓ |
| Cell | `T` | `InsertCellsLeft` | Insert cells to the left of the selected cells | ✓ |
| Cell | `>` | `IncrementText` | Increase indentation | ✓ |
| Cell | `<` | `DecrementText` | Decrease indentation | ✓ |
| Cell | `(` | `IncreaseDecimal` | Increase decimal places | ✓ |
| Cell | `)` | `DecreaseDecimal` | Decrease decimal places | ✓ |
| Cell | `<C-S-A>` | `AddNumber` | Add the number | ✓ |
| Cell | `<C-S-X>` | `SubtractNumber` | Subtract the number | ✓ |
| Cell | `g<C-A>` | `VisualAddNumber` | Add the number sequentially | ✓ |
| Cell | `g<C-X>` | `VisualSubtractNumber` | Subtract the number sequentially | ✓ |
| Cell | `zw` | `ToggleWrapText` | Toggle cell wrap text | |
| Cell | `&` | `ToggleMergeCells` | Toggle cell merge | |
| Cell | `f,` | `ApplyCommaStyle` | Apply comma style | |
| Cell | `<Space>` | `UnionSelectCells` | Add the current cell to the selection memory and select the remembered cells (allows selecting multiple cells) | |
| Cell | `<S-Space>` | `ExceptSelectCells` | Remove the current cell from the remembered selected cells | |
| Cell | `<S-BS>` | `ClearSelectCells` | Clear the remembered selected cells | |
| Cell | `gf` | `FollowHyperlinkOfActiveCell` | Open the hyperlink in the cell | |
| Cell | `FF`/`Ff` | `ApplyFlashFill` | Flash Fill (fallback to Auto Fill if not applicable) | |
| Cell | `FA`/`Fa` | `ApplyAutoFill` | Auto Fill | |
| Cell | `=s` | `AutoSum` | Auto SUM | |
| Cell | `=a` | `AutoAverage` | Auto SUM (average) | |
| Cell | `=c` | `AutoCount` | Auto SUM (count) | |
| Cell | `=m` | `AutoMax` | Auto SUM (maximum) | |
| Cell | `=i` | `AutoMin` | Auto SUM (minimum) | |
| Cell | `==` | `InsertFunction` | Insert function | |
| Mode | `v` | `ToggleVisualMode` | Toggle visual mode (extend selection) | |
| Mode | `V` | `ToggleVisualLine` | Toggle visual line mode (extend selection) | |
| Mode | `<C-.>` | `SwapVisualBase` | Swap the base cell for a range selection | |
| Border | `bb` | `ToggleBorderAll` | Set borders around and inside the selected cells (solid, thin line) | |
| Border | `ba` | `ToggleBorderAround` | Set borders around the selected cells (solid, thin line) | |
| Border | `bh` | `ToggleBorderLeft` | Set left borders of the selected cells (solid, thin line) | |
| Border | `bj` | `ToggleBorderBottom` | Set bottom borders of the selected cells (solid, thin line) | |
| Border | `bk` | `ToggleBorderTop` | Set top borders of the selected cells (solid, thin line) | |
| Border | `bl` | `ToggleBorderRight` | Set right borders of the selected cells (solid, thin line) | |
| Border | `bia` | `ToggleBorderInner` | Set borders inside the selected cells (solid, thin line) | |
| Border | `bis` | `ToggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, thin line) | |
| Border | `biv` | `ToggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, thin line) | |
| Border | `b/` | `ToggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, thin line) | |
| Border | `b\` | `ToggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, thin line) | |
| Border | `bB` | `ToggleBorderAll` | Set borders around and inside the selected cells (solid, thick line) | |
| Border | `bA` | `ToggleBorderAround` | Set borders around the selected cells (solid, thick line) | |
| Border | `bH` | `ToggleBorderLeft` | Set left borders of the selected cells (solid, thick line) | |
| Border | `bJ` | `ToggleBorderBottom` | Set bottom borders of the selected cells (solid, thick line) | |
| Border | `bK` | `ToggleBorderTop` | Set top borders of the selected cells (solid, thick line) | |
| Border | `bL` | `ToggleBorderRight` | Set right borders of the selected cells (solid, thick line) | |
| Border | `Bb` | `ToggleBorderAll` | Set borders around and inside the selected cells (solid, thick line) | |
| Border | `Ba` | `ToggleBorderAround` | Set borders around the selected cells (solid, thick line) | |
| Border | `Bh` | `ToggleBorderLeft` | Set left borders of the selected cells (solid, thick line) | |
| Border | `Bj` | `ToggleBorderBottom` | Set bottom borders of the selected cells (solid, thick line) | |
| Border | `Bk` | `ToggleBorderTop` | Set top borders of the selected cells (solid, thick line) | |
| Border | `Bl` | `ToggleBorderRight` | Set right borders of the selected cells (solid, thick line) | |
| Border | `Bia` | `ToggleBorderInner` | Set borders inside the selected cells (solid, thick line) | |
| Border | `Bis` | `ToggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, thick line) | |
| Border | `Biv` | `ToggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, thick line) | |
| Border | `B/` | `ToggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, thick line) | |
| Border | `B\` | `ToggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, thick line) | |
| Border | `bob` | `ToggleBorderAll` | Set borders around and inside the selected cells (solid, hairline) | |
| Border | `boa` | `ToggleBorderAround` | Set borders around the selected cells (solid, hairline) | |
| Border | `boh` | `ToggleBorderLeft` | Set left borders of the selected cells (solid, hairline) | |
| Border | `boj` | `ToggleBorderBottom` | Set bottom borders of the selected cells (solid, hairline) | |
| Border | `bok` | `ToggleBorderTop` | Set top borders of the selected cells (solid, hairline) | |
| Border | `bol` | `ToggleBorderRight` | Set right borders of the selected cells (solid, hairline) | |
| Border | `boia` | `ToggleBorderInner` | Set borders inside the selected cells (solid, hairline) | |
| Border | `bois` | `ToggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, hairline) | |
| Border | `boiv` | `ToggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, hairline) | |
| Border | `bo/` | `ToggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, hairline) | |
| Border | `bo\` | `ToggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, hairline) | |
| Border | `bmb` | `ToggleBorderAll` | Set borders around and inside the selected cells (solid, medium line) | |
| Border | `bma` | `ToggleBorderAround` | Set borders around the selected cells (solid, medium line) | |
| Border | `bmh` | `ToggleBorderLeft` | Set left borders of the selected cells (solid, medium line) | |
| Border | `bmj` | `ToggleBorderBottom` | Set bottom borders of the selected cells (solid, medium line) | |
| Border | `bmk` | `ToggleBorderTop` | Set top borders of the selected cells (solid, medium line) | |
| Border | `bml` | `ToggleBorderRight` | Set right borders of the selected cells (solid, medium line) | |
| Border | `bmia` | `ToggleBorderInner` | Set borders inside the selected cells (solid, medium line) | |
| Border | `bmis` | `ToggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (solid, medium line) | |
| Border | `bmiv` | `ToggleBorderInnerVertical` | Set vertical borders inside the selected cells (solid, medium line) | |
| Border | `bm/` | `ToggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (solid, medium line) | |
| Border | `bm\` | `ToggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (solid, medium line) | |
| Border | `btb` | `ToggleBorderAll` | Set borders around and inside the selected cells (double line, thick line) | |
| Border | `bta` | `ToggleBorderAround` | Set borders around the selected cells (double line, thick line) | |
| Border | `bth` | `ToggleBorderLeft` | Set left borders of the selected cells (double line, thick line) | |
| Border | `btj` | `ToggleBorderBottom` | Set bottom borders of the selected cells (double line, thick line) | |
| Border | `btk` | `ToggleBorderTop` | Set top borders of the selected cells (double line, thick line) | |
| Border | `btl` | `ToggleBorderRight` | Set right borders of the selected cells (double line, thick line) | |
| Border | `btia` | `ToggleBorderInner` | Set borders inside the selected cells (double line, thick line) | |
| Border | `btis` | `ToggleBorderInnerHorizontal` | Set horizontal borders inside the selected cells (double line, thick line) | |
| Border | `btiv` | `ToggleBorderInnerVertical` | Set vertical borders inside the selected cells (double line, thick line) | |
| Border | `bt/` | `ToggleBorderDiagonalUp` | Set diagonal up borders in the selected cells (double line, thick line) | |
| Border | `bt\` | `ToggleBorderDiagonalDown` | Set diagonal down borders in the selected cells (double line, thick line) | |
| Border | `bdd` | `DeleteBorderAll` | Delete all borders around and inside the selected cells | |
| Border | `bda` | `DeleteBorderAround` | Delete borders around the selected cells | |
| Border | `bdh` | `DeleteBorderLeft` | Delete left borders of the selected cells | |
| Border | `bdj` | `DeleteBorderBottom` | Delete bottom borders of the selected cells | |
| Border | `bdk` | `DeleteBorderTop` | Delete top borders of the selected cells | |
| Border | `bdl` | `DeleteBorderRight` | Delete right borders of the selected cells | |
| Border | `bdia` | `DeleteBorderInner` | Delete all inner borders of the selected cells | |
| Border | `bdis` | `DeleteBorderInnerHorizontal` | Delete horizontal inner borders of the selected cells | |
| Border | `bdiv` | `DeleteBorderInnerVertical` | Delete vertical inner borders of the selected cells | |
| Border | `bd/` | `DeleteBorderDiagonalUp` | Delete diagonal up borders in the selected cells | |
| Border | `bd\` | `DeleteBorderDiagonalDown` | Delete diagonal down borders in the selected cells | |
| Border | `bcc` | `SetBorderColorAll` | Set the color of all borders around and inside the selected cells | |
| Border | `bca` | `SetBorderColorAround` | Set the color of borders around the selected cells | |
| Border | `bch` | `SetBorderColorLeft` | Set the color of left borders of the selected cells | |
| Border | `bcj` | `SetBorderColorBottom` | Set the color of bottom borders of the selected cells | |
| Border | `bck` | `SetBorderColorTop` | Set the color of top borders of the selected cells | |
| Border | `bcl` | `SetBorderColorRight` | Set the color of right borders of the selected cells | |
| Border | `bcia` | `SetBorderColorInner` | Set the color of all inner borders of the selected cells | |
| Border | `bcis` | `SetBorderColorInnerHorizontal` | Set the color of horizontal inner borders of the selected cells | |
| Border | `bciv` | `SetBorderColorInnerVertical` | Set the color of vertical inner borders of the selected cells | |
| Border | `bc/` | `SetBorderColorDiagonalUp` | Set the color of diagonal up borders in the selected cells | |
| Border | `bc\` | `SetBorderColorDiagonalDown` | Set the color of diagonal down borders in the selected cells | |
| Row | `r-` | `NarrowRowsHeight` | Narrow the height of the row | ✓ |
| Row | `r+` | `WideRowsHeight` | Widen the height of the row | ✓ |
| Row | `rr` | `SelectRows` | Select rows | ✓ |
| Row | `ra` | `AppendRows` | Insert rows below the current row | ✓ |
| Row | `ri` | `InsertRows` | Insert rows above the current row | ✓ |
| Row | `rd` | `DeleteRows` | Delete the current row | ✓ |
| Row | `ry` | `YankRows` | Copy the current row | ✓ |
| Row | `rx` | `CutRows` | Cut the current row | ✓ |
| Row | `rh` | `HideRows` | Hide the current row | ✓ |
| Row | `rH` | `UnhideRows` | Unhide the current row | ✓ |
| Row | `rg` | `GroupRows` | Group the current row | ✓ |
| Row | `ru` | `UngroupRows` | Ungroup the current row | ✓ |
| Row | `rf` | `FoldRowsGroup` | Fold the current row group | ✓ |
| Row | `rs` | `SpreadRowsGroup` | Expand the folding of the current row group | ✓ |
| Row | `rj` | `AdjustRowsHeight` | Automatically adjust the height of the current row | ✓ |
| Row | `rw` | `SetRowsHeight` | Set the height of the current row to a custom value | ✓ |
| Row | `rl` | `ApplyRowsLock` | Locks to allow only specified rows to be selected | ✓ |
| Row | `rL` | `ClearRowsLock` | Remove locks applied by `ApplyRowsLock` |   |
| Column | `c-` | `NarrowColumnsWidth` | Narrow the width of the column | ✓ |
| Column | `c+` | `WideColumnsWidth` | Widen the width of the column | ✓ |
| Column | `cc` | `SelectColumns` | Select columns | ✓ |
| Column | `ca` | `AppendColumns` | Insert columns to the right of the current column | ✓ |
| Column | `ci` | `InsertColumns` | Insert columns to the left of the current column | ✓ |
| Column | `cd` | `DeleteColumns` | Delete the current column | ✓ |
| Column | `cy` | `YankColumns` | Copy the current column | ✓ |
| Column | `cx` | `CutColumns` | Cut the current column | ✓ |
| Column | `ch` | `HideColumns` | Hide the current column | ✓ |
| Column | `cH` | `UnhideColumns` | Unhide the current column | ✓ |
| Column | `cg` | `GroupColumns` | Group the current column | ✓ |
| Column | `cu` | `UngroupColumns` | Ungroup the current column | ✓ |
| Column | `cf` | `FoldColumnsGroup` | Fold the current column group | ✓ |
| Column | `cs` | `SpreadColumnsGroup` | Expand the folding of the current column group | ✓ |
| Column | `cj` | `AdjustColumnsWidth` | Automatically adjust the width of the current column | ✓ |
| Column | `cw` | `SetColumnsWidth` | Set the width of the current column to a custom value | ✓ |
| Column | `cl` | `ApplyColumnsLock` | Locks to allow only specified columns to be selected | ✓ |
| Column | `cL` | `ClearColumnsLock` | Remove locks applied by `ApplyColumnsLock` |   |
| Yank | `yr` | `YankRows` | Copy the current row | ✓ |
| Yank | `yc` | `YankColumns` | Copy the current column | ✓ |
| Yank | `ygg` | `YankRows` | Copy from the current row to the first row | |
| Yank | `yG` | `YankRows` | Copy from the current row to the last row in UsedRange | |
| Yank | `y{` | `YankRows` | Copy from the current row to the first row of the CurrentRegion | |
| Yank | `y}` | `YankRows` | Copy from the current row to the last row of the CurrentRegion | |
| Yank | `y0` | `YankColumns` | Copy from the current column to the first column in UsedRange | |
| Yank | `y$` | `YankColumns` | Copy from the current column to the last column in UsedRange | |
| Yank | `y^` | `YankColumns` | Copy from the current column to the first column of the CurrentRegion | |
| Yank | `yg$` | `YankColumns` | Copy from the current column to the last column of the CurrentRegion | |
| Yank | `yh` | `YankFromLeftCell` | Copy and paste the value from the cell to the left of the current cell | |
| Yank | `yj` | `YankFromDownCell` | Copy and paste the value from the cell below the current cell | |
| Yank | `yk` | `YankFromUpCell` | Copy and paste the value from the cell above the current cell | |
| Yank | `yl` | `YankFromRightCell` | Copy and paste the value from the cell to the right of the current cell | |
| Yank | `Y` | `YankAsPlaintext` | Copy the selected cells as plaintext | |
| Delete | `D`/`X` | `DeleteValue` | Delete the value in the cell | |
| Delete | `dx` | `DeleteRows` | Delete the current row | ✓ |
| Delete | `dd`/`dr` | `DeleteRows` | Delete the current row | ✓ |
| Delete | `dc` | `DeleteColumns` | Delete the current column | ✓ |
| Delete | `dgg` | `DeleteRows` | Delete from the current row to the top row | |
| Delete | `dG` | `DeleteRows` | Delete from the current row to the last row in UsedRange | |
| Delete | `d{` | `DeleteRows` | Delete from the current row to the first row of the CurrentRegion | |
| Delete | `d}` | `DeleteRows` | Delete from the current row to the last row of the CurrentRegion | |
| Delete | `d0` | `DeleteColumns` | Delete from the current column to the first column in UsedRange | |
| Delete | `d$` | `DeleteColumns` | Delete from the current column to the last column in UsedRange | |
| Delete | `d^` | `DeleteColumns` | Delete from the current column to the first column of the CurrentRegion | |
| Delete | `dg$` | `DeleteColumns` | Delete from the current column to the last column of the CurrentRegion | |
| Delete | `dh` | `DeleteToLeft` | Delete the current cell and shift left | ✓ |
| Delete | `dj` | `DeleteToUp` | Delete the current cell and shift up | ✓ |
| Delete | `dk` | `DeleteToUp` | Delete the current cell and shift up | ✓ |
| Delete | `dl` | `DeleteToLeft` | Delete the current cell and shift left | ✓ |
| Cut | `xr` | `CutRows` | Cut the current row | ✓ |
| Cut | `xc` | `CutColumns` | Cut the current column | ✓ |
| Cut | `xgg` | `CutRows` | Cut from the current row to the first row | |
| Cut | `xG` | `CutRows` | Cut from the current row to the last row in UsedRange | |
| Cut | `x{` | `CutRows` | Cut from the current row to the first row of the CurrentRegion | |
| Cut | `x}` | `CutRows` | Cut from the current row to the last row of the CurrentRegion | |
| Cut | `x0` | `CutColumns` | Cut from the current column to the first column in UsedRange | |
| Cut | `x$` | `CutColumns` | Cut from the current column to the last column in UsedRange | |
| Cut | `x^` | `CutColumns` | Cut from the current column to the first column of the CurrentRegion | |
| Cut | `xg$` | `CutColumns` | Cut from the current column to the last column of the CurrentRegion | |
| Paste | `p` | `PasteSmart` | Paste after copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Paste | `P` | `PasteSmart` | Paste before copying rows or columns; otherwise, send `Ctrl + V` | ✓ |
| Paste | `gp` | `PasteSpecial` | Show the paste special format dialog | |
| Paste | `U` | `PasteValue` | Paste values only | |
| Font | `-` | `DecreaseFontSize` | Decrease font size | ✓ |
| Font | `+` | `IncreaseFontSize` | Increase font size | ✓ |
| Font | `fn` | `ChangeFontName` | Focus on font name | |
| Font | `fs` | `ChangeFontSize` | Focus on font size | |
| Font | `fh` | `AlignLeft` | Align left | |
| Font | `fj` | `AlignBottom` | Align bottom | |
| Font | `fk` | `AlignTop` | Align top | |
| Font | `fl` | `AlignRight` | Align right | |
| Font | `fo` | `AlignCenter` | Align center horizontally | |
| Font | `fm` | `AlignMiddle` | Align center vertically | |
| Font | `fb` | `ToggleBold` | Toggle bold | |
| Font | `fi` | `ToggleItalic` | Toggle italic | |
| Font | `fu` | `ToggleUnderline` | Toggle underline | |
| Font | `f-` | `ToggleStrikethrough` | Toggle strikethrough | |
| Font | `ft` | `ChangeFormat` | Focus on cell format | |
| Font | `ff` | `ShowFontDialog` | Show the cell format dialog | |
| Color | `fc` | `SmartFontColor` | Show the font color selection dialog | |
| Color | `FC`/`Fc` | `SmartFillColor` | Show the fill color selection dialog | |
| Color | `bc` | `ChangeShapeBorderColor` | (when a shape is selected) Show the border color selection dialog | |
| Comment | `Ci`/`Cc` | `EditCellComment` | Edit the cell comment (add if none exists) | |
| Comment | `Ce`/`Cx`/`Cd` | `DeleteCellComment` | Delete the current cell's comment | |
| Comment | `CE`/`CD` | `DeleteCellCommentAll` | Delete all comments on the sheet | |
| Comment | `Ca` | `ToggleCellComment` | Toggle the display of the current cell's comment | |
| Comment | `Cr` | `ShowCellComment` | Show the current cell's comment | |
| Comment | `Cm` | `HideCellComment` | Hide the current cell's comment | |
| Comment | `CA` | `ToggleCellCommentAll` | Toggle the display of all comments | |
| Comment | `CR` | `ShowCellCommentAll` | Show all comments | |
| Comment | `CM` | `HideCellCommentAll` | Hide all comments | |
| Comment | `CH` | `HideCellCommentIndicator` | Hide the current cell's comment indicator | |
| Comment | `Cn` | `NextComment` | Select the next commented cell | ✓ |
| Comment | `Cp` | `PrevComment` | Select the previous commented cell | ✓ |
| Find & Replace | `/` | `ShowFindFollowLang` | Show the find dialog, following the language mode of IME | |
| Find & Replace | `?` | `ShowFindNotFollowLang` | Show the find dialog without following the language mode of IME | |
| Find & Replace | `n` | `NextFoundCell` | Select the next found cell | ✓ |
| Find & Replace | `N` | `PreviousFoundCell` | Select the previous found cell | ✓ |
| Find & Replace | `R` | `ShowReplaceWindow` | Show the find and replace dialog | |
| Find & Replace | `*` | `FindActiveValueNext` | Find the next cell with the active cell's value and select it | ✓ |
| Find & Replace | `#` | `FindActiveValuePrev` | Find the previous cell with the active cell's value and select it | ✓ |
| Find & Replace | `]c` | `NextSpecialCells` | Select the next cell with a comment | ✓ |
| Find & Replace | `[c` | `PrevSpecialCells` | Select the previous cell with a comment | ✓ |
| Find & Replace | `]o` | `NextSpecialCells` | Select the next cell with a constant value | ✓ |
| Find & Replace | `[o` | `PrevSpecialCells` | Select the previous cell with a constant value | ✓ |
| Find & Replace | `]f` | `NextSpecialCells` | Select the next cell with a formula | ✓ |
| Find & Replace | `[f` | `PrevSpecialCells` | Select the previous cell with a formula | ✓ |
| Find & Replace | `]k` | `NextSpecialCells` | Select the next empty cell | ✓ |
| Find & Replace | `[k` | `PrevSpecialCells` | Select the previous empty cell | ✓ |
| Find & Replace | `]t` | `NextSpecialCells` | Select the next cell with conditional formatting | ✓ |
| Find & Replace | `[t` | `PrevSpecialCells` | Select the previous cell with conditional formatting | ✓ |
| Find & Replace | `]v` | `NextSpecialCells` | Select the next cell with data validation | ✓ |
| Find & Replace | `[v` | `PrevSpecialCells` | Select the previous cell with data validation | ✓ |
| Find & Replace | `]s` | `NextShape` | Select the next shape | ✓ |
| Find & Replace | `[s` | `PrevShape` | Select the previous shape | ✓ |
| Scrolling | `<C-u>` | `ScrollUpHalf` | Scroll up by half a page | ✓ |
| Scrolling | `<C-d>` | `ScrollDownHalf` | Scroll down by half a page | ✓ |
| Scrolling | `<C-b>` | `ScrollUp` | Scroll up by one page | ✓ |
| Scrolling | `<C-f>` | `ScrollDown` | Scroll down by one page | ✓ |
| Scrolling | `<C-y>` | `ScrollUp1Row` | Scroll up by one row | ✓ |
| Scrolling | `<C-e>` | `ScrollDown1Row` | Scroll down by one row | ✓ |
| Scrolling | `,` | `ScrollLeftHalf` | Scroll left by half a page | ✓ |
| Scrolling | `;` | `ScrollRightHalf` | Scroll right by half a page | ✓ |
| Scrolling | `zh` | `ScrollLeft1Column` | Scroll left by one column | ✓ |
| Scrolling | `zl` | `ScrollRight1Column` | Scroll right by one column | ✓ |
| Scrolling | `zH` | `ScrollLeft` | Scroll left by one page | ✓ |
| Scrolling | `zL` | `ScrollRight` | Scroll right by one page | ✓ |
| Scrolling | `zt` | `ScrollCurrentTop` | Scroll to make the current row at the top (`SCREEN_OFFSET` pts of padding) | ✓ |
| Scrolling | `zz` | `ScrollCurrentMiddle` | Scroll to make the current row in the middle | ✓ |
| Scrolling | `zb` | `ScrollCurrentBottom` | Scroll to make the current row at the bottom (`SCREEN_OFFSET` pts of padding) | ✓ |
| Scrolling | `zs` | `ScrollCurrentLeft` | Scroll to make the current column at the left | ✓ |
| Scrolling | `zm` | `ScrollCurrentCenter` | Scroll to make the current column in the center | ✓ |
| Scrolling | `ze` | `ScrollCurrentRight` | Scroll to make the current column at the right | ✓ |
| Sheet | `e`/`wn` | `NextSheet` | Select the next sheet | ✓ |
| Sheet | `E`/`wp` | `PreviousSheet` | Select the previous sheet | ✓ |
| Sheet | `ww`/`ws` | `ShowSheetPicker` | Launch the Sheet Picker | |
| Sheet | `wr` | `RenameSheet` | Change the name of the active sheet | |
| Sheet | `wh` | `MoveSheetBack` | Move the active sheet one position to the front | ✓ |
| Sheet | `wl` | `MoveSheetForward` | Move the active sheet one position to the back | ✓ |
| Sheet | `wi` | `InsertWorksheet` | Insert a new worksheet in front of the active worksheet | |
| Sheet | `wa` | `AppendWorksheet` | Insert a new worksheet after the active worksheet | |
| Sheet | `wd` | `DeleteSheet` | Delete the active sheet | |
| Sheet | `w0` | `ActivateLastSheet` | Select the last sheet | |
| Sheet | `w$` | `ActivateLastSheet` | Select the last sheet | |
| Sheet | `wc` | `ChangeSheetTabColor` | Change the color of the active sheet tab | |
| Sheet | `wy` | `CloneSheet` | Clone the active sheet | |
| Sheet | `we` | `ExportSheet` | Show the move or copy sheet dialog | |
| Sheet | `w[num]` | `ActivateSheet` | Select the sheet at position `[num]` (only 1-9) | |
| Sheet | `:preview` | `PrintPreviewOfActiveSheet` | Show the print preview of the active sheet | |
| Workbook | `:e [path]` | `OpenWorkbook` | Open a workbook | |
| Workbook | `:e!` | `ReopenActiveWorkbook` | Discard changes to the active workbook and reopen it | |
| Workbook | `:w` | `SaveWorkbook` | Save the active workbook | |
| Workbook | `:q` | `CloseAskSaving` | Close the active workbook (show a dialog if there are unsaved changes) | |
| Workbook | `:q!`/`ZQ` | `CloseWithoutSaving` | Close the active workbook without saving | |
| Workbook | `:wq`/`:x`/`ZZ` | `CloseWithSaving` | Save and close the active workbook | |
| Workbook | `:saveas` | `SaveAsNewWorkbook` | Save as a new workbook | |
| Workbook | `:b [num]` | `ActivateWorkbook` | Select the workbook at position `[num]` | |
| Workbook | `]b`/`:bnext` | `NextWorkbook` | Select the next workbook | ✓ |
| Workbook | `[b`/`:bprevious` | `PreviousWorkbook` | Select the previous workbook | ✓ |
| Workbook | `~` | `ToggleReadOnly` | Toggle read-only mode | |
| Other | `u` | `Undo_CtrlZ` | Undo (send `Ctrl + Z`) | |
| Other | `<C-r>` | `RedoExecute` | Redo | |
| Other | `.` | `RepeatAction` | Repeat the previous action (limited to commands where `repeatRegister` is called) | |
| Other | `m` | `ZoomIn` | Zoom in by 10% or `[count]`% | ✓ |
| Other | `M` | `ZoomOut` | Zoom out by 10% or `[count]`% | ✓ |
| Other | `%` | `ZoomSpecifiedScale` | Set the zoom level to `[count]`% (digits `1`-`9` correspond to predefined zoom levels) | ✓ |
| Other | `\` | `ShowContextMenu` | Show the context menu | |
| Other | `:sort` | `Sort` | Sort in ascending order | |
| Other | `:sort!` | `Sort` | Sort in descenging order | |
| Other | `:unique` | `RemoveDuplicates` | Delete duplicates rows from sheet | |
| Other | `:opendir` | `OpenActiveBookDir` | Open file location | |
| Other | `:fullpath` | `YankActiveBookPath` | Copy full path to clipboard | |
| Other | `<C-i>` | `JumpNext` | Move to the next cell in the jump list | ✓ |
| Other | `<C-o>` | `JumpPrev` | Move to the previous cell in the jump list | ✓ |
| Other | `:clearjumps` | `ClearJumps` | Clear the jump list | |
| Other | `:help <KEY>` | `SearchHelp` | Search help for a given `<KEY>` | |
| Other | `zf` | `ToggleFreezePanes` | Toggle freeze panes on/off | |
| Other | `=v` | `ToggleFormulaBar` | Toggle the visibility of the formula bar | |
| Other | `gb` | `ToggleGridlines` | Toggle the visibility of the gridlines | |
| Other | `gh` | `ToggleHeadings` | Toggle the visibility of the headings | |
| Other | `gs` | `ShowSummaryInfo` | Show the file properties | |
| Other | `zp` | `SetPrintArea` | Set the selected cells as the print area | |
| Other | `zP` | `ClearPrintArea` | Clear the print area | |
| Other | `@@` | `ShowMacroDialog` | Show the macro dialog | |
| Other | `1-9` | `ShowCmdForm` | Specify `[count]` (only works with features marked with ✓ in Count) | |
| CmdLine | `<Tab>` | `ShowSuggest` | Show command suggestions if possible | |

</div></details>

\* The default keymaps are defined with `Map` method in [DefaultConfig.bas](./src/DefaultConfig.bas).

### 🔧 Custom Key Mapping

- Normal Mode
    - `<C-[>` → `<Esc>`
- Cmdline Mode
    - `<C-w>` → `<C-BS>`
    - `<C-u>` → `<S-Home><BS>`
    - `<C-k>` → `<S-End><Del>`
    - `<C-h>` → `<Left>`
    - `<C-l>` → `<Right>`
    - `<C-a>` → `<Home>`
    - `<C-e>` → `<End>`

## ⚙️ Customization

By placing the [configuration file](./config/_vimxlamrc) in the directory where vim.xlam is saved, you can load the settings at startup. The file must be named `_vimxlamrc`. Please save the file with cp932 encoding.

### 🔤 Syntax

- Lines starting with `#` or blank lines are ignored.
- Lines starting with `set` allow you to modify defined configuration values.
- Lines containing `map` or `unmap` allow you to modify key mappings.

### 🛠️ Options

You can configure using the same syntax as Vim's `set`. For configuration examples, refer to the [configuration file](./config/_vimxlamrc).

| Option Key | Type | Description | Default |
| ---------- | ---- | ----------- | ------- |
| `statusprefix` | string | Prefix for temporary messages displayed in the status bar | `vim.xlam: ` |
| `togglekey` | string | Key to toggle Vim mode on/off (Vim-style key specification) | `<C-m>` |
| `scrolloff` | float | Up and down offset for `ScrollCurrentXXX` series (px) | `54.0` |
| `jumplisthistory` | int | Maximum number of items to keep in the jump list | `100` |
| `[no]japanese` | bool | Japanese mode / English mode | `True` |
| `[no]jiskeyboard` | bool | JIS keyboard / US keyboard | `True` |
| `[no]quitapp` | bool | Quit Excel or not when closing the last workbook | `True` |
| `[no]numpadcount` | bool |  Whether NumPad is used as `[count]` or not | `False` |
| `suggestwait` | int | Delay time to display suggestions (ms, 0 to disable) | `1000` |
| `suggestlabels` | string | Shortcut labels for suggestion | Ommited |
| `colorpickersize` | float | ColorPicker form size (px) | `12.0` |
| `customcolor1` | string | Custom color #1 in ColorPicker | `#ff6600` ![#ff6600](https://placehold.co/15/ff6600/ff6600) |
| `customcolor2` | string | Custom color #2 in ColorPicker | `#ff9966` ![#ff9966](https://placehold.co/15/ff9966/ff9966) |
| `customcolor3` | string | Custom color #3 in ColorPicker | `#ff00ff` ![#ff00ff](https://placehold.co/15/ff00ff/ff00ff) |
| `customcolor4` | string | Custom color #4 in ColorPicker | `#008000` ![#008000](https://placehold.co/15/008000/008000) |
| `customcolor5` | string | Custom color #5 in ColorPicker | `#0000ff` ![#0000ff](https://placehold.co/15/0000ff/0000ff) |
| `[no]debug` | bool | Enable / disable debug mode | `False` |

#### Notes on `numpadcount`

- When `set numpadcount` is specified, the `ShowCmdForm` is automatically set for NumPad 1-9. You do not need to manually map keys using `nmap`.
- If you explicitly specify `set nonumpadcount`:
    - Key mappings for NumPad 1-9 set before this command will be cleared.
    - If you want to set key mappings for NumPad 1-9, please specify them after this setting.

By default, no key mappings are applied to the NumPad. This is useful if you want to input numbers without exiting Vim mode. Additionally, by setting key mappings, you can use the NumPad as a launcher to trigger useful functions with a single key.

However, with the default configuration, NumPad keys cannot be used as `[count]`. By setting `set numpadcount`, NumPad keys can be used as `[count]`, but they will no longer function as a launcher. (You can still assign keys like `0` and other symbol keys).

**Launcher Configuration Example**

```vim
nmap <kplus> AddNumber
nmap <kminus> SubtractNumber
nmap <k1> AlignLeft
nmap <k2> AlignCenter
nmap <k3> AlignRight
```

### 🗺️ Keymap

You can add/change/remove key mappings.

- `{lhs}` specifies the key mapping in the original Vim style.
  - If `<cmd>` is specified, it is treated as a command mode specification and processed as a simple string.
- `{rhs}` specifies the function name you want to execute.
  - If `<key>` is specified, another key is specified similar to `{lhs}` in Vim key mapping.

There are currently four modes.

- `n` (Normal): Normal mode. Basically, always map in mode.
- `v` (Visual): Visual mode. Specify keys to return to Normal mode, etc.
- `c` (Cmdline): Command-line mode. Effective in command-line mode with characters such as `:` or `/`.
- `i` (Shape_Insert): Shape Insert mode. Effective when pressing `i`/`a`, etc., during shape selection.

**Syntax for `map` and `unmap`**

```
[n|v|c|i]map [<cmd>]{lhs} [<key>]{rhs} [arg1] [arg2] [...]
 ^^^^^^^      ^^^^^ ^^^^^  ^^^^^ ^^^^^  ^^^^^^^^^^^^^^^^^
   |            |     |      |     |     `- args: Arguments of the function specified by {rhs}
   |            |     |      |     `------- rhs : Function name to be execute
   |            |     |      `------------- key : Flag to simulate keys with {rhs}
   |            |     `-------------------- lhs : Key sequence (vim style)
   |            `-------------------------- cmd : Flag to enable in command mode (plain text)
   `--------------------------------------- mode: Specify pre-defined mode ("n" if omitted)

[n|v|c|i]unmap [<cmd>]{lhs}
         ^^^^^
          `--- disable mapping
```

\* It follows the same syntax as that specified in [DefaultConfig.bas](./src/DefaultConfig.bas), so refer to it as needed.

## 🚀 Contributing

[Issues](https://github.com/sha5010/vim.xlam/issues) and [Pull Requests](https://github.com/sha5010/vim.xlam/pulls) are welcome. If you've developed your own features and would like to contribute, I'd appreciate your help.

English version of the README was generated by ChatGPT. If you come across any errors or have suggestions for improvements, please don't hesitate to let me know. Your feedback is highly appreciated.

## 😎 Author

[@sha_5010](https://twitter.com/sha_5010)

## 💡 Related Projects

- [ExcelLikeVim](https://github.com/kjnh10/ExcelLikeVim)
- [VimExcel](https://www.vector.co.jp/soft/winnt/business/se494158.html) (Japanese only)
- [vixcel](https://github.com/codetsar/vixcel)
- [Excel\_Vim\_Keys](https://github.com/treatmesubj/Excel_Vim_Keys)

## 🔒 License

[MIT](./LICENSE)
