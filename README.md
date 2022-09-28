<p align="center">
  <img alt="vim xlam_banner" src="https://user-images.githubusercontent.com/95682647/175554011-5f9b5a37-a08d-47f7-ac63-b162620cc99d.png" width="600">
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

[VimExcel](https://www.vector.co.jp/soft/winnt/business/se494158.html) を参考に、ExcelでVimのキーバインドが使えるようにした Excelアドインです。

This is an Excel add-in that allows Vim keybindings to be used in Excel, with reference to [Vimexcel](https://www.vector.co.jp/soft/winnt/business/se494158.html).

## Description

vim.xlam は vim のような使用感で Excel 上でもキーボード主体で操作できるようにするための Excelアドインです。

拡張性を意識して作成しており、自身でメソッドを作成し `map` メソッドでキー割り当てを行うことで、簡単にカスタマイズできます。またデフォルトのキーバインドから簡単に変えられるように設計していますので、ご自分に取って最適なキーバインドに設定することができます。

***Demo:***

![demo](https://user-images.githubusercontent.com/95682647/175773473-50376812-afcc-4ced-b436-7150d7b97872.gif)

\* サンプルファイルは [https://atelierkobato.com](https://atelierkobato.com/download/) 様のものを使用しております。

## Features

- `hjkl` を基本としたセル移動だけでなく、`gg`、`G`、`^`、`$` といったジャンプコマンドも多数使用可能
- フォント、背景色、罫線などの設定もマウス操作なしで効率的に実施可能
- コメント操作、スクロール操作、ワークシート操作などの機能も搭載
- 最後に編集したセルやジャンプ前のセルを記憶し、ジャンプする機能も搭載 (ジャンプリスト)
- 容易なカスタマイズ性を追求しており、どなたでもカスタマイズ可能

## Installation

1. [リリースページ](https://github.com/sha5010/vim.xlam/releases/latest)から最新の vim.xlam をダウンロードしてください。(または[ここ](https://github.com/sha5010/vim.xlam/releases/latest/download/vim.xlam)から直接ダウンロードできます）
2. ダウンロードした vim.xlam を `C:\Users\<USERNAME>\AppData\Roaming\Microsoft\AddIns` 配下に保存してください。
3. Excel を起動し、ファイル &gt; オプション &gt; アドイン と進み、画面下部の **設定...** ボタンをクリックしてください。
4. **参照...** のボタンをクリックし、保存した vim.xlam を選択してアドインを追加すれば完了です。

## Usage

- デフォルトの設定では `Ctrl + M` キーを押すことで Vimモードのオン/オフを切り替え可能
- セルの移動は `hjkl` で実施できるほか、`a` や `i` などでセルの編集が可能
- その他、多数のコマンドが使用可能

### Default Keybindings

**主なコマンド**

| Type | Keystroke | Action | Description |
| ---- | --------- | ------ | ----------- |
| Core | `<C-m>` | `toggleVim` | Vimモードの切替 |
| InsertMode | `a` | `appendFollowLangMode` | IMEを言語モードに合わせてセルを末尾から編集 |
| InsertMode | `i` | `insertFollowLangMode` | IMEを言語モードに合わせてセルを先頭から編集 |
| InsertMode | `s` | `substituteFollowLangMode` | IMEを言語モードに合わせてセルをクリアして編集 |
| Moving | `h` | `moveLeft` | ← |
| Moving | `j` | `moveDown` | ↓ |
| Moving | `k` | `moveUp` | ↑ |
| Moving | `l` | `moveRight` | → |
| Moving | `gg` | `moveToTopRow` | 1行目に移動。`[count]` ありなら `[count]` 行へ移動 |
| Border | `bb` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 細線) |
| Border | `ba` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 細線) |
| Border | `bia` | `toggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 細線) |
| Row | `ra` | `appendRows` | 現在行の後に行を挿入。`[count]` が与えられたときは `[count]` 行挿入 |
| Row | `ri` | `insertRows` | 現在行の前に行を挿入。`[count]` が与えられたときは `[count]` 行挿入 |
| Row | `rd` | `deleteRows` | 現在行を削除。`[count]` が与えられたときは `[count]` 行削除 |
| Column | `ca` | `appendColumns` | 現在列の後に列を挿入。`[count]` が与えられたときは `[count]` 列挿入 |
| Column | `ci` | `insertColumns` | 現在列の前に列を挿入。`[count]` が与えられたときは `[count]` 列挿入 |
| Column | `cd` | `deleteColumns` | 現在列を削除。`[count]` が与えられたときは `[count]` 列削除 |
| Delete | `D` | `deleteValue` | セルの値を削除 |
| Paste | `p` | `pasteSmart` | 行や列がコピーされたときは挿入。それ以外は `Ctrl + V` で貼り付け |
| Font | `<` | `decreaseFontSize` | フォントサイズの縮小 |
| Font | `>` | `increaseFontSize` | フォントサイズの拡大 |
| Color | `fc` | `smartFontColor` | フォントの色を選択するダイアログを表示 |
| Find & Replace | `/` | `showFindFollowLang` | IMEを言語モードに合わせて検索ダイアログを表示 |
| Find & Replace | `n` | `nextFoundCell` | 検索結果の次のセルを選択 |
| Find & Replace | `N` | `previousFoundCell` | 検索結果の前のセルを選択 |
| Scrolling | `<C-u>` | `scrollUpHalf` | 半ページ上スクロール |
| Scrolling | `<C-d>` | `scrollDownHalf` | 半ページ下スクロール |
| Scrolling | `zt` | `scrollCurrentTop` | 現在行が最上部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる)|
| Scrolling | `zz` | `scrollCurrentMiddle` | 現在行が中央に来るように縦スクロール |
| Scrolling | `zb` | `scrollCurrentBottom` | 現在行が最下部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる)|
| Worksheet | `v` | `nextWorksheet` | 次のシートを選択 |
| Worksheet | `V` | `previousWorksheet` | 前のシートを選択 |
| Worksheet | `wr` | `renameWorksheet` | アクティブなシート名を変更 |
| Worksheet | `ww` | `showSheetPicker` | SheetPicker を起動 |
| Workbook | `:w` | `saveWorkbook` | アクティブブックを保存 |
| Workbook | `:q` | `closeAskSaving` | アクティブブックを閉じる(未保存時はダイアログを表示) |
| Workbook | `:q!` | `closeWithoutSaving` | アクティブブックを保存せずに閉じる |
| Workbook | `:wq` | `closeWithSaving` | アクティブブックを保存して閉じる |
| Other | `u` | `undo_CtrlZ` | 元に戻す (`Ctrl + Z` を送出)|
| Other | `<C-r>` | `redoExecute` | やり直し |

<details><summary>全てのコマンドはこちらを展開</summary><div>

| Type | Keystroke | Action | Description |
| ---- | --------- | ------ | ----------- |
| Core | `<C-m>` | `toggleVim` | Vimモードの切替 |
| Core | `<C-p>` | `toggleLang` | 言語モードの切替 (日本語/英語) |
| Core | `:r` | `reloadVim` | vim.xlam をリロード |
| Core | `:r!` | `reloadVim` | vim.xlam をリロード (キーバインドを再適用) |
| Core | `:debug` | `toggleDebugMode` | デバッグモードを切り替える |
| InsertMode | `a` | `appendFollowLangMode` | IMEを言語モードに合わせてセルを末尾から編集 |
| InsertMode | `A` | `appendNotFollowLangMode` | IMEを言語モードに合わせずセルを末尾から編集 |
| InsertMode | `i` | `insertFollowLangMode` | IMEを言語モードに合わせてセルを先頭から編集 |
| InsertMode | `I` | `insertNotFollowLangMode` | IMEを言語モードに合わせずセルを先頭から編集 |
| InsertMode | `s` | `substituteFollowLangMode` | IMEを言語モードに合わせてセルをクリアして編集 |
| InsertMode | `S` | `substituteNotFollowLangMode` | IMEを言語モードに合わせずセルをクリアして編集 |
| Moving | `h` | `moveLeft` | ← |
| Moving | `j` | `moveDown` | ↓ |
| Moving | `k` | `moveUp` | ↑ |
| Moving | `l` | `moveRight` | → |
| Moving | `H` | `moveLeft` | Shift + ← |
| Moving | `J` | `moveDown` | Shift + ↓ |
| Moving | `K` | `moveUp` | Shift + ↑ |
| Moving | `L` | `moveRight` | Shift + → |
| Moving | `<C-h>` | `moveLeft` | Ctrl + ← |
| Moving | `<C-j>` | `moveDown` | Ctrl + ↓ |
| Moving | `<C-k>` | `moveUp` | Ctrl + ↑ |
| Moving | `<C-l>` | `moveRight` | Ctrl + → |
| Moving | `<C-S-H>` | `moveLeft` | Ctrl + Shift + ← |
| Moving | `<C-S-J>` | `moveDown` | Ctrl + Shift + ↓ |
| Moving | `<C-S-K>` | `moveUp` | Ctrl + Shift + ↑ |
| Moving | `<C-S-L>` | `moveRight` | Ctrl + Shift + → |
| Moving | `gg` | `moveToTopRow` | 1行目に移動。`[count]` ありなら `[count]` 行へ移動 |
| Moving | `G` | `moveToLastRow` | UsedRange の最終行に移動。`[count]` ありなら `[count]` 行へ移動 |
| Moving | `\|` | `moveToNthColumn` | `[count]` 列目に移動 |
| Moving | `0` | `moveToFirstColumn` | 1列目に移動 |
| Moving | `^` | `moveToLeftEnd` | UsedRange の最初の列に移動 |
| Moving | `$` | `moveToRightEnd` | UsedRange の最後の列に移動 |
| Moving | `g0` | `moveToA1` | A1セルに移動 |
| Moving | `{` | `moveToTopOfCurrentRegion` | CurrentRegion 内で最初の行に移動 |
| Moving | `}` | `moveToBottomOfCurrentRegion` | CurrentRegion 内で最後の行に移動 |
| Moving | `W[cell]` | `moveToSpecifiedCell` | 指定された `[cell]` へ移動 |
| Moving | `:[num]` | `moveToSpecifiedRow` | 指定された `[num]` 行目に移動 |
| Cell | `xx` | `cutCell` | セルを切り取り |
| Cell | `yy` | `yankCell` | セルをコピー|
| Cell | `o` | `insertCellsDown` | 選択セルの下にセルを挿入|
| Cell | `O` | `insertCellsUp` | 選択セルの上にセルを挿入|
| Cell | `t` | `insertCellsRight` | 選択セルの右にセルを挿入 |
| Cell | `T` | `insertCellsLeft` | 選択セルの左にセルを挿入 |
| Cell | `+` | `incrementText` | インデントを増やす |
| Cell | `-` | `decrementText` | インデントを減らす |
| Cell | `[` | `increaseDecimal` | 小数点表示桁上げ |
| Cell | `]` | `decreaseDecimal` | 小数点表示桁下げ |
| Cell | `<Space>` | `unionSelectCells` | 現在セルを記憶に追加し、記憶したセルを選択 (複数セルの選択が可能) |
| Cell | `<S-Space>` | `exceptSelectCells` | 記憶された選択済みセルから現在セルを取り除く |
| Cell | `@f` | `followHyperlinkOfActiveCell` | セルのハイパーリンクを開く |
| Border | `bb` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 細線) |
| Border | `ba` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 細線) |
| Border | `bh` | `toggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 細線) |
| Border | `bj` | `toggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 細線) |
| Border | `bk` | `toggleBorderTop` | 選択セルの上に罫線を設定 (実線, 細線) |
| Border | `bl` | `toggleBorderRight` | 選択セルの右に罫線を設定 (実線, 細線) |
| Border | `bia` | `toggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 細線) |
| Border | `bis` | `toggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 細線) |
| Border | `biv` | `toggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 細線) |
| Border | `b/` | `toggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 細線) |
| Border | `b\` | `toggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 細線) |
| Border | `bB` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 太線) |
| Border | `bA` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 太線) |
| Border | `bH` | `toggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 太線) |
| Border | `bJ` | `toggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 太線) |
| Border | `bK` | `toggleBorderTop` | 選択セルの上に罫線を設定 (実線, 太線) |
| Border | `bL` | `toggleBorderRight` | 選択セルの右に罫線を設定 (実線, 太線) |
| Border | `Bb` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 太線) |
| Border | `Ba` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 太線) |
| Border | `Bh` | `toggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 太線) |
| Border | `Bj` | `toggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 太線) |
| Border | `Bk` | `toggleBorderTop` | 選択セルの上に罫線を設定 (実線, 太線) |
| Border | `Bl` | `toggleBorderRight` | 選択セルの右に罫線を設定 (実線, 太線) |
| Border | `Bia` | `toggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 太線) |
| Border | `Bis` | `toggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 太線) |
| Border | `Biv` | `toggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 太線) |
| Border | `B/` | `toggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 太線) |
| Border | `B\` | `toggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 太線) |
| Border | `bob` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 極細線) |
| Border | `boa` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 極細線) |
| Border | `boh` | `toggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 極細線) |
| Border | `boj` | `toggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 極細線) |
| Border | `bok` | `toggleBorderTop` | 選択セルの上に罫線を設定 (実線, 極細線) |
| Border | `bol` | `toggleBorderRight` | 選択セルの右に罫線を設定 (実線, 極細線) |
| Border | `boia` | `toggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 極細線) |
| Border | `bois` | `toggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 極細線) |
| Border | `boiv` | `toggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 極細線) |
| Border | `bo/` | `toggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 極細線) |
| Border | `bo\` | `toggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 極細線) |
| Border | `bmb` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 中線) |
| Border | `bma` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 中線) |
| Border | `bmh` | `toggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 中線) |
| Border | `bmj` | `toggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 中線) |
| Border | `bmk` | `toggleBorderTop` | 選択セルの上に罫線を設定 (実線, 中線) |
| Border | `bml` | `toggleBorderRight` | 選択セルの右に罫線を設定 (実線, 中線) |
| Border | `bmia` | `toggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 中線) |
| Border | `bmis` | `toggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 中線) |
| Border | `bmiv` | `toggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 中線) |
| Border | `bm/` | `toggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 中線) |
| Border | `bm\` | `toggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 中線) |
| Border | `btb` | `toggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (二重線, 太線) |
| Border | `bta` | `toggleBorderAround` | 選択セルの周りに罫線を設定 (二重線, 太線) |
| Border | `bth` | `toggleBorderLeft` | 選択セルの左に罫線を設定 (二重線, 太線) |
| Border | `btj` | `toggleBorderBottom` | 選択セルの下に罫線を設定 (二重線, 太線) |
| Border | `btk` | `toggleBorderTop` | 選択セルの上に罫線を設定 (二重線, 太線) |
| Border | `btl` | `toggleBorderRight` | 選択セルの右に罫線を設定 (二重線, 太線) |
| Border | `btia` | `toggleBorderInner` | 選択セルの内側全てに罫線を設定 (二重線, 太線) |
| Border | `btis` | `toggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (二重線, 太線) |
| Border | `btiv` | `toggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (二重線, 太線) |
| Border | `bt/` | `toggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (二重線, 太線) |
| Border | `bt\` | `toggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (二重線, 太線) |
| Border | `bdd` | `deleteBorderAll` | 選択セルの周りと内側の全ての罫線を削除 |
| Border | `bda` | `deleteBorderAround` | 選択セルの周りの罫線を削除 |
| Border | `bdh` | `deleteBorderLeft` | 選択セルの左の罫線を削除 |
| Border | `bdj` | `deleteBorderBottom` | 選択セルの下の罫線を削除 |
| Border | `bdk` | `deleteBorderTop` | 選択セルの上の罫線を削除 |
| Border | `bdl` | `deleteBorderRight` | 選択セルの右の罫線を削除 |
| Border | `bdia` | `deleteBorderInner` | 選択セルの内側全ての罫線を削除 |
| Border | `bdis` | `deleteBorderInnerHorizontal` | 選択セルの内側水平の罫線を削除 |
| Border | `bdiv` | `deleteBorderInnerVertical` | 選択セルの内側垂直の罫線を削除 |
| Border | `bd/` | `deleteBorderDiagonalUp` | 選択セルの `/` 方向の罫線を削除 |
| Border | `bd\` | `deleteBorderDiagonalDown` | 選択セルの `\` 方向の罫線を削除 |
| Border | `bcc` | `setBorderColorAll` | 選択セルの周りと内側の全ての罫線の色を設定 |
| Border | `bca` | `setBorderColorAround` | 選択セルの周りの罫線の色を設定 |
| Border | `bch` | `setBorderColorLeft` | 選択セルの左の罫線の色を設定 |
| Border | `bcj` | `setBorderColorBottom` | 選択セルの下の罫線の色を設定 |
| Border | `bck` | `setBorderColorTop` | 選択セルの上の罫線の色を設定 |
| Border | `bcl` | `setBorderColorRight` | 選択セルの右の罫線の色を設定 |
| Border | `bcia` | `setBorderColorInner` | 選択セルの内側全ての罫線の色を設定 |
| Border | `bcis` | `setBorderColorInnerHorizontal` | 選択セルの内側水平の罫線の色を設定 |
| Border | `bciv` | `setBorderColorInnerVertical` | 選択セルの内側垂直の罫線の色を設定 |
| Border | `bc/` | `setBorderColorDiagonalUp` | 選択セルの `/` 方向の罫線の色を設定 |
| Border | `bc\` | `setBorderColorDiagonalDown` | 選択セルの `\` 方向の罫線の色を設定 |
| Row | `r-` | `narrowRowsHeight` | 行の高さを1pt狭くする |
| Row | `r+` | `wideRowsHeight` | 行の高さを1pt広くする |
| Row | `rr` | `selectRows` | 行を選択。`[count]` が与えられたときは `[count]` 行選択 |
| Row | `ra` | `appendRows` | 現在行の後に行を挿入。`[count]` が与えられたときは `[count]` 行挿入 |
| Row | `ri` | `insertRows` | 現在行の前に行を挿入。`[count]` が与えられたときは `[count]` 行挿入 |
| Row | `rd` | `deleteRows` | 現在行を削除。`[count]` が与えられたときは `[count]` 行削除 |
| Row | `ry` | `yankRows` | 現在行をコピー。`[count]` が与えられたときは `[count]` 行コピー |
| Row | `rx` | `cutRows` | 現在行を切り取り。`[count]` が与えられたときは `[count]` 行切り取り |
| Row | `rh` | `hideRows` | 現在行を非表示化。`[count]` が与えられたときは `[count]` 行非表示化 |
| Row | `rH` | `unhideRows` | 現在行を再表示。`[count]` が与えられたときは `[count]` 行再表示 |
| Row | `rg` | `groupRows` | 現在行をグループ化。`[count]` が与えられたときは `[count]` 行グループ化 |
| Row | `ru` | `ungroupRows` | 現在行をグループ化解除。`[count]` が与えられたときは `[count]` 行グループ化解除 |
| Row | `rf` | `foldRowsGroup` | 現在行を畳む。`[count]` が与えられたときは `[count]` 行畳む |
| Row | `rs` | `spreadRowsGroup` | 現在行の折り畳みを開く。`[count]` が与えられたときは `[count]` 行開く |
| Row | `rj` | `adjustRowsHeight` | 現在行の高さを自動調整。`[count]` が与えられたときは `[count]` 行自動調整 |
| Row | `rw` | `setRowsHeight` | 現在行の高さを任意に設定。`[count]` が与えられたときは `[count]` 行設定 |
| Column | `c-` | `narrowColumnsWidth` | 列幅を1pt狭くする |
| Column | `c+` | `wideColumnsWidth` | 列幅を1pt広くする |
| Column | `cc` | `selectColumns` | 列を選択。`[count]` が与えられたときは `[count]` 列選択 |
| Column | `ca` | `appendColumns` | 現在列の後に列を挿入。`[count]` が与えられたときは `[count]` 列挿入 |
| Column | `ci` | `insertColumns` | 現在列の前に列を挿入。`[count]` が与えられたときは `[count]` 列挿入 |
| Column | `cd` | `deleteColumns` | 現在列を削除。`[count]` が与えられたときは `[count]` 列削除 |
| Column | `cy` | `yankColumns` | 現在列をコピー。`[count]` が与えられたときは `[count]` 列コピー |
| Column | `cx` | `cutColumns` | 現在列を切り取り。`[count]` が与えられたときは `[count]` 列切り取り |
| Column | `ch` | `hideColumns` | 現在列を非表示化。`[count]` が与えられたときは `[count]` 列非表示化 |
| Column | `cH` | `unhideColumns` | 現在列を再表示。`[count]` が与えられたときは `[count]` 列再表示 |
| Column | `cg` | `groupColumns` | 現在列をグループ化。`[count]` が与えられたときは `[count]` 列グループ化 |
| Column | `cu` | `ungroupColumns` | 現在列をグループ化解除。`[count]` が与えられたときは `[count]` 列グループ化解除 |
| Column | `cf` | `foldColumnsGroup` | 現在列を畳む。`[count]` が与えられたときは `[count]` 列畳む |
| Column | `cs` | `spreadColumnsGroup` | 現在列の折り畳みを開く。`[count]` が与えられたときは `[count]` 列開く |
| Column | `cj` | `adjustColumnsWidth` | 現在列の幅を自動調整。`[count]` が与えられたときは `[count]` 列自動調整 |
| Column | `cw` | `setColumnsWidth` | 現在列の幅を任意に設定。`[count]` が与えられたときは `[count]` 列設定 |
| Yank | `yr` | `yankRows` | 現在行をコピー。`[count]` が与えられたときは `[count]` 行コピー |
| Yank | `yc` | `yankColumns` | 現在列をコピー。`[count]` が与えられたときは `[count]` 列コピー |
| Yank | `ygg` | `yankToTopRows` | 現在行から1行目までをコピー |
| Yank | `yG` | `yankToBottomRows` | 現在行から UsedRange の最終行までをコピー |
| Yank | `y{` | `yankToTopOfCurrentRegionRows` | 現在行から CurrentRegion の最初の行までをコピー |
| Yank | `y}` | `yankToBottomOfCurrentRegionRows` | 現在行から CurrentRegion の最後の行までをコピー |
| Yank | `y0` | `yankToLeftEndColumns` | 現在列から UsedRange の最初の列までをコピー |
| Yank | `y$` | `yankToRightEndColumns` | 現在列から UsedRange の最後の列までをコピー |
| Yank | `y^` | `yankToLeftOfCurrentRegionColumns` | 現在列から CurrentRegion  の最初の列までをコピー |
| Yank | `yg$` | `yankToRightOfCurrentRegionColumns` | 現在列から CurrentRegion の最後の列までをコピー |
| Yank | `yh` | `yankFromLeftCell` | 現在のセルの左の値をコピーして貼り付け |
| Yank | `yj` | `yankFromDownCell` | 現在のセルの下の値をコピーして貼り付け |
| Yank | `yk` | `yankFromUpCell` | 現在のセルの上の値をコピーして貼り付け |
| Yank | `yl` | `yankFromRightCell` | 現在のセルの右の値をコピーして貼り付け |
| Delete | `X` | `deleteValue` | セルの値を削除 |
| Delete | `D` | `deleteValue` | セルの値を削除 |
| Delete | `dx` | `deleteRows` | 現在行を削除。`[count]` が与えられたときは `[count]` 行削除 |
| Delete | `dr` | `deleteRows` | 現在行を削除。`[count]` が与えられたときは `[count]` 行削除 |
| Delete | `dc` | `deleteColumns` | 現在列を削除。`[count]` が与えられたときは `[count]` 列削除 |
| Delete | `dgg` | `deleteToTopRows` | 現在行から先頭行までを削除 |
| Delete | `dG` | `deleteToBottomRows` | 現在行から UsedRange の最終行までを削除 |
| Delete | `d{` | `deleteToTopOfCurrentRegionRows` | 現在行から CurrentRegion の最初の行までを削除 |
| Delete | `d}` | `deleteToBottomOfCurrentRegionRows` | 現在行から CurrentRegion の最後の行までを削除 |
| Delete | `d0` | `deleteToLeftEndColumns` | 現在列から UsedRange の最初の列までを削除 |
| Delete | `d$` | `deleteToRightEndColumns` | 現在列から UsedRange の最後の列までを削除 |
| Delete | `d^` | `deleteToLeftOfCurrentRegionColumns` | 現在列から CurrentRegion  の最初の列までを削除 |
| Delete | `dg$` | `deleteToRightOfCurrentRegionColumns` | 現在列から CurrentRegion の最後の列までを削除 |
| Delete | `dh` | `deleteToLeft` | 現在のセルを削除し左方向へシフト
| Delete | `dj` | `deleteToUp` | 現在のセルを削除し上方向へシフト |
| Delete | `dk` | `deleteToUp` | 現在のセルを削除し上方向へシフト |
| Delete | `dl` | `deleteToLeft` | 現在のセルを削除し左方向へシフト |
| Cut | `xr` | `cutRows` | 現在行を切り取り。`[count]` が与えられたときは `[count]` 行切り取り |
| Cut | `xc` | `cutColumns` | 現在列を切り取り。`[count]` が与えられたときは `[count]` 列切り取り |
| Cut | `xgg` | `cutToTopRows` | 現在行から1行目までを切り取り |
| Cut | `xG` | `cutToBottomRows` | 現在行から UsedRange の最後の行までを切り取り |
| Cut | `x{` | `cutToTopOfCurrentRegionRows` | 現在行から CurrentRegion  の最初の列までを切り取り |
| Cut | `x}` | `cutToBottomOfCurrentRegionRows` | 現在行から CurrentRegion の最後の行までを切り取り |
| Cut | `x0` | `cutToLeftEndColumns` | 現在列から UsedRange の最初の列までを切り取り |
| Cut | `x$` | `cutToRightEndColumns` | 現在列から UsedRange の最後の列までを切り取り |
| Cut | `x^` | `cutToLeftOfCurrentRegionColumns` | 現在列から CurrentRegion  の最初の列までを切り取り |
| Cut | `xg$` | `cutToRightOfCurrentRegionColumns` | 現在列から CurrentRegion の最後の列までを切り取り |
| Paste | `p` | `pasteSmart` | 行や列がコピーされたときは挿入。それ以外は `Ctrl + V` で貼り付け |
| Paste | `P` | `pasteSpecial` | 形式を選択して貼り付けのダイアログを表示 |
| Paste | `U` | `pasteValue` | 値のみ貼り付け |
| Font | `<` | `decreaseFontSize` | フォントサイズの縮小 |
| Font | `>` | `increaseFontSize` | フォントサイズの拡大 |
| Font | `fn` | `changeFontName` | フォント名にフォーカス |
| Font | `fs` | `changeFontSize` | フォントサイズにフォーカス |
| Font | `fh` | `alignLeft` | 左揃え |
| Font | `fj` | `alignBottom` | 下揃え |
| Font | `fk` | `alignTop` | 上揃え |
| Font | `fl` | `alignRight` | 右揃え |
| Font | `fo` | `alignCenter` | 文字列中央揃え |
| Font | `fm` | `alignMiddle` | 上下中央揃え |
| Font | `fb` | `toggleBold` | 太字 |
| Font | `fi` | `toggleItalic` | 斜体 |
| Font | `fu` | `toggleUnderline` | 下線 |
| Font | `f-` | `toggleStrikethrough` | 取り消し線 |
| Font | `ft` | `changeFormat` | 表示形式にフォーカス |
| Font | `ff` | `showFontDialog` | セルの書式設定のダイアログを表示 |
| Color | `fc` | `smartFontColor` | フォントの色を選択するダイアログを表示 |
| Color | `Fc` | `smartFillColor` | 塗りつぶしの色を選択すダイアログを表示 |
| Color | `bc` | `changeShapeBorderColor` | (図形選択時) 枠線の色を選択するダイアログを表示 |
| Comment | `Ci` | `editCellComment` | コメントを編集 (ない場合は追加) |
| Comment | `Cc` | `editCellComment` | コメントを編集 (ない場合は追加) |
| Comment | `Ce` | `deleteCellComment` | 現在セルのコメントを削除 |
| Comment | `Cx` | `deleteCellComment` | 現在セルのコメントを削除 |
| Comment | `Cd` | `deleteCellComment` | 現在セルのコメントを削除 |
| Comment | `CE` | `deleteCellCommentAll` | シート上のコメントを全て削除 |
| Comment | `CD` | `deleteCellCommentAll` | シート上のコメントを全て削除 |
| Comment | `Ca` | `toggleCellComment` | 現在セルのコメントの表示/非表示を切り替え |
| Comment | `Cr` | `showCellComment` | 現在セルのコメントを表示 |
| Comment | `Cm` | `hideCellComment` | 現在セルのコメントを非表示 |
| Comment | `CA` | `toggleCellCommentAll` | すべてのコメントの表示/非表示を切り替え |
| Comment | `CR` | `showCellCommentAll` | すべてのコメントを表示 |
| Comment | `CM` | `hideCellCommentAll` | すべてのコメントを非表示 |
| Comment | `CH` | `hideCellCommentIndicator` | 現在セルのコメントインジケータを非表示 |
| Comment | `Cn` | `nextCommentedCell` | 次のコメントを選択 |
| Comment | `Cp` | `prevCommentedCell` | 前のコメントを選択 |
| Find & Replace | `/` | `showFindFollowLang` | IMEを言語モードに合わせて検索ダイアログを表示 |
| Find & Replace | `?` | `showFindNotFollowLang` | IMEを言語モードに合わせず検索ダイアログを表示 |
| Find & Replace | `n` | `nextFoundCell` | 検索結果の次のセルを選択 |
| Find & Replace | `N` | `previousFoundCell` | 検索結果の前のセルを選択 |
| Find & Replace | `R` | `showReplaceWindow` | 検索と置換のダイアログを表示 |
| Find & Replace | `*` | `findActiveValueNext` | 現在セルの値で検索し次のセルを選択 |
| Find & Replace | `#` | `findActiveValuePrev` | 現在セルの値で検索し前のセルを選択 |
| Scrolling | `<C-u>` | `scrollUpHalf` | 半ページ上スクロール |
| Scrolling | `<C-d>` | `scrollDownHalf` | 半ページ下スクロール |
| Scrolling | `<C-b>` | `scrollUp` | 1ページ上スクロール |
| Scrolling | `<C-f>` | `scrollDown` | 1ページ下スクロール |
| Scrolling | `<C-y>` | `scrollUp1Row` | 1行上スクロール |
| Scrolling | `<C-e>` | `scrollDown1Row` | 1行下スクロール |
| Scrolling | `zh` | `scrollLeft1Column` | 1列左スクロール |
| Scrolling | `zl` | `scrollRight1Column` | 1列右スクロール |
| Scrolling | `zH` | `scrollLeft` | 1ページ左スクロール |
| Scrolling | `zL` | `scrollRight` | 1ページ右スクロール |
| Scrolling | `zt` | `scrollCurrentTop` | 現在行が最上部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる)|
| Scrolling | `zz` | `scrollCurrentMiddle` | 現在行が中央に来るように縦スクロール |
| Scrolling | `zb` | `scrollCurrentBottom` | 現在行が最下部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる)|
| Scrolling | `zs` | `scrollCurrentLeft` | 現在列が一番左に来るように横スクロール |
| Scrolling | `zm` | `scrollCurrentCenter` | 現在列が中央に来るように横スクロール |
| Scrolling | `ze` | `scrollCurrentRight` | 現在列が一番右に来るように横スクロール |
| Worksheet | `v` | `nextWorksheet` | 次のシートを選択 |
| Worksheet | `V` | `previousWorksheet` | 前のシートを選択 |
| Worksheet | `ww` | `showSheetPicker` | SheetPicker を起動 |
| Worksheet | `ws` | `showSheetPicker` | SheetPicker を起動 |
| Worksheet | `wn` | `nextWorksheet` | 次のシートを選択 |
| Worksheet | `wp` | `previousWorksheet` | 前のシートを選択 |
| Worksheet | `wr` | `renameWorksheet` | アクティブなシート名を変更 |
| Worksheet | `wh` | `moveWorksheetBack` | アクティブなシートを前に移動 |
| Worksheet | `wl` | `moveWorksheetForward` | アクティブなシートを次に移動 |
| Worksheet | `wi` | `insertWorksheet` | アクティブなシートの前に新しいシートを挿入 |
| Worksheet | `wa` | `appendWorksheet` | アクティブなシートの次に新しいシートを挿入 |
| Worksheet | `wd` | `deleteWorksheet` | アクティブなシートを削除 |
| Worksheet | `w0` | `activateLastWorksheet` | 一番最後のシートを選択 |
| Worksheet | `w$` | `activateLastWorksheet` | 一番最後のシートを選択 |
| Worksheet | `wc` | `changeWorksheetTabColor` | アクティブなシートの色を変更 |
| Worksheet | `wy` | `cloneWorksheet` | アクティブなシートを複製 |
| Worksheet | `we` | `exportWorksheet` | シートの移動またはコピーダイアログを表示 |
| Worksheet | `w[num]` | `activateWorksheet` | `[num]` 番目のシートを選択 (1-9 のみ)|
| Workbook | `:e` | `openWorkbook` | ブックを開く |
| Workbook | `:e!` | `reopenActiveWorkbook` | アクティブなブックの変更を破棄し開き直す |
| Workbook | `:w` | `saveWorkbook` | アクティブブックを保存 |
| Workbook | `:q` | `closeAskSaving` | アクティブブックを閉じる(未保存時はダイアログを表示) |
| Workbook | `:q!` | `closeWithoutSaving` | アクティブブックを保存せずに閉じる |
| Workbook | `:wq` | `closeWithSaving` | アクティブブックを保存して閉じる |
| Workbook | `:x` | `closeWithSaving` | アクティブブックを保存して閉じる |
| Workbook | `:b[num]` | `activateWorkbook` | `[num]` 番目のブックを選択 |
| Workbook | `:bn` | `nextWorkbook` | 次のワークブックを選択 |
| Workbook | `:bp` | `previousWorkbook` | 前のワークブックを選択 |
| Workbook | `@a` | `toggleReadOnly` | 読み取り専用を切り替える |
| Workbook | `@n` | `nextWorkbook` | 次のワークブックを選択 |
| Workbook | `@N` | `previousWorkbook` | 前のワークブックを選択 |
| Other | `u` | `undo_CtrlZ` | 元に戻す (`Ctrl + Z` を送出)|
| Other | `<C-r>` | `redoExecute` | やり直し |
| Other | `.` | `repeatAction` | 以前の動作を繰り返す (`repeatRegister` が呼ばれるコマンド限定)|
| Other | `m` | `zoomIn` | 10% ズームイン。`[count]` が与えられたときは `[count]`% ズームイン |
| Other | `M` | `zoomOut` | 10% ズームアウト。`[count]` が与えられたときは `[count]`% ズームアウト |
| Other | `%` | `zoomSpecifiedScale` | 表示倍率を `[count]`% に設定。`1-9` は決まった値に倍率変更 |
| Other | `\` | `showContextMenu` | コンテキストメニューを表示 |
| Other | `<C-i>` | `jumpNext` | ジャンプリストの次のセルへ移動 |
| Other | `<C-o>` | `jumpPrev` | ジャンプリストの前のセルへ移動 |
| Other | `:cle` | `clearJumps` | ジャンプリストをクリア |
| Other | `@w` | `toggleFreezePanes` | ウィンドウ枠の固定のオン/オフを切り替え |
| Other | `@r` | `toggleWrapText` | セルの折り返しのオン/オフを切り替え |
| Other | `@m` | `toggleMergeCells` | セル結合のオン/オフを切り替え |
| Other | `@x` | `toggleFormulaBar` | 関数バーの表示/非表示を切り替え |
| Other | `@s` | `showSummaryInfo` | ファイルのプロパティを表示 |
| Other | `@@` | `showMacroDialog` | マクロダイアログを表示 |
| Other | `@p` | `setPrintArea` | 選択セルを印刷範囲に設定 |
| Other | `@P` | `clearPrintArea` | 印刷範囲をクリア |
| Other | `1-9` | `showCmdForm` | `[count]` を指定 (`5ri` なら5行挿入) |

</div></details>

\* [UserConfig.bas](./src/UserConfig.bas) の中で `map` メソッドを使って定義しています。

### Custom Key Mapping

- `<C-[>` → `<Esc>`

## Customization

Under construction...

## Contributing

[Issue](https://github.com/sha5010/vim.xlam/issues) や [Pull Request](https://github.com/sha5010/vim.xlam/pulls) は大歓迎です。もしご自身で開発された機能がありましたら、開発にご協力いただけますと幸いです。

## Author

[@sha_5010](https://twitter.com/sha_5010)

## License

[MIT](./LICENSE)
