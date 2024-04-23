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
  [<a href="README.md">English</a>]
  [日本語]
</p>

# vim.xlam

Excelでも最高のVim体験を。Excel上でvimのキーバインドを使って操作できるようにするExcelアドインです。

## Description

vim.xlam は vim のような使用感で Excel 上でもキーボード主体で操作できるようにするための Excelアドインです。

拡張性を意識して作成しており、自身でメソッドを作成し `Map` メソッドでキー割り当てを行うことで、簡単にカスタマイズできます。またデフォルトのキーバインドから簡単に変えられるように設計していますので、ご自分に取って最適なキーバインドに設定することができます。

***Demo:***

![demo](https://user-images.githubusercontent.com/95682647/175773473-50376812-afcc-4ced-b436-7150d7b97872.gif)

\* サンプルファイルは [https://atelierkobato.com](https://atelierkobato.com/download/) 様のものを使用しております。

## Features

- `hjkl` を基本としたセル移動だけでなく、`gg`、`G`、`^`、`$` といったジャンプコマンドも多数使用可能
- フォント、背景色、罫線などの設定もマウス操作なしで効率的に実施可能
- コメント操作、スクロール操作、ワークシート操作などの機能も搭載
- 最後に編集したセルやジャンプ前のセルを記憶し、ジャンプする機能も搭載 (ジャンプリスト)
- `count` 指定や `.` リピートなど、Vim の強みを活かした操作
- 容易なカスタマイズ性を追求しており、どなたでもカスタマイズ可能

## Installation

1. [リリースページ](https://github.com/sha5010/vim.xlam/releases/latest)から最新の vim.xlam をダウンロードしてください。(または[最新版を直接ダウンロード](https://github.com/sha5010/vim.xlam/releases/latest/download/vim.xlam)）
2. ダウンロードした vim.xlam を `C:\Users\<USERNAME>\AppData\Roaming\Microsoft\AddIns` 配下に保存してください。
3. Excel を起動し、ファイル &gt; オプション &gt; アドイン と進み、画面下部の **設定...** ボタンをクリックしてください。
4. **参照...** のボタンをクリックし、保存した vim.xlam を選択してアドインを追加すれば完了です。

| :exclamation: **注意** |
| ---- |
| **このプロジェクトは現在開発中です**。リリースの際にしばしば破壊的変更が加わる可能性があります。更新する前にリリースノートをご確認ください。 |

## Usage

- アドインに追加することで自動的に起動するので、あとは最高のVim体験をお楽しみください！
- セル移動/編集、行列の追加/削除、色や罫線の設定 等、割り当てられたキーで操作可能
  - 実装済みの全てのコマンドは、下の表をご覧ください
- [設定ファイル](./config/_vimxlamrc)を配置することで設定やキーマップの[カスタマイズ](#Customization)も可能

### Default Keybindings

**主なコマンド**

| Type | Keystroke | Action | Description | Count |
| ---- | --------- | ------ | ----------- | ----- |
| Core | `<C-m>` | `ToggleVim` | Vimモードの切替 | |
| InsertMode | `a` | `AppendFollowLangMode` | IMEを言語モードに合わせてセルを末尾から編集 | |
| InsertMode | `i` | `InsertFollowLangMode` | IMEを言語モードに合わせてセルを先頭から編集 | |
| InsertMode | `s` | `SubstituteFollowLangMode` | IMEを言語モードに合わせてセルをクリアして編集 | |
| Moving | `h` | `MoveLeft` | ← | ✓ |
| Moving | `j` | `MoveDown` | ↓ | ✓ |
| Moving | `k` | `MoveUp` | ↑ | ✓ |
| Moving | `l` | `MoveRight` | → | ✓ |
| Moving | `gg` | `MoveToTopRow` | 1行目 または `[count]` 行目へ移動 | ✓ |
| Cell | `FF`/`Ff` | `ApplyFlashFill` | フラッシュフィル(適用不可の際はオートフィル) | |
| Cell | `v` | `ToggleVisualMode` | ビジュアルモード(選択範囲の拡張)を切り替え | |
| Border | `bb` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 細線) |
| Border | `ba` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 細線) | |
| Border | `bia` | `ToggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 細線) | |
| Row | `ra` | `AppendRows` | 現在行の後に行を挿入 | ✓ |
| Row | `ri` | `InsertRows` | 現在行の前に行を挿入 | ✓ |
| Row | `rd` | `DeleteRows` | 現在行を削除 | ✓ |
| Column | `ca` | `AppendColumns` | 現在列の後に列を挿入 | ✓ |
| Column | `ci` | `InsertColumns` | 現在列の前に列を挿入 | ✓ |
| Column | `cd` | `DeleteColumns` | 現在列を削除 | ✓ |
| Delete | `D`/`X` | `DeleteValue` | セルの値を削除 | |
| Paste | `p` | `PasteSmart` | 行や列がコピーされたときは後に追加。それ以外は `Ctrl + V` を送出 | ✓ |
| Paste | `P` | `PasteSmart` | 行や列がコピーされたときは前に挿入。それ以外は `Ctrl + V` を送出 | ✓ |
| Font | `-` | `DecreaseFontSize` | フォントサイズの縮小 | ✓ |
| Font | `+` | `IncreaseFontSize` | フォントサイズの拡大 | ✓ |
| Color | `fc` | `SmartFontColor` | フォントの色を選択するダイアログを表示 | |
| Find & Replace | `/` | `ShowFindFollowLang` | IMEを言語モードに合わせて検索ダイアログを表示 | |
| Find & Replace | `n` | `NextFoundCell` | 検索結果の次のセルを選択 | ✓ |
| Find & Replace | `N` | `PreviousFoundCell` | 検索結果の前のセルを選択 | ✓ |
| Scrolling | `<C-u>` | `ScrollUpHalf` | 半ページ上スクロール | ✓ |
| Scrolling | `<C-d>` | `ScrollDownHalf` | 半ページ下スクロール | ✓ |
| Scrolling | `zt` | `ScrollCurrentTop` | 現在行が最上部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる) | ✓ |
| Scrolling | `zz` | `ScrollCurrentMiddle` | 現在行が中央に来るように縦スクロール | ✓ |
| Scrolling | `zb` | `ScrollCurrentBottom` | 現在行が最下部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる) | ✓ |
| Worksheet | `e` | `NextWorksheet` | 次のシートを選択 | ✓ |
| Worksheet | `E` | `PreviousWorksheet` | 前のシートを選択 | ✓ |
| Worksheet | `ww` | `ShowSheetPicker` | SheetPicker を起動 | |
| Worksheet | `wr` | `RenameWorksheet` | アクティブなシート名を変更 | |
| Workbook | `:w` | `SaveWorkbook` | アクティブブックを保存 | |
| Workbook | `:q` | `CloseAskSaving` | アクティブブックを閉じる(未保存時はダイアログを表示) | |
| Workbook | `:q!`/`ZQ` | `CloseWithoutSaving` | アクティブブックを保存せずに閉じる | |
| Workbook | `:wq`/`x`/`ZZ` | `CloseWithSaving` | アクティブブックを保存して閉じる | |
| Other | `u` | `Undo_CtrlZ` | 元に戻す (`Ctrl + Z` を送出) | |
| Other | `<C-r>` | `RedoExecute` | やり直し | |

<details><summary>全てのコマンドはこちらを展開</summary><div>

| Type | Keystroke | Action | Description | Count |
| ---- | --------- | ------ | ----------- | ----- |
| Core | `<C-m>` | `ToggleVim` | Vimモードの切替 | |
| Core | `<C-p>` | `ToggleLang` | 言語モードの切替 (日本語/英語) | |
| Core | `:` | `EnterCmdlineMode` | コマンドラインモードに入る | |
| Core | `:debug` | `ToggleDebugMode` | デバッグモードの切り替え | |
| InsertMode | `a` | `AppendFollowLangMode` | IMEを言語モードに合わせてセルを末尾から編集 | |
| InsertMode | `A` | `AppendNotFollowLangMode` | IMEを言語モードに合わせずセルを末尾から編集 | |
| InsertMode | `i` | `InsertFollowLangMode` | IMEを言語モードに合わせてセルを先頭から編集 | |
| InsertMode | `I` | `InsertNotFollowLangMode` | IMEを言語モードに合わせずセルを先頭から編集 | |
| InsertMode | `s` | `SubstituteFollowLangMode` | IMEを言語モードに合わせてセルをクリアして編集 | |
| InsertMode | `S` | `SubstituteNotFollowLangMode` | IMEを言語モードに合わせずセルをクリアして編集 | |
| Moving | `h` | `MoveLeft` | ← | ✓ |
| Moving | `j` | `MoveDown` | ↓ | ✓ |
| Moving | `k` | `MoveUp` | ↑ | ✓ |
| Moving | `l` | `MoveRight` | → | ✓ |
| Moving | `H` | `MoveLeft` | Shift + ← | ✓ |
| Moving | `J` | `MoveDown` | Shift + ↓ | ✓ |
| Moving | `K` | `MoveUp` | Shift + ↑ | ✓ |
| Moving | `L` | `MoveRight` | Shift + → | ✓ |
| Moving | `<C-h>` | `MoveLeft` | Ctrl + ← | |
| Moving | `<C-j>` | `MoveDown` | Ctrl + ↓ | |
| Moving | `<C-k>` | `MoveUp` | Ctrl + ↑ | |
| Moving | `<C-l>` | `MoveRight` | Ctrl + → | |
| Moving | `<C-S-H>` | `MoveLeft` | Ctrl + Shift + ← | |
| Moving | `<C-S-J>` | `MoveDown` | Ctrl + Shift + ↓ | |
| Moving | `<C-S-K>` | `MoveUp` | Ctrl + Shift + ↑ | |
| Moving | `<C-S-L>` | `MoveRight` | Ctrl + Shift + → | |
| Moving | `gg` | `MoveToTopRow` | 1行目 または `[count]` 行目へ移動 | ✓ |
| Moving | `G` | `MoveToLastRow` | UsedRange の最終行 または `[count]` 行目へ移動 | ✓ |
| Moving | `\|` | `MoveToNthColumn` | `[count]` 列目に移動 | ✓ |
| Moving | `0` | `MoveToFirstColumn` | 1列目に移動 | |
| Moving | `^` | `MoveToLeftEnd` | UsedRange の最初の列に移動 | |
| Moving | `$` | `MoveToRightEnd` | UsedRange の最後の列に移動 | |
| Moving | `g0` | `MoveToA1` | A1セルに移動 | |
| Moving | `{` | `MoveToTopOfCurrentRegion` | CurrentRegion 内で最初の行に移動 | |
| Moving | `}` | `MoveToBottomOfCurrentRegion` | CurrentRegion 内で最後の行に移動 | |
| Moving | `W[cell]` | `MoveToSpecifiedCell` | 指定された `[cell]` へ移動 | |
| Moving | `:[num]` | `MoveToSpecifiedRow` | 指定された `[num]` 行目に移動 | |
| Cell | `xx` | `CutCell` | セルを切り取り | |
| Cell | `yy` | `YankCell` | セルをコピー | |
| Cell | `o` | `InsertCellsDown` | 選択セルの下にセルを挿入 | ✓ |
| Cell | `O` | `InsertCellsUp` | 選択セルの上にセルを挿入 | ✓ |
| Cell | `t` | `InsertCellsRight` | 選択セルの右にセルを挿入 | ✓ |
| Cell | `T` | `InsertCellsLeft` | 選択セルの左にセルを挿入 | ✓ |
| Cell | `>` | `IncrementText` | インデントを増やす | ✓ |
| Cell | `<` | `DecrementText` | インデントを減らす | ✓ |
| Cell | `(` | `IncreaseDecimal` | 小数点表示桁上げ | ✓ |
| Cell | `)` | `DecreaseDecimal` | 小数点表示桁下げ | ✓ |
| Cell | `zw` | `ToggleWrapText` | セルの折り返しのオン/オフを切り替え | |
| Cell | `&` | `ToggleMergeCells` | セル結合のオン/オフを切り替え | |
| Cell | `f,` | `ApplyCommaStyle` | 桁区切りスタイルを適用 | |
| Cell | `<Space>` | `UnionSelectCells` | 現在セルを記憶に追加し、記憶したセルを選択 (複数セルの選択が可能) | |
| Cell | `<S-Space>` | `ExceptSelectCells` | 記憶された選択済みセルから現在セルを取り除く | |
| Cell | `<S-BS>` | `ClearSelectCells` | 記憶された選択済みセルをクリアする | |
| Cell | `gf` | `FollowHyperlinkOfActiveCell` | セルのハイパーリンクを開く | |
| Cell | `FF`/`Ff` | `ApplyFlashFill` | フラッシュフィル(適用不可の際はオートフィル) | |
| Cell | `FA`/`Fa` | `ApplyAutoFill` | オートフィル | |
| Cell | `=s` | `AutoSum` | オートSUM (合計) | |
| Cell | `=a` | `AutoAverage` | オートSUM (平均) | |
| Cell | `=c` | `AutoCount` | オートSUM (数値の個数) | |
| Cell | `=m` | `AutoMax` | オートSUM (最大値) | |
| Cell | `=i` | `AutoMin` | オートSUM (最小値) | |
| Cell | `==` | `InsertFunction` | 関数の挿入 | |
| Mode | `v` | `ToggleVisualMode` | ビジュアルモード(選択範囲の拡張)を切り替え | |
| Mode | `V` | `ToggleVisualLine` | ビジュアル行モード(選択範囲の拡張)を切り替え | |
| Mode | `<C-.>` | `SwapVisualBase` | 範囲指定の基準セルを入れ替える | |
| Border | `bb` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 細線) | |
| Border | `ba` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 細線) | |
| Border | `bh` | `ToggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 細線) | |
| Border | `bj` | `ToggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 細線) | |
| Border | `bk` | `ToggleBorderTop` | 選択セルの上に罫線を設定 (実線, 細線) | |
| Border | `bl` | `ToggleBorderRight` | 選択セルの右に罫線を設定 (実線, 細線) | |
| Border | `bia` | `ToggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 細線) | |
| Border | `bis` | `ToggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 細線) | |
| Border | `biv` | `ToggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 細線) | |
| Border | `b/` | `ToggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 細線) | |
| Border | `b\` | `ToggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 細線) | |
| Border | `bB` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 太線) | |
| Border | `bA` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 太線) | |
| Border | `bH` | `ToggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 太線) | |
| Border | `bJ` | `ToggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 太線) | |
| Border | `bK` | `ToggleBorderTop` | 選択セルの上に罫線を設定 (実線, 太線) | |
| Border | `bL` | `ToggleBorderRight` | 選択セルの右に罫線を設定 (実線, 太線) | |
| Border | `Bb` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 太線) | |
| Border | `Ba` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 太線) | |
| Border | `Bh` | `ToggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 太線) | |
| Border | `Bj` | `ToggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 太線) | |
| Border | `Bk` | `ToggleBorderTop` | 選択セルの上に罫線を設定 (実線, 太線) | |
| Border | `Bl` | `ToggleBorderRight` | 選択セルの右に罫線を設定 (実線, 太線) | |
| Border | `Bia` | `ToggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 太線) | |
| Border | `Bis` | `ToggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 太線) | |
| Border | `Biv` | `ToggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 太線) | |
| Border | `B/` | `ToggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 太線) | |
| Border | `B\` | `ToggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 太線) | |
| Border | `bob` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 極細線) | |
| Border | `boa` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 極細線) | |
| Border | `boh` | `ToggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 極細線) | |
| Border | `boj` | `ToggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 極細線) | |
| Border | `bok` | `ToggleBorderTop` | 選択セルの上に罫線を設定 (実線, 極細線) | |
| Border | `bol` | `ToggleBorderRight` | 選択セルの右に罫線を設定 (実線, 極細線) | |
| Border | `boia` | `ToggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 極細線) | |
| Border | `bois` | `ToggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 極細線) | |
| Border | `boiv` | `ToggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 極細線) | |
| Border | `bo/` | `ToggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 極細線) | |
| Border | `bo\` | `ToggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 極細線) | |
| Border | `bmb` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (実線, 中線) | |
| Border | `bma` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (実線, 中線) | |
| Border | `bmh` | `ToggleBorderLeft` | 選択セルの左に罫線を設定 (実線, 中線) | |
| Border | `bmj` | `ToggleBorderBottom` | 選択セルの下に罫線を設定 (実線, 中線) | |
| Border | `bmk` | `ToggleBorderTop` | 選択セルの上に罫線を設定 (実線, 中線) | |
| Border | `bml` | `ToggleBorderRight` | 選択セルの右に罫線を設定 (実線, 中線) | |
| Border | `bmia` | `ToggleBorderInner` | 選択セルの内側全てに罫線を設定 (実線, 中線) | |
| Border | `bmis` | `ToggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (実線, 中線) | |
| Border | `bmiv` | `ToggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (実線, 中線) | |
| Border | `bm/` | `ToggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (実線, 中線) | |
| Border | `bm\` | `ToggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (実線, 中線) | |
| Border | `btb` | `ToggleBorderAll` | 選択セルの周りと内側の全てに罫線を設定 (二重線, 太線) | |
| Border | `bta` | `ToggleBorderAround` | 選択セルの周りに罫線を設定 (二重線, 太線) | |
| Border | `bth` | `ToggleBorderLeft` | 選択セルの左に罫線を設定 (二重線, 太線) | |
| Border | `btj` | `ToggleBorderBottom` | 選択セルの下に罫線を設定 (二重線, 太線) | |
| Border | `btk` | `ToggleBorderTop` | 選択セルの上に罫線を設定 (二重線, 太線) | |
| Border | `btl` | `ToggleBorderRight` | 選択セルの右に罫線を設定 (二重線, 太線) | |
| Border | `btia` | `ToggleBorderInner` | 選択セルの内側全てに罫線を設定 (二重線, 太線) | |
| Border | `btis` | `ToggleBorderInnerHorizontal` | 選択セルの内側水平に罫線を設定 (二重線, 太線) | |
| Border | `btiv` | `ToggleBorderInnerVertical` | 選択セルの内側垂直に罫線を設定 (二重線, 太線) | |
| Border | `bt/` | `ToggleBorderDiagonalUp` | 選択セルに `/` 方向の罫線を設定 (二重線, 太線) | |
| Border | `bt\` | `ToggleBorderDiagonalDown` | 選択セルに `\` 方向の罫線を設定 (二重線, 太線) | |
| Border | `bdd` | `DeleteBorderAll` | 選択セルの周りと内側の全ての罫線を削除 | |
| Border | `bda` | `DeleteBorderAround` | 選択セルの周りの罫線を削除 | |
| Border | `bdh` | `DeleteBorderLeft` | 選択セルの左の罫線を削除 | |
| Border | `bdj` | `DeleteBorderBottom` | 選択セルの下の罫線を削除 | |
| Border | `bdk` | `DeleteBorderTop` | 選択セルの上の罫線を削除 | |
| Border | `bdl` | `DeleteBorderRight` | 選択セルの右の罫線を削除 | |
| Border | `bdia` | `DeleteBorderInner` | 選択セルの内側全ての罫線を削除 | |
| Border | `bdis` | `DeleteBorderInnerHorizontal` | 選択セルの内側水平の罫線を削除 | |
| Border | `bdiv` | `DeleteBorderInnerVertical` | 選択セルの内側垂直の罫線を削除 | |
| Border | `bd/` | `DeleteBorderDiagonalUp` | 選択セルの `/` 方向の罫線を削除 | |
| Border | `bd\` | `DeleteBorderDiagonalDown` | 選択セルの `\` 方向の罫線を削除 | |
| Border | `bcc` | `SetBorderColorAll` | 選択セルの周りと内側の全ての罫線の色を設定 | |
| Border | `bca` | `SetBorderColorAround` | 選択セルの周りの罫線の色を設定 | |
| Border | `bch` | `SetBorderColorLeft` | 選択セルの左の罫線の色を設定 | |
| Border | `bcj` | `SetBorderColorBottom` | 選択セルの下の罫線の色を設定 | |
| Border | `bck` | `SetBorderColorTop` | 選択セルの上の罫線の色を設定 | |
| Border | `bcl` | `SetBorderColorRight` | 選択セルの右の罫線の色を設定 | |
| Border | `bcia` | `SetBorderColorInner` | 選択セルの内側全ての罫線の色を設定 | |
| Border | `bcis` | `SetBorderColorInnerHorizontal` | 選択セルの内側水平の罫線の色を設定 | |
| Border | `bciv` | `SetBorderColorInnerVertical` | 選択セルの内側垂直の罫線の色を設定 | |
| Border | `bc/` | `SetBorderColorDiagonalUp` | 選択セルの `/` 方向の罫線の色を設定 | |
| Border | `bc\` | `SetBorderColorDiagonalDown` | 選択セルの `\` 方向の罫線の色を設定 | |
| Row | `r-` | `NarrowRowsHeight` | 行の高さを狭くする | ✓ |
| Row | `r+` | `WideRowsHeight` | 行の高さを広くする | ✓ |
| Row | `rr` | `SelectRows` | 行を選択 | ✓ |
| Row | `ra` | `AppendRows` | 現在行の後に行を挿入 | ✓ |
| Row | `ri` | `InsertRows` | 現在行の前に行を挿入 | ✓ |
| Row | `rd` | `DeleteRows` | 現在行を削除 | ✓ |
| Row | `ry` | `YankRows` | 現在行をコピー | ✓ |
| Row | `rx` | `CutRows` | 現在行を切り取り | ✓ |
| Row | `rh` | `HideRows` | 現在行を非表示化 | ✓ |
| Row | `rH` | `UnhideRows` | 現在行を再表示 | ✓ |
| Row | `rg` | `GroupRows` | 現在行をグループ化 | ✓ |
| Row | `ru` | `UngroupRows` | 現在行をグループ化解除 | ✓ |
| Row | `rf` | `FoldRowsGroup` | 現在行を畳む | ✓ |
| Row | `rs` | `SpreadRowsGroup` | 現在行の折り畳みを開く | ✓ |
| Row | `rj` | `AdjustRowsHeight` | 現在行の高さを自動調整 | ✓ |
| Row | `rw` | `SetRowsHeight` | 現在行の高さを任意に設定 | ✓ |
| Column | `c-` | `NarrowColumnsWidth` | 列幅を狭くする | ✓ |
| Column | `c+` | `WideColumnsWidth` | 列幅を広くする | ✓ |
| Column | `cc` | `SelectColumns` | 列を選択 | ✓ |
| Column | `ca` | `AppendColumns` | 現在列の後に列を挿入 | ✓ |
| Column | `ci` | `InsertColumns` | 現在列の前に列を挿入 | ✓ |
| Column | `cd` | `DeleteColumns` | 現在列を削除 | ✓ |
| Column | `cy` | `YankColumns` | 現在列をコピー | ✓ |
| Column | `cx` | `CutColumns` | 現在列を切り取り | ✓ |
| Column | `ch` | `HideColumns` | 現在列を非表示化 | ✓ |
| Column | `cH` | `UnhideColumns` | 現在列を再表示 | ✓ |
| Column | `cg` | `GroupColumns` | 現在列をグループ化 | ✓ |
| Column | `cu` | `UngroupColumns` | 現在列をグループ化解除 | ✓ |
| Column | `cf` | `FoldColumnsGroup` | 現在列を畳む | ✓ |
| Column | `cs` | `SpreadColumnsGroup` | 現在列の折り畳みを開く | ✓ |
| Column | `cj` | `AdjustColumnsWidth` | 現在列の幅を自動調整 | ✓ |
| Column | `cw` | `SetColumnsWidth` | 現在列の幅を任意に設定 | ✓ |
| Yank | `yr` | `YankRows` | 現在行をコピー | ✓ |
| Yank | `yc` | `YankColumns` | 現在列をコピー | ✓ |
| Yank | `ygg` | `YankToTopRows` | 現在行から1行目までをコピー | |
| Yank | `yG` | `YankToBottomRows` | 現在行から UsedRange の最終行までをコピー | |
| Yank | `y{` | `YankToTopOfCurrentRegionRows` | 現在行から CurrentRegion の最初の行までをコピー | |
| Yank | `y}` | `YankToBottomOfCurrentRegionRows` | 現在行から CurrentRegion の最後の行までをコピー | |
| Yank | `y0` | `YankToLeftEndColumns` | 現在列から UsedRange の最初の列までをコピー | |
| Yank | `y$` | `YankToRightEndColumns` | 現在列から UsedRange の最後の列までをコピー | |
| Yank | `y^` | `YankToLeftOfCurrentRegionColumns` | 現在列から CurrentRegion  の最初の列までをコピー | |
| Yank | `yg$` | `YankToRightOfCurrentRegionColumns` | 現在列から CurrentRegion の最後の列までをコピー | |
| Yank | `yh` | `YankFromLeftCell` | 現在のセルの左の値をコピーして貼り付け | |
| Yank | `yj` | `YankFromDownCell` | 現在のセルの下の値をコピーして貼り付け | |
| Yank | `yk` | `YankFromUpCell` | 現在のセルの上の値をコピーして貼り付け | |
| Yank | `yl` | `YankFromRightCell` | 現在のセルの右の値をコピーして貼り付け | |
| Yank | `Y` | `YankAsPlaintext` | 選択中のセルをプレーンテキストとしてコピー | |
| Delete | `D`/`X` | `DeleteValue` | セルの値を削除 | |
| Delete | `dx` | `DeleteRows` | 現在行を削除 | ✓ |
| Delete | `dr` | `DeleteRows` | 現在行を削除 | ✓ |
| Delete | `dc` | `DeleteColumns` | 現在列を削除 | ✓ |
| Delete | `dgg` | `DeleteToTopRows` | 現在行から先頭行までを削除 | |
| Delete | `dG` | `DeleteToBottomRows` | 現在行から UsedRange の最終行までを削除 | |
| Delete | `d{` | `DeleteToTopOfCurrentRegionRows` | 現在行から CurrentRegion の最初の行までを削除 | |
| Delete | `d}` | `DeleteToBottomOfCurrentRegionRows` | 現在行から CurrentRegion の最後の行までを削除 | |
| Delete | `d0` | `DeleteToLeftEndColumns` | 現在列から UsedRange の最初の列までを削除 | |
| Delete | `d$` | `DeleteToRightEndColumns` | 現在列から UsedRange の最後の列までを削除 | |
| Delete | `d^` | `DeleteToLeftOfCurrentRegionColumns` | 現在列から CurrentRegion  の最初の列までを削除 | |
| Delete | `dg$` | `DeleteToRightOfCurrentRegionColumns` | 現在列から CurrentRegion の最後の列までを削除 | |
| Delete | `dh` | `DeleteToLeft` | 現在のセルを削除し左方向へシフト | ✓ |
| Delete | `dj` | `DeleteToUp` | 現在のセルを削除し上方向へシフト | ✓ |
| Delete | `dk` | `DeleteToUp` | 現在のセルを削除し上方向へシフト | ✓ |
| Delete | `dl` | `DeleteToLeft` | 現在のセルを削除し左方向へシフト | ✓ |
| Cut | `xr` | `CutRows` | 現在行を切り取り | ✓ |
| Cut | `xc` | `CutColumns` | 現在列を切り取り | ✓ |
| Cut | `xgg` | `CutToTopRows` | 現在行から1行目までを切り取り | |
| Cut | `xG` | `CutToBottomRows` | 現在行から UsedRange の最後の行までを切り取り | |
| Cut | `x{` | `CutToTopOfCurrentRegionRows` | 現在行から CurrentRegion  の最初の列までを切り取り | |
| Cut | `x}` | `CutToBottomOfCurrentRegionRows` | 現在行から CurrentRegion の最後の行までを切り取り | |
| Cut | `x0` | `CutToLeftEndColumns` | 現在列から UsedRange の最初の列までを切り取り | |
| Cut | `x$` | `CutToRightEndColumns` | 現在列から UsedRange の最後の列までを切り取り | |
| Cut | `x^` | `CutToLeftOfCurrentRegionColumns` | 現在列から CurrentRegion  の最初の列までを切り取り | |
| Cut | `xg$` | `CutToRightOfCurrentRegionColumns` | 現在列から CurrentRegion の最後の列までを切り取り | |
| Paste | `p` | `PasteSmart` | 行や列がコピーされたときは後に追加。それ以外は `Ctrl + V` を送出 | ✓ |
| Paste | `P` | `PasteSmart` | 行や列がコピーされたときは前に挿入。それ以外は `Ctrl + V` を送出 | ✓ |
| Paste | `gp` | `PasteSpecial` | 形式を選択して貼り付けのダイアログを表示 | |
| Paste | `U` | `PasteValue` | 値のみ貼り付け | |
| Font | `-` | `DecreaseFontSize` | フォントサイズの縮小 | ✓ |
| Font | `+` | `IncreaseFontSize` | フォントサイズの拡大 | ✓ |
| Font | `fn` | `ChangeFontName` | フォント名にフォーカス | |
| Font | `fs` | `ChangeFontSize` | フォントサイズにフォーカス | |
| Font | `fh` | `AlignLeft` | 左揃え | |
| Font | `fj` | `AlignBottom` | 下揃え | |
| Font | `fk` | `AlignTop` | 上揃え | |
| Font | `fl` | `AlignRight` | 右揃え | |
| Font | `fo` | `AlignCenter` | 文字列中央揃え | |
| Font | `fm` | `AlignMiddle` | 上下中央揃え | |
| Font | `fb` | `ToggleBold` | 太字 | |
| Font | `fi` | `ToggleItalic` | 斜体 | |
| Font | `fu` | `ToggleUnderline` | 下線 | |
| Font | `f-` | `ToggleStrikethrough` | 取り消し線 | |
| Font | `ft` | `ChangeFormat` | 表示形式にフォーカス | |
| Font | `ff` | `ShowFontDialog` | セルの書式設定のダイアログを表示 | |
| Color | `fc` | `SmartFontColor` | フォントの色を選択するダイアログを表示 | |
| Color | `FC`/`Fc` | `SmartFillColor` | 塗りつぶしの色を選択すダイアログを表示 | |
| Color | `bc` | `ChangeShapeBorderColor` | (図形選択時) 枠線の色を選択するダイアログを表示 | |
| Comment | `Ci`/`Cc` | `EditCellComment` | コメントを編集 (ない場合は追加) | |
| Comment | `Ce`/`Cx`/`Cd` | `DeleteCellComment` | 現在セルのコメントを削除 | |
| Comment | `CE`/`CD` | `DeleteCellCommentAll` | シート上のコメントを全て削除 | |
| Comment | `Ca` | `ToggleCellComment` | 現在セルのコメントの表示/非表示を切り替え | |
| Comment | `Cr` | `ShowCellComment` | 現在セルのコメントを表示 | |
| Comment | `Cm` | `HideCellComment` | 現在セルのコメントを非表示 | |
| Comment | `CA` | `ToggleCellCommentAll` | すべてのコメントの表示/非表示を切り替え | |
| Comment | `CR` | `ShowCellCommentAll` | すべてのコメントを表示 | |
| Comment | `CM` | `HideCellCommentAll` | すべてのコメントを非表示 | |
| Comment | `CH` | `HideCellCommentIndicator` | 現在セルのコメントインジケータを非表示 | |
| Comment | `Cn` | `NextComment` | 次のコメントを選択 | ✓ |
| Comment | `Cp` | `PrevComment` | 前のコメントを選択 | ✓ |
| Find & Replace | `/` | `ShowFindFollowLang` | IMEを言語モードに合わせて検索ダイアログを表示 | |
| Find & Replace | `?` | `ShowFindNotFollowLang` | IMEを言語モードに合わせず検索ダイアログを表示 | |
| Find & Replace | `n` | `NextFoundCell` | 検索結果の次のセルを選択 | ✓ |
| Find & Replace | `N` | `PreviousFoundCell` | 検索結果の前のセルを選択 | ✓ |
| Find & Replace | `R` | `ShowReplaceWindow` | 検索と置換のダイアログを表示 | |
| Find & Replace | `*` | `FindActiveValueNext` | 現在セルの値で検索し次のセルを選択 | ✓ |
| Find & Replace | `#` | `FindActiveValuePrev` | 現在セルの値で検索し前のセルを選択 | ✓ |
| Find & Replace | `]c` | `NextSpecialCells` | 次のコメントがあるセルを選択 | ✓ |
| Find & Replace | `[c` | `PrevSpecialCells` | 前のコメントがあるセルを選択 | ✓ |
| Find & Replace | `]o` | `NextSpecialCells` | 次の定数があるセルを選択 | ✓ |
| Find & Replace | `[o` | `PrevSpecialCells` | 前の定数があるセルを選択 | ✓ |
| Find & Replace | `]f` | `NextSpecialCells` | 次の数式があるセルを選択 | ✓ |
| Find & Replace | `[f` | `PrevSpecialCells` | 前の数式があるセルを選択 | ✓ |
| Find & Replace | `]k` | `NextSpecialCells` | 次の空白セルを選択 | ✓ |
| Find & Replace | `[k` | `PrevSpecialCells` | 前の空白セルを選択 | ✓ |
| Find & Replace | `]t` | `NextSpecialCells` | 次の条件付き書式があるセルを選択 | ✓ |
| Find & Replace | `[t` | `PrevSpecialCells` | 前の条件付き書式があるセルを選択 | ✓ |
| Find & Replace | `]v` | `NextSpecialCells` | 次の入力規則があるセルを選択 | ✓ |
| Find & Replace | `[v` | `PrevSpecialCells` | 前の入力規則があるセルを選択 | ✓ |
| Find & Replace | `]s` | `NextShape` | 次の図形を選択 | ✓ |
| Find & Replace | `[s` | `PrevShape` | 前の図形を選択 | ✓ |
| Scrolling | `<C-u>` | `ScrollUpHalf` | 半ページ上スクロール | ✓ |
| Scrolling | `<C-d>` | `ScrollDownHalf` | 半ページ下スクロール | ✓ |
| Scrolling | `<C-b>` | `ScrollUp` | 1ページ上スクロール | ✓ |
| Scrolling | `<C-f>` | `ScrollDown` | 1ページ下スクロール | ✓ |
| Scrolling | `<C-y>` | `ScrollUp1Row` | 1行上スクロール | ✓ |
| Scrolling | `<C-e>` | `ScrollDown1Row` | 1行下スクロール | ✓ |
| Scrolling | `zh` | `ScrollLeft1Column` | 1列左スクロール | ✓ |
| Scrolling | `zl` | `ScrollRight1Column` | 1列右スクロール | ✓ |
| Scrolling | `zH` | `ScrollLeft` | 1ページ左スクロール | ✓ |
| Scrolling | `zL` | `ScrollRight` | 1ページ右スクロール | ✓ |
| Scrolling | `zt` | `ScrollCurrentTop` | 現在行が最上部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる) | ✓ |
| Scrolling | `zz` | `ScrollCurrentMiddle` | 現在行が中央に来るように縦スクロール | ✓ |
| Scrolling | `zb` | `ScrollCurrentBottom` | 現在行が最下部に来るように縦スクロール (`SCREEN_OFFSET` pt分余裕をもたせる) | ✓ |
| Scrolling | `zs` | `ScrollCurrentLeft` | 現在列が一番左に来るように横スクロール | ✓ |
| Scrolling | `zm` | `ScrollCurrentCenter` | 現在列が中央に来るように横スクロール | ✓ |
| Scrolling | `ze` | `ScrollCurrentRight` | 現在列が一番右に来るように横スクロール | ✓ |
| Worksheet | `e`/`wn` | `NextWorksheet` | 次のシートを選択 | ✓ |
| Worksheet | `E`/`wp` | `PreviousWorksheet` | 前のシートを選択 | ✓ |
| Worksheet | `ww`/`ws` | `ShowSheetPicker` | SheetPicker を起動 | |
| Worksheet | `wr` | `RenameWorksheet` | アクティブなシート名を変更 | |
| Worksheet | `wh` | `MoveWorksheetBack` | アクティブなシートを前に移動 | ✓ |
| Worksheet | `wl` | `MoveWorksheetForward` | アクティブなシートを次に移動 | ✓ |
| Worksheet | `wi` | `InsertWorksheet` | アクティブなシートの前に新しいシートを挿入 | |
| Worksheet | `wa` | `AppendWorksheet` | アクティブなシートの次に新しいシートを挿入 | |
| Worksheet | `wd` | `DeleteWorksheet` | アクティブなシートを削除 | |
| Worksheet | `w0`/`w$` | `ActivateLastWorksheet` | 一番最後のシートを選択 | |
| Worksheet | `wc` | `ChangeWorksheetTabColor` | アクティブなシートの色を変更 | |
| Worksheet | `wy` | `CloneWorksheet` | アクティブなシートを複製 | |
| Worksheet | `we` | `ExportWorksheet` | シートの移動またはコピーダイアログを表示 | |
| Worksheet | `w[num]` | `ActivateWorksheet` | `[num]` 番目のシートを選択 (1-9 のみ) | |
| Worksheet | `:printpreview` | `PrintPreviewOfActiveSheet` | 印刷プレビューを表示 | |
| Workbook | `:e` | `OpenWorkbook` | ブックを開く | |
| Workbook | `:e!` | `ReopenActiveWorkbook` | アクティブなブックの変更を破棄し開き直す | |
| Workbook | `:w` | `SaveWorkbook` | アクティブブックを保存 | |
| Workbook | `:q` | `CloseAskSaving` | アクティブブックを閉じる(未保存時はダイアログを表示) | |
| Workbook | `:q!`/`ZQ` | `CloseWithoutSaving` | アクティブブックを保存せずに閉じる | |
| Workbook | `:wq`/`:x`/`ZZ` | `CloseWithSaving` | アクティブブックを保存して閉じる | |
| Workbook | `:b [num]` | `ActivateWorkbook` | `[num]` 番目のブックを選択 | |
| Workbook | `]b`/`:bn` | `NextWorkbook` | 次のワークブックを選択 | ✓ |
| Workbook | `[b`/`:bp` | `PreviousWorkbook` | 前のワークブックを選択 | ✓ |
| Workbook | `~` | `ToggleReadOnly` | 読み取り専用を切り替える | |
| Other | `u` | `Undo_CtrlZ` | 元に戻す (`Ctrl + Z` を送出) | |
| Other | `<C-r>` | `RedoExecute` | やり直し | |
| Other | `.` | `RepeatAction` | 以前の動作を繰り返す (`RepeatRegister` が呼ばれるコマンド限定) | |
| Other | `m` | `ZoomIn` | 10% または `[count]`% ズームイン | ✓ |
| Other | `M` | `ZoomOut` | 10% または `[count]`% ズームアウト | ✓ |
| Other | `%` | `ZoomSpecifiedScale` | 表示倍率を `[count]`% に設定。`1`-`9` は決まった値に倍率変更 | ✓ |
| Other | `\` | `ShowContextMenu` | コンテキストメニューを表示 | |
| Other | `<C-i>` | `JumpNext` | ジャンプリストの次のセルへ移動 | ✓ |
| Other | `<C-o>` | `JumpPrev` | ジャンプリストの前のセルへ移動 | ✓ |
| Other | `:cle` | `ClearJumps` | ジャンプリストをクリア | |
| Other | `zf` | `ToggleFreezePanes` | ウィンドウ枠の固定のオン/オフを切り替え | |
| Other | `=v` | `ToggleFormulaBar` | 関数バーの表示/非表示を切り替え | |
| Other | `gs` | `ShowSummaryInfo` | ファイルのプロパティを表示 | |
| Other | `zp` | `SetPrintArea` | 選択セルを印刷範囲に設定 | |
| Other | `zP` | `ClearPrintArea` | 印刷範囲をクリア | |
| Other | `@@` | `ShowMacroDialog` | マクロダイアログを表示 | |
| Other | `1-9` | `ShowCmdForm` | `[count]` を指定 (Count に ✓ がついている機能で有効) | |

</div></details>

\* デフォルトの設定は [DefaultConfig.bas](./src/DefaultConfig.bas) の中で `Map` メソッドを使って定義しています。

### Custom Key Mapping

- `<C-[>` → `<Esc>`

## Customization

[設定ファイル](./config/_vimxlamrc)を vim.xlam が保存されているディレクトリに置くことで、起動時に設定を読み込むことができます。読み込むファイル名は `_vimxlamrc` のみです。ファイルエンコーディングは Shift-JIS で保存してください。

### Syntax

- `#` で始まる行や空行は無視されます
- `set` で始まる行は、定義された設定値を変更できます
- `map` または `unmap` を含む行は、キーマップの設定を変更できます

### Options

Vim の `set` と同じシンタックスで設定できます。設定例は[設定ファイル](./config/_vimxlamrc)をご覧ください。

| Option Key | Type | Description | Default |
| ---------- | ---- | ----------- | ------- |
| `statusprefix` | string | ステータスバーに表示される一時的なメッセージのプレフィックス | `vim.xlam: ` |
| `togglekey` | string | Vimモードの有効/無効を切り替えるキー (Vim風のキー指定) | `<C-m>` |
| `scrolloff` | float | `ScrollCurrentXXX` 系の上下オフセット量 (px) | `54.0` |
| `jumplisthistory` | int | ジャンプリストの最大保持数 | `100` |
| `[no]japanese` | bool | 日本語モード / 英語モード | `True` |
| `[no]jiskeyboard` | bool | JISキーボード / USキーボード | `True` |
| `colorpickersize` | float | ColorPicker のフォームサイズ (px) | `12.0` |
| `customcolor1` | string | ColorPicker のカスタム色 #1 | `#ff6600` ![#ff6600](https://via.placeholder.com/15/ff6600/000000?text=+) |
| `customcolor2` | string | ColorPicker のカスタム色 #2 | `#ff9966` ![#ff9966](https://via.placeholder.com/15/ff9966/000000?text=+) |
| `customcolor3` | string | ColorPicker のカスタム色 #3 | `#ff00ff` ![#ff00ff](https://via.placeholder.com/15/ff00ff/000000?text=+) |
| `customcolor4` | string | ColorPicker のカスタム色 #4 | `#008000` ![#008000](https://via.placeholder.com/15/008000/000000?text=+) |
| `customcolor5` | string | ColorPicker のカスタム色 #5 | `#0000ff` ![#0000ff](https://via.placeholder.com/15/0000ff/000000?text=+) |
| `[no]debug` | bool | デバッグモードの有効 / 無効 | `False` |

### Keymap

キーマップの追加/変更/削除が可能です。

- `{lhs}` には本家 Vim のキーマップ指定
  - `<cmd>` を指定した場合はコマンドモードの指定となり、単純な文字列として扱われる
- `{rhs}` には実行したい関数名を指定
  - `<key>` を指定した場合は別のキーを `{lhs}` と同様に Vim のキーマップを指定

モードは現在4種類あります。

- `n` (Normal): 通常モード。基本的には常にモードでマッピングする。
- `v` (Visual): 範囲選択モード。Normalモードに戻るキーなどを指定。
- `c` (Cmdline): コマンドラインモード。`:` や `/` などのコマンドラインモードで有効。
- `i` (Shape_Insert): 図形挿入モード。図形選択中に `i`/`a` などを押した場合などで有効。

**`map` や `unmap` のシンタックス**

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

\* [DefaultConfig.bas](./src/DefaultConfig.bas) 内に記載されているものと同じシンタックスになりますので、必要に応じて参考にしてください。

## Contributing

[Issue](https://github.com/sha5010/vim.xlam/issues) や [Pull Request](https://github.com/sha5010/vim.xlam/pulls) は大歓迎です。もしご自身で開発された機能がありましたら、開発にご協力いただけますと幸いです。

## Author

[@sha_5010](https://twitter.com/sha_5010)

## Related Projects

- [ExcelLikeVim](https://github.com/kjnh10/ExcelLikeVim)
- [VimExcel](https://www.vector.co.jp/soft/winnt/business/se494158.html)
- [vixcel](https://github.com/codetsar/vixcel)
- [Excel\_Vim\_Keys](https://github.com/treatmesubj/Excel_Vim_Keys)

## License

[MIT](./LICENSE)
