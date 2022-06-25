<p align="center">
  <a href="/sha5010/vim.xlam/releases/latest">
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

Under construction...

## Customization

Under construction...

## Contributing

[Issue](https://github.com/sha5010/vim.xlam/issues) や [Pull Request](https://github.com/sha5010/vim.xlam/pulls) は大歓迎です。もしご自身で開発された機能がありましたら、開発にご協力いただけますと幸いです。

## Author

[@sha_5010](https://twitter.com/sha_5010)

## License

[MIT](./LICENSE)
