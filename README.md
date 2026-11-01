# xlsx-editor.html

ブラウザだけで動く、Excel 形式（.xlsx / .xls）のスプレッドシートエディタです。
**HTML ファイル 1 つ**で完結し、インストール・サーバー・インターネット接続は不要です。ファイルの中身が外部に送信されることもありません。

- 値・数式の編集と、ブラウザ上での再計算
- グラフの作成（9 種類）、画像の挿入
- 既存の xlsx に含まれる画像・グラフを保持したまま保存

## 使い方

1. `xlsx-editor.html` をダウンロードし、ダブルクリックしてブラウザで開きます。
2. 「開く」ボタン、またはファイルのドラッグ＆ドロップで .xlsx / .xlsm / .xls / .csv を読み込みます（「新規」で空のブックからも始められます）。
3. 編集して、`Ctrl+S` で保存します。

### 対応ブラウザ

| ブラウザ | 動作 | 保存 |
|---|---|---|
| Chrome / Edge（推奨） | ○ | 元のファイルへの上書き保存、保存先の選択に対応 |
| Firefox / Safari | ○ | 保存のたびにダウンロード |
| Android（Chrome） | ○ | 保存のたびにダウンロード（「ダウンロード」フォルダに保存） |
| iPhone / iPad（Safari） | ○ | 保存のたびにダウンロード（「ファイル」アプリの「ダウンロード」に保存） |

### スマートフォン・タブレットで使う

スマートフォンでは、ダウンロードした HTML ファイルをそのまま開けないことがあります（iPhone の「ファイル」アプリはプレビュー表示になり、Android の Chrome は端末内の HTML を開けない場合があります）。
**GitHub Pages などで公開した URL をブラウザで開く**のが確実です。GitHub Pages の場合、リポジトリの Settings → Pages で公開すると `https://<ユーザー名>.github.io/<リポジトリ名>/xlsx-editor.html` で開けます。

ページを開いたあとは、ファイルの読み込み・編集・保存はすべて端末の中で行われ、ファイルの中身が送信されることはありません。

| 操作 | 方法 |
|---|---|
| ファイルを開く | 「開く」→ 端末内や「ファイル」アプリ、Google ドライブなどから選ぶ |
| セルを選ぶ | タップ |
| セルを編集 | 選んだセルをもう一度タップ（または数式バーをタップ）。確定は改行キー |
| 範囲を選ぶ | 「範囲選択」ボタンを押してから、範囲の端のセルをタップ（行・列の番号のタップで行・列全体） |
| コピー・貼り付け・行列の挿入削除など | セルを長押ししてメニューから選ぶ |
| シート名の変更 | シートのタブを長押し |
| 画像・グラフの移動 | タップで選んでからドラッグ。右下の ○ でサイズ変更 |
| グラフの編集・画像やグラフの削除 | 長押ししてメニューから選ぶ |
| 保存 | 「保存」または「xlsx で保存」でダウンロード |

## 機能

### セルと数式

- セルに直接入力（日本語入力に対応）、ダブルクリックや `F2` で修正、数式バーでも編集できます
- 数式は編集のたびに、影響するセルだけをブラウザ上で再計算します（ほかのシートへの参照にも対応）
- ファイルを開いた直後は、Excel が保存した計算結果をそのまま表示します。「再計算」（`F9`）ですべて計算し直します
- 範囲選択（ドラッグ、Shift+クリック、Shift+矢印、行・列の見出しのクリック）
- 範囲のコピー・貼り付け（数式は相対参照をずらして貼り付け。Excel からの複数セルの貼り付けにも対応）
- 下方向・右方向へのコピー（`Ctrl+D` / `Ctrl+R`）、オート SUM（`Alt+=`）
- 右クリックで行・列の挿入と削除（数式やグラフの参照も合わせて調整）
- 範囲を選ぶと、ステータスバーに合計・平均・データの個数を表示
- シートの切り替えと追加
- シート名の変更（タブをダブルクリック、または右クリック →「シート名の変更」）。ほかのセルの数式やグラフがそのシートを参照していれば、参照先のシート名も合わせて書き換えます

### 対応している関数

| 分類 | 関数 |
|---|---|
| 数学・集計 | SUM, PRODUCT, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, SUMPRODUCT, ROUND, ROUNDUP, ROUNDDOWN, TRUNC, INT, ABS, MOD, POWER, SQRT, EXP, LN, LOG, LOG10, PI, SIGN, CEILING, FLOOR, RAND |
| 統計 | MEDIAN, STDEV(.S/.P), VAR(.S/.P), LARGE, SMALL, RANK(.EQ) |
| 条件付き集計 | SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS |
| 論理 | IF, IFS, IFERROR, IFNA, AND, OR, NOT, SWITCH, CHOOSE, ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA |
| 検索 | VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, ROW, COLUMN, ROWS, COLUMNS |
| 文字列 | LEN, LEFT, RIGHT, MID, UPPER, LOWER, PROPER, TRIM, CONCAT, CONCATENATE, TEXTJOIN, SUBSTITUTE, FIND, SEARCH, REPT, VALUE, TEXT, EXACT |
| 日付・時刻 | TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, TIME, HOUR, MINUTE, SECOND, EDATE, EOMONTH, DATEDIF |

未対応の関数を含むセルは、ファイルに保存されていた値をそのまま表示します（点線の下線で表示し、ステータスバーでもお知らせします）。

### グラフ

- データ範囲を選んで「グラフ」ボタン（1 セルだけ選んでいる場合は、周りのデータの塊を自動で範囲にします）
- 縦棒・積み上げ縦棒・横棒・折れ線・面・円・ドーナツ・散布図・レーダー
- セルの値を変えるとすぐに更新されます。ダブルクリックで種類・タイトル・範囲を編集できます
- xlsx に保存すると、Excel のネイティブなグラフ（セル範囲を参照するグラフ）として書き出されます

### 画像

- 「画像」ボタン、ドラッグ＆ドロップ、`Ctrl+V` で挿入できます（PNG / JPEG / GIF / BMP。WebP などは自動で PNG に変換）
- 画像・グラフはドラッグで移動、選択すると右下に出る □ でサイズ変更、`Delete` で削除できます
- 既存の xlsx に含まれる画像・グラフも表示され、保存時は元の書式のまま残ります

### 保存

| 操作 | 内容 |
|---|---|
| `Ctrl+S` / 「保存」 | 開いたファイルに上書き。新規の場合は保存先を選んだあと、以降は上書き |
| `Ctrl+Shift+S` / 「xlsx で保存」 | 名前を付けて保存 |
| 「xls で保存」 | Excel 97-2003 形式で保存（下の制限を参照） |

## キーボードショートカット

| キー | 動作 |
|---|---|
| 矢印 / Enter / Tab | セルの移動 |
| Shift + 矢印 | 範囲選択 |
| Ctrl + 矢印 | データの端へ移動 |
| Ctrl + A | 使用範囲をすべて選択 |
| F2 / ダブルクリック | セルの編集 |
| シートタブをダブルクリック | シート名の変更（Enter で確定、Esc で取り消し） |
| Delete | セル（範囲）や、選択中の画像・グラフを削除 |
| Ctrl + C / X / V | コピー / 切り取り / 貼り付け |
| Ctrl + D / Ctrl + R | 下方向 / 右方向へコピー |
| Alt + = | オート SUM |
| Ctrl + Z / Ctrl + Y | 元に戻す / やり直し |
| F9 | すべての数式を再計算 |
| Ctrl + S / Ctrl + Shift + S | 保存 / 名前を付けて保存 |
| Ctrl + O | ファイルを開く |

## 制限事項

- **xls 形式で保存すると、数式・画像・グラフは保存されません**（数式は計算結果の値だけが残ります）。保存前に確認メッセージが表示されます。
- **セルの書式（色・罫線・フォントなど）は、保存すると失われます。**
- Excel で作成されたグラフは、移動・サイズ変更・削除のみできます（内容の編集はできません）。ブラウザ上の表示は近い見た目で、Excel と完全には同じになりません。また、行・列を挿入・削除しても、これらのグラフの参照範囲は調整されません。
- セルの結合の表示、図形・SmartArt の表示、条件付き書式、ピボットテーブル、マクロの実行には対応していません（図形などは保存時にそのまま残ります）。
- 配列数式・名前付き範囲・INDIRECT / OFFSET などは計算できません（ファイルに保存されていた値を表示します）。

- スマートフォンでは、開いたファイルへの上書き保存はできません（保存のたびに新しいファイルとしてダウンロードされます）。

大切なファイルは、編集前にコピーを取ってからお使いください。

## ライセンス

本体のコード（ライブラリ以外の部分）は MIT License です。

```
MIT License

Copyright (c) 2026 OkiaYoshisada

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### 埋め込んでいるライブラリ

| ライブラリ | バージョン | ライセンス |
|---|---|---|
| [SheetJS Community Edition](https://sheetjs.com) | 0.20.3 | Apache License 2.0 |
| [Chart.js](https://www.chartjs.org) | 4.4.1 | MIT License |
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | MIT License（MIT / GPLv3 のデュアルライセンスから MIT を選択） |

各ライブラリのライセンス全文は [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md) を参照してください。

Microsoft Excel は Microsoft Corporation の商標です。本ソフトウェアは Microsoft とは関係ありません。



### 最後に
小生、ITエンジニアを生業としております。
AIに作らせたこのツールを、気に入っていただけましたらDMでお仕事いただけると幸いです。