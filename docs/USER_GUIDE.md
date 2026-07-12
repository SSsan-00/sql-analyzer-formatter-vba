# SQL Analysis Formatter 利用ガイド

Excel 上で SQL の識別子を和名へ変換し、解析結果を `SQL解析` シートと `アウトプット` シートへ出力するマクロです。

## 同梱ファイル

- `SqlAnalysisFormatter.bas`
- `SqlAnalysisFormatter.Parser.exe`
- `README.md`

## 導入手順

1. Excel で利用先ブックを開きます。
2. `Alt + F11` で VBA エディターを開きます。
3. `ファイル` -> `ファイルのインポート` から `SqlAnalysisFormatter.bas` を取り込みます。
4. VBA エディターで `SetupWorkbook` を実行します。
5. ブックを `.xlsm` 形式で保存します。
6. 保存したブックと同じフォルダへ `SqlAnalysisFormatter.Parser.exe` を置きます。

`SqlAnalysisFormatter.Parser.exe` が見つからない場合でも、VBA 内蔵の暫定解析で動作します。

## 利用手順

1. `変換定義` シートに変換定義を入力します。
2. `SQL解析` シートの A 列 2 行目以降に SQL を入力します。
3. `解析` ボタンを押します。
4. `SQL解析` シートの B 列と C 列以降、`アウトプット` シートを確認します。

SQL は A5:SQL Mk-2（A5M2）の `Ctrl+q` で整形した結果を貼り付ける前提です。

## クリア

`クリア` ボタンを押すと確認ダイアログを表示します。
承認すると、`変換定義` と `SQL解析` の 2 行目以降、`アウトプット` の内容をクリアします。
