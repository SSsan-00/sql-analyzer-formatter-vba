# SQL Analysis Formatter

Excel マクロで SQL の識別子を対応表の和名に変換するツールです。

## 機能

- `変換定義` シートの 2 行目以降を対応表として読み込みます。
- `SQL解析` シートの A 列 2 行目以降に入力された SQL を確認します。
- 変換後の SQL を `SQL解析` シートの B 列へ出力します。
- 変換が発生した行は、変換後の内容を `SQL解析` シートの C 列以降へ 1 件ずつ出力します。
- `アウトプット` シートへ、和名変換後クエリを AST 解析した定義表を出力します。
- `解析` ボタンで変換処理を実行します。
- `クリア` ボタンで確認後、`変換定義` と `SQL解析` の 2 行目以降、`アウトプット` の内容をクリアし、1 行目のヘッダーを復元します。

## シート構成

### 変換定義

| 列 | 項目 |
| --- | --- |
| A | 所属テーブルID |
| B | 所属テーブル和名 |
| C | フィールドID |
| D | フィールド和名 |

例: `tb1` / `ユーザー` / `name` / `名前`

`tb1.name` のようなテーブル修飾付き識別子は `tb1.名前` に変換されます。
所属テーブルIDは変換せず、フィールドIDだけを和名に変換します。
A 列が `-` の行は、単独のフィールド ID として和名に変換します。
`AS` 直後の単独 ID はエイリアスとして扱い、変換しません。
`(和名未取得)` や空欄は変換対象外です。

### SQL解析

| 列 | 項目 |
| --- | --- |
| A | SQLクエリ |
| B | 和名変換後クエリ |
| C 以降 | 変換内容 |

SQL は A 列の 2 行目以降へ、1 行ずつ入力します。
解析対象SQLは、A5:SQL Mk-2（A5M2）の `Ctrl+q` で整形した結果を貼り付ける前提です。
改行とインデントを含むSQLも、1つのクエリとして同じセルに貼り付けて扱います。
同じ行で複数の変換が発生した場合、C 列、D 列、E 列のように右方向へ変換後の内容だけを出力します。

### アウトプット

解析結果をもとにした別フォーマット出力用のシートです。
SELECT 全体は `＜DB入出力項目定義＞`、サブクエリは同じ表形式で `サブクエリ[名前]`、INSERT SELECT・DELETE・UPDATE は `＜データ移送表＞` として A 列から CL 列へ出力します。
WITH 句は名前付きサブクエリとして扱い、ネストしたサブクエリは内側から外側、最後にクエリ全体の順で出力します。
無名サブクエリには出力順に `SQ1`、`SQ2` の名前を付けます。
各表には参照テーブル、取得項目、条件、結合、並べ替えなど、文法に応じた解析結果を配置します。
アウトプットは MS ゴシック 9 ポイント、列幅 1.14、行高 13.5、目盛り線なしで整形します。
未対応のクエリは分解せず、和名変換後クエリをそのまま出力します。

## 導入手順

### マクロ付きブックを使う場合

1. `SqlAnalysisFormatter.xlsm` をダウンロードします。
2. `SqlAnalysisFormatter.Parser.exe` をブックと同じフォルダへ配置します。
3. Excel で開き、必要に応じてマクロを有効化します。
4. `変換定義` シートに変換定義を入力します。
5. `SQL解析` シートの A 列 2 行目以降に SQL を入力します。
6. `解析` ボタンを押して B 列と C 列以降、`アウトプット` シートの出力を確認します。

確定フォーマットの生成には、ScriptDom を内包した `SqlAnalysisFormatter.Parser.exe` を使います。
exe が存在しない場合や実行に失敗した場合も SQL 解析は継続し、`アウトプット` には和名変換後クエリをそのまま出力します。

### ローカルBootstrapを使う場合

GitHubへ通信せずに導入するための貼り付け用 bootstrap は、開発者が生成します。
ユーザー向け bootstrap は、成果物フォルダへ `SqlAnalysisFormatter.bas`、`SqlAnalysisFormatter.Parser.exe`、利用者向け `README.md` だけを展開します。
開発者向け bootstrap は、それに加えてプロダクションコード、テストコード、開発者向けドキュメントを展開します。
展開後の成果物フォルダには bootstrap 自身のソースを含めません。

### `.bas` から導入する場合

1. Excel で対象ブックを開き、`Alt + F11` で VBA エディターを開きます。
2. `ファイル` -> `ファイルのインポート` から `src/vba/SqlAnalysisFormatter.bas` を取り込みます。
3. VBA エディターで `SetupWorkbook` を実行します。
4. ブックを `.xlsm` 形式で保存します。

利用者が取り込む `.bas` は `src/vba/SqlAnalysisFormatter.bas` だけです。
`src/vba/SqlAnalysisFormatterTests.bas` は開発者用のテストモジュールなので、通常利用時は取り込みません。

## 開発者向け

VBA ソースは `src/vba/SqlAnalysisFormatter.bas` で管理しています。
Excel ブックへ反映する場合は、既存の `SqlAnalysisFormatter` モジュールを削除してから再インポートしてください。
配布用ブックにはテストモジュールを含めず、本体モジュールだけを入れます。

### テスト

開発時は変更後に必ずテストを実行します。
テスト用VBAソースは `src/vba/SqlAnalysisFormatterTests.bas` です。
Rubberduck VBA のテスト注釈を付けています。
Rubberduckを使わない場合は、次のスクリプトを実行します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1
```

parser exe 経由も含めて確認する場合は、先に parser を publish してから次を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1 -ParserExePath dist/parser/SqlAnalysisFormatter.Parser.exe
powershell -ExecutionPolicy Bypass -File tools/run-output-golden-tests.ps1
```

このスクリプトは一時コピーしたブックに本体モジュールとテストモジュールを取り込み、`RunAllSqlAnalysisFormatterTests` を実行します。
保存済みの `SqlAnalysisFormatter.xlsm` にはテストモジュールを残しません。

CRUDテストケースの内容は `tests/CRUD_TEST_CASES.md`、確定済みの47出力ケースは `tests/OutputReportCases.json` と `tests/SqlAnalysisFormatter.OutputExpectations.xlsx` にまとめています。

### C# parser

T-SQL の AST 解析には `tools/SqlAnalysisFormatter.Parser` を使います。
対象フレームワークは `.NET 8.0` です。
問題が出た場合のみ、`.NET 9.0` 以降への更新を検討します。

```powershell
dotnet test SqlAnalysisFormatter.sln
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
```

publish 結果は `dist/parser/SqlAnalysisFormatter.Parser.exe` です。
単一 exe、self-contained、win-x64 として出力します。

### Bootstrap生成

貼り付け用 bootstrap は次のコマンドで生成します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience User
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience Developer
powershell -ExecutionPolicy Bypass -File tools/test-bootstrap.ps1
```

生成先は `dist/bootstrap/SqlAnalysisFormatter.user.bootstrap.ps1` と `dist/bootstrap/SqlAnalysisFormatter.developer.bootstrap.ps1` です。
この生成物はサイズが大きいため、通常のソース管理対象には含めません。
