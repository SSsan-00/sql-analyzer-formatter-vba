# SQL Analysis Formatter 開発ガイド

## 構成

- `src/vba/SqlAnalysisFormatter.bas`: 利用者へ配布する VBA 本体
- `src/vba/SqlAnalysisFormatterTests.bas`: 開発者用 VBA テスト
- `docs/USER_GUIDE.md`: ユーザー向けBootstrapで`README.md`として配布する利用ガイド
- `tools/SqlAnalysisFormatter.Parser`: ScriptDom を使う C# parser
- `tests/SqlAnalysisFormatter.Parser.Tests`: MSTest による C# テスト
- `tests/CRUD_TEST_CASES.md`: SQL 変換ケース資料
- `tests/OutputReportCases.json`: 確定済み47ケースの入力 SQL と和名定義
- `tests/SqlAnalysisFormatter.OutputExpectations.xlsx`: 確定済み47ケースの期待値ブック
- `tools/run-output-golden-tests.ps1`: 実 Excel による値・書式回帰テスト

## テスト

変更後は次を実行します。

```powershell
dotnet test SqlAnalysisFormatter.sln
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1
```

parser exe 経由も確認する場合は、先に publish します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1 -ParserExePath dist/parser/SqlAnalysisFormatter.Parser.exe
powershell -ExecutionPolicy Bypass -File tools/run-output-golden-tests.ps1
```

`run-output-golden-tests.ps1` は47ケースについてセル値、主要罫線、塗り、フォント、折り返し、行高、列幅、目盛り線を実 Excel で比較します。
機能追加は、失敗するテストを先に追加し、最小実装で成功させ、全回帰テストを維持したまま整理する TDD サイクルで進めます。

`test-bootstrap.ps1`は、ユーザー向けREADMEに導入、初回セットアップ、トラブル対応の各セクションが含まれることも確認します。
利用手順を変更した場合は、`README.md`と`docs/USER_GUIDE.md`を同時に更新します。

## Publish

parser は .NET 8.0、win-x64、self-contained、単一 exe として publish します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
```

## Bootstrap

bootstrap 生成は利用者向けと開発者向けを分けます。

```powershell
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience User
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience Developer
powershell -ExecutionPolicy Bypass -File tools/test-bootstrap.ps1
```

生成済み bootstrap は `dist/bootstrap` に出力します。
生成物はサイズが大きいためソース管理に含めません。
