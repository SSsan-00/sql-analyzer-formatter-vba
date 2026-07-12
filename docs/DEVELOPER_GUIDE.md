# SQL Analysis Formatter 開発ガイド

## 構成

- `src/vba/SqlAnalysisFormatter.bas`: 利用者へ配布する VBA 本体
- `src/vba/SqlAnalysisFormatterTests.bas`: 開発者用 VBA テスト
- `tools/SqlAnalysisFormatter.Parser`: ScriptDom を使う C# parser
- `tests/SqlAnalysisFormatter.Parser.Tests`: MSTest による C# テスト
- `tests/CRUD_TEST_CASES.md`: SQL 変換ケース資料

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
```

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
