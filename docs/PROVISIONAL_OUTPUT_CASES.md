# 暫定実装ケース

この資料は、ユーザーレビュー前の推測期待値または現行実装出力を管理する。
2026-07-20時点で、登録済み81ケースはすべて最終レビュー済みであり、SEL-082を追加レビュー中である。

## 共通ルール

- 更新系SQLにSELECT処理が含まれる場合、SELECT表を先、最終的な`＜データ移送表＞`を後に出力する。
- SELECTがネストする場合、内側から外側へ出力する。
- 無名SELECTには`SQ1`、`SQ2`を付け、派生テーブルやCTEの既存名は維持する。
- SELECTを含まないINSERT、UPDATE、DELETEは`＜データ移送表＞`だけを出力する。

## レビュー対象

| ケース | SQL概要 | 現行実装の出力 |
| --- | --- | --- |
| SEL-082 | 複合条件を持つTHEN・ELSE両側のネストCASE | 最上位、THEN側、ELSE側の3つのCASEについて、括弧を保持しながらAND、OR、末端条件を2列ずつ階層表示する |

SEL-082は最上位CASEのTHEN側とELSE側に1つずつCASEを置いた合計3CASEの構造である。3つのWHENをAND、OR、括弧が混在する複合条件にしている。`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`のSEL-082シートには、仕様確定前の基準として現行実装の解析結果を保存している。

## 現在の制約

- INSERT VALUESは対象列を明示した単一行と複数行を表形式へ変換する。複数行は`＜VALUES n行目＞`ごとの移送表へ分ける。
- DEFAULT VALUES、INSERT EXECUTEは原因付きフォールバックとする。
- INSERT SELECTのUNIONとUNION ALLは各SELECTを移送パターンへ分けて出力する。EXCEPTとINTERSECTを直接ソースにする場合は原因付きフォールバックとする。

## レビュー手順

新しい暫定ケースを追加した場合は、次のコマンドで元SQLと和名定義を開発用ブックへ投入する。

```powershell
powershell -ExecutionPolicy Bypass -File tools/Set-ManualOutputCase.ps1 -CaseId SEL-082
```

`解析`を実行し、`アウトプット`シートを確認する。暫定期待値と異なる場合は、従来どおりシートへ正しい期待値を記入する。
レビュー確定後は、`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`、`ManualOutputCases.json`の`review_status`、`OutputReportCases.json`への登録、本資料を同じ変更で更新する。
