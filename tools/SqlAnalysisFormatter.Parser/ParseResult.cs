namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// SQL解析結果
/// </summary>
public sealed record ParseResult(
    bool Ok,
    string QueryType,
    IReadOnlyList<QueryBlock> Blocks,
    IReadOnlyList<string> ParseErrors);

/// <summary>
/// アウトプットへ出力するクエリ断片
/// </summary>
public sealed record QueryBlock(
    string Kind,
    string Text,
    int StartOffset,
    int Length);
