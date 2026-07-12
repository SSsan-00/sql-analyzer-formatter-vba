namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// 解析結果の出力形式変換
/// </summary>
public static class ParserOutputFormatter
{
    public const char VbaBlockSeparator = '\u001e';

    /// <summary>
    /// VBAがSplitで扱えるブロック区切り文字列へ変換
    /// </summary>
    public static string ToVbaBlocks(ParseResult result)
    {
        return string.Join(VbaBlockSeparator, result.Blocks.Select(block => block.Text));
    }
}
