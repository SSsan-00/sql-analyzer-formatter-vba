using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// VBAから呼び出すSQL解析CLI
/// </summary>
internal static class Program
{
    private const string VersionText = "SqlAnalysisFormatter.Parser 0.1.0";

    /// <summary>
    /// SQLを読み取り、解析結果JSONを標準出力へ書き出す
    /// </summary>
    private static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        if (args.Length == 1 && args[0] == "--version")
        {
            Console.WriteLine(VersionText);
            return 0;
        }

        string sql;
        try
        {
            sql = await ReadSqlAsync(args);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }

        var result = TsqlAstParser.Parse(sql);
        var json = JsonSerializer.Serialize(result, JsonOptions());
        Console.WriteLine(json);
        return 0;
    }

    private static async Task<string> ReadSqlAsync(string[] args)
    {
        if (args.Length == 0)
        {
            using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
            return await reader.ReadToEndAsync();
        }

        if (args.Length == 2 && args[0] == "--input")
        {
            return await File.ReadAllTextAsync(args[1], Encoding.UTF8);
        }

        throw new ArgumentException("Usage: SqlAnalysisFormatter.Parser.exe [--input <path>] [--version]");
    }

    private static JsonSerializerOptions JsonOptions()
    {
        return new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = false
        };
    }
}
