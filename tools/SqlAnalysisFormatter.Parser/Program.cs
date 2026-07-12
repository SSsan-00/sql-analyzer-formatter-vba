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
        CliOptions options;
        try
        {
            options = ParseOptions(args);
            sql = await ReadSqlAsync(options);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }

        var result = TsqlAstParser.Parse(sql);
        await WriteResultAsync(result, options);
        return 0;
    }

    private static CliOptions ParseOptions(string[] args)
    {
        var options = new CliOptions();
        for (var index = 0; index < args.Length; index++)
        {
            switch (args[index])
            {
                case "--input":
                    options.InputPath = ReadOptionValue(args, ref index, "--input");
                    break;
                case "--output":
                    options.OutputPath = ReadOptionValue(args, ref index, "--output");
                    break;
                case "--format":
                    options.Format = ReadOptionValue(args, ref index, "--format");
                    break;
                default:
                    throw new ArgumentException("Usage: SqlAnalysisFormatter.Parser.exe [--input <path>] [--output <path>] [--format json|vba-blocks] [--version]");
            }
        }

        if (options.Format is not "json" and not "vba-blocks")
        {
            throw new ArgumentException("Unknown format: " + options.Format);
        }

        return options;
    }

    private static string ReadOptionValue(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException("Missing value for " + optionName);
        }

        index++;
        return args[index];
    }

    private static async Task<string> ReadSqlAsync(CliOptions options)
    {
        if (options.InputPath.Length == 0)
        {
            using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
            return await reader.ReadToEndAsync();
        }

        return await File.ReadAllTextAsync(options.InputPath, Encoding.UTF8);
    }

    private static async Task WriteResultAsync(ParseResult result, CliOptions options)
    {
        var outputText = options.Format == "vba-blocks"
            ? ParserOutputFormatter.ToVbaBlocks(result)
            : JsonSerializer.Serialize(result, JsonOptions());

        if (options.OutputPath.Length == 0)
        {
            Console.WriteLine(outputText);
            return;
        }

        var encoding = options.Format == "vba-blocks"
            ? Encoding.Unicode
            : new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        await File.WriteAllTextAsync(options.OutputPath, outputText, encoding);
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

    private sealed class CliOptions
    {
        public string InputPath { get; set; } = string.Empty;

        public string OutputPath { get; set; } = string.Empty;

        public string Format { get; set; } = "json";
    }
}
