using System.Text;
using System.Text.RegularExpressions;

namespace FormTranslator.Services;

public sealed class CodeBehindLocalizationExporter
{
    private static readonly Regex StringLiteralRegex = new Regex(
        "@\"(?:\"\"|[^\"])*\"|\"(?:\\\\.|[^\"\\\\])*\"",
        RegexOptions.Compiled);
    private static readonly Regex PhraseRegex = new Regex(
        "\\p{L}+(?:[/'&-]\\p{L}+)*(?:[ \\t]+\\p{L}+(?:[/'&-]\\p{L}+)*)*",
        RegexOptions.Compiled);

    public List<string> ExtractAllCodeBehindStrings(string folder)
    {
        var result = new List<string>();
        var files = Directory.GetFiles(folder, "*.cs", SearchOption.AllDirectories)
            .Where(path => !path.EndsWith(".Designer.cs", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.EndsWith(".g.cs", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.Contains("\\bin\\", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.Contains("\\obj\\", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);

        foreach (var file in files)
        {
            var text = File.ReadAllText(file);
            foreach (Match literalMatch in StringLiteralRegex.Matches(text))
            {
                var value = DecodeCSharpStringLiteral(literalMatch.Value);
                foreach (var piece in SplitLiteralIntoPieces(value))
                {
                    var trimmed = piece.Trim();
                    if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        result.Add(trimmed);
                    }
                }
            }
        }

        return result;
    }

    private static IEnumerable<string> SplitLiteralIntoPieces(string literal)
    {
        var pieces = new List<string>();
        var index = 0;

        foreach (Match match in PhraseRegex.Matches(literal))
        {
            if (match.Index > index)
            {
                var nonPhrase = literal.Substring(index, match.Index - index);
                if (!string.IsNullOrWhiteSpace(nonPhrase))
                {
                    pieces.Add(nonPhrase);
                }
            }

            if (!string.IsNullOrWhiteSpace(match.Value))
            {
                pieces.Add(match.Value);
            }

            index = match.Index + match.Length;
        }

        if (index < literal.Length)
        {
            var tail = literal[index..];
            if (!string.IsNullOrWhiteSpace(tail))
            {
                pieces.Add(tail);
            }
        }

        if (pieces.Count == 0 && !string.IsNullOrWhiteSpace(literal))
        {
            pieces.Add(literal);
        }

        return pieces;
    }

    private static string DecodeCSharpStringLiteral(string literal)
    {
        if (literal.StartsWith("@\"", StringComparison.Ordinal) && literal.EndsWith('"'))
        {
            var body = literal[2..^1];
            return body.Replace("\"\"", "\"");
        }

        if (literal.Length < 2 || literal[0] != '"' || literal[^1] != '"')
        {
            return literal;
        }

        var bodyText = literal[1..^1];
        var sb = new StringBuilder(bodyText.Length);
        for (var i = 0; i < bodyText.Length; i++)
        {
            if (bodyText[i] == '\\' && i + 1 < bodyText.Length)
            {
                i++;
                sb.Append(bodyText[i] switch
                {
                    '\\' => '\\',
                    '"' => '"',
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    _ => bodyText[i]
                });
            }
            else
            {
                sb.Append(bodyText[i]);
            }
        }

        return sb.ToString();
    }
}
