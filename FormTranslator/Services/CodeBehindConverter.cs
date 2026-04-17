using System.Text;
using System.Text.RegularExpressions;

namespace FormTranslator.Services;

public sealed class CodeBehindConverter
{
    private static readonly Regex MessageBoxCallRegex = new Regex(
        "MessageBox\\.Show\\s*\\(",
        RegexOptions.Compiled);
    private static readonly Regex CatchRegex = new Regex(
        "catch\\s*\\(\\s*Exception(?:\\s+(?<ex>[A-Za-z_][A-Za-z0-9_]*))?\\s*\\)\\s*\\{",
        RegexOptions.Compiled | RegexOptions.Singleline);
    private static readonly Regex ErrorProcessingBlockRegex = new Regex(
        "(?ms)^\\s*ErrorProcessing\\s+\\w+\\s*=\\s*new\\s+ErrorProcessing\\(\\)\\s*;\\s*\\r?\\n\\s*\\w+\\.Exception\\s*=\\s*\\w+\\s*;\\s*\\r?\\n\\s*\\w+\\.PostError\\(\\)\\s*;\\s*\\r?\\n?",
        RegexOptions.Compiled);
    private static readonly Regex ShowErrorCallRegex = new Regex(
        "Utilites\\.ShowErrorMessage\\s*\\(",
        RegexOptions.Compiled);
    private static readonly Regex ShowNonErrorCallRegex = new Regex(
        "Utilites\\.Show(?:WarningMessage|InfoMessage|YesNo)\\s*\\(",
        RegexOptions.Compiled);
    private static readonly Regex TranslateWordLiteralRegex = new Regex(
        "^DataTranslator\\.TranslateWord\\s*\\(\\s*(?<literal>@\"(?:\"\"|[^\"])*\"|\"(?:\\\\.|[^\"\\\\])*\")\\s*\\)$",
        RegexOptions.Compiled);
    private static readonly Regex TextAssignmentRegex = new Regex(
        "^(?<prefix>\\s*[A-Za-z_][A-Za-z0-9_\\.]*\\.(?:Text|ToolTipText|HeaderText|Caption)\\s*=\\s*)(?<expr>.+?)(?<suffix>;\\s*(?://.*)?)$",
        RegexOptions.Compiled);
    private static readonly Regex PhraseRegex = new Regex(
        "\\p{L}+(?:[/'&-]\\p{L}+)*(?:[ \\t]+\\p{L}+(?:[/'&-]\\p{L}+)*)*",
        RegexOptions.Compiled);

    public void ConvertFolder(string folder, Action<string> log)
    {
        var files = Directory.GetFiles(folder, "*.cs", SearchOption.AllDirectories)
            .Where(path => !path.EndsWith(".Designer.cs", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.EndsWith(".g.cs", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.Contains("\\bin\\", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.Contains("\\obj\\", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var path in files)
        {
            var original = File.ReadAllText(path);
            var rewritten = ConvertText(original, out var replacements);

            if (!ReferenceEquals(original, rewritten) && !string.Equals(original, rewritten, StringComparison.Ordinal))
            {
                var fileInfo = new FileInfo(path);
                if (fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                }

                File.WriteAllText(path, rewritten, Encoding.UTF8);
                log($"Code-behind converted: {Path.GetFileName(path)} ({replacements} replacements)");
            }
        }
    }

    private static string ConvertText(string source, out int replacements)
    {
        var sb = new StringBuilder(source);
        var delta = 0;
        replacements = 0;

        var matches = MessageBoxCallRegex.Matches(source)
            .Cast<Match>()
            .Select(m => m.Index)
            .ToList();

        foreach (var originalStart in matches)
        {
            var start = originalStart + delta;
            if (!TryLocateCall(sb.ToString(), start, out var callStart, out var callEnd, out var argsText))
            {
                continue;
            }

            var args = SplitTopLevelArguments(argsText);
            if (args.Count == 0)
            {
                continue;
            }

            var messageIndex = DetermineMessageArgumentIndex(args);
            if (messageIndex < 0 || messageIndex >= args.Count)
            {
                continue;
            }

            var messageExpression = args[messageIndex].Trim();
            var inCatch = TryGetActiveCatchException(sb.ToString(), callStart, out var catchExceptionName);
            var classification = ClassifyCall(sb.ToString(), callStart, callEnd, inCatch);

            var exceptionName = catchExceptionName;
            if (string.IsNullOrWhiteSpace(exceptionName))
            {
                exceptionName = "ex";
            }

            var messageForWrapper = PrepareMessageExpression(messageExpression, classification, exceptionName);
            var wrapperCall = BuildWrapperCall(classification, messageForWrapper, exceptionName);

            sb.Remove(callStart, callEnd - callStart + 1);
            sb.Insert(callStart, wrapperCall);
            delta += wrapperCall.Length - (callEnd - callStart + 1);
            replacements++;
        }

        var converted = RemoveErrorProcessingBlocks(sb.ToString());
        converted = NormalizeErrorWrapperCalls(converted);
        converted = NormalizeNonErrorWrapperCalls(converted);
        converted = ApplyRuntimeTextTranslations(converted);
        converted = NormalizeCatchIndentation(converted);
        return converted;
    }

    private static string RemoveErrorProcessingBlocks(string text)
    {
        return ErrorProcessingBlockRegex.Replace(text, string.Empty);
    }

    private static string NormalizeErrorWrapperCalls(string text)
    {
        var sb = new StringBuilder(text);
        var delta = 0;
        var matches = ShowErrorCallRegex.Matches(text).Cast<Match>().Select(m => m.Index).ToList();

        foreach (var originalStart in matches)
        {
            var start = originalStart + delta;
            if (!TryLocateCall(sb.ToString(), start, out var callStart, out var callEnd, out var argsText))
            {
                continue;
            }

            var args = SplitTopLevelArguments(argsText);
            if (args.Count == 0)
            {
                continue;
            }

            var messageIndex = DetermineMessageArgumentIndex(args);
            if (messageIndex < 0 || messageIndex >= args.Count)
            {
                continue;
            }

            var sanitizedMessage = SanitizeErrorMessageExpression(args[messageIndex]);
            var exceptionArg = args.Count > messageIndex + 1 ? args[messageIndex + 1].Trim() : "ex";
            var replacement = $"Utilites.ShowErrorMessage({sanitizedMessage}, {exceptionArg})";

            sb.Remove(callStart, callEnd - callStart + 1);
            sb.Insert(callStart, replacement);
            delta += replacement.Length - (callEnd - callStart + 1);
        }

        return sb.ToString();
    }

    private static string NormalizeNonErrorWrapperCalls(string text)
    {
        var sb = new StringBuilder(text);
        var delta = 0;
        var matches = ShowNonErrorCallRegex.Matches(text).Cast<Match>().Select(m => m.Index).ToList();

        foreach (var originalStart in matches)
        {
            var start = originalStart + delta;
            if (!TryLocateCall(sb.ToString(), start, out var callStart, out var callEnd, out var argsText))
            {
                continue;
            }

            var args = SplitTopLevelArguments(argsText);
            if (args.Count == 0)
            {
                continue;
            }

            var messageIndex = DetermineMessageArgumentIndex(args);
            if (messageIndex < 0 || messageIndex >= args.Count)
            {
                continue;
            }

            var originalMessage = args[messageIndex].Trim();
            var rewrittenMessage = RewriteUtilityMessageExpression(originalMessage);
            if (string.Equals(originalMessage, rewrittenMessage, StringComparison.Ordinal))
            {
                continue;
            }

            args[messageIndex] = rewrittenMessage;
            var replacement = sb.ToString().Substring(callStart, callEnd - callStart + 1);

            var openParen = replacement.IndexOf('(');
            var prefix = replacement[..(openParen + 1)];
            var rebuilt = prefix + string.Join(", ", args) + ")";

            sb.Remove(callStart, callEnd - callStart + 1);
            sb.Insert(callStart, rebuilt);
            delta += rebuilt.Length - (callEnd - callStart + 1);
        }

        return sb.ToString();
    }

    private static string RewriteUtilityMessageExpression(string expression)
    {
        var parts = SplitTopLevelConcat(expression.Trim());
        if (parts.Count <= 1)
        {
            return expression.Trim();
        }

        var rewritten = new List<string>(parts.Count);
        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (trimmed.StartsWith("DataTranslator.TranslateWord(", StringComparison.Ordinal))
            {
                rewritten.Add(trimmed);
                continue;
            }

            if (IsSimpleStringLiteral(trimmed) && ContainsLetters(trimmed))
            {
                rewritten.Add(ExpandStringLiteralForTranslation(trimmed));
            }
            else
            {
                rewritten.Add(trimmed);
            }
        }

        return string.Join(" + ", rewritten);
    }

    private static string ApplyRuntimeTextTranslations(string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            if (line.Contains("Utilites.Show", StringComparison.Ordinal)
                || line.Contains("MessageBox.Show", StringComparison.Ordinal)
                || line.Contains("DataTranslator.TranslateWord", StringComparison.Ordinal))
            {
                continue;
            }

            var match = TextAssignmentRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            var wrapped = WrapTranslatableLiteralsInExpression(match.Groups["expr"].Value);
            lines[i] = match.Groups["prefix"].Value + wrapped + match.Groups["suffix"].Value;
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static string WrapTranslatableLiteralsInExpression(string expression)
    {
        var parts = SplitTopLevelConcat(expression);
        var rewritten = new List<string>(parts.Count);

        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (trimmed.StartsWith("DataTranslator.TranslateWord(", StringComparison.Ordinal))
            {
                rewritten.Add(trimmed);
                continue;
            }

            if (IsSimpleStringLiteral(trimmed) && ContainsLetters(trimmed))
            {
                rewritten.Add(ExpandStringLiteralForTranslation(trimmed));
            }
            else
            {
                rewritten.Add(trimmed);
            }
        }

        return string.Join(" + ", rewritten);
    }

    private static bool ContainsLetters(string stringLiteral)
    {
        var value = DecodeStringLiteralValue(stringLiteral);
        return value.Any(char.IsLetter);
    }

    private static string ExpandStringLiteralForTranslation(string stringLiteral)
    {
        var raw = DecodeStringLiteralValue(stringLiteral);
        var pieces = new List<string>();
        var index = 0;

        foreach (Match match in PhraseRegex.Matches(raw))
        {
            if (match.Index > index)
            {
                var nonPhrase = raw.Substring(index, match.Index - index);
                if (nonPhrase.Length > 0)
                {
                    pieces.Add(ToStringLiteral(nonPhrase));
                }
            }

            var phrase = match.Value.Trim();
            if (phrase.Length > 0)
            {
                pieces.Add($"DataTranslator.TranslateWord({ToStringLiteral(phrase)})");
            }

            var consumed = match.Value.Length;
            var trailingSpaceCount = match.Value.Length - match.Value.TrimEnd(' ', '\t').Length;
            if (trailingSpaceCount > 0)
            {
                var trailing = match.Value[^trailingSpaceCount..];
                pieces.Add(ToStringLiteral(trailing));
            }

            index = match.Index + consumed;
        }

        if (index < raw.Length)
        {
            pieces.Add(ToStringLiteral(raw[index..]));
        }

        if (pieces.Count == 0)
        {
            return stringLiteral;
        }

        return string.Join(" + ", pieces.Where(piece => piece != "\"\""));
    }

    private static string ToStringLiteral(string value)
    {
        var escaped = value
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\r", "\\r", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal)
            .Replace("\t", "\\t", StringComparison.Ordinal);
        return $"\"{escaped}\"";
    }

    private static string DecodeStringLiteralValue(string literal)
    {
        if (literal.StartsWith("@\"", StringComparison.Ordinal) && literal.EndsWith('"'))
        {
            return literal[2..^1].Replace("\"\"", "\"");
        }

        if (!literal.StartsWith("\"", StringComparison.Ordinal) || !literal.EndsWith('"'))
        {
            return literal;
        }

        var body = literal[1..^1];
        var sb = new StringBuilder(body.Length);
        for (var i = 0; i < body.Length; i++)
        {
            if (body[i] == '\\' && i + 1 < body.Length)
            {
                i++;
                sb.Append(body[i] switch
                {
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    '"' => '"',
                    '\\' => '\\',
                    _ => body[i]
                });
            }
            else
            {
                sb.Append(body[i]);
            }
        }

        return sb.ToString();
    }

    private static string NormalizeCatchIndentation(string text)
    {
        var normalizedText = text.Replace("\r\n", "\n");
        var lines = normalizedText.Split('\n').ToList();

        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            var trimmed = line.TrimStart();
            if (!trimmed.StartsWith("catch", StringComparison.Ordinal))
            {
                continue;
            }

            var catchIndent = line[..(line.Length - trimmed.Length)];
            var braceLine = i;
            if (!line.Contains('{'))
            {
                var foundBrace = false;
                for (var j = i + 1; j < lines.Count; j++)
                {
                    if (lines[j].Trim().Length == 0)
                    {
                        continue;
                    }

                    if (lines[j].TrimStart().StartsWith("{", StringComparison.Ordinal))
                    {
                        braceLine = j;
                        lines[j] = catchIndent + lines[j].TrimStart();
                        foundBrace = true;
                    }

                    break;
                }

                if (!foundBrace)
                {
                    continue;
                }
            }

            var depth = 1;
            for (var j = braceLine + 1; j < lines.Count && depth > 0; j++)
            {
                var current = lines[j];
                var currentTrimmed = current.TrimStart();
                var closeCount = current.Count(ch => ch == '}');
                var openCount = current.Count(ch => ch == '{');

                var effectiveDepth = depth - closeCount;
                if (effectiveDepth < 0)
                {
                    effectiveDepth = 0;
                }

                if (currentTrimmed.Length > 0)
                {
                    var expectedIndent = catchIndent + new string(' ', effectiveDepth * 4);
                    lines[j] = expectedIndent + currentTrimmed;
                }

                depth += openCount - closeCount;
            }
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static bool TryLocateCall(string text, int start, out int callStart, out int callEnd, out string argsText)
    {
        callStart = start;
        callEnd = -1;
        argsText = string.Empty;

        var openParen = text.IndexOf('(', start);
        if (openParen < 0)
        {
            return false;
        }

        var depth = 1;
        var inString = false;
        var inVerbatim = false;

        for (var i = openParen + 1; i < text.Length; i++)
        {
            var c = text[i];
            var prev = i > 0 ? text[i - 1] : '\0';

            if (!inString && c == '@' && i + 1 < text.Length && text[i + 1] == '"')
            {
                inString = true;
                inVerbatim = true;
                i++;
                continue;
            }

            if (c == '"' && !inVerbatim && prev != '\\')
            {
                inString = !inString;
                continue;
            }

            if (inVerbatim && c == '"')
            {
                if (i + 1 < text.Length && text[i + 1] == '"')
                {
                    i++;
                    continue;
                }

                inString = false;
                inVerbatim = false;
                continue;
            }

            if (inString)
            {
                continue;
            }

            if (c == '(')
            {
                depth++;
            }
            else if (c == ')')
            {
                depth--;
                if (depth == 0)
                {
                    callEnd = i;
                    argsText = text.Substring(openParen + 1, i - openParen - 1);
                    return true;
                }
            }
        }

        return false;
    }

    private static List<string> SplitTopLevelArguments(string argsText)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        var depthParen = 0;
        var depthBracket = 0;
        var depthBrace = 0;
        var inString = false;
        var inVerbatim = false;

        for (var i = 0; i < argsText.Length; i++)
        {
            var c = argsText[i];
            var prev = i > 0 ? argsText[i - 1] : '\0';

            if (!inString && c == '@' && i + 1 < argsText.Length && argsText[i + 1] == '"')
            {
                inString = true;
                inVerbatim = true;
                current.Append(c);
                continue;
            }

            if (c == '"' && !inVerbatim && prev != '\\')
            {
                inString = !inString;
                current.Append(c);
                continue;
            }

            if (inVerbatim && c == '"')
            {
                current.Append(c);
                if (i + 1 < argsText.Length && argsText[i + 1] == '"')
                {
                    i++;
                    current.Append(argsText[i]);
                    continue;
                }

                inString = false;
                inVerbatim = false;
                continue;
            }

            if (!inString)
            {
                if (c == '(') depthParen++;
                else if (c == ')') depthParen--;
                else if (c == '[') depthBracket++;
                else if (c == ']') depthBracket--;
                else if (c == '{') depthBrace++;
                else if (c == '}') depthBrace--;
                else if (c == ',' && depthParen == 0 && depthBracket == 0 && depthBrace == 0)
                {
                    result.Add(current.ToString().Trim());
                    current.Clear();
                    continue;
                }
            }

            current.Append(c);
        }

        var tail = current.ToString().Trim();
        if (tail.Length > 0)
        {
            result.Add(tail);
        }

        return result;
    }

    private static int DetermineMessageArgumentIndex(IReadOnlyList<string> args)
    {
        if (args.Count == 0)
        {
            return -1;
        }

        return string.Equals(args[0].Trim(), "this", StringComparison.Ordinal) && args.Count > 1
            ? 1
            : 0;
    }

    private static bool TryGetActiveCatchException(string text, int callStart, out string exceptionName)
    {
        exceptionName = string.Empty;
        var start = Math.Max(0, callStart - 8000);
        var snippet = text.Substring(start, callStart - start);
        var matches = CatchRegex.Matches(snippet);
        if (matches.Count == 0)
        {
            return false;
        }

        var match = matches[^1];
        var depth = 1;
        for (var i = match.Index + match.Length; i < snippet.Length; i++)
        {
            if (snippet[i] == '{') depth++;
            if (snippet[i] == '}') depth--;
            if (depth <= 0)
            {
                return false;
            }
        }

        exceptionName = match.Groups["ex"].Success ? match.Groups["ex"].Value : "ex";
        return true;
    }

    private static string ClassifyCall(string text, int callStart, int callEnd, bool inCatch)
    {
        var call = text.Substring(callStart, callEnd - callStart + 1);
        if (call.Contains("MessageBoxButtons.YesNo", StringComparison.Ordinal))
        {
            return "yesno";
        }

        if (call.Contains("MessageBoxIcon.Error", StringComparison.Ordinal) || inCatch)
        {
            return "error";
        }

        if (call.Contains("MessageBoxIcon.Warning", StringComparison.Ordinal))
        {
            return "warning";
        }

        return "info";
    }

    private static string PrepareMessageExpression(string expression, string classification, string exceptionName)
    {
        var trimmed = expression.Trim();
        if (classification != "error" && IsSimpleStringLiteral(trimmed))
        {
            return trimmed;
        }

        var parts = SplitTopLevelConcat(trimmed);
        if (classification == "error")
        {
            parts = parts
                .Where(part => !ContainsExceptionText(part, exceptionName))
                .ToList();
        }

        parts = parts
            .Select(part => StripTranslateWordLiteral(part.Trim()))
            .Where(part => !IsNewLinePart(part))
            .Where(part => !IsEmptyStringLiteral(part))
            .ToList();

        if (parts.Count == 0)
        {
            return "\"Error occurred.\"";
        }

        return parts.Count == 1 ? parts[0] : string.Join(" + ", parts);
    }

    private static List<string> SplitTopLevelConcat(string expression)
    {
        var parts = new List<string>();
        var sb = new StringBuilder();
        var depthParen = 0;
        var depthBracket = 0;
        var depthBrace = 0;
        var inString = false;
        var inVerbatim = false;

        for (var i = 0; i < expression.Length; i++)
        {
            var c = expression[i];
            var prev = i > 0 ? expression[i - 1] : '\0';

            if (!inString && c == '@' && i + 1 < expression.Length && expression[i + 1] == '"')
            {
                inString = true;
                inVerbatim = true;
                sb.Append(c);
                continue;
            }

            if (c == '"' && !inVerbatim && prev != '\\')
            {
                inString = !inString;
                sb.Append(c);
                continue;
            }

            if (inVerbatim && c == '"')
            {
                sb.Append(c);
                if (i + 1 < expression.Length && expression[i + 1] == '"')
                {
                    i++;
                    sb.Append(expression[i]);
                    continue;
                }

                inString = false;
                inVerbatim = false;
                continue;
            }

            if (!inString)
            {
                if (c == '(') depthParen++;
                else if (c == ')') depthParen--;
                else if (c == '[') depthBracket++;
                else if (c == ']') depthBracket--;
                else if (c == '{') depthBrace++;
                else if (c == '}') depthBrace--;
                else if (c == '+' && depthParen == 0 && depthBracket == 0 && depthBrace == 0)
                {
                    parts.Add(sb.ToString().Trim());
                    sb.Clear();
                    continue;
                }
            }

            sb.Append(c);
        }

        var last = sb.ToString().Trim();
        if (last.Length > 0)
        {
            parts.Add(last);
        }

        return parts;
    }

    private static bool ContainsExceptionText(string expressionPart, string exceptionName)
    {
        var part = expressionPart.Trim();
        return part.Contains(exceptionName + ".Message", StringComparison.Ordinal)
            || part.Contains(exceptionName + ".ToString()", StringComparison.Ordinal)
            || part.Contains("ex.Message", StringComparison.Ordinal)
            || part.Contains("ex.ToString()", StringComparison.Ordinal);
    }

    private static string SanitizeErrorMessageExpression(string expression)
    {
        var parts = SplitTopLevelConcat(expression)
            .Select(part => StripTranslateWordLiteral(part.Trim()))
            .Where(part => !IsNewLinePart(part))
            .Where(part => !IsEmptyStringLiteral(part))
            .ToList();

        if (parts.Count == 0)
        {
            return "\"Error occurred.\"";
        }

        return parts.Count == 1 ? parts[0] : string.Join(" + ", parts);
    }

    private static bool IsNewLinePart(string part)
    {
        return string.Equals(part, "Environment.NewLine", StringComparison.Ordinal)
            || string.Equals(part, "System.Environment.NewLine", StringComparison.Ordinal);
    }

    private static string StripTranslateWordLiteral(string part)
    {
        var match = TranslateWordLiteralRegex.Match(part);
        if (match.Success)
        {
            return match.Groups["literal"].Value;
        }

        return part;
    }

    private static bool IsEmptyStringLiteral(string part)
    {
        if (part == "\"\"" || part == "@\"\"")
        {
            return true;
        }

        return false;
    }

    private static bool IsSimpleStringLiteral(string expression)
    {
        var value = expression.Trim();
        return (value.StartsWith("\"", StringComparison.Ordinal) && value.EndsWith("\"", StringComparison.Ordinal))
            || (value.StartsWith("@\"", StringComparison.Ordinal) && value.EndsWith("\"", StringComparison.Ordinal));
    }

    private static string BuildWrapperCall(string classification, string messageExpression, string exceptionName)
    {
        return classification switch
        {
            "error" => $"Utilites.ShowErrorMessage({messageExpression}, {exceptionName})",
            "warning" => $"Utilites.ShowWarningMessage({messageExpression})",
            "yesno" => $"Utilites.ShowYesNo({messageExpression})",
            _ => $"Utilites.ShowInfoMessage({messageExpression})"
        };
    }
}
