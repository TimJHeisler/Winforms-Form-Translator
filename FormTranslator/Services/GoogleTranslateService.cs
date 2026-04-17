using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace FormTranslator.Services;

public sealed class GoogleTranslateService
{
    public const int BatchSize = 50;
    private static readonly HttpClient Http = new HttpClient();
    private static readonly Regex MarkerRegex = new Regex(
        "__FT(?<id>\\d+)__(?<text>.*?)__END\\k<id>__",
        RegexOptions.Compiled | RegexOptions.Singleline);

    public async Task<Dictionary<string, string>> TranslateAllAsync(
        IEnumerable<string> uniqueEnglishTexts,
        string targetLanguageCode,
        Action onItemTranslated)
    {
        var englishList = uniqueEnglishTexts.Where(text => !string.IsNullOrWhiteSpace(text)).ToList();
        var map = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var batch in englishList.Chunk(BatchSize))
        {
            var batchTranslations = await TranslateBatchAsync(batch, targetLanguageCode, onItemTranslated);
            foreach (var pair in batchTranslations)
            {
                map[pair.Key] = pair.Value;
            }
        }

        return map;
    }

    private static async Task<Dictionary<string, string>> TranslateBatchAsync(
        IEnumerable<string> englishBatch,
        string targetLanguageCode,
        Action onItemTranslated)
    {
        var items = englishBatch.ToList();
        try
        {
            var translations = await TranslateBatchRequestAsync(items, targetLanguageCode);

            foreach (var _ in items)
            {
                onItemTranslated();
            }

            return translations;
        }
        catch
        {
            var fallback = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var item in items)
            {
                fallback[item] = item;
                onItemTranslated();
            }

            return fallback;
        }
    }

    private static async Task<Dictionary<string, string>> TranslateBatchRequestAsync(
        IReadOnlyList<string> items,
        string targetLanguageCode)
    {
        var payload = BuildMarkedPayload(items);
        var url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl="
            + WebUtility.UrlEncode(targetLanguageCode)
            + "&format=html"
            + "&dt=t&q="
            + WebUtility.UrlEncode(payload);

        var response = await Http.GetStringAsync(url);
        using var document = JsonDocument.Parse(response);
        var translatedPayload = FlattenTranslatedPayload(document.RootElement);

        var translations = new Dictionary<string, string>(StringComparer.Ordinal);
        var parsed = ParseMarkedPayload(translatedPayload);
        for (var i = 0; i < items.Count; i++)
        {
            translations[items[i]] = parsed.TryGetValue(i, out var translated) && !string.IsNullOrWhiteSpace(translated)
                ? translated
                : items[i];
        }

        return translations;
    }

    private static string BuildMarkedPayload(IReadOnlyList<string> items)
    {
        var sb = new StringBuilder(items.Count * 64);
        for (var i = 0; i < items.Count; i++)
        {
            sb.Append("__FT");
            sb.Append(i);
            sb.Append("__");
            sb.Append(WebUtility.HtmlEncode(items[i]));
            sb.Append("__END");
            sb.Append(i);
            sb.Append("__");
            sb.Append('\n');
        }

        return sb.ToString();
    }

    private static string FlattenTranslatedPayload(JsonElement rootElement)
    {
        var sb = new StringBuilder();
        foreach (var segment in rootElement[0].EnumerateArray())
        {
            if (segment.GetArrayLength() > 0)
            {
                sb.Append(segment[0].GetString());
            }
        }

        return sb.ToString();
    }

    private static Dictionary<int, string> ParseMarkedPayload(string translatedPayload)
    {
        var result = new Dictionary<int, string>();
        foreach (Match match in MarkerRegex.Matches(translatedPayload))
        {
            if (!int.TryParse(match.Groups["id"].Value, out var id))
            {
                continue;
            }

            result[id] = WebUtility.HtmlDecode(match.Groups["text"].Value).Trim();
        }

        return result;
    }
}
