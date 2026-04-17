namespace FormTranslator;

public static class AppModels
{
    public sealed class AppSettings
    {
        public string TargetLanguageCode { get; set; } = "es-MX";
        public string LastFolder { get; set; } = string.Empty;
        public bool ExportBatchCsv { get; set; } = true;
        public bool ConvertCodeBehind { get; set; } = false;
    }

    public sealed record LanguageOption(string Name, string Code)
    {
        public override string ToString() => Name;
    }

    public sealed record ResxEntry(string Key, string English);

    public sealed record FormTranslationData(string DesignerPath, List<ResxEntry> Entries);
}
