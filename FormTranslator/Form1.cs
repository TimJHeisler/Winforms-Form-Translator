using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Resources;
using FormTranslator.Services;
using AppSettings = FormTranslator.AppModels.AppSettings;
using FormTranslationData = FormTranslator.AppModels.FormTranslationData;
using LanguageOption = FormTranslator.AppModels.LanguageOption;
using ResxEntry = FormTranslator.AppModels.ResxEntry;

namespace FormTranslator
{
    public partial class Form1 : Form
    {
        private static readonly Color Surface = Color.FromArgb(24, 26, 31);
        private static readonly Color Panel = Color.FromArgb(34, 37, 44);
        private static readonly Color Input = Color.FromArgb(20, 22, 28);
        private static readonly Color Accent = Color.FromArgb(0, 122, 204);
        private static readonly Color Foreground = Color.FromArgb(231, 234, 239);
        private readonly GoogleTranslateService _translateService = new GoogleTranslateService();
        private readonly CodeBehindConverter _codeBehindConverter = new CodeBehindConverter();
        private readonly CodeBehindLocalizationExporter _codeBehindLocalizationExporter = new CodeBehindLocalizationExporter();
        private bool _suppressLanguageChange;
        private AppSettings _settings = AppSettingsStore.Load();
        private string _currentRunTimestamp = string.Empty;
        private static readonly LanguageOption[] LanguageOptions =
        {
            new LanguageOption("Spanish (Mexico)", "es-MX"),
            new LanguageOption("Spanish", "es"),
            new LanguageOption("French", "fr"),
            new LanguageOption("German", "de"),
            new LanguageOption("Italian", "it"),
            new LanguageOption("Portuguese (Brazil)", "pt-BR")
        };
        private static readonly Regex AssignmentRegex = new Regex(
            "^\\s*(?<target>[A-Za-z_][A-Za-z0-9_\\.]*)\\.(?<property>Text|ToolTipText|HeaderText|Caption|PlaceholderText)\\s*=\\s*(?<value>\"(?:\\\\.|[^\"\\\\])*\")\\s*;",
            RegexOptions.Compiled);
        private static readonly Regex FormTextRegex = new Regex(
            "^\\s*Text\\s*=\\s*(?<value>\"(?:\\\\.|[^\"\\\\])*\")\\s*;",
            RegexOptions.Compiled);
        private static readonly Regex SetToolTipRegex = new Regex(
            "^\\s*[A-Za-z_][A-Za-z0-9_]*\\.SetToolTip\\(\\s*(?<target>[A-Za-z_][A-Za-z0-9_]*)\\s*,\\s*(?<value>\"(?:\\\\.|[^\"\\\\])*\")\\s*\\)\\s*;",
            RegexOptions.Compiled);
        private static readonly Regex ItemsAddRegex = new Regex(
            "^\\s*(?<target>[A-Za-z_][A-Za-z0-9_\\.]*)\\.Items\\.Add\\(\\s*(?<value>\"(?:\\\\.|[^\"\\\\])*\")\\s*\\)\\s*;",
            RegexOptions.Compiled);
        private static readonly Regex AddRangeStartRegex = new Regex(
            "^\\s*(?<target>[A-Za-z_][A-Za-z0-9_\\.]*)\\.(?<collection>Items|DropDownItems)\\.AddRange\\(\\s*new\\s+object\\[\\]\\s*\\{",
            RegexOptions.Compiled);
        private static readonly Regex StatementEndRegex = new Regex(
            "\\)\\s*;\\s*$",
            RegexOptions.Compiled);
        private static readonly Regex MenuAddRegex = new Regex(
            "^\\s*(?<target>[A-Za-z_][A-Za-z0-9_\\.]*)\\.(?<collection>Items|DropDownItems)\\.Add\\(\\s*(?<value>\"(?:\\\\.|[^\"\\\\])*\")\\s*\\)\\s*;",
            RegexOptions.Compiled);
        private static readonly Regex StringLiteralRegex = new Regex(
            "\"(?:\\\\.|[^\"\\\\])*\"",
            RegexOptions.Compiled);

        public Form1()
        {
            InitializeComponent();
            InitializeLanguageOptions();
            ApplyLoadedSettings();
            ApplyDarkTheme();
        }

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog(this) == DialogResult.OK)
            {
                txtFolderPath.Text = folderBrowserDialog1.SelectedPath;
                _settings.LastFolder = folderBrowserDialog1.SelectedPath;
                AppSettingsStore.Save(_settings);
            }
        }

        private void CboLanguage_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_suppressLanguageChange)
            {
                return;
            }

            var selected = GetSelectedLanguage();
            _settings.TargetLanguageCode = selected.Code;
            UpdateTranslateButtonText(selected.Code);
            AppSettingsStore.Save(_settings);
        }

        private void ChkExportCsv_CheckedChanged(object? sender, EventArgs e)
        {
            if (_suppressLanguageChange)
            {
                return;
            }

            _settings.ExportBatchCsv = chkExportCsv.Checked;
            AppSettingsStore.Save(_settings);
        }

        private void ChkConvertCodeBehind_CheckedChanged(object? sender, EventArgs e)
        {
            if (_suppressLanguageChange)
            {
                return;
            }

            _settings.ConvertCodeBehind = chkConvertCodeBehind.Checked;
            AppSettingsStore.Save(_settings);
        }

        private async void BtnTranslate_Click(object? sender, EventArgs e)
        {
            var folder = txtFolderPath.Text.Trim();
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show(this, "Pick a valid folder first.", "Missing folder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnTranslate.Enabled = false;
            txtLog.Clear();
            progressBar1.Value = 0;
            lblProgress.Text = "0%";
            _currentRunTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);

            try
            {
                var languageCode = GetSelectedLanguage().Code;
                await TranslateFolderAsync(folder, languageCode);
                Log("Done.");
            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex.Message}");
                MessageBox.Show(this, ex.ToString(), "Translation failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnTranslate.Enabled = true;
            }
        }

        private async Task TranslateFolderAsync(string folder, string targetLanguageCode)
        {
            var designerFiles = Directory.GetFiles(folder, "*.Designer.cs", SearchOption.AllDirectories)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var formEntries = new List<FormTranslationData>();
            var uniqueTextsSet = new HashSet<string>(StringComparer.Ordinal);

            if (designerFiles.Count == 0)
            {
                Log("No .Designer.cs files found for form resource translation.");
            }
            else
            {
                foreach (var designerPath in designerFiles)
                {
                    var entries = ExtractEntries(designerPath);
                    if (entries.Count == 0)
                    {
                        continue;
                    }

                    var normalizedEntries = entries
                        .GroupBy(e => e.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(group => group.Last())
                        .ToList();

                    foreach (var entry in normalizedEntries)
                    {
                        if (!string.IsNullOrWhiteSpace(entry.English))
                        {
                            uniqueTextsSet.Add(entry.English);
                        }
                    }

                    formEntries.Add(new FormTranslationData(designerPath, normalizedEntries));
                    Log($"Found {normalizedEntries.Count} translatable entries in {Path.GetFileName(designerPath)}");
                }
            }

            if (formEntries.Count > 0)
            {
                var uniqueTexts = uniqueTextsSet
                    .OrderBy(text => text, StringComparer.Ordinal)
                    .ToList();

                Log($"Translating {uniqueTexts.Count} unique texts to {targetLanguageCode} using Google endpoint...");
                var totalSteps = uniqueTexts.Count + formEntries.Count;
                var completedSteps = 0;
                UpdateProgress(completedSteps, totalSteps);

                var translations = await _translateService.TranslateAllAsync(uniqueTexts, targetLanguageCode, () =>
                {
                    completedSteps++;
                    UpdateProgress(completedSteps, totalSteps);
                });

                Log($"Translated {uniqueTexts.Count} items");

                foreach (var formData in formEntries)
                {
                    WriteResx(formData, targetLanguageCode, translations);
                    completedSteps++;
                    UpdateProgress(completedSteps, totalSteps);
                }

                if (chkExportCsv.Checked)
                {
                    WriteFolderCsvFiles(formEntries, targetLanguageCode, translations);
                }
            }
            else
            {
                Log("No translatable form entries found for resource generation.");
            }

            if (chkConvertCodeBehind.Checked)
            {
                _codeBehindConverter.ConvertFolder(folder, Log);
                await WriteCodeBehindCsvAsync(folder, targetLanguageCode);
            }

            if (formEntries.Count == 0 && chkConvertCodeBehind.Checked)
            {
                UpdateProgress(1, 1);
            }
        }

        private static List<ResxEntry> ExtractEntries(string designerPath)
        {
            var result = new List<ResxEntry>();
            var lines = File.ReadAllLines(designerPath);
            var itemIndexByTarget = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var collectingAddRange = false;
            var addRangeTarget = string.Empty;
            var addRangeCollection = string.Empty;
            var addRangeBuffer = new StringBuilder();

            foreach (var line in lines)
            {
                if (collectingAddRange)
                {
                    var nestedAddRangeStart = AddRangeStartRegex.Match(line);
                    if (nestedAddRangeStart.Success)
                    {
                        FlushAddRangeBuffer(result, itemIndexByTarget, addRangeTarget, addRangeCollection, addRangeBuffer);
                        addRangeTarget = nestedAddRangeStart.Groups["target"].Value;
                        addRangeCollection = nestedAddRangeStart.Groups["collection"].Value;
                        addRangeBuffer.Clear();
                        addRangeBuffer.AppendLine(line);
                        if (StatementEndRegex.IsMatch(line))
                        {
                            FlushAddRangeBuffer(result, itemIndexByTarget, addRangeTarget, addRangeCollection, addRangeBuffer);
                            collectingAddRange = false;
                            addRangeTarget = string.Empty;
                            addRangeCollection = string.Empty;
                            addRangeBuffer.Clear();
                        }

                        continue;
                    }

                    addRangeBuffer.AppendLine(line);
                    if (StatementEndRegex.IsMatch(line))
                    {
                        FlushAddRangeBuffer(result, itemIndexByTarget, addRangeTarget, addRangeCollection, addRangeBuffer);

                        collectingAddRange = false;
                        addRangeTarget = string.Empty;
                        addRangeCollection = string.Empty;
                        addRangeBuffer.Clear();
                    }

                    continue;
                }

                var assignmentMatch = AssignmentRegex.Match(line);
                if (assignmentMatch.Success)
                {
                    var target = assignmentMatch.Groups["target"].Value;
                    var property = assignmentMatch.Groups["property"].Value;
                    var rawValue = assignmentMatch.Groups["value"].Value;
                    var english = DecodeCSharpString(rawValue);

                    if (!string.IsNullOrWhiteSpace(english))
                    {
                        result.Add(new ResxEntry($"{target}.{property}", english));
                    }

                    continue;
                }

                var formTextMatch = FormTextRegex.Match(line);
                if (formTextMatch.Success)
                {
                    var rawValue = formTextMatch.Groups["value"].Value;
                    var english = DecodeCSharpString(rawValue);
                    if (!string.IsNullOrWhiteSpace(english))
                    {
                        result.Add(new ResxEntry("$this.Text", english));
                    }

                    continue;
                }

                var tooltipMatch = SetToolTipRegex.Match(line);
                if (tooltipMatch.Success)
                {
                    var target = tooltipMatch.Groups["target"].Value;
                    var rawValue = tooltipMatch.Groups["value"].Value;
                    var english = DecodeCSharpString(rawValue);
                    if (!string.IsNullOrWhiteSpace(english))
                    {
                        result.Add(new ResxEntry($"{target}.ToolTip", english));
                    }

                    continue;
                }

                var itemAddMatch = ItemsAddRegex.Match(line);
                if (itemAddMatch.Success)
                {
                    var target = itemAddMatch.Groups["target"].Value;
                    var rawValue = itemAddMatch.Groups["value"].Value;
                    var english = DecodeCSharpString(rawValue);
                    if (!string.IsNullOrWhiteSpace(english))
                    {
                        AddCollectionItemEntry(result, itemIndexByTarget, target, "Items", english);
                    }

                    continue;
                }

                var menuAddMatch = MenuAddRegex.Match(line);
                if (menuAddMatch.Success)
                {
                    var target = menuAddMatch.Groups["target"].Value;
                    var collection = menuAddMatch.Groups["collection"].Value;
                    var rawValue = menuAddMatch.Groups["value"].Value;
                    var english = DecodeCSharpString(rawValue);
                    if (!string.IsNullOrWhiteSpace(english))
                    {
                        AddCollectionItemEntry(result, itemIndexByTarget, target, collection, english);
                    }
                }

                var addRangeStartMatch = AddRangeStartRegex.Match(line);
                if (addRangeStartMatch.Success)
                {
                    addRangeTarget = addRangeStartMatch.Groups["target"].Value;
                    addRangeCollection = addRangeStartMatch.Groups["collection"].Value;
                    collectingAddRange = true;
                    addRangeBuffer.Clear();
                    addRangeBuffer.AppendLine(line);

                    if (StatementEndRegex.IsMatch(line))
                    {
                        FlushAddRangeBuffer(result, itemIndexByTarget, addRangeTarget, addRangeCollection, addRangeBuffer);

                        collectingAddRange = false;
                        addRangeTarget = string.Empty;
                        addRangeCollection = string.Empty;
                        addRangeBuffer.Clear();
                    }
                }
            }

            return result;
        }

        private static void AddCollectionItemEntry(
            List<ResxEntry> result,
            Dictionary<string, int> itemIndexByTarget,
            string target,
            string collectionName,
            string english)
        {
            var counterKey = $"{target}.{collectionName}";
            itemIndexByTarget.TryGetValue(counterKey, out var index);
            var key = index == 0 ? counterKey : $"{counterKey}{index}";
            itemIndexByTarget[counterKey] = index + 1;
            result.Add(new ResxEntry(key, english));
        }

        private static void FlushAddRangeBuffer(
            List<ResxEntry> result,
            Dictionary<string, int> itemIndexByTarget,
            string target,
            string collectionName,
            StringBuilder buffer)
        {
            foreach (Match itemMatch in StringLiteralRegex.Matches(buffer.ToString()))
            {
                var english = DecodeCSharpString(itemMatch.Value);
                if (!string.IsNullOrWhiteSpace(english))
                {
                    AddCollectionItemEntry(result, itemIndexByTarget, target, collectionName, english);
                }
            }
        }

        private void WriteResx(FormTranslationData formData, string targetLanguageCode, IReadOnlyDictionary<string, string> translations)
        {
            var designerPath = formData.DesignerPath;
            var baseName = Path.GetFileNameWithoutExtension(designerPath);
            if (baseName.EndsWith(".Designer", StringComparison.OrdinalIgnoreCase))
            {
                baseName = baseName[..^".Designer".Length];
            }

            var outputFolder = GetOutputFolder(designerPath, targetLanguageCode);
            var outputResx = Path.Combine(outputFolder, $"{baseName}.{targetLanguageCode}.resx");

            using var writer = new ResXResourceWriter(outputResx);
            foreach (var entry in formData.Entries)
            {
                var translated = translations.TryGetValue(entry.English, out var value)
                    ? value
                    : entry.English;

                writer.AddResource(entry.Key, translated);
            }

            writer.Generate();
            Log($"Wrote {Path.GetFileName(outputResx)} ({formData.Entries.Count} entries)");
        }

        private void WriteFolderCsvFiles(
            IReadOnlyList<FormTranslationData> formEntries,
            string targetLanguageCode,
            IReadOnlyDictionary<string, string> translations)
        {
            var folderGroups = formEntries
                .GroupBy(entry => Path.GetDirectoryName(entry.DesignerPath) ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase);

            foreach (var folderGroup in folderGroups)
            {
                var englishTexts = folderGroup
                    .SelectMany(form => form.Entries)
                    .Select(entry => entry.English)
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(text => text, StringComparer.Ordinal)
                    .ToList();

                if (englishTexts.Count == 0)
                {
                    continue;
                }

                var outputFolder = GetRunOutputFolder(folderGroup.Key, targetLanguageCode);
                var outputCsv = Path.Combine(outputFolder, "all_forms.translations.csv");
                var sb = new StringBuilder();
                sb.AppendLine("English,Translated");

                foreach (var english in englishTexts)
                {
                    var translated = translations.TryGetValue(english, out var value)
                        ? value
                        : english;

                    sb.Append(EscapeCsv(english));
                    sb.Append(',');
                    sb.Append(EscapeCsv(translated));
                    sb.AppendLine();
                }

                File.WriteAllText(outputCsv, sb.ToString(), Encoding.UTF8);
                Log($"Wrote {Path.GetFileName(outputCsv)} ({englishTexts.Count} rows) for {folderGroup.Key}");
            }
        }

        private string GetOutputFolder(string designerPath, string targetLanguageCode)
        {
            var baseFolder = Path.GetDirectoryName(designerPath) ?? string.Empty;
            return GetRunOutputFolder(baseFolder, targetLanguageCode);
        }

        private string GetRunOutputFolder(string baseFolder, string targetLanguageCode)
        {
            var timestamp = string.IsNullOrWhiteSpace(_currentRunTimestamp)
                ? DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
                : _currentRunTimestamp;
            var outputFolder = Path.Combine(baseFolder, "translations", targetLanguageCode, timestamp);
            Directory.CreateDirectory(outputFolder);
            return outputFolder;
        }

        private static string EscapeCsv(string value)
        {
            var text = value.Replace("\r\n", "\\n").Replace("\n", "\\n").Replace("\r", "\\n");
            var needsQuotes = text.Contains(',') || text.Contains('"') || text.Contains(' ');
            if (!needsQuotes)
            {
                return text;
            }

            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }

        private async Task WriteCodeBehindCsvAsync(string rootFolder, string targetLanguageCode)
        {
            var strings = _codeBehindLocalizationExporter
                .ExtractAllCodeBehindStrings(rootFolder)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(text => text, StringComparer.Ordinal)
                .ToList();

            if (strings.Count == 0)
            {
                Log("No code-behind utility strings found for CSV export.");
                return;
            }

            Log($"Translating {strings.Count} code-behind strings to {targetLanguageCode}...");
            var translations = await _translateService.TranslateAllAsync(strings, targetLanguageCode, () => { });

            var outputFolder = GetRunOutputFolder(rootFolder, targetLanguageCode);
            var outputCsv = Path.Combine(outputFolder, "codebehind_strings.translations.csv");

            var sb = new StringBuilder();
            sb.AppendLine("English,Translated");
            foreach (var english in strings)
            {
                var translated = translations.TryGetValue(english, out var value) ? value : english;
                sb.Append(EscapeCsv(english));
                sb.Append(',');
                sb.Append(EscapeCsv(translated));
                sb.AppendLine();
            }

            File.WriteAllText(outputCsv, sb.ToString(), Encoding.UTF8);
            Log($"Wrote {Path.GetFileName(outputCsv)} ({strings.Count} rows)");
        }

        private void Log(string message)
        {
            txtLog.AppendText($"{DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture)}  {message}{Environment.NewLine}");
        }

        private void UpdateProgress(int completed, int total)
        {
            var safeTotal = Math.Max(total, 1);
            var percent = (int)Math.Round(completed * 100.0 / safeTotal, MidpointRounding.AwayFromZero);
            percent = Math.Clamp(percent, 0, 100);

            progressBar1.Value = percent;
            lblProgress.Text = $"{percent}%";
        }

        private void ApplyDarkTheme()
        {
            BackColor = Surface;
            ForeColor = Foreground;
            txtLog.Font = new Font("Consolas", 10F, FontStyle.Regular, GraphicsUnit.Point);

            ApplyThemeToControl(this);
        }

        private void InitializeLanguageOptions()
        {
            cboLanguage.DisplayMember = nameof(LanguageOption.Name);
            cboLanguage.ValueMember = nameof(LanguageOption.Code);
            cboLanguage.Items.AddRange(LanguageOptions.Cast<object>().ToArray());
            cboLanguage.SelectedIndex = 0;
            UpdateTranslateButtonText(LanguageOptions[0].Code);
        }

        private LanguageOption GetSelectedLanguage()
        {
            if (cboLanguage.SelectedItem is LanguageOption selected)
            {
                return selected;
            }

            return LanguageOptions[0];
        }

        private void UpdateTranslateButtonText(string languageCode)
        {
            btnTranslate.Text = $"Translate to {languageCode}";
        }

        private void ApplyLoadedSettings()
        {
            if (!string.IsNullOrWhiteSpace(_settings.LastFolder) && Directory.Exists(_settings.LastFolder))
            {
                txtFolderPath.Text = _settings.LastFolder;
            }

            var savedCode = string.IsNullOrWhiteSpace(_settings.TargetLanguageCode)
                ? LanguageOptions[0].Code
                : _settings.TargetLanguageCode;

            var index = Array.FindIndex(LanguageOptions, option =>
                string.Equals(option.Code, savedCode, StringComparison.OrdinalIgnoreCase));

            _suppressLanguageChange = true;
            cboLanguage.SelectedIndex = index >= 0 ? index : 0;
            chkExportCsv.Checked = _settings.ExportBatchCsv;
            chkConvertCodeBehind.Checked = _settings.ConvertCodeBehind;
            _suppressLanguageChange = false;

            var selected = GetSelectedLanguage();
            _settings.TargetLanguageCode = selected.Code;
            UpdateTranslateButtonText(selected.Code);
            AppSettingsStore.Save(_settings);
        }

        private void ApplyThemeToControl(Control control)
        {
            switch (control)
            {
                case Form form:
                    form.BackColor = Surface;
                    form.ForeColor = Foreground;
                    break;
                case TextBox textBox:
                    textBox.BackColor = Input;
                    textBox.ForeColor = Foreground;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                    break;
                case Button button:
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderColor = Accent;
                    button.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 95, 160);
                    button.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 108, 182);
                    button.BackColor = Accent;
                    button.ForeColor = Color.White;
                    break;
                case Label label:
                    label.ForeColor = Foreground;
                    break;
                case ComboBox comboBox:
                    comboBox.BackColor = Input;
                    comboBox.ForeColor = Foreground;
                    comboBox.FlatStyle = FlatStyle.Flat;
                    break;
                case CheckBox checkBox:
                    checkBox.ForeColor = Foreground;
                    break;
                case ProgressBar progressBar:
                    progressBar.BackColor = Panel;
                    progressBar.ForeColor = Accent;
                    break;
                default:
                    control.BackColor = Panel;
                    control.ForeColor = Foreground;
                    break;
            }

            foreach (Control child in control.Controls)
            {
                ApplyThemeToControl(child);
            }
        }

        private static string DecodeCSharpString(string literal)
        {
            if (literal.Length < 2 || literal[0] != '"' || literal[^1] != '"')
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
                        '\\' => '\\',
                        '"' => '"',
                        'n' => '\n',
                        'r' => '\r',
                        't' => '\t',
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

    }
}
