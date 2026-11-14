using System.IO;
using System.Net;
using System.Text.Json;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;
using CvGenerator.Core.Settings;
using CvGenerator.Docx.Services;
using Microsoft.Win32;

namespace CvGenerator.App.Services;

public sealed class AppState : IDisposable
{
    private readonly CvJsonService _json;
    private readonly OutputFolderService _folders;
    private readonly TemplateCatalogService _templates;
    private readonly PlaceholderScanner _scanner;
    private readonly GenerationService _generation;
    private readonly TemplateSyntaxValidator _templateValidator = new();
    private readonly LiveJsonValidator _liveValidator = new();
    private readonly PdfConversionService _pdf = new();
    private SettingsStore _settingsStore = null!;
    private AppSettings _settings = null!;
    private IReadOnlyList<GenerationIssue> _templateIssues = Array.Empty<GenerationIssue>();
    private FileSystemWatcher? _templateWatcher;
    private readonly object _watchLock = new();
    private DateTime _lastWatchReload = DateTime.MinValue;

    public AppState(
        CvJsonService json,
        OutputFolderService folders,
        TemplateCatalogService templates,
        PlaceholderScanner scanner,
        GenerationService generation)
    {
        _json = json;
        _folders = folders;
        _templates = templates;
        _scanner = scanner;
        _generation = generation;
    }

    public event Action? Changed;

    public string JsonText { get; set; } = SampleJson;
    public string OutputRoot { get; private set; } = "";
    public string TemplatesRoot { get; private set; } = "";
    public string? LibreOfficePath { get; private set; }
    public string? ResolvedLibreOfficePath { get; private set; }
    public bool IsLibreOfficeReady => !string.IsNullOrEmpty(ResolvedLibreOfficePath);
    public IReadOnlyList<TemplateInfo> Templates { get; private set; } = Array.Empty<TemplateInfo>();
    public string? SelectedTemplateId { get; set; }
    public IReadOnlyList<PlaceholderInfo> Placeholders { get; private set; } = Array.Empty<PlaceholderInfo>();
    public string OutputPreview { get; private set; } = "";
    public string JsonHighlightHtml { get; private set; } = WebUtility.HtmlEncode(SampleJson);
    public IReadOnlyList<JsonMark> JsonMarks { get; private set; } = Array.Empty<JsonMark>();
    public IReadOnlyList<GenerationIssue> Issues { get; private set; } = Array.Empty<GenerationIssue>();
    public GenerationStatus Status { get; private set; } = GenerationStatus.Ready;
    public string? LastOutputDirectory { get; private set; }
    public bool IsBusy { get; private set; }
    public string IssueFilter { get; set; } = "All";
    public bool IsCopyArmed { get; private set; } = true;
    public bool ShowCopyCheckIcon { get; private set; }
    public IReadOnlyList<PlaceholderBinding> PlaceholderBindings { get; private set; } = Array.Empty<PlaceholderBinding>();

    public void Initialize(AppSettings settings, SettingsStore store, string templatesRoot)
    {
        _settings = settings;
        _settingsStore = store;
        OutputRoot = settings.OutputRoot;
        TemplatesRoot = string.IsNullOrWhiteSpace(settings.TemplatesRoot) ? templatesRoot : settings.TemplatesRoot;
        LibreOfficePath = settings.LibreOfficePath;
        RefreshLibreOfficeStatus();
        RefreshTemplates();
        RefreshPreviewAndLiveValidation();
        StartTemplateWatcher();
        Notify();
    }

    public void Dispose()
    {
        StopTemplateWatcher();
    }

    private void StartTemplateWatcher()
    {
        StopTemplateWatcher();
        if (string.IsNullOrWhiteSpace(TemplatesRoot) || !Directory.Exists(TemplatesRoot))
            return;

        try
        {
            _templateWatcher = new FileSystemWatcher(TemplatesRoot)
            {
                Filter = "*.docx",
                IncludeSubdirectories = true,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.Size
            };
            _templateWatcher.Changed += OnTemplateFileChanged;
            _templateWatcher.Created += OnTemplateFileChanged;
            _templateWatcher.Deleted += OnTemplateFileChanged;
            _templateWatcher.Renamed += OnTemplateFileRenamed;
            _templateWatcher.EnableRaisingEvents = true;
        }
        catch
        {
            StopTemplateWatcher();
        }
    }

    private void StopTemplateWatcher()
    {
        if (_templateWatcher is null)
            return;
        _templateWatcher.EnableRaisingEvents = false;
        _templateWatcher.Changed -= OnTemplateFileChanged;
        _templateWatcher.Created -= OnTemplateFileChanged;
        _templateWatcher.Deleted -= OnTemplateFileChanged;
        _templateWatcher.Renamed -= OnTemplateFileRenamed;
        _templateWatcher.Dispose();
        _templateWatcher = null;
    }

    private void OnTemplateFileRenamed(object sender, RenamedEventArgs e) => ReloadFromWatcher();

    private void OnTemplateFileChanged(object sender, FileSystemEventArgs e) => ReloadFromWatcher();

    private void ReloadFromWatcher()
    {
        lock (_watchLock)
        {
            // Debounce Word temp saves (~$ files + burst writes).
            if ((DateTime.UtcNow - _lastWatchReload).TotalMilliseconds < 400)
                return;
            _lastWatchReload = DateTime.UtcNow;
        }

        try
        {
            // Slight delay so the file is fully written.
            Task.Run(async () =>
            {
                await Task.Delay(350);
                RefreshTemplates();
                Notify();
            });
        }
        catch
        {
            // ignore
        }
    }

    public void RefreshLibreOfficeStatus()
    {
        ResolvedLibreOfficePath = _pdf.FindLibreOffice(LibreOfficePath);
    }

    public void RefreshTemplates()
    {
        Templates = _templates.ListTemplates(TemplatesRoot);
        if (SelectedTemplateId is null || Templates.All(t => t.Id != SelectedTemplateId))
            SelectedTemplateId = Templates.FirstOrDefault()?.Id;
        RefreshPlaceholders();
        RefreshPreviewAndLiveValidation();
    }

    public void RefreshPlaceholders()
    {
        var path = _templates.ResolvePath(TemplatesRoot, SelectedTemplateId);
        if (path is null)
        {
            Placeholders = Array.Empty<PlaceholderInfo>();
            _templateIssues = Array.Empty<GenerationIssue>();
            return;
        }

        try
        {
            Placeholders = _scanner.ScanFile(path);
            _templateIssues = _templateValidator.ValidateFile(path);
        }
        catch
        {
            Placeholders = Array.Empty<PlaceholderInfo>();
            _templateIssues = Array.Empty<GenerationIssue>();
        }
    }

    public void RefreshPreviewAndLiveValidation()
    {
        JsonElement? root = null;
        try
        {
            using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(JsonText) ? "{}" : JsonText);
            root = doc.RootElement.Clone();
            OutputPreview = _folders.BuildPreviewPath(OutputRoot, doc.RootElement);
        }
        catch
        {
            OutputPreview = Path.Combine(OutputRoot, DateTime.Now.ToString("yyyy-MM-dd"), "(invalid JSON)");
        }

        // Prefer scoped ValidationPath from the scanner (same alias name can mean
        // different arrays in different loops, e.g. skill under experience vs skills).
        var livePlaceholders = Placeholders.Select(p =>
            (p.Kind, (string?)(p.ValidationPath ?? p.JsonPath), p.Token));

        var live = _liveValidator.Validate(JsonText, livePlaceholders, _templateIssues);

        JsonHighlightHtml = live.HighlightHtml;
        JsonMarks = live.Marks;
        PlaceholderBindings = BuildPlaceholderBindings(root);
        if (!IsBusy)
        {
            Issues = live.Issues;
            Status = live.Issues.Any(i => i.Severity == IssueSeverity.Error)
                ? GenerationStatus.Failed
                : GenerationStatus.Ready;
        }
    }

    private IReadOnlyList<PlaceholderBinding> BuildPlaceholderBindings(JsonElement? root)
    {
        var resolver = new JsonPathResolver();
        var list = new List<PlaceholderBinding>();

        foreach (var p in Placeholders)
        {
            if (p.Kind is "array-start" or "array-end")
            {
                list.Add(new PlaceholderBinding(p, "marker", p.Kind == "array-start" ? "loop start" : "loop end"));
                continue;
            }

            if (root is null)
            {
                list.Add(new PlaceholderBinding(p, "missing", "invalid JSON"));
                continue;
            }

            var data = root.Value;
            var path = p.ValidationPath ?? p.JsonPath ?? "";
            var kind = p.Kind;

            if (kind == "loop-field" && !string.IsNullOrWhiteSpace(path))
            {
                list.Add(BindLoopField(p, path, data, resolver));
                continue;
            }

            // simple field
            if (string.IsNullOrWhiteSpace(path) || !resolver.TryGetElement(data, path, out var el))
            {
                list.Add(new PlaceholderBinding(p, "missing", "missing"));
                continue;
            }

            if (resolver.TryResolveValue(data, path, out var value))
            {
                if (string.IsNullOrWhiteSpace(value))
                    list.Add(new PlaceholderBinding(p, "empty", "empty"));
                else
                    list.Add(new PlaceholderBinding(p, "ok", TrimPreview(value)));
            }
            else
            {
                list.Add(new PlaceholderBinding(p, "missing", "missing"));
            }
        }

        return list;
    }

    private static PlaceholderBinding BindLoopField(
        PlaceholderInfo p,
        string loopPath,
        JsonElement root,
        JsonPathResolver resolver)
    {
        // skills_i.category | skills_i.detail_i | skills_i.detail_i.name
        if (!TrySplitLoopPath(loopPath, out var chain, out var field))
            return new PlaceholderBinding(p, "missing", "bad loop field");

        if (!resolver.TryGetElement(root, chain[0], out var cursor) || cursor.ValueKind != JsonValueKind.Array)
            return new PlaceholderBinding(p, "missing", $"no array '{chain[0]}'");

        var levelItems = cursor.EnumerateArray().ToList();
        if (levelItems.Count == 0)
            return new PlaceholderBinding(p, "empty", "array empty");

        // Descend parent arrays using the first item at each level; sample the innermost array.
        var focus = levelItems[0];
        for (var i = 1; i < chain.Count; i++)
        {
            if (!resolver.TryGetElement(focus, chain[i], out var next) || next.ValueKind != JsonValueKind.Array)
                return new PlaceholderBinding(p, "missing", $"no nested '{chain[i]}'");

            levelItems = next.EnumerateArray().ToList();
            if (levelItems.Count == 0)
                return new PlaceholderBinding(p, "empty", $"'{chain[i]}' empty");

            focus = levelItems[0];
        }

        var values = new List<string>();
        var anyMissing = false;
        foreach (var item in levelItems.Take(5))
        {
            if (string.IsNullOrEmpty(field))
            {
                var scalar = item.ValueKind == JsonValueKind.String ? item.GetString() : item.ToString();
                if (string.IsNullOrWhiteSpace(scalar)) anyMissing = true;
                else values.Add(TrimPreview(scalar!));
            }
            else if (resolver.TryResolveValue(item, field, out var v) && !string.IsNullOrWhiteSpace(v))
                values.Add(TrimPreview(v));
            else
                anyMissing = true;
        }

        if (values.Count == 0)
            return new PlaceholderBinding(p, "missing", string.IsNullOrEmpty(field) ? "no values" : $"no '{field}'");
        if (anyMissing)
            return new PlaceholderBinding(p, "partial", string.Join(" · ", values) + (levelItems.Count > 5 ? "…" : ""));
        return new PlaceholderBinding(p, "ok", string.Join(" · ", values) + (levelItems.Count > 5 ? "…" : ""));
    }

    private static bool TrySplitLoopPath(string path, out List<string> chain, out string? field)
    {
        chain = new List<string>();
        field = null;
        var s = path.Trim();
        var i = 0;
        while (i < s.Length)
        {
            var idx = s.IndexOf("_i", i, StringComparison.Ordinal);
            if (idx < 0)
                break;
            chain.Add(s[i..idx]);
            i = idx + 2;
            if (i < s.Length && s[i] == '.')
                i++;
            else if (i >= s.Length)
                return chain.Count > 0;
            else
                break;
        }

        if (chain.Count == 0)
            return false;
        field = i < s.Length ? s[i..] : null;
        if (string.IsNullOrEmpty(field))
            field = null;
        return true;
    }

    private static string TrimPreview(string value)
    {
        var oneLine = value.Replace("\r", " ").Replace("\n", " ").Trim();
        return oneLine.Length <= 48 ? oneLine : oneLine[..45] + "…";
    }

    public void SetOutputRoot(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        OutputRoot = path;
        _settings.OutputRoot = path;
        _settingsStore.Save(_settings);
        Directory.CreateDirectory(path);
        RefreshPreviewAndLiveValidation();
        Notify();
    }

    public void BrowseOutputRoot()
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Select output root directory",
            InitialDirectory = Directory.Exists(OutputRoot) ? OutputRoot : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };
        if (dlg.ShowDialog() == true)
            SetOutputRoot(dlg.FolderName);
    }

    public void BrowseLibreOffice()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select LibreOffice soffice.exe (usually …\\LibreOffice\\program\\soffice.exe)",
            Filter = "LibreOffice (soffice.exe)|soffice.exe|Executables (*.exe)|*.exe|All files (*.*)|*.*",
            InitialDirectory = @"C:\Program Files\LibreOffice\program"
        };
        if (!Directory.Exists(dlg.InitialDirectory))
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);

        if (dlg.ShowDialog() != true)
            return;

        LibreOfficePath = dlg.FileName;
        _settings.LibreOfficePath = LibreOfficePath;
        _settingsStore.Save(_settings);
        RefreshLibreOfficeStatus();
        Notify();
    }

    public void ClearLibreOfficePath()
    {
        LibreOfficePath = null;
        _settings.LibreOfficePath = null;
        _settingsStore.Save(_settings);
        RefreshLibreOfficeStatus();
        Notify();
    }

    public void OpenLibreOfficeDownload()
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "https://www.libreoffice.org/download/download-libreoffice/",
                UseShellExecute = true
            });
        }
        catch
        {
            // ignore
        }
    }

    public void OpenJsonFile()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            Title = "Open CV JSON"
        };
        if (dlg.ShowDialog() != true)
            return;
        JsonText = File.ReadAllText(dlg.FileName);
        RefreshPreviewAndLiveValidation();
        Notify();
    }

    public void ImportTemplate()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Word documents (*.docx)|*.docx",
            Title = "Import DOCX template"
        };
        if (dlg.ShowDialog() != true)
            return;
        _templates.ImportTemplate(TemplatesRoot, dlg.FileName);
        RefreshTemplates();
        Notify();
    }

    public void FormatJson()
    {
        try
        {
            JsonText = _json.Format(JsonText);
            RefreshPreviewAndLiveValidation();
            Notify();
        }
        catch (Exception ex)
        {
            Issues = new[]
            {
                new GenerationIssue(IssueSeverity.Error, "InvalidJson", "ParseJson", ex.Message)
            };
            Status = GenerationStatus.Failed;
            JsonHighlightHtml = "";
            Notify();
        }
    }

    public void SelectTemplate(string? id)
    {
        SelectedTemplateId = id;
        RefreshPlaceholders();
        RefreshPreviewAndLiveValidation();
        Notify();
    }

    public void ClearIssues()
    {
        RefreshPreviewAndLiveValidation();
        Notify();
    }

    public string GetFilteredIssuesText()
    {
        var items = FilteredIssues();
        var collector = new IssueCollector();
        foreach (var i in items)
            collector.Add(i);
        return collector.ToPlainText();
    }

    public IEnumerable<GenerationIssue> FilteredIssues()
    {
        return IssueFilter switch
        {
            "Errors" => Issues.Where(i => i.Severity == IssueSeverity.Error),
            "Warnings" => Issues.Where(i => i.Severity == IssueSeverity.Warning),
            _ => Issues
        };
    }

    public void CopyOutputDirectoryPath()
    {
        if (!IsCopyArmed)
            return;

        var path = !string.IsNullOrWhiteSpace(LastOutputDirectory)
            ? LastOutputDirectory
            : OutputPreview;

        try
        {
            System.Windows.Clipboard.SetText(path);
            IsCopyArmed = false;
            ShowCopyCheckIcon = true;
            Notify();
            _ = ResetCopyButtonLater();
        }
        catch
        {
            // keep copy icon; no text toast
            IsCopyArmed = true;
            ShowCopyCheckIcon = false;
            Notify();
        }
    }

    private async Task ResetCopyButtonLater()
    {
        await Task.Delay(2000);
        IsCopyArmed = true;
        ShowCopyCheckIcon = false;
        Notify();
    }

    public async Task GenerateAllAsync()
    {
        if (IsBusy)
            return;

        IsBusy = true;
        Status = GenerationStatus.Ready;
        Issues = Array.Empty<GenerationIssue>();
        Notify();

        try
        {
            await Task.Run(() =>
            {
                var issues = new IssueCollector();
                var templatePath = _templates.ResolvePath(TemplatesRoot, SelectedTemplateId);
                var result = _generation.GenerateAll(new GenerationRequest
                {
                    JsonText = JsonText,
                    TemplatePath = templatePath ?? string.Empty,
                    OutputRoot = OutputRoot,
                    LibreOfficePath = LibreOfficePath
                }, issues);

                Issues = result.Issues;
                Status = result.Status;
                LastOutputDirectory = result.OutputDirectory;
            });
        }
        finally
        {
            IsBusy = false;
            Notify();
        }
    }

    public void OpenLastOutput()
    {
        GenerationService.OpenDirectory(LastOutputDirectory);
    }

    private void Notify() => Changed?.Invoke();

    private const string SampleJson =
        """
        {
          "name": "Cody Lee Berry",
          "role": "Senior Engineer",
          "linkedin": "https://linkedin.com/in/example",
          "topSkills": ["React", "Node.js", "C#"],
          "notes": "Tailor summary for Platform team. Emphasize API design.",
          "summary": "Engineer focused on reliable product delivery.",
          "experience": [
            {
              "title": "Senior Engineer",
              "company": "Example Corp",
              "years": "2022-Present",
              "description": "Led resume tooling and document pipelines."
            },
            {
              "title": "Engineer",
              "company": "Prior Co",
              "years": "2019-2022",
              "description": "Built APIs and frontend features."
            }
          ]
        }
        """;
}

public sealed record PlaceholderBinding(
    PlaceholderInfo Placeholder,
    string State,   // ok | missing | empty | partial | marker
    string Display);
