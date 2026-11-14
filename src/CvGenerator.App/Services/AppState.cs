using System.IO;
using System.Text.Json;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;
using CvGenerator.Core.Settings;
using CvGenerator.Docx.Services;
using Microsoft.Win32;

namespace CvGenerator.App.Services;

public sealed class AppState
{
    private readonly CvJsonService _json;
    private readonly OutputFolderService _folders;
    private readonly TemplateCatalogService _templates;
    private readonly PlaceholderScanner _scanner;
    private readonly GenerationService _generation;
    private SettingsStore _settingsStore = null!;
    private AppSettings _settings = null!;

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
    private readonly PdfConversionService _pdf = new();
    public string? SelectedTemplateId { get; set; }
    public IReadOnlyList<PlaceholderInfo> Placeholders { get; private set; } = Array.Empty<PlaceholderInfo>();
    public string OutputPreview { get; private set; } = "";
    public IReadOnlyList<GenerationIssue> Issues { get; private set; } = Array.Empty<GenerationIssue>();
    public GenerationStatus Status { get; private set; } = GenerationStatus.Ready;
    public string? LastOutputDirectory { get; private set; }
    public bool IsBusy { get; private set; }
    public string IssueFilter { get; set; } = "All";

    public void Initialize(AppSettings settings, SettingsStore store, string templatesRoot)
    {
        _settings = settings;
        _settingsStore = store;
        OutputRoot = settings.OutputRoot;
        TemplatesRoot = string.IsNullOrWhiteSpace(settings.TemplatesRoot) ? templatesRoot : settings.TemplatesRoot;
        LibreOfficePath = settings.LibreOfficePath;
        RefreshLibreOfficeStatus();
        RefreshTemplates();
        RefreshPreview();
        Notify();
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
    }

    public void RefreshPlaceholders()
    {
        var path = _templates.ResolvePath(TemplatesRoot, SelectedTemplateId);
        if (path is null)
        {
            Placeholders = Array.Empty<PlaceholderInfo>();
            return;
        }

        try
        {
            Placeholders = _scanner.ScanFile(path);
        }
        catch
        {
            Placeholders = Array.Empty<PlaceholderInfo>();
        }
    }

    public void RefreshPreview()
    {
        try
        {
            using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(JsonText) ? "{}" : JsonText);
            OutputPreview = _folders.BuildPreviewPath(OutputRoot, doc.RootElement);
        }
        catch
        {
            OutputPreview = Path.Combine(OutputRoot, DateTime.Now.ToString("yyyy-MM-dd"), "(invalid JSON)");
        }
    }

    public void SetOutputRoot(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        OutputRoot = path;
        _settings.OutputRoot = path;
        _settingsStore.Save(_settings);
        Directory.CreateDirectory(path);
        RefreshPreview();
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
        RefreshPreview();
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
            RefreshPreview();
            Notify();
        }
        catch (Exception ex)
        {
            Issues = new[]
            {
                new GenerationIssue(IssueSeverity.Error, "InvalidJson", "ParseJson", ex.Message)
            };
            Status = GenerationStatus.Failed;
            Notify();
        }
    }

    public void SelectTemplate(string? id)
    {
        SelectedTemplateId = id;
        RefreshPlaceholders();
        Notify();
    }

    public void ClearIssues()
    {
        Issues = Array.Empty<GenerationIssue>();
        if (Status is GenerationStatus.Failed or GenerationStatus.Partial or GenerationStatus.Success)
            Status = GenerationStatus.Ready;
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
            }
          ]
        }
        """;
}
