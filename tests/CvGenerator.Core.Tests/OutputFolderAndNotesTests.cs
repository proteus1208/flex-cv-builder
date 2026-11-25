using System.Text.Json;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;
using CvGenerator.Core.Settings;

namespace CvGenerator.Core.Tests;

public class OutputFolderAndNotesTests
{
    [Fact]
    public void Builds_path_from_tokyo_date_and_directory()
    {
        using var doc = JsonDocument.Parse("""
        {
          "directory": "CODY LEE BERRY_Volantsoft Inc_C# .NET REST API SQL",
          "name": "Cody Lee Berry",
          "role": "Senior Engineer"
        }
        """);
        var folders = new OutputFolderService();
        var when = new DateTimeOffset(2026, 11, 21, 14, 30, 5, TimeSpan.FromHours(9));
        var relative = folders.BuildRelativeOutputPath(doc.RootElement, when, out var fallback);
        Assert.False(fallback);
        Assert.Equal(Path.Combine("2026-11-21", "CODY LEE BERRY_Volantsoft Inc_C# .NET REST API SQL"), relative);
    }

    [Fact]
    public void Builds_path_from_tokyo_date_and_stamp_when_directory_missing()
    {
        using var doc = JsonDocument.Parse("""{"name":"Cody Lee Berry"}""");
        var folders = new OutputFolderService();
        var when = new DateTimeOffset(2026, 11, 21, 14, 30, 5, TimeSpan.FromHours(9));
        var relative = folders.BuildRelativeOutputPath(doc.RootElement, when, out var fallback);
        Assert.True(fallback);
        Assert.Equal(Path.Combine("2026-11-21", "14-30-05"), relative);
    }

    [Fact]
    public void Preview_path_nests_under_output_root_date_folder()
    {
        using var doc = JsonDocument.Parse("""{"directory":"Demo"}""");
        var folders = new OutputFolderService();
        var when = new DateTimeOffset(2026, 11, 25, 9, 5, 1, TimeSpan.FromHours(9));
        var root = Path.Combine(Path.GetTempPath(), "cv-out");
        var preview = folders.BuildPreviewPath(root, doc.RootElement, when);
        Assert.Equal(Path.Combine(root, "2026-11-25", "Demo"), preview);
    }

    [Fact]
    public void CreateOutputDirectory_creates_date_then_directory()
    {
        var root = Path.Combine(Path.GetTempPath(), "cvgen-out-" + Guid.NewGuid().ToString("N"));
        try
        {
            using var doc = JsonDocument.Parse("""{"directory":"CODY LEE BERRY_Volantsoft Inc"}""");
            var issues = new IssueCollector();
            var when = new DateTimeOffset(2026, 11, 26, 12, 0, 0, TimeSpan.FromHours(9));
            var created = new OutputFolderService().CreateOutputDirectory(root, doc.RootElement, issues, when);

            Assert.True(Directory.Exists(created));
            Assert.Equal(Path.Combine(root, "2026-11-26", "CODY LEE BERRY_Volantsoft Inc"), created);
            Assert.Equal("2026-11-26", Directory.GetParent(created)!.Name);
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, true);
        }
    }

    [Fact]
    public void Writes_notes_txt_and_warns_when_missing()
    {
        var dir = Path.Combine(Path.GetTempPath(), "cvgen-notes-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            using var doc = JsonDocument.Parse("""{"name":"A"}""");
            var issues = new IssueCollector();
            var path = new NotesExportService().WriteNotes(doc.RootElement, dir, issues);
            Assert.True(File.Exists(path));
            Assert.Equal(string.Empty, File.ReadAllText(path));
            Assert.Contains(issues.Issues, i => i.Code == "MissingNotes" && i.Severity == IssueSeverity.Warning);

            using var withNotes = JsonDocument.Parse("""{"notes":"hello"}""");
            var issues2 = new IssueCollector();
            var path2 = new NotesExportService().WriteNotes(withNotes.RootElement, dir, issues2);
            Assert.Equal("hello", File.ReadAllText(path2));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void Settings_persist_output_root()
    {
        var dir = Path.Combine(Path.GetTempPath(), "cvgen-settings-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var store = new SettingsStore(dir);
            store.Save(new AppSettings
            {
                OutputRoot = @"D:\CV-Out",
                TemplatesRoot = @"D:\Templates",
                LastJson = """{"directory":"demo"}"""
            });
            var loaded = store.Load("fallback-out", "fallback-tpl");
            Assert.Equal(@"D:\CV-Out", loaded.OutputRoot);
            Assert.Equal(@"D:\Templates", loaded.TemplatesRoot);
            Assert.Equal("""{"directory":"demo"}""", loaded.LastJson);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void IssueCollector_partial_when_pdf_fails()
    {
        var issues = new IssueCollector();
        issues.Warning("UnresolvedPlaceholder", "Docx", "missing phone");
        issues.Error("LibreOfficeMissing", "Pdf", "no soffice");
        Assert.Equal(GenerationStatus.Partial, issues.ResolveStatus(pdfAttempted: true, pdfSucceeded: false));
    }
}
