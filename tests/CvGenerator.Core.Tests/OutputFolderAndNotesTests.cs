using System.Text.Json;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;
using CvGenerator.Core.Settings;

namespace CvGenerator.Core.Tests;

public class OutputFolderAndNotesTests
{
    [Fact]
    public void Builds_folder_name_from_directory_field()
    {
        using var doc = JsonDocument.Parse("""
        {
          "directory": "Acme-Backend-Role",
          "name": "Cody Lee Berry",
          "role": "Senior Engineer"
        }
        """);
        var folders = new OutputFolderService();
        var leaf = folders.BuildLeafFolderName(doc.RootElement, out var fallback);
        Assert.False(fallback);
        Assert.Equal("Acme-Backend-Role", leaf);
    }

    [Fact]
    public void Builds_folder_name_from_datetime_when_directory_missing()
    {
        using var doc = JsonDocument.Parse("""{"name":"Cody Lee Berry"}""");
        var folders = new OutputFolderService();
        var when = new DateTime(2026, 11, 21, 14, 30, 5);
        var leaf = folders.BuildLeafFolderName(doc.RootElement, when, out var fallback);
        Assert.True(fallback);
        Assert.Equal("2026-11-21_14-30-05", leaf);
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
            store.Save(new AppSettings { OutputRoot = @"D:\CV-Out", TemplatesRoot = @"D:\Templates" });
            var loaded = store.Load("fallback-out", "fallback-tpl");
            Assert.Equal(@"D:\CV-Out", loaded.OutputRoot);
            Assert.Equal(@"D:\Templates", loaded.TemplatesRoot);
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
