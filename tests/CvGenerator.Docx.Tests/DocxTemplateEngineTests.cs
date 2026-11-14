using System.Text.Json;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Tests;

public class DocxTemplateEngineTests
{
    [Fact]
    public void Replaces_simple_placeholders_including_split_runs()
    {
        var template = CreateTempDocx(doc =>
        {
            // Split "${name}" across runs: "${" + "name" + "}"
            var p = new Paragraph(
                new Run(new Text("${")),
                new Run(new Text("name")),
                new Run(new Text("}")),
                new Run(new Text(" / ")),
                new Run(new Text("${linkedin}")));
            doc.MainDocumentPart!.Document.Body!.AppendChild(p);
        });

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""{"name":"Ada","linkedin":"https://li.test"}""");
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, json.RootElement, issues);

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Ada", text);
            Assert.Contains("https://li.test", text);
            Assert.DoesNotContain("${name}", text);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Expands_repeating_sections()
    {
        var template = CreateTempDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("Header"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ start experience}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${experience.i.title} at ${experience.i.company}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ end experience}"))));
            body.AppendChild(new Paragraph(new Run(new Text("Footer"))));
        });

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "experience": [
                {"title":"Eng","company":"A"},
                {"title":"Lead","company":"B"}
              ]
            }
            """);
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, json.RootElement, issues);

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Eng at A", text);
            Assert.Contains("Lead at B", text);
            Assert.DoesNotContain("start experience", text);
            Assert.Contains("Header", text);
            Assert.Contains("Footer", text);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Unresolved_placeholders_become_warnings()
    {
        var template = CreateTempDocx(doc =>
        {
            doc.MainDocumentPart!.Document.Body!.AppendChild(
                new Paragraph(new Run(new Text("Phone ${phone}"))));
        });
        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""{"name":"Ada"}""");
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, json.RootElement, issues);
            Assert.Contains(issues.Issues, i => i.Code == "UnresolvedPlaceholder" && i.Path == "phone");
            using var result = WordprocessingDocument.Open(output, false);
            Assert.Equal("Phone ", result.MainDocumentPart!.Document.Body!.InnerText);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Scanner_detects_tokens()
    {
        var scanner = new PlaceholderScanner();
        var found = scanner.ScanText("${name} ${ start exp} ${exp.i.title} ${ end exp}");
        Assert.Contains(found, p => p.Kind == "simple" && p.JsonPath == "name");
        Assert.Contains(found, p => p.Kind == "loop-start" && p.JsonPath == "exp");
        Assert.Contains(found, p => p.Kind == "loop-field");
        Assert.Contains(found, p => p.Kind == "loop-end");
    }

    [Fact]
    public void GenerateAll_writes_json_docx_notes_and_reports_pdf_partial()
    {
        var template = CreateTempDocx(doc =>
        {
            doc.MainDocumentPart!.Document.Body!.AppendChild(
                new Paragraph(new Run(new Text("${name} - ${role}"))));
        });
        var root = Path.Combine(Path.GetTempPath(), "cvgen-all-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        // Force PDF failure with a real path that is not a working LibreOffice binary.
        var fakeSoffice = Path.Combine(root, "fake-soffice");
        File.WriteAllText(fakeSoffice, "#!/bin/sh\nexit 1\n");
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(fakeSoffice, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);

        try
        {
            var issues = new IssueCollector();
            var result = new GenerationService().GenerateAll(new GenerationRequest
            {
                JsonText = """
                {
                  "name": "Cody Lee Berry",
                  "role": "Senior Engineer",
                  "topSkills": ["React", "Node.js", "C#"],
                  "notes": "note body"
                }
                """,
                TemplatePath = template,
                OutputRoot = root,
                LibreOfficePath = fakeSoffice
            }, issues);

            Assert.Equal(Core.Models.GenerationStatus.Partial, result.Status);
            Assert.True(File.Exists(result.DocxPath));
            Assert.True(File.Exists(result.JsonPath));
            Assert.True(File.Exists(result.NotesPath));
            Assert.Equal("note body", File.ReadAllText(result.NotesPath!));
            Assert.Contains(result.Issues, i => i.Step == "Pdf" && i.Severity == Core.Models.IssueSeverity.Error);
            Assert.Contains("Cody Lee Berry - Senior Engineer - React-Node.js-C#", result.OutputDirectory!);
        }
        finally
        {
            File.Delete(template);
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    private static string CreateTempDocx(Action<WordprocessingDocument> build)
    {
        var path = Path.Combine(Path.GetTempPath(), "cvgen-tpl-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            build(doc);
            main.Document.Save();
        }

        return path;
    }
}
