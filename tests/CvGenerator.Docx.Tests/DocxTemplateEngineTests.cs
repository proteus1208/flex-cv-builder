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
    public void Replaces_simple_placeholders_including_split_runs_and_keeps_run_props()
    {
        var template = CreateTempDocx(doc =>
        {
            var run1 = new Run(new RunProperties(new Bold()), new Text("${"));
            var run2 = new Run(new RunProperties(new Bold()), new Text("name"));
            var run3 = new Run(new RunProperties(new Bold()), new Text("}"));
            var run4 = new Run(new Text(" / "));
            var run5 = new Run(new RunProperties(new Color { Val = "FF0000" }), new Text("${linkedin}"));
            doc.MainDocumentPart!.Document.Body!.AppendChild(new Paragraph(run1, run2, run3, run4, run5));
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

            var boldRun = result.MainDocumentPart.Document.Body.Descendants<Run>()
                .FirstOrDefault(r => r.InnerText.Contains("Ada"));
            Assert.NotNull(boldRun?.RunProperties?.Bold);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Expands_all_array_items_with_ArrayStart_End_and_arrayname_i()
    {
        var template = CreateTempDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("Header"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart experience}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${experience_i.title} at ${experience_i.company}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd experience}"))));
            body.AppendChild(new Paragraph(new Run(new Text("Footer"))));
        });

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "experience": [
                {"title":"Eng","company":"A"},
                {"title":"Lead","company":"B"},
                {"title":"Intern","company":"C"}
              ]
            }
            """);
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, json.RootElement, issues);

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Eng at A", text);
            Assert.Contains("Lead at B", text);
            Assert.Contains("Intern at C", text);
            Assert.DoesNotContain("ArrayStart", text);
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
    public void Nested_ArrayStart_expands_inner_arrays()
    {
        var template = CreateTempDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart experience}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${experience_i.title}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart highlights}"))));
            body.AppendChild(new Paragraph(new Run(new Text("- ${highlights_i.text}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd highlights}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd experience}"))));
        });

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "experience": [
                {"title":"A", "highlights":[{"text":"a1"},{"text":"a2"}]},
                {"title":"B", "highlights":[{"text":"b1"}]}
              ]
            }
            """);
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, json.RootElement, issues);

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("A", text);
            Assert.Contains("a1", text);
            Assert.Contains("a2", text);
            Assert.Contains("B", text);
            Assert.Contains("b1", text);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void ArrayStart_with_iterate_alias_uses_alias_instead_of_name_i()
    {
        var template = CreateTempDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart experience exp}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${exp.title} at ${exp.company}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd experience}"))));
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
            new DocxTemplateEngine().Generate(template, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Eng at A", text);
            Assert.Contains("Lead at B", text);
            Assert.DoesNotContain("exp.", text);
            Assert.DoesNotContain("ArrayStart", text);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Nested_ArrayStart_via_parent_alias()
    {
        var template = CreateTempDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart skills skill}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${skill.category}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart skill.detail d}"))));
            body.AppendChild(new Paragraph(new Run(new Text("- ${d.name}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd skill.detail}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart skill.tags tag}"))));
            body.AppendChild(new Paragraph(new Run(new Text("#${tag}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd skill.tags}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd skills}"))));
        });

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "skills": [
                {
                  "category": "Backend",
                  "detail": [ {"name":"C#"}, {"name":"SQL"} ],
                  "tags": ["api", "dotnet"]
                }
              ]
            }
            """);
            new DocxTemplateEngine().Generate(template, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Backend", text);
            Assert.Contains("- C#", text);
            Assert.Contains("- SQL", text);
            Assert.Contains("#api", text);
            Assert.Contains("#dotnet", text);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Scanner_rewrites_alias_fields_to_canonical_loop_paths()
    {
        var scanner = new PlaceholderScanner();
        var found = scanner.ScanText(
            "${ArrayStart experience exp} ${exp.title} ${ArrayEnd experience}");
        Assert.Contains(found, p => p.Kind == "array-start" && p.JsonPath == "experience" && p.Alias == "exp");
        Assert.Contains(found, p => p.Kind == "loop-field" && p.ValidationPath == "experience_i.title");
    }

    [Fact]
    public void Alias_normalizer_maps_nested_iterate_names_for_live_validation()
    {
        var scanner = new PlaceholderScanner();
        var found = scanner.ScanText("""
            ${ArrayStart skills skill}
            ${skill.category}
            ${ArrayStart skill.detail d}
            ${d.name}
            ${ArrayEnd skill.detail}
            ${ArrayEnd skills}
            """);

        Assert.Contains(found, p => p.Kind == "loop-field" && p.ValidationPath == "skills_i.category");
        Assert.Contains(found, p => p.Kind == "loop-field" && p.ValidationPath == "skills_i.detail_i.name");
        Assert.Contains(found, p => p.Kind == "array-start" && p.Alias == "d" && p.ValidationPath == "skills_i.detail");
    }

    [Fact]
    public void Reused_iterate_alias_resolves_to_correct_source_per_scope()
    {
        var scanner = new PlaceholderScanner();
        var found = scanner.ScanText("""
            ${ArrayStart experience exp}
            ${ArrayStart exp.skills skill}${skill}, ${ArrayEnd exp.skills}
            ${ArrayEnd experience}
            ${ArrayStart skills skill}
            ${skill.category}
            ${ArrayStart skill.detail detail}${detail}, ${ArrayEnd skill.detail}
            ${ArrayEnd skills}
            """);

        Assert.Contains(found, p => p.Token == "${skill}" && p.ValidationPath == "experience_i.skills_i");
        Assert.Contains(found, p => p.Token == "${skill.category}" && p.ValidationPath == "skills_i.category");
        Assert.Contains(found, p =>
            p.Kind == "array-start" &&
            p.JsonPath == "skill.detail" &&
            p.ValidationPath == "skills_i.detail");
    }

    [Fact]
    public void Parent_qualified_nested_ArrayStart_expands_object_and_scalar_items()
    {
        var template = CreateTempDocx(doc =>
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart skills}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${skills_i.category}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart skills_i.detail}"))));
            body.AppendChild(new Paragraph(new Run(new Text("- ${skills_i.detail_i.name} (${skills_i.detail_i.level})"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd skills_i.detail}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayStart skills_i.tags}"))));
            body.AppendChild(new Paragraph(new Run(new Text("#${skills_i.tags_i}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd skills_i.tags}"))));
            body.AppendChild(new Paragraph(new Run(new Text("${ArrayEnd skills}"))));
        });

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "skills": [
                {
                  "category": "Backend",
                  "detail": [
                    {"name":"C#","level":"expert"},
                    {"name":"SQL","level":"solid"}
                  ],
                  "tags": ["api", "dotnet"]
                },
                {
                  "category": "Frontend",
                  "detail": [
                    {"name":"React","level":"strong"}
                  ],
                  "tags": ["ui"]
                }
              ]
            }
            """);
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, json.RootElement, issues);

            using var result = WordprocessingDocument.Open(output, false);
            var text = result.MainDocumentPart!.Document.Body!.InnerText;
            Assert.Contains("Backend", text);
            Assert.Contains("C# (expert)", text);
            Assert.Contains("SQL (solid)", text);
            Assert.Contains("#api", text);
            Assert.Contains("#dotnet", text);
            Assert.Contains("Frontend", text);
            Assert.Contains("React (strong)", text);
            Assert.Contains("#ui", text);
            Assert.DoesNotContain("ArrayStart", text);
            Assert.DoesNotContain("skills_i", text);
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
    public void Scanner_detects_ArrayStart_tokens()
    {
        var scanner = new PlaceholderScanner();
        var found = scanner.ScanText("${name} ${ArrayStart exp} ${exp_i.title} ${ArrayEnd exp}");
        Assert.Contains(found, p => p.Kind == "simple" && p.JsonPath == "name");
        Assert.Contains(found, p => p.Kind == "array-start" && p.JsonPath == "exp");
        Assert.Contains(found, p => p.Kind == "loop-field");
        Assert.Contains(found, p => p.Kind == "array-end");
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
