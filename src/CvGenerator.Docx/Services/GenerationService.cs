using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;

namespace CvGenerator.Docx.Services;

public sealed class GenerationRequest
{
    public required string JsonText { get; init; }
    public required string TemplatePath { get; init; }
    public required string OutputRoot { get; init; }
    public string? LibreOfficePath { get; init; }
}

public sealed class GenerationService
{
    private readonly CvJsonService _json = new();
    private readonly OutputFolderService _folders = new();
    private readonly NotesExportService _notes = new();
    private readonly DocxTemplateEngine _docx = new();
    private readonly PdfConversionService _pdf = new();
    private readonly TemplateSyntaxValidator _templateValidator = new();

    public GenerationResult GenerateAll(GenerationRequest request, IssueCollector issues)
    {
        issues.Clear();

        if (!_json.TryParse(request.JsonText, issues, out var document) || document is null)
        {
            return new GenerationResult
            {
                Status = GenerationStatus.Failed,
                Issues = issues.Issues.ToList()
            };
        }

        using (document)
        {
            var root = document.RootElement;

            if (string.IsNullOrWhiteSpace(request.TemplatePath) || !File.Exists(request.TemplatePath))
            {
                issues.Error("TemplateMissing", "Docx", "DOCX template file was not found.", request.TemplatePath);
                return Failed(issues);
            }

            foreach (var issue in _templateValidator.ValidateFile(request.TemplatePath))
                issues.Add(issue);

            if (issues.HasErrorsInSteps("Template"))
                return Failed(issues);

            var outputDir = _folders.CreateOutputDirectory(request.OutputRoot, root, issues);
            if (string.IsNullOrEmpty(outputDir) || issues.HasErrorsInSteps("Folder"))
                return Failed(issues);

            var notesPath = _notes.WriteNotes(root, outputDir, issues);

            var docxPath = Path.Combine(outputDir, "resume.docx");
            _docx.Generate(request.TemplatePath, docxPath, root, issues);

            string? jsonPath = null;
            if (!issues.HasErrorsInSteps("Docx"))
                jsonPath = _json.WriteResumeJson(root, outputDir, issues);
            else
            {
                issues.Info("PdfSkipped", "Pdf", "PDF skipped because DOCX generation failed.");
                return new GenerationResult
                {
                    Status = GenerationStatus.Failed,
                    OutputDirectory = outputDir,
                    NotesPath = string.IsNullOrEmpty(notesPath) ? null : notesPath,
                    Issues = issues.Issues.ToList()
                };
            }

            var pdfPath = _pdf.ConvertToPdf(docxPath, outputDir, request.LibreOfficePath, issues);
            var pdfOk = !string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath);
            var status = issues.ResolveStatus(pdfAttempted: true, pdfSucceeded: pdfOk);

            return new GenerationResult
            {
                Status = status,
                OutputDirectory = outputDir,
                JsonPath = string.IsNullOrEmpty(jsonPath) ? null : jsonPath,
                DocxPath = File.Exists(docxPath) ? docxPath : null,
                PdfPath = pdfOk ? pdfPath : null,
                NotesPath = string.IsNullOrEmpty(notesPath) ? null : notesPath,
                Issues = issues.Issues.ToList()
            };
        }
    }

    public static void OpenDirectory(string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            return;

        try
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = directory,
                    UseShellExecute = true
                });
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                Process.Start("open", directory);
            }
            else
            {
                Process.Start("xdg-open", directory);
            }
        }
        catch
        {
            // ignore open failures
        }
    }

    private static GenerationResult Failed(IssueCollector issues) => new()
    {
        Status = GenerationStatus.Failed,
        Issues = issues.Issues.ToList()
    };
}
