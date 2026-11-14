namespace CvGenerator.Core.Models;

public enum IssueSeverity
{
    Info,
    Warning,
    Error
}

public sealed record GenerationIssue(
    IssueSeverity Severity,
    string Code,
    string Step,
    string Message,
    string? Path = null,
    string? Detail = null);

public enum GenerationStatus
{
    Ready,
    Success,
    Partial,
    Failed
}

public sealed class GenerationResult
{
    public GenerationStatus Status { get; init; } = GenerationStatus.Ready;
    public string? OutputDirectory { get; init; }
    public string? JsonPath { get; init; }
    public string? DocxPath { get; init; }
    public string? PdfPath { get; init; }
    public string? NotesPath { get; init; }
    public IReadOnlyList<GenerationIssue> Issues { get; init; } = Array.Empty<GenerationIssue>();
}
