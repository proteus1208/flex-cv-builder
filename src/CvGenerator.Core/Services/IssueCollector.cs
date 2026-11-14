using CvGenerator.Core.Models;

namespace CvGenerator.Core.Services;

public sealed class IssueCollector
{
    private readonly List<GenerationIssue> _issues = new();

    public IReadOnlyList<GenerationIssue> Issues => _issues;

    public void Clear() => _issues.Clear();

    public void Add(GenerationIssue issue) => _issues.Add(issue);

    public void Info(string code, string step, string message, string? path = null, string? detail = null)
        => Add(new GenerationIssue(IssueSeverity.Info, code, step, message, path, detail));

    public void Warning(string code, string step, string message, string? path = null, string? detail = null)
        => Add(new GenerationIssue(IssueSeverity.Warning, code, step, message, path, detail));

    public void Error(string code, string step, string message, string? path = null, string? detail = null)
        => Add(new GenerationIssue(IssueSeverity.Error, code, step, message, path, detail));

    public bool HasErrors => _issues.Any(i => i.Severity == IssueSeverity.Error);

    public bool HasErrorsInSteps(params string[] steps)
        => _issues.Any(i => i.Severity == IssueSeverity.Error && steps.Contains(i.Step));

    public GenerationStatus ResolveStatus(bool pdfAttempted, bool pdfSucceeded)
    {
        if (HasErrorsInSteps("ParseJson", "Folder", "Docx", "Notes", "JsonOut"))
            return GenerationStatus.Failed;

        if (pdfAttempted && !pdfSucceeded)
            return GenerationStatus.Partial;

        if (HasErrors)
            return GenerationStatus.Failed;

        return GenerationStatus.Success;
    }

    public string ToPlainText(IEnumerable<GenerationIssue>? source = null)
    {
        var items = source ?? _issues;
        return string.Join(Environment.NewLine,
            items.Select(i =>
            {
                var path = string.IsNullOrEmpty(i.Path) ? "" : $" [{i.Path}]";
                var detail = string.IsNullOrEmpty(i.Detail) ? "" : $" | {i.Detail}";
                return $"{i.Severity.ToString().ToUpperInvariant()}\t{i.Step}\t{i.Message}{path}{detail}";
            }));
    }
}
