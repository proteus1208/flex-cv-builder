using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using CvGenerator.Core.Models;

namespace CvGenerator.Core.Services;

public sealed class LiveValidationResult
{
    public IReadOnlyList<GenerationIssue> Issues { get; init; } = Array.Empty<GenerationIssue>();
    public string HighlightHtml { get; init; } = "";
    public IReadOnlyList<JsonMark> Marks { get; init; } = Array.Empty<JsonMark>();
}

public sealed record JsonMark(string Path, IssueSeverity Severity, string Reason);

public sealed class LiveJsonValidator
{
    private readonly JsonPathResolver _paths = new();

    public LiveValidationResult Validate(
        string jsonText,
        IEnumerable<(string Kind, string? JsonPath, string Token)> placeholders,
        IEnumerable<GenerationIssue> templateIssues)
    {
        var issues = new IssueCollector();
        foreach (var ti in templateIssues)
            issues.Add(ti);

        var marks = new List<JsonMark>();
        var placeholderList = placeholders.ToList();

        if (string.IsNullOrWhiteSpace(jsonText))
        {
            issues.Error("InvalidJson", "ParseJson", "JSON is empty.");
            return new LiveValidationResult
            {
                Issues = issues.Issues.ToList(),
                HighlightHtml = "",
                Marks = marks
            };
        }

        JsonDocument? doc = null;
        try
        {
            doc = JsonDocument.Parse(jsonText);
        }
        catch (JsonException ex)
        {
            issues.Error("InvalidJson", "ParseJson", $"Invalid JSON: {ex.Message}", detail: ex.ToString());
            return new LiveValidationResult
            {
                Issues = issues.Issues.ToList(),
                HighlightHtml = BuildHighlight(jsonText, marks),
                Marks = marks
            };
        }

        using (doc)
        {
            var root = doc.RootElement;

            foreach (var p in placeholderList)
            {
                if (string.IsNullOrWhiteSpace(p.JsonPath))
                    continue;

                if (p.Kind == "simple")
                {
                    if (!_paths.TryGetElement(root, p.JsonPath, out var el))
                    {
                        issues.Warning("MissingField", "ResolvePaths", $"Missing JSON field for {p.Token}.", p.JsonPath);
                        marks.Add(new JsonMark(p.JsonPath, IssueSeverity.Warning, "missing"));
                    }
                    else if (IsEmpty(el))
                    {
                        issues.Warning("EmptyField", "ResolvePaths", $"JSON field '{p.JsonPath}' is empty.", p.JsonPath);
                        marks.Add(new JsonMark(p.JsonPath, IssueSeverity.Warning, "empty"));
                    }
                }
                else if (p.Kind == "loop-start")
                {
                    var items = _paths.GetArrayItems(root, p.JsonPath, out _, out var missing);
                    if (missing)
                    {
                        issues.Warning(
                            "MissingLoopArray",
                            "ResolvePaths",
                            $"Loop array '{p.JsonPath}' is missing — output will be empty for that section.",
                            p.JsonPath);
                        marks.Add(new JsonMark(p.JsonPath, IssueSeverity.Warning, "missing"));
                    }
                    else if (items.Count == 0)
                    {
                        issues.Warning(
                            "EmptyLoopArray",
                            "ResolvePaths",
                            $"Loop array '{p.JsonPath}' is empty — output will be empty for that section.",
                            p.JsonPath);
                        marks.Add(new JsonMark(p.JsonPath, IssueSeverity.Warning, "empty"));
                    }
                }
                else if (p.Kind == "loop-field")
                {
                    // path like experience.i.title → check experience exists; field presence checked per item lightly
                    var idx = p.JsonPath.IndexOf(".i.", StringComparison.Ordinal);
                    if (idx <= 0)
                        continue;
                    var arrayPath = p.JsonPath[..idx];
                    var field = p.JsonPath[(idx + 3)..];
                    var items = _paths.GetArrayItems(root, arrayPath, out _, out var missing);
                    if (missing || items.Count == 0)
                        continue;

                    for (var i = 0; i < items.Count; i++)
                    {
                        if (!_paths.TryGetElement(items[i], field, out var el) || IsEmpty(el))
                        {
                            var markPath = $"{arrayPath}[{i}].{field}";
                            issues.Warning(
                                "MissingLoopField",
                                "ResolvePaths",
                                $"Missing/empty '{field}' in {arrayPath}[{i}].",
                                markPath);
                            marks.Add(new JsonMark(field, IssueSeverity.Warning, "empty"));
                        }
                    }
                }
            }

            // Folder fields soft warnings
            foreach (var required in new[] { "name", "role", "topSkills" })
            {
                if (!_paths.TryGetElement(root, required, out var el) || IsEmpty(el))
                {
                    if (marks.All(m => !m.Path.Equals(required, StringComparison.OrdinalIgnoreCase)))
                    {
                        issues.Warning(
                            "MissingFolderField",
                            "Folder",
                            $"Recommended field '{required}' missing/empty for output folder naming.",
                            required);
                        marks.Add(new JsonMark(required, IssueSeverity.Warning, "missing"));
                    }
                }
            }
        }

        return new LiveValidationResult
        {
            Issues = issues.Issues.ToList(),
            HighlightHtml = BuildHighlight(jsonText, marks),
            Marks = marks
                .GroupBy(m => m.Path + m.Reason)
                .Select(g => g.First())
                .ToList()
        };
    }

    private static bool IsEmpty(JsonElement el) =>
        el.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined
        || (el.ValueKind == JsonValueKind.String && string.IsNullOrWhiteSpace(el.GetString()))
        || (el.ValueKind == JsonValueKind.Array && !el.EnumerateArray().Any());

    private static string BuildHighlight(string jsonText, IReadOnlyList<JsonMark> marks)
    {
        if (string.IsNullOrEmpty(jsonText))
            return "";

        var severityByLeaf = new Dictionary<string, IssueSeverity>(StringComparer.OrdinalIgnoreCase);
        foreach (var mark in marks)
        {
            var leaf = mark.Path;
            var bracket = leaf.LastIndexOf(']');
            if (bracket >= 0 && bracket + 2 < leaf.Length && leaf[bracket + 1] == '.')
                leaf = leaf[(bracket + 2)..];
            else if (leaf.Contains('.'))
                leaf = leaf[(leaf.LastIndexOf('.') + 1)..];

            if (!severityByLeaf.TryGetValue(leaf, out var existing) || mark.Severity > existing)
                severityByLeaf[leaf] = mark.Severity;
        }

        // Escape then wrap "leaf" property names.
        var escaped = WebUtility.HtmlEncode(jsonText);
        foreach (var (leaf, severity) in severityByLeaf.OrderByDescending(k => k.Key.Length))
        {
            var cls = severity == IssueSeverity.Error ? "json-mark error" : "json-mark warn";
            var pattern = $"&quot;{Regex.Escape(leaf)}&quot;";
            escaped = Regex.Replace(
                escaped,
                pattern,
                $"<span class=\"{cls}\">&quot;{WebUtility.HtmlEncode(leaf)}&quot;</span>",
                RegexOptions.IgnoreCase);
        }

        // Preserve trailing newline for textarea sync.
        return escaped.Replace("\n", "<br/>");
    }
}
