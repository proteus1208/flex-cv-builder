using System.Net;
using System.Text.Json;
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

        // Always build a highlight mirror of the raw JSON so the editor is correct on first paint.
        if (string.IsNullOrWhiteSpace(jsonText))
        {
            issues.Error("InvalidJson", "ParseJson", "JSON is empty.");
            return Result(issues, marks, "");
        }

        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(jsonText);
        }
        catch (JsonException ex)
        {
            issues.Error("InvalidJson", "ParseJson", $"Invalid JSON: {ex.Message}", detail: ex.ToString());
            return Result(issues, marks, BuildPlainHighlight(jsonText));
        }

        using (doc)
        {
            var root = doc.RootElement;

            foreach (var p in placeholderList)
            {
                if (string.IsNullOrWhiteSpace(p.JsonPath))
                    continue;

                switch (p.Kind)
                {
                    case "simple":
                        if (!_paths.TryGetElement(root, p.JsonPath, out var el))
                        {
                            issues.Warning("MissingField", "ResolvePaths",
                                $"Template expects '{p.JsonPath}' ({p.Token}) but JSON has no matching key.",
                                p.JsonPath);
                            marks.Add(new JsonMark(p.JsonPath, IssueSeverity.Warning, "missing"));
                        }
                        else if (IsEmpty(el))
                        {
                            issues.Warning("EmptyField", "ResolvePaths",
                                $"JSON field '{p.JsonPath}' is empty.", p.JsonPath);
                            marks.Add(new JsonMark(p.JsonPath, IssueSeverity.Warning, "empty"));
                        }
                        break;

                    case "array-start":
                    case "loop-start":
                        ValidateArray(root, p.JsonPath, p.Token, issues, marks);
                        break;

                    case "loop-field":
                        ValidateLoopField(root, p.JsonPath, p.Token, issues, marks);
                        break;
                }
            }

            foreach (var required in new[] { "name", "role", "topSkills" })
            {
                if (!_paths.TryGetElement(root, required, out var el) || IsEmpty(el))
                {
                    if (marks.All(m => !m.Path.Equals(required, StringComparison.OrdinalIgnoreCase)))
                    {
                        issues.Warning("MissingFolderField", "Folder",
                            $"Recommended field '{required}' missing/empty for output folder naming.",
                            required);
                        marks.Add(new JsonMark(required, IssueSeverity.Warning, "missing"));
                    }
                }
            }
        }

        var uniqueMarks = marks
            .GroupBy(m => m.Path + "|" + m.Reason, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();

        return new LiveValidationResult
        {
            Issues = issues.Issues.ToList(),
            HighlightHtml = BuildHighlight(jsonText, uniqueMarks),
            Marks = uniqueMarks
        };
    }

    private void ValidateArray(
        JsonElement root,
        string arrayPath,
        string token,
        IssueCollector issues,
        List<JsonMark> marks)
    {
        var items = _paths.GetArrayItems(root, arrayPath, out _, out var missing);
        if (missing)
        {
            issues.Warning("MissingLoopArray", "ResolvePaths",
                $"Template array '{arrayPath}' ({token}) not found in JSON — section will be empty.",
                arrayPath);
            marks.Add(new JsonMark(arrayPath, IssueSeverity.Warning, "missing"));
        }
        else if (items.Count == 0)
        {
            issues.Warning("EmptyLoopArray", "ResolvePaths",
                $"JSON array '{arrayPath}' is empty — section will be empty.",
                arrayPath);
            marks.Add(new JsonMark(arrayPath, IssueSeverity.Warning, "empty"));
        }
    }

    private void ValidateLoopField(
        JsonElement root,
        string jsonPath,
        string token,
        IssueCollector issues,
        List<JsonMark> marks)
    {
        // experience_i.title
        var idx = jsonPath.IndexOf("_i.", StringComparison.Ordinal);
        if (idx <= 0)
            return;

        var arrayPath = jsonPath[..idx];
        var field = jsonPath[(idx + 3)..];
        var items = _paths.GetArrayItems(root, arrayPath, out _, out var missing);
        if (missing)
        {
            issues.Warning("MissingLoopArray", "ResolvePaths",
                $"Template field {token} needs array '{arrayPath}' which is missing in JSON.",
                arrayPath);
            marks.Add(new JsonMark(arrayPath, IssueSeverity.Warning, "missing"));
            return;
        }

        if (items.Count == 0)
            return;

        for (var i = 0; i < items.Count; i++)
        {
            if (!_paths.TryGetElement(items[i], field, out var el) || IsEmpty(el))
            {
                var markPath = $"{arrayPath}[{i}].{field}";
                issues.Warning("MissingLoopField", "ResolvePaths",
                    $"Missing/empty '{field}' in {arrayPath}[{i}] for {token}.",
                    markPath);
                // Do NOT mark bare leaf "field" — that wrongly highlights other keys.
                marks.Add(new JsonMark(markPath, IssueSeverity.Warning, "empty"));
            }
        }
    }

    private static LiveValidationResult Result(
        IssueCollector issues,
        List<JsonMark> marks,
        string html) => new()
    {
        Issues = issues.Issues.ToList(),
        HighlightHtml = html,
        Marks = marks
    };

    private static bool IsEmpty(JsonElement el) =>
        el.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined
        || (el.ValueKind == JsonValueKind.String && string.IsNullOrWhiteSpace(el.GetString()))
        || (el.ValueKind == JsonValueKind.Array && !el.EnumerateArray().Any());

    private static string BuildPlainHighlight(string jsonText)
        => WebUtility.HtmlEncode(jsonText);

    private static string BuildHighlight(string jsonText, IReadOnlyList<JsonMark> marks)
    {
        // Mirror exact text (no <br/>) so first paint matches the textarea with white-space:pre-wrap.
        var escaped = WebUtility.HtmlEncode(jsonText);

        // Only highlight keys that physically exist in JSON and are empty (top-level or dotted last segment once).
        foreach (var mark in marks.Where(m => m.Reason == "empty" && !m.Path.Contains('[')))
        {
            var leaf = mark.Path.Contains('.') ? mark.Path[(mark.Path.LastIndexOf('.') + 1)..] : mark.Path;
            var needle = $"&quot;{WebUtility.HtmlEncode(leaf)}&quot;";
            var styled =
                $"<span class=\"json-mark warn\">&quot;{WebUtility.HtmlEncode(leaf)}&quot;</span>";
            // Replace first occurrence only to reduce collateral highlights.
            var idx = escaped.IndexOf(needle, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
                escaped = escaped[..idx] + styled + escaped[(idx + needle.Length)..];
        }

        return escaped;
    }
}
