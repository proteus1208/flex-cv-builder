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
                        ValidateArrayStart(root, p.JsonPath!, p.Token, issues, marks);
                        break;

                    case "loop-field":
                        ValidateLoopField(root, p.JsonPath!, p.Token, issues, marks);
                        break;
                }
            }

            if (!_paths.TryGetElement(root, "directory", out var dirEl) || IsEmpty(dirEl))
            {
                if (marks.All(m => !m.Path.Equals("directory", StringComparison.OrdinalIgnoreCase)))
                {
                    issues.Info(
                        "MissingFolderField",
                        "Folder",
                        "No 'directory' field — output folder will use current date/time.",
                        "directory");
                    marks.Add(new JsonMark("directory", IssueSeverity.Info, "missing"));
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

    private void ValidateArrayStart(
        JsonElement root,
        string arrayToken,
        string token,
        IssueCollector issues,
        List<JsonMark> marks)
    {
        // skills → root.skills ; skills_i.detail → sample skills[0].detail
        if (!arrayToken.Contains("_i.", StringComparison.Ordinal))
        {
            var items = _paths.GetArrayItems(root, arrayToken, out _, out var missing);
            if (missing)
            {
                issues.Warning("MissingLoopArray", "ResolvePaths",
                    $"Template array '{arrayToken}' ({token}) not found in JSON — section will be empty.",
                    arrayToken);
                marks.Add(new JsonMark(arrayToken, IssueSeverity.Warning, "missing"));
            }
            else if (items.Count == 0)
            {
                issues.Warning("EmptyLoopArray", "ResolvePaths",
                    $"JSON array '{arrayToken}' is empty — section will be empty.",
                    arrayToken);
                marks.Add(new JsonMark(arrayToken, IssueSeverity.Warning, "empty"));
            }

            return;
        }

        if (!TrySplitLoopPath(arrayToken, out var chain, out var property) ||
            string.IsNullOrEmpty(property) ||
            property.Contains('.'))
        {
            return;
        }

        if (!_paths.TryGetElement(root, chain[0], out var cursor) || cursor.ValueKind != JsonValueKind.Array)
        {
            issues.Warning("MissingLoopArray", "ResolvePaths",
                $"Template array '{arrayToken}' needs '{chain[0]}' which is missing in JSON.",
                chain[0]);
            marks.Add(new JsonMark(chain[0], IssueSeverity.Warning, "missing"));
            return;
        }

        var level = cursor.EnumerateArray().ToList();
        if (level.Count == 0)
        {
            issues.Warning("EmptyLoopArray", "ResolvePaths",
                $"JSON array '{chain[0]}' is empty — section will be empty.",
                chain[0]);
            marks.Add(new JsonMark(chain[0], IssueSeverity.Warning, "empty"));
            return;
        }

        var parentItem = level[0];
        for (var c = 1; c < chain.Count; c++)
        {
            if (!_paths.TryGetElement(parentItem, chain[c], out var next) || next.ValueKind != JsonValueKind.Array)
            {
                var path = string.Join(".", chain.Take(c + 1));
                issues.Warning("MissingLoopArray", "ResolvePaths",
                    $"Nested array '{chain[c]}' missing for {token}.",
                    path);
                marks.Add(new JsonMark(path, IssueSeverity.Warning, "missing"));
                return;
            }

            level = next.EnumerateArray().ToList();
            if (level.Count == 0)
            {
                var path = string.Join(".", chain.Take(c + 1));
                issues.Warning("EmptyLoopArray", "ResolvePaths",
                    $"JSON array '{path}' is empty — section will be empty.",
                    path);
                marks.Add(new JsonMark(path, IssueSeverity.Warning, "empty"));
                return;
            }

            parentItem = level[0];
        }

        var markPath = $"{chain[0]}[0].{property}";
        if (!_paths.TryGetElement(parentItem, property, out var child) ||
            (child.ValueKind is not JsonValueKind.Array and not JsonValueKind.Object))
        {
            issues.Warning("MissingLoopArray", "ResolvePaths",
                $"Template array '{arrayToken}' ({token}) not found under {chain[0]}[0].",
                markPath);
            marks.Add(new JsonMark(markPath, IssueSeverity.Warning, "missing"));
        }
        else if (child.ValueKind == JsonValueKind.Array && !child.EnumerateArray().Any())
        {
            issues.Warning("EmptyLoopArray", "ResolvePaths",
                $"JSON array for '{arrayToken}' is empty — section will be empty.",
                markPath);
            marks.Add(new JsonMark(markPath, IssueSeverity.Warning, "empty"));
        }
    }

    private void ValidateLoopField(
        JsonElement root,
        string jsonPath,
        string token,
        IssueCollector issues,
        List<JsonMark> marks)
    {
        // skills_i.category | skills_i.detail_i | skills_i.detail_i.name
        if (!TrySplitLoopPath(jsonPath, out var chain, out var field))
            return;

        // chain[0] may be a dotted absolute path (experience.company_1.industries_i → chain[0]=experience.company_1.industries)
        if (!_paths.TryGetElement(root, chain[0], out var cursor) || cursor.ValueKind != JsonValueKind.Array)
        {
            issues.Warning("MissingLoopArray", "ResolvePaths",
                $"Template field {token} needs array '{chain[0]}' which is missing in JSON.",
                chain[0]);
            marks.Add(new JsonMark(chain[0], IssueSeverity.Warning, "missing"));
            return;
        }

        var items = cursor.EnumerateArray().ToList();
        if (items.Count == 0)
            return;

        // Scalar items (alias like ${industry} → path_i) or single-level object fields.
        if (chain.Count == 1)
        {
            for (var i = 0; i < items.Count; i++)
            {
                if (string.IsNullOrEmpty(field))
                    continue;

                if (!_paths.TryGetElement(items[i], field, out var el) || IsEmpty(el))
                {
                    var markPath = $"{chain[0]}[{i}].{field}";
                    issues.Warning("MissingLoopField", "ResolvePaths",
                        $"Missing/empty '{field}' in {chain[0]}[{i}] for {token}.",
                        markPath);
                    marks.Add(new JsonMark(markPath, IssueSeverity.Warning, "empty"));
                }
            }

            return;
        }

        // Nested: skills_i.detail_i.name — check under first skill item for detail array
        var parentItem = items[0];
        for (var c = 1; c < chain.Count; c++)
        {
            if (!_paths.TryGetElement(parentItem, chain[c], out var next) || next.ValueKind != JsonValueKind.Array)
            {
                var path = string.Join(".", chain.Take(c + 1));
                issues.Warning("MissingLoopArray", "ResolvePaths",
                    $"Nested array '{chain[c]}' missing for {token}.",
                    path);
                marks.Add(new JsonMark(path, IssueSeverity.Warning, "missing"));
                return;
            }

            var nested = next.EnumerateArray().ToList();
            if (nested.Count == 0)
                return;

            if (c == chain.Count - 1)
            {
                if (string.IsNullOrEmpty(field))
                    return; // scalar items OK if array exists

                for (var i = 0; i < nested.Count; i++)
                {
                    if (!_paths.TryGetElement(nested[i], field, out var el) || IsEmpty(el))
                    {
                        var markPath = $"{chain[0]}[0].{chain[c]}[{i}].{field}";
                        issues.Warning("MissingLoopField", "ResolvePaths",
                            $"Missing/empty '{field}' for {token}.",
                            markPath);
                        marks.Add(new JsonMark(markPath, IssueSeverity.Warning, "empty"));
                    }
                }

                return;
            }

            parentItem = nested[0];
        }
    }

    private static bool TrySplitLoopPath(string path, out List<string> chain, out string? field)
    {
        chain = new List<string>();
        field = null;
        var s = path.Trim();
        var i = 0;
        while (i < s.Length)
        {
            var idx = s.IndexOf("_i", i, StringComparison.Ordinal);
            if (idx < 0)
                break;
            chain.Add(s[i..idx]);
            i = idx + 2;
            if (i < s.Length && s[i] == '.')
                i++;
            else if (i >= s.Length)
            {
                field = null;
                return chain.Count > 0;
            }
            else
                break;
        }

        if (chain.Count == 0)
            return false;
        field = i < s.Length ? s[i..] : null;
        if (string.IsNullOrEmpty(field))
            field = null;
        return true;
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
