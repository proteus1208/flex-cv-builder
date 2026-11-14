using System.Text.RegularExpressions;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;
using DocumentFormat.OpenXml.Packaging;

namespace CvGenerator.Docx.Services;

public sealed partial class TemplateSyntaxValidator
{
    [GeneratedRegex(@"\$\{([^{}]*)\}")]
    private static partial Regex PlaceholderRegex();

    [GeneratedRegex(@"^\s*start\s+([A-Za-z_][A-Za-z0-9_.]*)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex StartRegex();

    [GeneratedRegex(@"^\s*end\s+([A-Za-z_][A-Za-z0-9_.]*)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex EndRegex();

    [GeneratedRegex(@"^\s*([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)\s*$")]
    private static partial Regex SimplePathRegex();

    [GeneratedRegex(@"^\s*([A-Za-z_][A-Za-z0-9_.]*)\.i\.([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)\s*$")]
    private static partial Regex LoopFieldRegex();

    public IReadOnlyList<GenerationIssue> ValidateFile(string docxPath)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(docxPath, false);
            var text = ExtractAllText(doc);
            return ValidateText(text);
        }
        catch (Exception ex)
        {
            return new[]
            {
                new GenerationIssue(
                    IssueSeverity.Error,
                    "TemplateReadFailed",
                    "Template",
                    $"Cannot read template for validation: {ex.Message}",
                    Detail: ex.ToString())
            };
        }
    }

    public IReadOnlyList<GenerationIssue> ValidateText(string text)
    {
        var issues = new IssueCollector();
        var starts = new List<(string Name, string Token)>();
        var ends = new List<(string Name, string Token)>();
        var loopFields = new List<(string Array, string Token)>();

        foreach (Match match in PlaceholderRegex().Matches(text ?? string.Empty))
        {
            var token = match.Value;
            var inner = match.Groups[1].Value;

            if (string.IsNullOrWhiteSpace(inner))
            {
                issues.Error("EmptyPlaceholder", "Template", "Empty placeholder ${} is invalid.", token);
                continue;
            }

            // Reject nested or broken forms with too many structural keywords.
            var parts = inner.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (parts.Length > 2 &&
                !StartRegex().IsMatch(inner) &&
                !EndRegex().IsMatch(inner))
            {
                issues.Error(
                    "InvalidPlaceholder",
                    "Template",
                    $"Too many parts inside placeholder: {token}",
                    token);
                continue;
            }

            if (LooksLikeKeywordMistake(inner, out var keywordHint))
            {
                issues.Error(
                    "InvalidPlaceholderKeyword",
                    "Template",
                    $"Invalid placeholder keyword in {token}. {keywordHint}",
                    token);
                continue;
            }

            var start = StartRegex().Match(inner);
            if (start.Success)
            {
                if (parts.Length != 2)
                {
                    issues.Error(
                        "InvalidLoopStart",
                        "Template",
                        $"Loop start must be exactly `${{ start name }}`: {token}",
                        token);
                }
                else
                {
                    starts.Add((start.Groups[1].Value, token));
                }

                continue;
            }

            var end = EndRegex().Match(inner);
            if (end.Success)
            {
                if (parts.Length != 2)
                {
                    issues.Error(
                        "InvalidLoopEnd",
                        "Template",
                        $"Loop end must be exactly `${{ end name }}`: {token}",
                        token);
                }
                else
                {
                    ends.Add((end.Groups[1].Value, token));
                }

                continue;
            }

            var loopField = LoopFieldRegex().Match(inner);
            if (loopField.Success)
            {
                loopFields.Add((loopField.Groups[1].Value, token));
                continue;
            }

            if (SimplePathRegex().IsMatch(inner))
                continue;

            issues.Error(
                "InvalidPlaceholder",
                "Template",
                $"Invalid placeholder syntax: {token}. Use ${{path}}, ${{ start name }}, ${{ end name }}, or ${{name.i.field}}.",
                token);
        }

        // Pair start/end by name (multiset).
        var startCounts = starts.GroupBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);
        var endCounts = ends.GroupBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

        foreach (var (name, token) in starts)
        {
            endCounts.TryGetValue(name, out var endCount);
            startCounts.TryGetValue(name, out var startCount);
            if (endCount == 0)
            {
                issues.Error(
                    "UnpairedLoopStart",
                    "Template",
                    $"Loop start has no matching end: {token}",
                    token);
            }
            else if (startCount != endCount)
            {
                issues.Error(
                    "UnbalancedLoop",
                    "Template",
                    $"Loop '{name}' has {startCount} start(s) and {endCount} end(s).",
                    name);
            }
        }

        foreach (var (name, token) in ends)
        {
            if (!startCounts.ContainsKey(name))
            {
                issues.Error(
                    "UnpairedLoopEnd",
                    "Template",
                    $"Loop end has no matching start: {token}",
                    token);
            }
        }

        var knownLoops = new HashSet<string>(startCounts.Keys, StringComparer.OrdinalIgnoreCase);
        foreach (var (array, token) in loopFields)
        {
            if (!knownLoops.Contains(array))
            {
                issues.Warning(
                    "OrphanLoopField",
                    "Template",
                    $"Loop field {token} has no matching `${{ start {array} }}`.",
                    token);
            }
        }

        // Deduplicate unbalanced messages (same name reported twice from starts).
        return issues.Issues
            .GroupBy(i => $"{i.Code}|{i.Path}|{i.Message}")
            .Select(g => g.First())
            .ToList();
    }

    private static bool LooksLikeKeywordMistake(string inner, out string hint)
    {
        hint = string.Empty;
        var first = inner.Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
        var lower = first.ToLowerInvariant();
        if (lower is "start" or "end")
            return false;

        string[] mistakes = ["starts", "begin", "beginning", "ends", "finish", "stop", "for", "each", "loop"];
        if (mistakes.Contains(lower))
        {
            hint = "Use only `start` / `end`.";
            return true;
        }

        return false;
    }

    private static string ExtractAllText(WordprocessingDocument doc)
    {
        var parts = new List<string>();
        if (doc.MainDocumentPart?.Document?.Body is { } body)
            parts.Add(body.InnerText);

        if (doc.MainDocumentPart != null)
        {
            foreach (var hp in doc.MainDocumentPart.HeaderParts)
                parts.Add(hp.Header?.InnerText ?? "");
            foreach (var fp in doc.MainDocumentPart.FooterParts)
                parts.Add(fp.Footer?.InnerText ?? "");
        }

        return string.Join('\n', parts);
    }
}
