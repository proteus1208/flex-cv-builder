using System.Text.RegularExpressions;
using CvGenerator.Core.Models;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace CvGenerator.Docx.Services;

public sealed partial class TemplateSyntaxValidator
{
    [GeneratedRegex(@"^\s*([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)*)\s*$")]
    private static partial Regex SimplePathRegex();

    public IReadOnlyList<GenerationIssue> ValidateFile(string docxPath)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(docxPath, false);
            return ValidateText(ExtractAllText(doc));
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
        var loopFields = new List<(LoopReference Ref, string Token)>();

        foreach (Match match in PlaceholderPatterns.AnyPlaceholder().Matches(text ?? string.Empty))
        {
            var token = match.Value;
            var inner = match.Groups[1].Value;

            if (string.IsNullOrWhiteSpace(inner))
            {
                issues.Error("EmptyPlaceholder", "Template", "Empty placeholder ${} is invalid.", token);
                continue;
            }

            if (LooksLikeLegacyOrBadKeyword(inner, out var hint))
            {
                issues.Error("InvalidPlaceholderKeyword", "Template", $"{hint} in {token}", token);
                continue;
            }

            if (PlaceholderPatterns.IsArrayStart(token, out ArrayMarker startMarker))
            {
                var name = startMarker.Source;
                var validSource =
                    PlaceholderPatterns.TryParseQualifiedArrayName(name, out _, out _) ||
                    PlaceholderPatterns.TryParseAliasQualifiedSource(name, out _, out _);
                if (!validSource)
                {
                    issues.Error("InvalidArrayStart", "Template",
                        $"Invalid ArrayStart source '{name}'. Use `skills`, `skills_i.detail`, or `skill.detail`.", token);
                }
                else if (startMarker.Alias != null &&
                         startMarker.Alias.Equals(name, StringComparison.OrdinalIgnoreCase) &&
                         !name.Contains('.'))
                {
                    issues.Warning("AliasSameAsSource", "Template",
                        $"Iterate alias '{startMarker.Alias}' matches the array name — prefer a short alias like `item`.", token);
                    starts.Add((name, token));
                }
                else
                {
                    starts.Add((name, token));
                }

                continue;
            }

            if (PlaceholderPatterns.IsArrayEnd(token, out ArrayMarker endMarker))
            {
                ends.Add((endMarker.Source, token));
                continue;
            }

            if (PlaceholderPatterns.TryParseLoopReference(inner, out var loopRef))
            {
                loopFields.Add((loopRef, token));
                continue;
            }

            var trimmedParts = inner.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (trimmedParts.Length > 1)
            {
                issues.Error("InvalidPlaceholder", "Template",
                    $"Too many parts inside placeholder: {token}", token);
                continue;
            }

            if (SimplePathRegex().IsMatch(inner))
                continue;

            if (inner.Contains(".i.", StringComparison.Ordinal))
            {
                issues.Error(
                    "LegacyLoopField",
                    "Template",
                    $"Use ${{array_i.field}} / ${{parent_i.child_i.field}} instead of .i. form: {token}",
                    token);
                continue;
            }

            issues.Error(
                "InvalidPlaceholder",
                "Template",
                $"Invalid placeholder: {token}.",
                token);
        }

        var startCounts = starts.GroupBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);
        var endCounts = ends.GroupBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);

        foreach (var name in startCounts.Keys.ToList())
        {
            endCounts.TryGetValue(name, out var ec);
            var sc = startCounts[name];
            if (ec == 0)
            {
                issues.Error("UnpairedArrayStart", "Template",
                    $"ArrayStart has no matching ArrayEnd for '{name}'.", name);
            }
            else if (sc != ec)
            {
                issues.Error("UnbalancedArray", "Template",
                    $"Array '{name}' has {sc} ArrayStart(s) and {ec} ArrayEnd(s).", name);
            }
        }

        foreach (var (name, token) in ends)
        {
            if (!startCounts.ContainsKey(name))
            {
                issues.Error("UnpairedArrayEnd", "Template",
                    $"ArrayEnd has no matching ArrayStart: {token}", token);
            }
        }

        foreach (var (loopRef, token) in loopFields)
        {
            // Outer array of the chain should appear as some ArrayStart local or qualified name.
            var outer = loopRef.Chain[0];
            var hasStart = startCounts.Keys.Any(k =>
                k.Equals(outer, StringComparison.OrdinalIgnoreCase) ||
                k.StartsWith(outer + "_i.", StringComparison.OrdinalIgnoreCase) ||
                k.EndsWith("." + outer, StringComparison.OrdinalIgnoreCase) ||
                // qualified start skills_i.detail implies skills exists as parent start
                (PlaceholderPatterns.TryParseQualifiedArrayName(k, out var parents, out var prop) &&
                 (parents.Contains(outer, StringComparer.OrdinalIgnoreCase) ||
                  prop.Equals(outer, StringComparison.OrdinalIgnoreCase) ||
                  parents.Count == 0 && prop.Equals(outer, StringComparison.OrdinalIgnoreCase))));

            // Also: ArrayStart skills covers skills_i.*
            hasStart = hasStart || startCounts.ContainsKey(outer);

            if (!hasStart)
            {
                issues.Warning("OrphanLoopField", "Template",
                    $"Loop field {token} may have no matching ArrayStart for '{outer}'.", token);
            }
        }

        return issues.Issues
            .GroupBy(i => $"{i.Code}|{i.Path}|{i.Message}")
            .Select(g => g.First())
            .ToList();
    }

    private static bool LooksLikeLegacyOrBadKeyword(string inner, out string hint)
    {
        hint = string.Empty;
        var first = inner.Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
        var lower = first.ToLowerInvariant();

        if (lower is "arraystart" or "arrayend")
            return false;

        if (lower is "start" or "end")
        {
            hint = "Use ArrayStart / ArrayEnd (not start / end)";
            return true;
        }

        string[] mistakes = ["starts", "begin", "ends", "finish", "stop", "for", "each", "loop"];
        if (mistakes.Contains(lower))
        {
            hint = "Invalid keyword — use ArrayStart / ArrayEnd";
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
