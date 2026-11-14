using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace CvGenerator.Docx.Services;

public sealed class PlaceholderInfo
{
    public required string Token { get; init; }
    public required string Kind { get; init; } // simple | array-start | array-end | loop-field
    public string? JsonPath { get; init; }
    /// <summary>Iterate alias from ${ArrayStart source alias}, when present.</summary>
    public string? Alias { get; init; }
}

public sealed class PlaceholderScanner
{
    public IReadOnlyList<PlaceholderInfo> ScanFile(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        return ScanText(ExtractAllText(doc));
    }

    public IReadOnlyList<PlaceholderInfo> ScanText(string text)
    {
        var results = new List<PlaceholderInfo>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var aliasToCanonical = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var matches = PlaceholderPatterns.AnyPlaceholder().Matches(text ?? "").Cast<System.Text.RegularExpressions.Match>().ToList();

        // Pass 1: collect ArrayStart aliases → canonical _i prefixes for live validation.
        foreach (var m in matches)
        {
            var raw = m.Value;
            if (!PlaceholderPatterns.IsArrayStart(raw, out ArrayMarker marker) || marker.Alias is null)
                continue;

            var canonical = ToCanonicalPrefix(marker.Source, aliasToCanonical);
            if (canonical != null)
                aliasToCanonical[marker.Alias] = canonical;
        }

        // Pass 2: classify placeholders.
        foreach (var m in matches)
        {
            var raw = m.Value;
            if (!seen.Add(raw))
                continue;

            var inner = m.Groups[1].Value.Trim();

            if (PlaceholderPatterns.IsArrayStart(raw, out ArrayMarker start))
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "array-start",
                    JsonPath = start.Source,
                    Alias = start.Alias
                });
                continue;
            }

            if (PlaceholderPatterns.IsArrayEnd(raw, out ArrayMarker end))
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "array-end",
                    JsonPath = end.Source,
                    Alias = end.Alias
                });
                continue;
            }

            if (PlaceholderPatterns.TryParseLoopReference(inner, out _))
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "loop-field",
                    JsonPath = inner
                });
                continue;
            }

            if (TryRewriteAliasPath(inner, aliasToCanonical, out var canonicalPath))
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "loop-field",
                    JsonPath = canonicalPath
                });
                continue;
            }

            results.Add(new PlaceholderInfo
            {
                Token = raw,
                Kind = "simple",
                JsonPath = inner
            });
        }

        return results;
    }

    private static string? ToCanonicalPrefix(
        string source,
        IReadOnlyDictionary<string, string> aliasToCanonical)
    {
        if (PlaceholderPatterns.TryParseQualifiedArrayName(source, out var parents, out var prop))
        {
            if (parents.Count == 0)
                return prop + "_i";

            return string.Join(".", parents.Select(p => p + "_i")) + "." + prop + "_i";
        }

        if (PlaceholderPatterns.TryParseAliasQualifiedSource(source, out var parentAlias, out var child) &&
            aliasToCanonical.TryGetValue(parentAlias, out var parentPrefix))
        {
            return parentPrefix + "." + child + "_i";
        }

        return null;
    }

    private static bool TryRewriteAliasPath(
        string inner,
        IReadOnlyDictionary<string, string> aliasToCanonical,
        out string canonical)
    {
        canonical = "";
        var s = inner.Trim();

        string? bestAlias = null;
        foreach (var alias in aliasToCanonical.Keys)
        {
            if (s.Equals(alias, StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith(alias + ".", StringComparison.OrdinalIgnoreCase))
            {
                if (bestAlias is null || alias.Length > bestAlias.Length)
                    bestAlias = alias;
            }
        }

        if (bestAlias is null)
            return false;

        var prefix = aliasToCanonical[bestAlias];
        if (s.Equals(bestAlias, StringComparison.OrdinalIgnoreCase))
        {
            canonical = prefix;
            return true;
        }

        canonical = prefix + "." + s[(bestAlias.Length + 1)..];
        return true;
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
