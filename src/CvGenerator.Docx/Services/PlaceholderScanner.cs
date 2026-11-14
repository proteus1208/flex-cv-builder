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
    /// <summary>Classic _i path for live JSON checks (scope-aware).</summary>
    public string? ValidationPath { get; init; }
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
        // Keep every occurrence in order so alias scopes can open/close correctly.
        var ordered = new List<PlaceholderInfo>();
        foreach (System.Text.RegularExpressions.Match m in PlaceholderPatterns.AnyPlaceholder().Matches(text ?? ""))
        {
            var raw = m.Value;
            var inner = m.Groups[1].Value.Trim();

            if (PlaceholderPatterns.IsArrayStart(raw, out ArrayMarker start))
            {
                ordered.Add(new PlaceholderInfo
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
                ordered.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "array-end",
                    JsonPath = end.Source,
                    Alias = end.Alias
                });
                continue;
            }

            ordered.Add(ParseNonArrayPlaceholder(raw, inner));
        }

        var scoped = AliasPathNormalizer.AssignScopedPaths(ordered);
        return AliasPathNormalizer.DedupeByToken(scoped);
    }

    /// <summary>All occurrences with scoped validation paths (for thorough live checks).</summary>
    public IReadOnlyList<PlaceholderInfo> ScanTextAllOccurrences(string text)
    {
        var ordered = new List<PlaceholderInfo>();
        foreach (System.Text.RegularExpressions.Match m in PlaceholderPatterns.AnyPlaceholder().Matches(text ?? ""))
        {
            var raw = m.Value;
            var inner = m.Groups[1].Value.Trim();

            if (PlaceholderPatterns.IsArrayStart(raw, out ArrayMarker start))
            {
                ordered.Add(new PlaceholderInfo
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
                ordered.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "array-end",
                    JsonPath = end.Source,
                    Alias = end.Alias
                });
                continue;
            }

            ordered.Add(ParseNonArrayPlaceholder(raw, inner));
        }

        return AliasPathNormalizer.AssignScopedPaths(ordered);
    }

    private static PlaceholderInfo ParseNonArrayPlaceholder(string raw, string inner)
    {
        if (PlaceholderPatterns.IsJoinKeyword(inner))
        {
            var kind = inner.Equals("JoinStart", StringComparison.OrdinalIgnoreCase)
                ? "join-start"
                : "join-end";
            return new PlaceholderInfo { Token = raw, Kind = kind, JsonPath = kind };
        }

        return new PlaceholderInfo
        {
            Token = raw,
            Kind = PlaceholderPatterns.TryParseLoopReference(inner, out _) ? "loop-field" : "simple",
            JsonPath = inner
        };
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
