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

        var matches = PlaceholderPatterns.AnyPlaceholder().Matches(text ?? "").Cast<System.Text.RegularExpressions.Match>().ToList();

        // Pass 1: collect markers so alias map can resolve nested starts.
        var provisional = new List<PlaceholderInfo>();
        foreach (var m in matches)
        {
            var raw = m.Value;
            if (!seen.Add(raw))
                continue;

            var inner = m.Groups[1].Value.Trim();

            if (PlaceholderPatterns.IsArrayStart(raw, out ArrayMarker start))
            {
                provisional.Add(new PlaceholderInfo
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
                provisional.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "array-end",
                    JsonPath = end.Source,
                    Alias = end.Alias
                });
                continue;
            }

            provisional.Add(new PlaceholderInfo
            {
                Token = raw,
                Kind = PlaceholderPatterns.TryParseLoopReference(inner, out _) ? "loop-field" : "simple",
                JsonPath = inner
            });
        }

        var aliasMap = AliasPathNormalizer.BuildAliasMap(provisional);

        // Pass 2: rewrite alias fields onto classic _i paths for validation/bindings.
        foreach (var p in provisional)
        {
            if (p.Kind is "array-start" or "array-end")
            {
                results.Add(p);
                continue;
            }

            var kind = AliasPathNormalizer.EffectiveKind(p.Kind, p.JsonPath, p.Token, aliasMap);
            var path = AliasPathNormalizer.ToValidationPath(kind, p.JsonPath, p.Token, aliasMap);
            results.Add(new PlaceholderInfo
            {
                Token = p.Token,
                Kind = kind,
                JsonPath = path,
                Alias = p.Alias
            });
        }

        return results;
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
