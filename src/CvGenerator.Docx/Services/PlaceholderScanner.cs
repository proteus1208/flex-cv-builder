using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Services;

public sealed class PlaceholderInfo
{
    public required string Token { get; init; }
    public required string Kind { get; init; } // simple | loop-start | loop-end | loop-field
    public string? JsonPath { get; init; }
}

public sealed partial class PlaceholderScanner
{
    [GeneratedRegex(@"\$\{\s*start\s+([A-Za-z0-9_.]+)\s*\}", RegexOptions.IgnoreCase)]
    private static partial Regex LoopStartRegex();

    [GeneratedRegex(@"\$\{\s*end\s+([A-Za-z0-9_.]+)\s*\}", RegexOptions.IgnoreCase)]
    private static partial Regex LoopEndRegex();

    [GeneratedRegex(@"\$\{([^{}]+)\}")]
    private static partial Regex AnyPlaceholderRegex();

    public IReadOnlyList<PlaceholderInfo> ScanFile(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var texts = ExtractAllText(doc);
        return ScanText(texts);
    }

    public IReadOnlyList<PlaceholderInfo> ScanText(string text)
    {
        var results = new List<PlaceholderInfo>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (Match m in AnyPlaceholderRegex().Matches(text))
        {
            var raw = m.Value;
            if (!seen.Add(raw))
                continue;

            var inner = m.Groups[1].Value.Trim();
            if (LoopStartRegex().IsMatch(raw))
            {
                var name = LoopStartRegex().Match(raw).Groups[1].Value;
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "loop-start",
                    JsonPath = name
                });
            }
            else if (LoopEndRegex().IsMatch(raw))
            {
                var name = LoopEndRegex().Match(raw).Groups[1].Value;
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "loop-end",
                    JsonPath = name
                });
            }
            else if (inner.Contains(".i.", StringComparison.Ordinal))
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "loop-field",
                    JsonPath = inner
                });
            }
            else
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "simple",
                    JsonPath = inner
                });
            }
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
