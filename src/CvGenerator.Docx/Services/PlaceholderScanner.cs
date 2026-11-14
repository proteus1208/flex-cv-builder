using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace CvGenerator.Docx.Services;

public sealed class PlaceholderInfo
{
    public required string Token { get; init; }
    public required string Kind { get; init; } // simple | array-start | array-end | loop-field
    public string? JsonPath { get; init; }
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

        foreach (System.Text.RegularExpressions.Match m in PlaceholderPatterns.AnyPlaceholder().Matches(text ?? ""))
        {
            var raw = m.Value;
            if (!seen.Add(raw))
                continue;

            var inner = m.Groups[1].Value.Trim();

            if (PlaceholderPatterns.IsArrayStart(raw, out var startName) ||
                PlaceholderPatterns.ArrayStart().IsMatch(raw))
            {
                var name = PlaceholderPatterns.ArrayStart().Match(raw).Groups[1].Value;
                results.Add(new PlaceholderInfo { Token = raw, Kind = "array-start", JsonPath = name });
                continue;
            }

            if (PlaceholderPatterns.ArrayEnd().IsMatch(raw))
            {
                var name = PlaceholderPatterns.ArrayEnd().Match(raw).Groups[1].Value;
                results.Add(new PlaceholderInfo { Token = raw, Kind = "array-end", JsonPath = name });
                continue;
            }

            if (PlaceholderPatterns.IsLoopField(inner, out var arrayName, out var field))
            {
                results.Add(new PlaceholderInfo
                {
                    Token = raw,
                    Kind = "loop-field",
                    JsonPath = $"{arrayName}_i.{field}"
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
