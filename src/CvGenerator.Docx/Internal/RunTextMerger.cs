using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Handles Word run-splitting and placeholder replacement while preserving run/paragraph formatting.
/// </summary>
internal static class RunTextMerger
{
    public static string GetParagraphText(Paragraph paragraph)
    {
        var sb = new StringBuilder();
        foreach (var text in paragraph.Descendants<Text>())
            sb.Append(text.Text ?? "");
        return sb.ToString();
    }

    /// <summary>
    /// Replace each ${...} in the paragraph using <paramref name="replaceInner"/>.
    /// Only the placeholder character range is rewritten; surrounding runs keep their properties.
    /// </summary>
    public static void ReplacePlaceholders(Paragraph paragraph, Func<string, string> replaceInner)
    {
        var original = GetParagraphText(paragraph);
        if (!original.Contains("${", StringComparison.Ordinal))
            return;

        var matches = PlaceholderPatterns.AnyPlaceholder().Matches(original);
        if (matches.Count == 0)
            return;

        // Apply from end to start so earlier offsets stay valid.
        for (var i = matches.Count - 1; i >= 0; i--)
        {
            var match = matches[i];
            var inner = match.Groups[1].Value.Trim();
            var replacement = replaceInner(inner) ?? string.Empty;
            ReplaceRange(paragraph, match.Index, match.Length, replacement);
        }
    }

    public static void ReplaceRange(Paragraph paragraph, int start, int length, string replacement)
    {
        if (length < 0)
            return;

        var segments = FlattenTextSegments(paragraph);
        if (segments.Count == 0)
        {
            // No text runs — create one after pPr, cloning nothing.
            var run = new Run(new Text(replacement) { Space = SpaceProcessingModeValues.Preserve });
            var pPr = paragraph.GetFirstChild<ParagraphProperties>();
            if (pPr != null) pPr.InsertAfterSelf(run);
            else paragraph.AppendChild(run);
            return;
        }

        var end = start + length;
        RunProperties? formatProps = null;
        Text? firstText = null;
        var firstLocal = 0;

        foreach (var seg in segments)
        {
            var segEnd = seg.Start + seg.Length;
            if (segEnd <= start || seg.Start >= end)
                continue;

            formatProps ??= seg.Run.RunProperties != null
                ? (RunProperties)seg.Run.RunProperties.CloneNode(true)
                : null;

            var overlapStart = Math.Max(start, seg.Start);
            var overlapEnd = Math.Min(end, segEnd);
            var localStart = overlapStart - seg.Start;
            var localLen = overlapEnd - overlapStart;
            var text = seg.Text.Text ?? "";

            if (firstText is null)
            {
                firstText = seg.Text;
                firstLocal = localStart;
                var before = text[..localStart];
                var after = text[(localStart + localLen)..];
                seg.Text.Text = before + replacement + after;
                seg.Text.Space = SpaceProcessingModeValues.Preserve;
            }
            else
            {
                // Clear overlapping portion in subsequent segments.
                seg.Text.Text = text.Remove(localStart, localLen);
                seg.Text.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        // If placeholder sat in a gap with no coverage, append.
        if (firstText is null && !string.IsNullOrEmpty(replacement))
        {
            var run = new Run();
            if (formatProps != null)
                run.RunProperties = formatProps;
            run.AppendChild(new Text(replacement) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    private static List<TextSegment> FlattenTextSegments(Paragraph paragraph)
    {
        var list = new List<TextSegment>();
        var offset = 0;
        foreach (var run in paragraph.Descendants<Run>())
        {
            foreach (var text in run.Elements<Text>())
            {
                var value = text.Text ?? "";
                list.Add(new TextSegment(run, text, offset, value.Length));
                offset += value.Length;
            }
        }

        return list;
    }

    private sealed record TextSegment(Run Run, Text Text, int Start, int Length);
}
