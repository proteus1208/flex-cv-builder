using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
    /// Replace each ${...} using <paramref name="replaceInner"/>.
    /// JSON values may include **bold**, *italic*, __underline__, and [label](url) markup.
    /// </summary>
    public static void ReplacePlaceholders(
        Paragraph paragraph,
        OpenXmlPart? part,
        Func<string, string> replaceInner)
    {
        var original = GetParagraphText(paragraph);
        if (!original.Contains("${", StringComparison.Ordinal))
            return;

        var matches = PlaceholderPatterns.AnyPlaceholder().Matches(original);
        if (matches.Count == 0)
            return;

        for (var i = matches.Count - 1; i >= 0; i--)
        {
            var match = matches[i];
            var inner = match.Groups[1].Value.Trim();
            var replacement = replaceInner(inner) ?? string.Empty;

            if (part != null && RichTextMarkup.ContainsMarkup(replacement))
                ReplaceRangeWithRich(paragraph, part, match.Index, match.Length, replacement);
            else
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
            var run = new Run(new Text(replacement) { Space = SpaceProcessingModeValues.Preserve });
            var pPr = paragraph.GetFirstChild<ParagraphProperties>();
            if (pPr != null) pPr.InsertAfterSelf(run);
            else paragraph.AppendChild(run);
            return;
        }

        var end = start + length;
        RunProperties? formatProps = null;
        Text? firstText = null;

        foreach (var seg in segments)
        {
            var segEnd = seg.Start + seg.Length;
            if (segEnd <= start || seg.Start >= end)
                continue;

            formatProps ??= CloneRunProperties(seg.Run);

            var overlapStart = Math.Max(start, seg.Start);
            var overlapEnd = Math.Min(end, segEnd);
            var localStart = overlapStart - seg.Start;
            var localLen = overlapEnd - overlapStart;
            var text = seg.Text.Text ?? "";

            if (firstText is null)
            {
                firstText = seg.Text;
                var before = text[..localStart];
                var after = text[(localStart + localLen)..];
                seg.Text.Text = before + replacement + after;
                seg.Text.Space = SpaceProcessingModeValues.Preserve;
            }
            else
            {
                seg.Text.Text = text.Remove(localStart, localLen);
                seg.Text.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        if (firstText is null && !string.IsNullOrEmpty(replacement))
        {
            var run = new Run();
            if (formatProps != null)
                run.RunProperties = formatProps;
            run.AppendChild(new Text(replacement) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    public static void ReplaceRangeWithRich(
        Paragraph paragraph,
        OpenXmlPart part,
        int start,
        int length,
        string richValue)
    {
        var segments = FlattenTextSegments(paragraph);
        var end = start + length;
        Run? anchorRun = null;
        RunProperties? baseProps = null;
        string before = "";
        string after = "";

        Text? firstText = null;
        foreach (var seg in segments)
        {
            var segEnd = seg.Start + seg.Length;
            if (segEnd <= start || seg.Start >= end)
                continue;

            baseProps ??= CloneRunProperties(seg.Run);
            anchorRun ??= seg.Run;

            var overlapStart = Math.Max(start, seg.Start);
            var overlapEnd = Math.Min(end, segEnd);
            var localStart = overlapStart - seg.Start;
            var localLen = overlapEnd - overlapStart;
            var text = seg.Text.Text ?? "";

            if (firstText is null)
            {
                firstText = seg.Text;
                before = text[..localStart];
                after = text[(localStart + localLen)..];
                seg.Text.Text = before; // keep prefix only; rich + after inserted after run
                seg.Text.Space = SpaceProcessingModeValues.Preserve;
            }
            else
            {
                seg.Text.Text = text.Remove(localStart, localLen);
                seg.Text.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        var richNodes = BuildRichNodes(part, richValue, baseProps);
        OpenXmlElement insertAfter = anchorRun ?? (OpenXmlElement?)paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph;

        if (anchorRun != null)
        {
            OpenXmlElement cursor = anchorRun;
            foreach (var node in richNodes)
            {
                cursor.InsertAfterSelf(node);
                cursor = node;
            }

            if (!string.IsNullOrEmpty(after))
            {
                var afterRun = new Run();
                if (baseProps != null)
                    afterRun.RunProperties = (RunProperties)baseProps.CloneNode(true);
                afterRun.AppendChild(new Text(after) { Space = SpaceProcessingModeValues.Preserve });
                cursor.InsertAfterSelf(afterRun);
            }
        }
        else
        {
            foreach (var node in richNodes)
                paragraph.AppendChild(node);
            if (!string.IsNullOrEmpty(after))
            {
                var afterRun = new Run(new Text(after) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.AppendChild(afterRun);
            }
        }
    }

    private static List<OpenXmlElement> BuildRichNodes(
        OpenXmlPart part,
        string richValue,
        RunProperties? baseProps)
    {
        var nodes = new List<OpenXmlElement>();
        foreach (var segment in RichTextMarkup.Parse(richValue))
        {
            if (string.IsNullOrEmpty(segment.Text) && segment.Kind != RichSegmentKind.Link)
                continue;

            if (segment.Kind == RichSegmentKind.Link && IsSafeUri(segment.Url, out var uri))
            {
                var rel = part.AddHyperlinkRelationship(uri, true);
                var hyperlink = new Hyperlink { Id = rel.Id };
                var run = CreateRun(segment.Text, baseProps, RichSegmentKind.Text);
                EnsureHyperlinkStyle(run.RunProperties ??= new RunProperties());
                hyperlink.AppendChild(run);
                nodes.Add(hyperlink);
                continue;
            }

            nodes.Add(CreateRun(segment.Text, baseProps, segment.Kind));
        }

        return nodes;
    }

    private static Run CreateRun(string text, RunProperties? baseProps, RichSegmentKind kind)
    {
        var run = new Run();
        var props = baseProps != null
            ? (RunProperties)baseProps.CloneNode(true)
            : new RunProperties();
        if (kind == RichSegmentKind.Bold)
            props.Bold ??= new Bold();
        if (kind == RichSegmentKind.Italic)
            props.Italic ??= new Italic();
        if (kind == RichSegmentKind.Underline)
            props.Underline ??= new Underline { Val = UnderlineValues.Single };
        if (props.ChildElements.Count > 0)
            run.RunProperties = props;
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void EnsureHyperlinkStyle(RunProperties props)
    {
        props.Color ??= new Color { Val = "0563C1" };
        props.Underline ??= new Underline { Val = UnderlineValues.Single };
    }

    private static bool IsSafeUri(string? url, out Uri uri)
    {
        uri = null!;
        if (string.IsNullOrWhiteSpace(url))
            return false;
        if (!Uri.TryCreate(url, UriKind.Absolute, out uri!))
            return false;
        return uri.Scheme is "http" or "https" or "mailto";
    }

    private static RunProperties? CloneRunProperties(Run run)
        => run.RunProperties != null ? (RunProperties)run.RunProperties.CloneNode(true) : null;

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
