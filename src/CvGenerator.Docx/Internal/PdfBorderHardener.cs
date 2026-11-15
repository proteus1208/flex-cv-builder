using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Word 2016 ExportAsFixedFormat often drops thin <c>w:pBdr</c> bottom borders
/// (section bars under PROFILE / EXPERIENCE, etc.) even though they appear in the DOCX.
/// Convert those borders into short shaded paragraphs, which Word PDF keeps reliably.
/// </summary>
internal static class PdfBorderHardener
{
    public static void Harden(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body != null)
            HardenContainer(doc.MainDocumentPart.Document.Body);

        if (doc.MainDocumentPart == null)
            return;

        foreach (var header in doc.MainDocumentPart.HeaderParts)
        {
            if (header.Header != null)
                HardenContainer(header.Header);
        }

        foreach (var footer in doc.MainDocumentPart.FooterParts)
        {
            if (footer.Footer != null)
                HardenContainer(footer.Footer);
        }
    }

    private static void HardenContainer(OpenXmlElement container)
    {
        // Snapshot — we insert siblings after matching paragraphs.
        foreach (var para in container.Elements<Paragraph>().ToList())
        {
            var borders = para.ParagraphProperties?.ParagraphBorders;
            var bottom = borders?.BottomBorder;
            if (!IsVisibleBorder(bottom))
                continue;

            var color = ResolveColor(bottom!);
            var heightTwips = BorderHeightTwips(bottom!);

            bottom!.Remove();
            if (borders != null && !borders.ChildElements.Any())
                borders.Remove();

            var bar = new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Before = "40",
                        After = "60",
                        Line = heightTwips.ToString(),
                        LineRule = LineSpacingRuleValues.Exact
                    },
                    new Shading
                    {
                        Val = ShadingPatternValues.Clear,
                        Color = "auto",
                        Fill = color
                    },
                    new KeepNext()
                ),
                new Run(
                    new RunProperties(new FontSize { Val = "2" }),
                    new Text("\u00A0") { Space = SpaceProcessingModeValues.Preserve }));

            para.InsertAfterSelf(bar);
        }
    }

    private static bool IsVisibleBorder(BorderType? border)
    {
        if (border?.Val?.Value is null)
            return false;
        var val = border.Val.Value;
        return val != BorderValues.Nil && val != BorderValues.None;
    }

    private static string ResolveColor(BorderType border)
    {
        var color = border.Color?.Value;
        if (string.IsNullOrWhiteSpace(color) ||
            color.Equals("auto", StringComparison.OrdinalIgnoreCase))
            return "000000";
        return color.Trim().TrimStart('#').ToUpperInvariant();
    }

    /// <summary>
    /// OOXML border sz is eighths of a point. Twips = points * 20.
    /// Word PDF tends to drop sub-point rules, so enforce a visible minimum.
    /// </summary>
    private static int BorderHeightTwips(BorderType border)
    {
        var eighths = border.Size?.Value ?? 6;
        var twips = (int)Math.Ceiling(eighths * 20.0 / 8.0);
        return Math.Clamp(Math.Max(twips, 24), 24, 120); // 1.2pt .. 6pt
    }
}
