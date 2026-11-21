using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Root cause: Word 2016 <c>ExportAsFixedFormat</c> drops <c>w:pBdr</c> paragraph
/// bottom borders (section bars). The DOCX is correct; the PDF engine omits them.
///
/// Fix (PDF temp copy only): replace each bottom border with a thin VML rectangle
/// of the same color/thickness and the same <c>space</c> gap. Never mutate resume.docx.
/// </summary>
internal static class PdfBorderHardener
{
    private const double DefaultContentWidthPt = 468; // letter 8.5" - 1" margins

    public static void HardenForPdfExport(WordprocessingDocument doc)
    {
        var widthPt = ResolveContentWidthPt(doc);

        if (doc.MainDocumentPart?.Document?.Body != null)
            HardenContainer(doc.MainDocumentPart.Document.Body, widthPt);

        if (doc.MainDocumentPart == null)
            return;

        foreach (var header in doc.MainDocumentPart.HeaderParts)
        {
            if (header.Header != null)
                HardenContainer(header.Header, widthPt);
        }

        foreach (var footer in doc.MainDocumentPart.FooterParts)
        {
            if (footer.Footer != null)
                HardenContainer(footer.Footer, widthPt);
        }

        doc.MainDocumentPart.Document?.Save();
    }

    private static void HardenContainer(OpenXmlElement container, double widthPt)
    {
        foreach (var para in container.Elements<Paragraph>().ToList())
        {
            var borders = para.ParagraphProperties?.ParagraphBorders;
            var bottom = borders?.BottomBorder;
            if (!IsVisibleBorder(bottom))
                continue;

            var color = ResolveColor(bottom!);
            var heightPt = BorderHeightPt(bottom!);
            var gapTwips = BorderGapTwips(bottom!);

            // Remove pBdr so Word doesn't try (and fail) to paint the border in PDF.
            bottom!.Remove();
            if (borders != null && !borders.ChildElements.Any())
                borders.Remove();

            var rule = new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Before = gapTwips.ToString(),
                        After = "0",
                        Line = "20",
                        LineRule = LineSpacingRuleValues.Exact
                    }),
                new Run(CreateVmlRule(color, widthPt, heightPt)));

            para.InsertAfterSelf(rule);
        }
    }

    private static Picture CreateVmlRule(string colorHex, double widthPt, double heightPt)
    {
        // VML filled rect — Word 2016 PDF keeps these; paragraph borders it drops.
        var rect = new V.Rectangle();
        rect.SetAttribute(new OpenXmlAttribute("id", null!, "SectionBar"));
        rect.SetAttribute(new OpenXmlAttribute(
            "style",
            null!,
            $"width:{FormatPt(widthPt)}pt;height:{FormatPt(heightPt)}pt;mso-wrap-style:none"));
        rect.SetAttribute(new OpenXmlAttribute("fillcolor", null!, "#" + colorHex));
        rect.SetAttribute(new OpenXmlAttribute("stroked", null!, "f"));
        rect.SetAttribute(new OpenXmlAttribute("filled", null!, "t"));

        var fill = new V.Fill();
        fill.SetAttribute(new OpenXmlAttribute("type", null!, "solid"));
        fill.SetAttribute(new OpenXmlAttribute("color", null!, "#" + colorHex));
        rect.Append(fill);

        var pict = new Picture();
        pict.Append(rect);
        return pict;
    }

    private static string FormatPt(double pt)
        => pt.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

    private static double ResolveContentWidthPt(WordprocessingDocument doc)
    {
        try
        {
            var sect = doc.MainDocumentPart?.Document?.Body?.Elements<SectionProperties>().LastOrDefault()
                       ?? doc.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sect?.GetFirstChild<PageSize>();
            var pgMar = sect?.GetFirstChild<PageMargin>();
            if (pgSz?.Width?.Value is { } pageTwips && pageTwips > 0)
            {
                var left = pgMar?.Left?.Value ?? 1440U;
                var right = pgMar?.Right?.Value ?? 1440U;
                var contentTwips = (double)pageTwips - left - right;
                if (contentTwips > 0)
                    return contentTwips / 20.0;
            }
        }
        catch
        {
            // fall through
        }

        return DefaultContentWidthPt;
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

    /// <summary>OOXML border sz is eighths of a point.</summary>
    private static double BorderHeightPt(BorderType border)
    {
        var eighths = border.Size?.Value ?? 8;
        if (eighths < 1)
            eighths = 1;
        return eighths / 8.0;
    }

    /// <summary>Border space is points → twips.</summary>
    private static int BorderGapTwips(BorderType border)
    {
        var points = border.Space?.Value ?? 0;
        return Math.Clamp((int)Math.Round(points * 20.0, MidpointRounding.AwayFromZero), 0, 240);
    }
}
