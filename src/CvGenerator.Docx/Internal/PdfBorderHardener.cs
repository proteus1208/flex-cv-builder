using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// PDF-export helper only. Never call during normal DOCX generation — output.docx
/// must keep the template's <c>w:pBdr</c> bars unchanged.
/// On a temp copy, rewrite bottom borders as a single-cell table bottom border
/// (thin rule), which Word 2016 PDF export keeps more reliably than paragraph borders.
/// </summary>
internal static class PdfBorderHardener
{
    /// <summary>
    /// Mutates <paramref name="doc"/> in place. Use only on a temp clone for PDF export.
    /// </summary>
    public static void HardenForPdfExport(WordprocessingDocument doc)
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
        foreach (var para in container.Elements<Paragraph>().ToList())
        {
            var borders = para.ParagraphProperties?.ParagraphBorders;
            var bottom = borders?.BottomBorder;
            if (!IsVisibleBorder(bottom))
                continue;

            var color = ResolveColor(bottom!);
            var size = bottom!.Size?.Value ?? 8;
            var gapTwips = BorderGapTwips(bottom);

            bottom.Remove();
            if (borders != null && !borders.ChildElements.Any())
                borders.Remove();

            // Full-width 1-row table with ONLY a bottom border = thin horizontal rule.
            var table = new Table(
                new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Nil },
                        new LeftBorder { Val = BorderValues.Nil },
                        new BottomBorder
                        {
                            Val = BorderValues.Single,
                            Size = size,
                            Color = color,
                            Space = 0
                        },
                        new RightBorder { Val = BorderValues.Nil },
                        new InsideHorizontalBorder { Val = BorderValues.Nil },
                        new InsideVerticalBorder { Val = BorderValues.Nil }),
                    new TableCellMarginDefault(
                        new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                        new TableCellLeftMargin { Width = 0, Type = TableWidthValues.Dxa },
                        new TableCellRightMargin { Width = 0, Type = TableWidthValues.Dxa })),
                new TableRow(
                    new TableRowProperties(
                        new TableRowHeight
                        {
                            Val = (uint)Math.Max(20, gapTwips > 0 ? gapTwips : 20),
                            HeightType = HeightRuleValues.Exact
                        }),
                    new TableCell(
                        new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                            new TableCellBorders(
                                new TopBorder { Val = BorderValues.Nil },
                                new LeftBorder { Val = BorderValues.Nil },
                                new BottomBorder
                                {
                                    Val = BorderValues.Single,
                                    Size = size,
                                    Color = color,
                                    Space = 0
                                },
                                new RightBorder { Val = BorderValues.Nil })),
                        new Paragraph(
                            new ParagraphProperties(
                                new SpacingBetweenLines
                                {
                                    Before = "0",
                                    After = "0",
                                    Line = "20",
                                    LineRule = LineSpacingRuleValues.Exact
                                })))));

            para.InsertAfterSelf(table);
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

    private static int BorderGapTwips(BorderType border)
    {
        var points = border.Space?.Value ?? 0;
        return Math.Clamp((int)Math.Round(points * 20.0, MidpointRounding.AwayFromZero), 0, 240);
    }
}
