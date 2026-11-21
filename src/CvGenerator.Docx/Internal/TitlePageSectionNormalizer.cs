using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Templates often use a title-page section with a large <c>w:pgMar w:top</c> so the
/// first-page header (name block) clears the body. That top margin applies to every
/// page of the section. When loops expand and the section spans multiple pages,
/// page 2+ inherit the same huge top padding. Move the extra space onto the first
/// body paragraph and use the continuation section's top margin for the section.
/// </summary>
internal static class TitlePageSectionNormalizer
{
    public static void Normalize(Body body)
    {
        var sections = CollectSections(body);
        if (sections.Count < 2)
            return;

        for (var i = 0; i < sections.Count; i++)
        {
            var section = sections[i];
            if (section.SectPr.GetFirstChild<TitlePage>() is null)
                continue;

            var margin = section.SectPr.GetFirstChild<PageMargin>();
            if (margin?.Top?.Value is null)
                continue;

            var top = margin.Top.Value;
            var continuationTop = ResolveContinuationTop(sections, i);
            if (top <= continuationTop)
                continue;

            var delta = top - continuationTop;
            margin.Top = continuationTop;

            // Match continuation header distance when title section used header=0
            // (padding was entirely in pgMar.top).
            if ((margin.Header?.Value ?? 0) == 0 && i + 1 < sections.Count)
            {
                var nextHeader = sections[i + 1].SectPr.GetFirstChild<PageMargin>()?.Header?.Value;
                if (nextHeader is > 0)
                    margin.Header = nextHeader;
            }

            EnsureDefaultHeaderReference(section.SectPr, sections, i);
            ApplyFirstPageSpacer(section.FirstContent, delta);
        }
    }

    private static int ResolveContinuationTop(IReadOnlyList<SectionSpan> sections, int index)
    {
        for (var j = index + 1; j < sections.Count; j++)
        {
            var top = sections[j].SectPr.GetFirstChild<PageMargin>()?.Top?.Value;
            if (top is not null)
                return top.Value;
        }

        return 720;
    }

    private static void EnsureDefaultHeaderReference(
        SectionProperties sectPr,
        IReadOnlyList<SectionSpan> sections,
        int index)
    {
        if (sectPr.Elements<HeaderReference>().Any(h => h.Type?.Value == HeaderFooterValues.Default))
            return;

        string? defaultId = null;
        for (var j = index + 1; j < sections.Count && defaultId is null; j++)
        {
            defaultId = sections[j].SectPr.Elements<HeaderReference>()
                .FirstOrDefault(h => h.Type?.Value == HeaderFooterValues.Default)
                ?.Id?.Value;
        }

        if (string.IsNullOrEmpty(defaultId))
            return;

        var firstHeader = sectPr.Elements<HeaderReference>().FirstOrDefault();
        var reference = new HeaderReference { Type = HeaderFooterValues.Default, Id = defaultId };
        if (firstHeader is null)
            sectPr.PrependChild(reference);
        else
            firstHeader.InsertBeforeSelf(reference);
    }

    private static void ApplyFirstPageSpacer(OpenXmlElement? firstContent, int deltaTwips)
    {
        if (deltaTwips <= 0 || firstContent is null)
            return;

        if (firstContent is not Paragraph firstPara)
        {
            firstPara = new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines { Before = deltaTwips.ToString(), After = "0" }));
            firstContent.InsertBeforeSelf(firstPara);
            return;
        }

        var pPr = firstPara.ParagraphProperties ??= new ParagraphProperties();
        var spacing = pPr.GetFirstChild<SpacingBetweenLines>();
        if (spacing is null)
        {
            spacing = new SpacingBetweenLines();
            pPr.PrependChild(spacing);
        }

        var existing = 0;
        if (spacing.Before?.Value is { } beforeText)
            int.TryParse(beforeText, out existing);

        spacing.Before = (existing + deltaTwips).ToString();
    }

    private static List<SectionSpan> CollectSections(Body body)
    {
        var sections = new List<SectionSpan>();
        OpenXmlElement? start = body.ChildElements.FirstOrDefault();

        foreach (var el in body.ChildElements)
        {
            if (el is not Paragraph para)
                continue;

            var sect = para.ParagraphProperties?.GetFirstChild<SectionProperties>();
            if (sect is null)
                continue;

            sections.Add(new SectionSpan(start, sect));
            start = para.NextSibling();
        }

        var final = body.Elements<SectionProperties>().LastOrDefault();
        if (final is not null)
            sections.Add(new SectionSpan(start, final));

        return sections;
    }

    private readonly record struct SectionSpan(OpenXmlElement? FirstContent, SectionProperties SectPr);
}
