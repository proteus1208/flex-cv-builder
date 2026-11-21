using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Templates often use a title-page section with a large <c>w:pgMar w:top</c> so the
/// first-page header (name block) clears the body. That top margin is real header
/// clearance — not body/content padding — but it applies to every page of the section.
///
/// When loops expand and the section spills past page 1, later pages inherit the huge
/// top gap. Fix: keep the large top margin on the title section (header padding stays
/// header padding), and insert a next-page section break with continuation margins
/// once estimated content no longer fits on the first page.
/// </summary>
internal static class TitlePageSectionNormalizer
{
    private const int DefaultPageHeightTwips = 15840;
    private const int DefaultParaTwips = 320;
    private const int MinSplitParas = 4;

    public static void Normalize(Body body)
    {
        var sections = CollectSections(body);
        if (sections.Count == 0)
            return;

        // Work back-to-front so inserted breaks do not shift later indices.
        for (var i = sections.Count - 1; i >= 0; i--)
        {
            var section = sections[i];
            if (section.SectPr.GetFirstChild<TitlePage>() is null)
                continue;

            var margin = section.SectPr.GetFirstChild<PageMargin>();
            if (margin?.Top?.Value is null)
                continue;

            var titleTop = margin.Top.Value;
            var continuationTop = ResolveContinuationTop(sections, i);
            if (titleTop <= continuationTop)
                continue;

            var elements = GetSectionElements(body, section).ToList();
            if (elements.Count < MinSplitParas)
                continue;

            var pageHeight = (int)(section.SectPr.GetFirstChild<PageSize>()?.Height?.Value
                                   ?? (uint)DefaultPageHeightTwips);
            var bottom = (int)(margin.Bottom?.Value ?? 720);
            // Stay conservative — under-estimating height leaves huge top margin on page 2.
            var usable = Math.Max(
                (int)((pageHeight - titleTop - bottom) * 0.82),
                DefaultParaTwips * 6);

            var splitAt = FindSplitIndex(elements, usable);
            if (splitAt is null)
                continue; // still fits on one title page

            SplitTitleSectionAtOverflow(
                elements[splitAt.Value],
                section.SectPr,
                sections,
                i,
                continuationTop);
        }
    }

    private static int? FindSplitIndex(IReadOnlyList<OpenXmlElement> elements, int usableTwips)
    {
        var cumulative = 0;
        // Keep at least one block on the title page; split before the overflowing element.
        for (var i = 0; i < elements.Count; i++)
        {
            cumulative += EstimateHeightTwips(elements[i]);
            if (cumulative <= usableTwips)
                continue;
            // Need a later element to become the start of the continuation section.
            if (i <= 0 || i >= elements.Count)
                return null;
            return i;
        }

        return null;
    }

    /// <summary>
    /// Ends the title-page section after the last paragraph that fits page 1 (keeps large
    /// header top margin + titlePg there), and converts the original title sectPr into a
    /// continuation section for overflow content.
    /// </summary>
    private static void SplitTitleSectionAtOverflow(
        OpenXmlElement splitElement,
        SectionProperties titleSectPr,
        IReadOnlyList<SectionSpan> sections,
        int titleIndex,
        int continuationTop)
    {
        SectionProperties? prototype = titleIndex + 1 < sections.Count
            ? sections[titleIndex + 1].SectPr
            : null;

        var splitPara = splitElement as Paragraph
                        ?? splitElement.ElementsBefore().OfType<Paragraph>().LastOrDefault();
        if (splitPara is null)
            return;

        var previous = splitPara.PreviousSibling() as Paragraph
                       ?? splitPara.ElementsBefore().OfType<Paragraph>().LastOrDefault();
        if (previous is null)
            return;

        if (previous.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
            return; // already a break here

        // Clone title-page sectPr onto the last in-page paragraph so section 1 keeps
        // header clearance (large top + titlePg + first header).
        var titleEnd = (SectionProperties)titleSectPr.CloneNode(true);
        if (titleEnd.GetFirstChild<SectionType>() is null)
            titleEnd.PrependChild(new SectionType { Val = SectionMarkValues.NextPage });
        else
            titleEnd.GetFirstChild<SectionType>()!.Val = SectionMarkValues.NextPage;

        var pPr = previous.ParagraphProperties ??= new ParagraphProperties();
        pPr.AppendChild(titleEnd);

        // Original break now starts the overflow/continuation region (normal top margin).
        ConvertToContinuation(titleSectPr, prototype, continuationTop);
    }

    private static void ConvertToContinuation(
        SectionProperties sectPr,
        SectionProperties? prototype,
        int continuationTop)
    {
        sectPr.GetFirstChild<TitlePage>()?.Remove();

        if (sectPr.GetFirstChild<SectionType>() is null)
            sectPr.PrependChild(new SectionType { Val = SectionMarkValues.NextPage });
        else
            sectPr.GetFirstChild<SectionType>()!.Val = SectionMarkValues.NextPage;

        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin != null)
        {
            margin.Top = continuationTop;
            var protoMar = prototype?.GetFirstChild<PageMargin>();
            if ((margin.Header?.Value ?? 0) == 0 && protoMar?.Header?.Value is > 0)
                margin.Header = protoMar.Header.Value;
        }

        if (!sectPr.Elements<HeaderReference>().Any(h => h.Type?.Value == HeaderFooterValues.Default))
        {
            var defaultId = prototype?.Elements<HeaderReference>()
                .FirstOrDefault(h => h.Type?.Value == HeaderFooterValues.Default)
                ?.Id?.Value;
            if (!string.IsNullOrEmpty(defaultId))
            {
                var first = sectPr.Elements<HeaderReference>().FirstOrDefault();
                var reference = new HeaderReference { Type = HeaderFooterValues.Default, Id = defaultId };
                if (first is null)
                    sectPr.PrependChild(reference);
                else
                    first.InsertBeforeSelf(reference);
            }
        }
    }

    private static int EstimateHeightTwips(OpenXmlElement el)
    {
        if (el is Table)
            return DefaultParaTwips * 3;

        if (el is not Paragraph para)
            return DefaultParaTwips;

        var textLen = para.InnerText.Length;
        var spacing = para.ParagraphProperties?.GetFirstChild<SpacingBetweenLines>();
        var before = ParseTwips(spacing?.Before?.Value);
        var after = ParseTwips(spacing?.After?.Value);
        var line = ParseTwips(spacing?.Line?.Value);
        if (line <= 0)
            line = 240;

        // Rough wrap estimate at ~90 chars/line for letter width with ~0.55" margins.
        var lines = Math.Max(1, (int)Math.Ceiling(Math.Max(textLen, 1) / 90.0));
        if (para.ParagraphProperties?.GetFirstChild<KeepLines>() != null && textLen > 90)
            lines = Math.Max(lines, (int)Math.Ceiling(textLen / 90.0));

        return before + after + lines * line;
    }

    private static int ParseTwips(string? value)
        => int.TryParse(value, out var n) ? n : 0;

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

    private static IEnumerable<OpenXmlElement> GetSectionElements(Body body, SectionSpan section)
    {
        var started = section.FirstContent is null;
        foreach (var el in body.ChildElements)
        {
            if (!started)
            {
                if (ReferenceEquals(el, section.FirstContent))
                    started = true;
                else
                    continue;
            }

            if (el is Paragraph p &&
                p.ParagraphProperties?.GetFirstChild<SectionProperties>() is { } sect &&
                ReferenceEquals(sect, section.SectPr))
            {
                yield break; // section break paragraph ends this section; exclude it
            }

            if (el is SectionProperties)
                yield break;

            yield return el;
        }
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
