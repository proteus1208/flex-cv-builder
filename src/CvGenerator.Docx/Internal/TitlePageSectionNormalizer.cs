using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Templates often use a title-page section with a large <c>w:pgMar w:top</c> so the
/// first-page header (name block) clears the body. That top margin applies to every
/// page of the section, so overflow onto page 2 can leave a large top gap.
///
/// Auto-inserting a section break at the estimated overflow point was tried, but Word
/// treats a margin-changing section break as a hard page break — which landed mid-body
/// (e.g. between experience bullets). That forced break is worse than the top gap, so
/// this normalizer does not inject section breaks. Templates that need different page-1
/// vs page-2 margins should place their own section break after first-page content.
/// </summary>
internal static class TitlePageSectionNormalizer
{
    public static void Normalize(Body body)
    {
        _ = body;
    }
}
