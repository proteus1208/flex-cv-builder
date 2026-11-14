using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Word often splits placeholder text across multiple runs. Helpers rebuild paragraph text
/// and rewrite run contents while preserving the first run's formatting where possible.
/// </summary>
internal static class RunTextMerger
{
    public static string GetParagraphText(Paragraph paragraph)
    {
        var sb = new StringBuilder();
        foreach (var text in paragraph.Descendants<Text>())
            sb.Append(text.Text);
        return sb.ToString();
    }

    /// <summary>
    /// Replaces the full visible text of a paragraph, keeping formatting from the first run that had text.
    /// </summary>
    public static void SetParagraphText(Paragraph paragraph, string newText)
    {
        var runs = paragraph.Elements<Run>().ToList();
        Run? formatRun = runs.FirstOrDefault(r => r.GetFirstChild<Text>() != null) ?? runs.FirstOrDefault();

        foreach (var run in runs)
            run.Remove();

        var newRun = new Run();
        if (formatRun?.RunProperties != null)
            newRun.RunProperties = (RunProperties)formatRun.RunProperties.CloneNode(true);

        var text = new Text(newText)
        {
            Space = SpaceProcessingModeValues.Preserve
        };
        newRun.AppendChild(text);

        var pPr = paragraph.GetFirstChild<ParagraphProperties>();
        if (pPr != null)
            pPr.InsertAfterSelf(newRun);
        else
            paragraph.PrependChild(newRun);
    }

    public static IEnumerable<Paragraph> GetParagraphs(OpenXmlElement root)
        => root.Descendants<Paragraph>();
}
