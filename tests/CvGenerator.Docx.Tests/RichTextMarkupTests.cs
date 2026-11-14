using System.Text.Json;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Internal;
using CvGenerator.Docx.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Tests;

public class RichTextMarkupTests
{
    [Fact]
    public void Parses_bold_and_link_markup()
    {
        var segments = RichTextMarkup.Parse("Hello **Ada** and [site](https://example.com)");
        Assert.Equal(4, segments.Count);
        Assert.Equal(RichSegmentKind.Text, segments[0].Kind);
        Assert.Equal(RichSegmentKind.Bold, segments[1].Kind);
        Assert.Equal("Ada", segments[1].Text);
        Assert.Equal(RichSegmentKind.Link, segments[3].Kind);
        Assert.Equal("https://example.com", segments[3].Url);
    }

    [Fact]
    public void Docx_renders_bold_and_hyperlink_from_json()
    {
        var template = Path.Combine(Path.GetTempPath(), "cvgen-rich-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var doc = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("${summary}"))),
                new Paragraph(new Run(new Text("${linkedin}")))));
            main.Document.Save();
        }

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "summary": "Shipped **PDF** export",
              "linkedin": "[Cody](https://linkedin.com/in/cody)"
            }
            """);
            new DocxTemplateEngine().Generate(template, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var body = result.MainDocumentPart!.Document.Body!;
            Assert.Contains(body.Descendants<Bold>(), _ => true);
            Assert.Contains(body.Descendants<Hyperlink>(), h => h.Id != null);
            Assert.Contains("PDF", body.InnerText);
            Assert.Contains("Cody", body.InnerText);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
