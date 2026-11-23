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
    public void Parses_bold_italic_underline_and_link_markup()
    {
        var segments = RichTextMarkup.Parse(
            "Hello **Ada**, *nice*, __ready__, and [site](https://example.com)");
        Assert.Equal(8, segments.Count);
        Assert.Equal(RichSegmentKind.Text, segments[0].Kind);
        Assert.Equal(RichSegmentKind.Bold, segments[1].Kind);
        Assert.Equal("Ada", segments[1].Text);
        Assert.Equal(RichSegmentKind.Italic, segments[3].Kind);
        Assert.Equal("nice", segments[3].Text);
        Assert.Equal(RichSegmentKind.Underline, segments[5].Kind);
        Assert.Equal("ready", segments[5].Text);
        Assert.Equal(RichSegmentKind.Link, segments[7].Kind);
        Assert.Equal("https://example.com", segments[7].Url);
    }

    [Fact]
    public void Inline_array_preserves_markdown_link_labels()
    {
        var template = Path.Combine(Path.GetTempPath(), "cvgen-rich-arr-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var doc = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(
                    "${ArrayStart extra_infos info}${info}${JoinStart} | ${JoinEnd}${ArrayEnd extra_infos}")))));
            main.Document.Save();
        }

        var output = Path.ChangeExtension(template, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""
            {
              "extra_infos": [
                "codyleeberry1999@gmail.com",
                "[LinkedIn](https://www.linkedin.com/in/cody-lee-berry-603179449/)",
                "[GitHub](https://github.com/proteus1208)",
                "[Portfolio](https://cody-lee-berry.vercel.app/)"
              ]
            }
            """);
            new DocxTemplateEngine().Generate(template, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var body = result.MainDocumentPart!.Document.Body!;
            var text = body.InnerText;
            Assert.Contains("LinkedIn", text);
            Assert.Contains("GitHub", text);
            Assert.Contains("Portfolio", text);
            Assert.DoesNotContain("[LinkedIn]", text);
            Assert.DoesNotContain("https://github.com", text);

            var labels = body.Descendants<Hyperlink>()
                .Select(h => h.InnerText)
                .ToList();
            Assert.Contains("LinkedIn", labels);
            Assert.Contains("GitHub", labels);
            Assert.Contains("Portfolio", labels);
            Assert.Equal(3, body.Descendants<Hyperlink>().Count(h => h.Id != null));
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Docx_renders_rich_markup_from_json()
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
              "summary": "Shipped **PDF** with *speed* and __care__",
              "linkedin": "[Cody](https://linkedin.com/in/cody)"
            }
            """);
            new DocxTemplateEngine().Generate(template, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var body = result.MainDocumentPart!.Document.Body!;
            Assert.Contains(body.Descendants<Bold>(), _ => true);
            Assert.Contains(body.Descendants<Italic>(), _ => true);
            Assert.Contains(body.Descendants<Underline>(), u => u.Val?.Value == UnderlineValues.Single);
            Assert.Contains(body.Descendants<Hyperlink>(), h => h.Id != null);
            Assert.Contains("PDF", body.InnerText);
            Assert.Contains("speed", body.InnerText);
            Assert.Contains("care", body.InnerText);
            Assert.Contains("Cody", body.InnerText);
        }
        finally
        {
            File.Delete(template);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
