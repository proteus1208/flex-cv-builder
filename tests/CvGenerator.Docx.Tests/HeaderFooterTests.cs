using System.Text.Json;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Tests;

public class HeaderFooterTests
{
    [Fact]
    public void Title_page_large_top_margin_moves_to_first_paragraph_after_expand()
    {
        var path = Path.Combine(Path.GetTempPath(), "cvgen-title-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var header1 = main.AddNewPart<HeaderPart>("rIdFirst");
            header1.Header = new Header(new Paragraph(new Run(new Text("Name"))));
            var header2 = main.AddNewPart<HeaderPart>("rIdDefault");
            header2.Header = new Header(new Paragraph(new Run(new Text(""))));

            // Section 1 (title page): large top margin + first header only.
            // Section 2: normal top margin + default header.
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("${ArrayStart items}"))),
                new Paragraph(new Run(new Text("- ${items_i}"))),
                new Paragraph(new Run(new Text("${ArrayEnd items}"))),
                new Paragraph(
                    new ParagraphProperties(
                        new SectionProperties(
                            new HeaderReference { Type = HeaderFooterValues.First, Id = "rIdFirst" },
                            new PageMargin
                            {
                                Top = 2419,
                                Right = 720,
                                Bottom = 720,
                                Left = 720,
                                Header = 0,
                                Footer = 320
                            },
                            new TitlePage())),
                    new Run(new Text("Education"))),
                new Paragraph(new Run(new Text("Skills"))),
                new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdDefault" },
                    new PageMargin
                    {
                        Top = 792,
                        Right = 720,
                        Bottom = 720,
                        Left = 720,
                        Header = 288,
                        Footer = 320
                    })));
            main.Document.Save();
        }

        var output = Path.ChangeExtension(path, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""{"items":["a","b","c","d","e","f"]}""");
            new DocxTemplateEngine().Generate(path, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var body = result.MainDocumentPart!.Document.Body!;
            var breakPara = body.Elements<Paragraph>()
                .First(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null);
            var titleSect = breakPara.ParagraphProperties!.GetFirstChild<SectionProperties>()!;
            var titleTop = titleSect.GetFirstChild<PageMargin>()!.Top!.Value;
            Assert.Equal(792, titleTop);

            // Extra first-page padding moved onto the first body paragraph.
            var first = body.Elements<Paragraph>().First();
            var before = first.ParagraphProperties?.GetFirstChild<SpacingBetweenLines>()?.Before?.Value;
            Assert.Equal((2419 - 792).ToString(), before);

            // Default header wired so later pages of the title section stay empty.
            Assert.Contains(
                titleSect.Elements<HeaderReference>(),
                h => h.Type?.Value == HeaderFooterValues.Default && h.Id == "rIdDefault");
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void Replaces_placeholders_in_headers_and_footers()
    {
        var path = Path.Combine(Path.GetTempPath(), "cvgen-hf-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body ${name}"))),
                new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdH" },
                    new FooterReference { Type = HeaderFooterValues.Default, Id = "rIdF" })));

            var headerPart = main.AddNewPart<HeaderPart>("rIdH");
            headerPart.Header = new Header(new Paragraph(new Run(new Text("H ${name}"))));

            var footerPart = main.AddNewPart<FooterPart>("rIdF");
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text("F ${role}"))));

            main.Document.Save();
        }

        var output = Path.ChangeExtension(path, ".out.docx");
        try
        {
            using var json = JsonDocument.Parse("""{"name":"Ada","role":"Engineer"}""");
            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(path, output, json.RootElement, issues);

            using var result = WordprocessingDocument.Open(output, false);
            Assert.Contains("Ada", result.MainDocumentPart!.Document.Body!.InnerText);
            Assert.Contains("Ada", result.MainDocumentPart.HeaderParts.First().Header!.InnerText);
            Assert.Contains("Engineer", result.MainDocumentPart.FooterParts.First().Footer!.InnerText);
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
