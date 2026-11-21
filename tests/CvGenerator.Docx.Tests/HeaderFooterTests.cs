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
    public void Title_page_keeps_header_top_margin_and_splits_overflow_to_continuation()
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
            // Many array items force the title section past one page.
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("${ArrayStart items}"))),
                new Paragraph(new Run(new Text("- ${items_i}"))),
                new Paragraph(new Run(new Text("detail line for ${items_i} with more wrapping text here"))),
                new Paragraph(new Run(new Text("location ${items_i}"))),
                new Paragraph(new Run(new Text("${ArrayEnd items}"))),
                new Paragraph(
                    new ParagraphProperties(
                        new SectionProperties(
                            new HeaderReference { Type = HeaderFooterValues.First, Id = "rIdFirst" },
                            new PageSize { Width = 12240, Height = 15840 },
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
                    new PageSize { Width = 12240, Height = 15840 },
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
            using var json = JsonDocument.Parse(
                """{"items":["a1","a2","a3","a4","a5","a6","a7","a8","a9","a10","a11","a12","a13","a14","a15","a16","a17","a18","a19","a20"]}""");
            new DocxTemplateEngine().Generate(path, output, json.RootElement, new IssueCollector());

            using var result = WordprocessingDocument.Open(output, false);
            var body = result.MainDocumentPart!.Document.Body!;
            var sectionBreaks = body.Elements<Paragraph>()
                .Select(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>())
                .Where(s => s != null)
                .Cast<SectionProperties>()
                .ToList();

            Assert.True(sectionBreaks.Count >= 2, "overflow should insert a continuation section break");

            // First section break must still be the title page (header padding via pgMar.top).
            var titleSect = sectionBreaks.First(s => s.GetFirstChild<TitlePage>() != null);
            Assert.Equal(2419, titleSect.GetFirstChild<PageMargin>()!.Top!.Value);
            Assert.Equal(
                SectionMarkValues.NextPage,
                titleSect.GetFirstChild<SectionType>()?.Val?.Value);

            // First body paragraph must NOT receive content spacing.before hacks.
            var first = body.Elements<Paragraph>().First();
            var before = first.ParagraphProperties?.GetFirstChild<SpacingBetweenLines>()?.Before?.Value;
            Assert.True(string.IsNullOrEmpty(before) || before == "0" || int.Parse(before!) < 500);

            // Later section(s) use continuation top margin (no titlePg).
            var continuation = sectionBreaks.First(s => s.GetFirstChild<TitlePage>() is null);
            Assert.Equal(792, continuation.GetFirstChild<PageMargin>()!.Top!.Value);
            Assert.Contains(
                continuation.Elements<HeaderReference>(),
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
