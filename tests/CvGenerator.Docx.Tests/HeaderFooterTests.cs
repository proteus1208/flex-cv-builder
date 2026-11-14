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
