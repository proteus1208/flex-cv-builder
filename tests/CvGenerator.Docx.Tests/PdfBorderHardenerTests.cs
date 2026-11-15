using System.IO;
using System.Linq;
using System.Text.Json;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CvGenerator.Docx.Tests;

public class PdfBorderHardenerTests
{
    [Fact]
    public void Generate_Converts_Paragraph_Bottom_Border_To_Shaded_Bar()
    {
        var template = Path.Combine(Path.GetTempPath(), "cvgen-bar-" + Path.GetRandomFileName() + ".docx");
        var output = Path.Combine(Path.GetTempPath(), "cvgen-bar-out-" + Path.GetRandomFileName() + ".docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphBorders(
                                new BottomBorder
                                {
                                    Val = BorderValues.Single,
                                    Size = 4,
                                    Color = "CCD7E1",
                                    Space = 3
                                })),
                        new Run(new Text("PROFILE"))),
                    new Paragraph(new Run(new Text("Body text")))));
                main.Document.Save();
            }

            var issues = new IssueCollector();
            new DocxTemplateEngine().Generate(template, output, JsonDocument.Parse("{}").RootElement, issues);

            using var outDoc = WordprocessingDocument.Open(output, false);
            var body = outDoc.MainDocumentPart!.Document.Body!;
            var paragraphs = body.Elements<Paragraph>().ToList();

            Assert.Equal(3, paragraphs.Count); // heading + shaded bar + body
            Assert.DoesNotContain(
                paragraphs[0].Descendants<BottomBorder>(),
                b => b.Val?.Value == BorderValues.Single);

            var bar = paragraphs[1];
            Assert.Equal("CCD7E1", bar.ParagraphProperties?.Shading?.Fill?.Value);
            Assert.Equal(LineSpacingRuleValues.Exact, bar.ParagraphProperties?.SpacingBetweenLines?.LineRule?.Value);
            Assert.Equal("Body text", bar.NextSibling<Paragraph>()?.InnerText);
        }
        finally
        {
            try { File.Delete(template); } catch { /* ignore */ }
            try { File.Delete(output); } catch { /* ignore */ }
        }
    }
}
