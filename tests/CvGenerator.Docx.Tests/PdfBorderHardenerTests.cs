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
    public void Generate_Preserves_Paragraph_Bottom_Border_In_OutputDocx()
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

            // Output DOCX must keep the real border — no shaded filler paragraph.
            Assert.Equal(2, paragraphs.Count);
            var bottom = paragraphs[0].ParagraphProperties?.ParagraphBorders?.BottomBorder;
            Assert.NotNull(bottom);
            Assert.Equal(BorderValues.Single, bottom!.Val?.Value);
            Assert.Equal("CCD7E1", bottom.Color?.Value);
            Assert.Equal(4U, bottom.Size?.Value);
            Assert.Equal(3U, bottom.Space?.Value);
            Assert.Null(paragraphs[1].ParagraphProperties?.Shading?.Fill?.Value);
            Assert.Equal("Body text", paragraphs[1].InnerText);
        }
        finally
        {
            try { File.Delete(template); } catch { /* ignore */ }
            try { File.Delete(output); } catch { /* ignore */ }
        }
    }
}
