using System.IO;
using System.Linq;
using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CvGenerator.Docx.Tests;

public class PdfBorderHardenerTests
{
    [Fact]
    public void HardenForPdfExport_Replaces_BottomBorder_With_VmlRule_SameColorAndThickness()
    {
        var path = Path.Combine(Path.GetTempPath(), "cvgen-pdf-bar-" + Path.GetRandomFileName() + ".docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines { Before = "100", After = "60" },
                            new ParagraphBorders(
                                new BottomBorder
                                {
                                    Val = BorderValues.Single,
                                    Size = 9,
                                    Color = "2F5597",
                                    Space = 2
                                })),
                        new Run(new Text("Professional Summary"))),
                    new Paragraph(new Run(new Text("Body")))));
                main.Document.Save();
            }

            using (var doc = WordprocessingDocument.Open(path, true))
            {
                PdfBorderHardener.HardenForPdfExport(doc);
            }

            using var outDoc = WordprocessingDocument.Open(path, false);
            var paras = outDoc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
            Assert.Equal(3, paras.Count); // heading + VML rule + body
            Assert.Null(paras[0].ParagraphProperties?.ParagraphBorders?.BottomBorder);
            // Heading after must move below the bar (not sit above it).
            Assert.Equal("0", paras[0].ParagraphProperties?.SpacingBetweenLines?.After?.Value);

            var pict = paras[1].Descendants<Picture>().FirstOrDefault();
            Assert.NotNull(pict);
            var xml = pict!.OuterXml;
            Assert.Contains("2F5597", xml, System.StringComparison.OrdinalIgnoreCase);
            // sz=9 → 9/8 = 1.125pt
            Assert.Contains("1.125pt", xml, System.StringComparison.Ordinal);
            Assert.Equal("40", paras[1].ParagraphProperties?.SpacingBetweenLines?.Before?.Value); // space=2pt
            Assert.Equal("60", paras[1].ParagraphProperties?.SpacingBetweenLines?.After?.Value); // original after
            Assert.Equal("Body", paras[2].InnerText);
        }
        finally
        {
            try { File.Delete(path); } catch { /* ignore */ }
        }
    }
}
