using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Internal;

/// <summary>
/// Word templates often store fonts as theme refs (minorHAnsi → Calibri) instead of
/// concrete names. LibreOffice resolves those differently (or substitutes Carlito),
/// so the PDF looks "wrong". Expand theme refs to real typeface names before PDF export.
/// </summary>
internal static class ThemeFontMaterializer
{
    public static void Materialize(WordprocessingDocument doc)
    {
        var map = LoadThemeFonts(doc);
        if (map.IsEmpty)
            return;

        if (doc.MainDocumentPart?.Document != null)
            MaterializeTree(doc.MainDocumentPart.Document, map);

        if (doc.MainDocumentPart?.StyleDefinitionsPart?.Styles != null)
            MaterializeTree(doc.MainDocumentPart.StyleDefinitionsPart.Styles, map);

        if (doc.MainDocumentPart != null)
        {
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                if (header.Header != null)
                    MaterializeTree(header.Header, map);
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                if (footer.Footer != null)
                    MaterializeTree(footer.Footer, map);
            }
        }
    }

    private static void MaterializeTree(DocumentFormat.OpenXml.OpenXmlElement root, ThemeMap map)
    {
        foreach (var rFonts in root.Descendants<W.RunFonts>())
            MaterializeRunFonts(rFonts, map);
    }

    private static void MaterializeRunFonts(W.RunFonts rFonts, ThemeMap map)
    {
        if (ThemeRefText(rFonts.AsciiTheme) is { } asciiTheme)
        {
            var name = map.Resolve(asciiTheme);
            if (!string.IsNullOrWhiteSpace(name))
            {
                rFonts.Ascii = name;
                rFonts.AsciiTheme = null;
            }
        }

        if (ThemeRefText(rFonts.HighAnsiTheme) is { } hAnsiTheme)
        {
            var name = map.Resolve(hAnsiTheme);
            if (!string.IsNullOrWhiteSpace(name))
            {
                rFonts.HighAnsi = name;
                rFonts.HighAnsiTheme = null;
            }
        }

        if (ThemeRefText(rFonts.EastAsiaTheme) is { } eaTheme)
        {
            var name = map.Resolve(eaTheme);
            if (!string.IsNullOrWhiteSpace(name))
            {
                rFonts.EastAsia = name;
                rFonts.EastAsiaTheme = null;
            }
        }

        if (ThemeRefText(rFonts.ComplexScriptTheme) is { } csTheme)
        {
            var name = map.Resolve(csTheme);
            if (!string.IsNullOrWhiteSpace(name))
            {
                rFonts.ComplexScript = name;
                rFonts.ComplexScriptTheme = null;
            }
        }
    }

    private static string? ThemeRefText(EnumValue<ThemeFontValues>? value)
    {
        if (value is null || !value.HasValue)
            return null;
        // Prefer the XML token (majorAscii / minorHAnsi) over ToString().
        if (!string.IsNullOrWhiteSpace(value.InnerText))
            return value.InnerText;
        return value.Value.ToString();
    }

    private static ThemeMap LoadThemeFonts(WordprocessingDocument doc)
    {
        var map = new ThemeMap();
        var themePart = doc.MainDocumentPart?.ThemePart;
        var scheme = themePart?.Theme?.ThemeElements?.GetFirstChild<A.FontScheme>();
        if (scheme is null)
            return map;

        var major = scheme.MajorFont;
        var minor = scheme.MinorFont;

        map.MajorLatin = major?.LatinFont?.Typeface?.Value;
        map.MajorEastAsia = major?.EastAsianFont?.Typeface?.Value;
        map.MajorCs = major?.ComplexScriptFont?.Typeface?.Value;

        map.MinorLatin = minor?.LatinFont?.Typeface?.Value;
        map.MinorEastAsia = minor?.EastAsianFont?.Typeface?.Value;
        map.MinorCs = minor?.ComplexScriptFont?.Typeface?.Value;

        return map;
    }

    private sealed class ThemeMap
    {
        public string? MajorLatin { get; set; }
        public string? MajorEastAsia { get; set; }
        public string? MajorCs { get; set; }
        public string? MinorLatin { get; set; }
        public string? MinorEastAsia { get; set; }
        public string? MinorCs { get; set; }

        public bool IsEmpty =>
            string.IsNullOrWhiteSpace(MajorLatin) &&
            string.IsNullOrWhiteSpace(MinorLatin) &&
            string.IsNullOrWhiteSpace(MajorEastAsia) &&
            string.IsNullOrWhiteSpace(MinorEastAsia) &&
            string.IsNullOrWhiteSpace(MajorCs) &&
            string.IsNullOrWhiteSpace(MinorCs);

        public string? Resolve(string themeRef) => themeRef switch
        {
            "majorAscii" or "majorHAnsi" => MajorLatin,
            "minorAscii" or "minorHAnsi" => MinorLatin,
            "majorEastAsia" => MajorEastAsia,
            "minorEastAsia" => MinorEastAsia,
            "majorBidi" => MajorCs,
            "minorBidi" => MinorCs,
            _ => null
        };
    }
}
