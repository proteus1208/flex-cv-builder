using System.Text.RegularExpressions;

namespace CvGenerator.Docx.Internal;

internal static partial class PlaceholderPatterns
{
    // ${ArrayStart name} / ${ArrayEnd name}
    [GeneratedRegex(@"\$\{\s*ArrayStart\s+([A-Za-z_][A-Za-z0-9_.]*)\s*\}", RegexOptions.IgnoreCase)]
    public static partial Regex ArrayStart();

    [GeneratedRegex(@"\$\{\s*ArrayEnd\s+([A-Za-z_][A-Za-z0-9_.]*)\s*\}", RegexOptions.IgnoreCase)]
    public static partial Regex ArrayEnd();

    // ${arrayname_i.field.sub}
    [GeneratedRegex(@"^([A-Za-z_][A-Za-z0-9_]*)_i\.([A-Za-z_][A-Za-z0-9_.]*)$")]
    public static partial Regex LoopField();

    [GeneratedRegex(@"\$\{([^{}]+)\}")]
    public static partial Regex AnyPlaceholder();

    public static bool IsArrayStart(string text, out string arrayName)
    {
        var m = ArrayStart().Match(text);
        arrayName = m.Success ? m.Groups[1].Value : "";
        return m.Success;
    }

    public static bool IsArrayEnd(string text, out string arrayName)
    {
        var m = ArrayEnd().Match(text);
        arrayName = m.Success ? m.Groups[1].Value : "";
        return m.Success;
    }

    public static bool IsLoopField(string inner, out string arrayName, out string fieldPath)
    {
        var m = LoopField().Match(inner.Trim());
        if (!m.Success)
        {
            arrayName = "";
            fieldPath = "";
            return false;
        }

        arrayName = m.Groups[1].Value;
        fieldPath = m.Groups[2].Value;
        return true;
    }
}
