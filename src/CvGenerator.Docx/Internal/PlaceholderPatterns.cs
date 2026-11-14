using System.Text.RegularExpressions;

namespace CvGenerator.Docx.Internal;

internal sealed record LoopReference(IReadOnlyList<string> Chain, string? FieldPath);

internal static partial class PlaceholderPatterns
{
    // ${ArrayStart name} / ${ArrayEnd name} — name may be "skills" or "skills_i.detail"
    [GeneratedRegex(@"\$\{\s*ArrayStart\s+([A-Za-z_][A-Za-z0-9_.]*)\s*\}", RegexOptions.IgnoreCase)]
    public static partial Regex ArrayStart();

    [GeneratedRegex(@"\$\{\s*ArrayEnd\s+([A-Za-z_][A-Za-z0-9_.]*)\s*\}", RegexOptions.IgnoreCase)]
    public static partial Regex ArrayEnd();

    [GeneratedRegex(@"\$\{([^{}]+)\}")]
    public static partial Regex AnyPlaceholder();

    // No ^ — Match(input, startAt) still anchors ^ at index 0, so nested
    // segments like detail_i in skills_i.detail_i.name must be unanchored.
    [GeneratedRegex(@"([A-Za-z_][A-Za-z0-9_]*)_i")]
    private static partial Regex LoopSegment();

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

    /// <summary>
    /// Parses loop field refs:
    ///   skills_i.category
    ///   skills_i.detail_i
    ///   skills_i.detail_i.name
    /// </summary>
    public static bool TryParseLoopReference(string inner, out LoopReference reference)
    {
        reference = null!;
        var s = inner.Trim();
        if (!s.Contains("_i", StringComparison.Ordinal))
            return false;

        var chain = new List<string>();
        var i = 0;
        while (i < s.Length)
        {
            var m = LoopSegment().Match(s, i);
            if (!m.Success || m.Index != i)
                break;

            chain.Add(m.Groups[1].Value);
            i = m.Index + m.Length;
            if (i >= s.Length)
            {
                reference = new LoopReference(chain, null);
                return chain.Count > 0;
            }

            if (s[i] == '.')
            {
                i++;
                continue;
            }

            break;
        }

        if (chain.Count == 0)
            return false;

        var field = i < s.Length ? s[i..] : null;
        if (string.IsNullOrEmpty(field))
        {
            reference = new LoopReference(chain, null);
            return true;
        }

        // Remaining must be a normal dotted path (no _i).
        if (field.Contains("_i", StringComparison.Ordinal))
            return false;

        reference = new LoopReference(chain, field);
        return true;
    }

    /// <summary>
    /// ArrayStart/ArrayEnd name:
    ///   skills              → parents=[], property=skills
    ///   skills_i.detail     → parents=[skills], property=detail
    ///   a_i.b_i.c           → parents=[a,b], property=c
    /// </summary>
    public static bool TryParseQualifiedArrayName(
        string arrayStartName,
        out IReadOnlyList<string> parentChain,
        out string propertyName)
    {
        parentChain = Array.Empty<string>();
        propertyName = "";
        var s = arrayStartName.Trim();
        if (string.IsNullOrEmpty(s))
            return false;

        // No parent qualification
        if (!s.Contains("_i.", StringComparison.Ordinal))
        {
            if (s.Contains("_i", StringComparison.Ordinal))
                return false; // e.g. trailing skills_i alone is invalid as array name
            propertyName = s;
            return true;
        }

        // Reuse loop ref parser: skills_i.detail → chain=[skills], field=detail
        if (!TryParseLoopReference(s, out var reference) || string.IsNullOrEmpty(reference.FieldPath))
            return false;

        // fieldPath must be a single identifier (the child array property), not a.b
        if (reference.FieldPath.Contains('.'))
            return false;

        parentChain = reference.Chain;
        propertyName = reference.FieldPath;
        return true;
    }

    public static bool IsLoopField(string inner)
        => TryParseLoopReference(inner, out _);
}
