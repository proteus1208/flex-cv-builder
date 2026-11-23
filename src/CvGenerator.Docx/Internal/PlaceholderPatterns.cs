using System.Text.RegularExpressions;

namespace CvGenerator.Docx.Internal;

internal sealed record LoopReference(IReadOnlyList<string> Chain, string? FieldPath);

/// <summary>
/// ${ArrayStart source [alias]} / ${ArrayEnd source [alias]}
/// source: experience | skills_i.detail | skill.detail
/// alias: optional iterate name used instead of source_i
/// </summary>
internal sealed record ArrayMarker(string Source, string? Alias);

internal static partial class PlaceholderPatterns
{
    [GeneratedRegex(
        @"\$\{\s*ArrayStart\s+([A-Za-z_][A-Za-z0-9_.]*)(?:\s+([A-Za-z_][A-Za-z0-9_]*))?\s*\}",
        RegexOptions.IgnoreCase)]
    public static partial Regex ArrayStart();

    [GeneratedRegex(
        @"\$\{\s*ArrayEnd\s+([A-Za-z_][A-Za-z0-9_.]*)(?:\s+([A-Za-z_][A-Za-z0-9_]*))?\s*\}",
        RegexOptions.IgnoreCase)]
    public static partial Regex ArrayEnd();

    // ${JoinStart}separator${JoinEnd} — separator only between array items (no trailing)
    [GeneratedRegex(@"\$\{\s*JoinStart\s*\}", RegexOptions.IgnoreCase)]
    public static partial Regex JoinStart();

    [GeneratedRegex(@"\$\{\s*JoinEnd\s*\}", RegexOptions.IgnoreCase)]
    public static partial Regex JoinEnd();

    [GeneratedRegex(@"\$\{([^{}]+)\}")]
    public static partial Regex AnyPlaceholder();

    // No ^ — Match(input, startAt) still anchors ^ at index 0, so nested
    // segments like detail_i in skills_i.detail_i.name must be unanchored.
    [GeneratedRegex(@"([A-Za-z_][A-Za-z0-9_]*)_i")]
    private static partial Regex LoopSegment();

    public static bool IsArrayStart(string text, out ArrayMarker marker)
    {
        marker = null!;
        var m = ArrayStart().Match(text);
        if (!m.Success)
            return false;
        marker = new ArrayMarker(
            m.Groups[1].Value,
            m.Groups[2].Success ? m.Groups[2].Value : null);
        return true;
    }

    public static bool IsArrayEnd(string text, out ArrayMarker marker)
    {
        marker = null!;
        var m = ArrayEnd().Match(text);
        if (!m.Success)
            return false;
        marker = new ArrayMarker(
            m.Groups[1].Value,
            m.Groups[2].Success ? m.Groups[2].Value : null);
        return true;
    }

    /// <summary>
    /// Parses classic loop field refs:
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
    /// Classic ArrayStart/ArrayEnd source with _i parents:
    ///   skills              → parents=[], property=skills
    ///   skills_i.detail     → parents=[skills], property=detail
    ///   a_i.b_i.c           → parents=[a,b], property=c
    /// Returns false for alias-qualified sources like skill.detail (no _i).
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

        // Plain identifier (any root array name, e.g. extra_infos — not a false "_i" hit).
        if (!s.Contains('.', StringComparison.Ordinal) && !HasLoopMarker(s))
        {
            propertyName = s;
            return true;
        }

        // Classic parent qualification requires a real loop marker segment (foo_i.bar).
        if (!HasLoopMarker(s) || !s.Contains('.', StringComparison.Ordinal))
            return false;

        if (!TryParseLoopReference(s, out var reference) || string.IsNullOrEmpty(reference.FieldPath))
            return false;

        if (reference.FieldPath.Contains('.'))
            return false;

        parentChain = reference.Chain;
        propertyName = reference.FieldPath;
        return true;
    }

    /// <summary>
    /// Alias-qualified nested source under an active iterate alias:
    ///   skill.detail              → alias=skill, relative=detail
    ///   exp.skills.primary        → alias=exp, relative=skills.primary
    /// </summary>
    public static bool TryParseAliasQualifiedSource(
        string source,
        out string parentAlias,
        out string relativePath)
    {
        parentAlias = "";
        relativePath = "";
        var s = source.Trim();
        var dot = s.IndexOf('.');
        if (dot <= 0 || dot >= s.Length - 1)
            return false;
        if (HasLoopMarker(s))
            return false;

        parentAlias = s[..dot];
        relativePath = s[(dot + 1)..];
        if (!IsIdentifier(parentAlias))
            return false;

        var parts = relativePath.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || parts.Any(p => !IsIdentifier(p)))
            return false;

        relativePath = string.Join(".", parts);
        return true;
    }

    /// <summary>
    /// Absolute dotted JSON path from the document root (no loop <c>_i</c> markers), e.g.
    /// <c>experience.company_1.skills.primary</c>, <c>industries</c>, or <c>extra_infos</c>.
    /// </summary>
    public static bool TryParseRootArrayPath(string source, out string arrayPath, out string propertyName)
    {
        arrayPath = "";
        propertyName = "";
        var s = source.Trim();
        if (string.IsNullOrEmpty(s) || HasLoopMarker(s))
            return false;

        var parts = s.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0 || parts.Any(p => !IsIdentifier(p)))
            return false;

        arrayPath = string.Join(".", parts);
        propertyName = parts[^1];
        return true;
    }

    /// <summary>
    /// True when <paramref name="source"/> is a supported ArrayStart/ArrayEnd path:
    /// any root array name, <c>parent_i.child</c>, <c>alias.child[.nested]</c>, or absolute dotted root path.
    /// </summary>
    public static bool IsValidArraySource(string source)
        => TryParseQualifiedArrayName(source, out _, out _)
           || TryParseAliasQualifiedSource(source, out _, out _)
           || TryParseRootArrayPath(source, out _, out _);

    public static bool IsIdentifier(string value)
        => value.Length > 0
           && (char.IsLetter(value[0]) || value[0] == '_')
           && value.All(c => char.IsLetterOrDigit(c) || c == '_');

    /// <summary>
    /// True when <paramref name="source"/> contains a real loop marker segment
    /// (<c>foo_i</c> or <c>foo_i.bar</c>), not a coincidental substring like <c>extra_infos</c>.
    /// </summary>
    public static bool HasLoopMarker(string source)
    {
        var s = source.AsSpan();
        for (var i = 0; i < s.Length - 1; i++)
        {
            if (s[i] != '_' || s[i + 1] != 'i')
                continue;

            var after = i + 2;
            if (after < s.Length && s[after] != '.')
                continue; // e.g. "_in" inside extra_infos

            // Require an identifier character before "_i"
            if (i == 0)
                continue;
            var prev = s[i - 1];
            if (char.IsLetterOrDigit(prev) || prev == '_')
                return true;
        }

        return false;
    }

    public static bool IsJoinStart(string text) => JoinStart().IsMatch(text);
    public static bool IsJoinEnd(string text) => JoinEnd().IsMatch(text);

    public static bool IsJoinKeyword(string inner)
    {
        var t = inner.Trim();
        return t.Equals("JoinStart", StringComparison.OrdinalIgnoreCase) ||
               t.Equals("JoinEnd", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// ${item}${JoinStart}, ${JoinEnd} → body=${item}, separator=", "
    /// </summary>
    public static bool TryExtractJoin(
        string loopBody,
        out string itemTemplate,
        out string separator)
    {
        itemTemplate = loopBody;
        separator = "";
        var start = JoinStart().Match(loopBody);
        if (!start.Success)
            return false;

        var end = JoinEnd().Match(loopBody, start.Index + start.Length);
        if (!end.Success)
            return false;

        separator = loopBody[(start.Index + start.Length)..end.Index];
        itemTemplate = loopBody[..start.Index] + loopBody[(end.Index + end.Length)..];
        return true;
    }

    public static bool IsLoopField(string inner)
        => TryParseLoopReference(inner, out _);
}
