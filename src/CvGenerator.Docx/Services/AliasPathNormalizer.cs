using CvGenerator.Docx.Internal;

namespace CvGenerator.Docx.Services;

/// <summary>
/// Maps iterate aliases (${ArrayStart experience exp} → exp)
/// onto classic _i paths for live JSON validation / bindings.
/// </summary>
public static class AliasPathNormalizer
{
    public static IReadOnlyDictionary<string, string> BuildAliasMap(IEnumerable<PlaceholderInfo> placeholders)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Repeat until nested alias parents resolve (document order may vary).
        for (var pass = 0; pass < 8; pass++)
        {
            var added = false;
            foreach (var p in placeholders)
            {
                if (p.Kind != "array-start" || string.IsNullOrEmpty(p.Alias) || string.IsNullOrEmpty(p.JsonPath))
                    continue;
                if (map.ContainsKey(p.Alias))
                    continue;

                var canonical = ToItemPrefix(p.JsonPath, map);
                if (canonical is null)
                    continue;

                map[p.Alias] = canonical;
                added = true;
            }

            if (!added)
                break;
        }

        return map;
    }

    /// <summary>
    /// Returns a path LiveJsonValidator understands (classic _i form), or the original path.
    /// </summary>
    public static string ToValidationPath(
        string kind,
        string? jsonPath,
        string? token,
        IReadOnlyDictionary<string, string> aliasMap)
    {
        if (string.IsNullOrWhiteSpace(jsonPath))
            return jsonPath ?? "";

        var path = jsonPath.Trim();

        if (kind is "array-start" or "array-end" or "loop-start")
            return ToArraySource(path, aliasMap) ?? path;

        if (kind == "loop-field")
        {
            // Already classic, or still alias form
            if (path.Contains("_i", StringComparison.Ordinal))
                return path;
            return RewriteAliasField(path, aliasMap) ?? path;
        }

        // simple — may actually be an alias field that was misclassified
        var rewritten = RewriteAliasField(path, aliasMap);
        if (rewritten != null)
            return rewritten;

        // Also try rewriting from token inner text
        if (!string.IsNullOrEmpty(token) &&
            token.StartsWith("${", StringComparison.Ordinal) &&
            token.EndsWith('}'))
        {
            var inner = token[2..^1].Trim();
            rewritten = RewriteAliasField(inner, aliasMap);
            if (rewritten != null)
                return rewritten;
        }

        return path;
    }

    public static string EffectiveKind(
        string kind,
        string? jsonPath,
        string? token,
        IReadOnlyDictionary<string, string> aliasMap)
    {
        if (kind is "array-start" or "array-end" or "loop-start" or "loop-field")
            return kind;

        if (kind != "simple" || string.IsNullOrWhiteSpace(jsonPath))
            return kind;

        var path = jsonPath.Trim();
        if (RewriteAliasField(path, aliasMap) != null)
            return "loop-field";

        if (!string.IsNullOrEmpty(token) &&
            token.StartsWith("${", StringComparison.Ordinal) &&
            token.EndsWith('}'))
        {
            var inner = token[2..^1].Trim();
            if (RewriteAliasField(inner, aliasMap) != null)
                return "loop-field";
        }

        return kind;
    }

    private static string? ToItemPrefix(string source, IReadOnlyDictionary<string, string> aliasMap)
    {
        if (PlaceholderPatterns.TryParseQualifiedArrayName(source, out var parents, out var prop))
        {
            if (parents.Count == 0)
                return prop + "_i";
            return string.Join(".", parents.Select(p => p + "_i")) + "." + prop + "_i";
        }

        if (PlaceholderPatterns.TryParseAliasQualifiedSource(source, out var parentAlias, out var child) &&
            aliasMap.TryGetValue(parentAlias, out var parentPrefix))
        {
            return parentPrefix + "." + child + "_i";
        }

        return null;
    }

    private static string? ToArraySource(string source, IReadOnlyDictionary<string, string> aliasMap)
    {
        if (PlaceholderPatterns.TryParseQualifiedArrayName(source, out var parents, out var prop))
        {
            if (parents.Count == 0)
                return prop;
            return string.Join(".", parents.Select(p => p + "_i")) + "." + prop;
        }

        if (PlaceholderPatterns.TryParseAliasQualifiedSource(source, out var parentAlias, out var child) &&
            aliasMap.TryGetValue(parentAlias, out var parentPrefix))
        {
            return parentPrefix + "." + child;
        }

        return null;
    }

    private static string? RewriteAliasField(string inner, IReadOnlyDictionary<string, string> aliasMap)
    {
        var s = inner.Trim();
        string? bestAlias = null;
        foreach (var alias in aliasMap.Keys)
        {
            if (s.Equals(alias, StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith(alias + ".", StringComparison.OrdinalIgnoreCase))
            {
                if (bestAlias is null || alias.Length > bestAlias.Length)
                    bestAlias = alias;
            }
        }

        if (bestAlias is null)
            return null;

        var prefix = aliasMap[bestAlias];
        if (s.Equals(bestAlias, StringComparison.OrdinalIgnoreCase))
            return prefix;

        return prefix + "." + s[(bestAlias.Length + 1)..];
    }
}
