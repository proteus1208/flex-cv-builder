using CvGenerator.Docx.Internal;

namespace CvGenerator.Docx.Services;

/// <summary>
/// Resolves iterate aliases with proper nesting scope so the same alias
/// (e.g. skill) can be reused in different loops without colliding.
/// </summary>
public static class AliasPathNormalizer
{
    private sealed record ScopeFrame(string Source, string? Alias, string ItemPrefix);

    /// <summary>
    /// Walk placeholders in document order, assign ValidationPath using an alias stack.
    /// </summary>
    public static IReadOnlyList<PlaceholderInfo> AssignScopedPaths(IReadOnlyList<PlaceholderInfo> ordered)
    {
        var stack = new List<ScopeFrame>();
        var results = new List<PlaceholderInfo>(ordered.Count);

        foreach (var p in ordered)
        {
            if (p.Kind is "array-start" or "loop-start")
            {
                var map = AliasMapFromStack(stack);
                var itemPrefix = ToItemPrefix(p.JsonPath ?? "", map);
                var arrayPath = ToArraySource(p.JsonPath ?? "", map) ?? (p.JsonPath ?? "");
                if (itemPrefix != null)
                    stack.Add(new ScopeFrame(p.JsonPath ?? "", p.Alias, itemPrefix));

                results.Add(Clone(p, p.Kind, p.JsonPath, arrayPath));
                continue;
            }

            if (p.Kind == "array-end")
            {
                PopSource(stack, p.JsonPath);
                var map = AliasMapFromStack(stack);
                var arrayPath = ToArraySource(p.JsonPath ?? "", map) ?? (p.JsonPath ?? "");
                results.Add(Clone(p, p.Kind, p.JsonPath, arrayPath));
                continue;
            }

            var aliases = AliasMapFromStack(stack);
            var kind = p.Kind;
            var displayPath = p.JsonPath;
            var validationPath = p.JsonPath ?? "";

            if (kind == "loop-field" || kind == "simple")
            {
                var inner = (p.JsonPath ?? "").Trim();
                if (kind == "simple" &&
                    !string.IsNullOrEmpty(p.Token) &&
                    p.Token.StartsWith("${", StringComparison.Ordinal) &&
                    p.Token.EndsWith('}'))
                {
                    inner = p.Token[2..^1].Trim();
                }

                if (!inner.Contains("_i", StringComparison.Ordinal))
                {
                    var rewritten = RewriteAliasField(inner, aliases);
                    if (rewritten != null)
                    {
                        kind = "loop-field";
                        validationPath = rewritten;
                        displayPath = rewritten;
                    }
                }
                else
                {
                    kind = "loop-field";
                    validationPath = inner;
                    displayPath = inner;
                }
            }

            results.Add(Clone(p, kind, displayPath, validationPath));
        }

        return results;
    }

    /// <summary>Dedupe by token, keeping first scoped occurrence (correct for that context).</summary>
    public static IReadOnlyList<PlaceholderInfo> DedupeByToken(IEnumerable<PlaceholderInfo> scoped)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var list = new List<PlaceholderInfo>();
        foreach (var p in scoped)
        {
            if (!seen.Add(p.Token))
                continue;
            list.Add(p);
        }

        return list;
    }

    private static PlaceholderInfo Clone(
        PlaceholderInfo p,
        string kind,
        string? jsonPath,
        string? validationPath)
        => new()
        {
            Token = p.Token,
            Kind = kind,
            JsonPath = jsonPath,
            Alias = p.Alias,
            ValidationPath = validationPath
        };

    private static Dictionary<string, string> AliasMapFromStack(List<ScopeFrame> stack)
    {
        // Outer → inner; inner reuses of the same alias win.
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var frame in stack)
        {
            if (!string.IsNullOrEmpty(frame.Alias))
                map[frame.Alias] = frame.ItemPrefix;
        }

        return map;
    }

    private static void PopSource(List<ScopeFrame> stack, string? source)
    {
        if (string.IsNullOrEmpty(source) || stack.Count == 0)
            return;

        for (var i = stack.Count - 1; i >= 0; i--)
        {
            var match = stack[i].Source.Equals(source, StringComparison.OrdinalIgnoreCase);
            stack.RemoveAt(i);
            if (match)
                return;
        }
    }

    private static string? ToItemPrefix(string source, IReadOnlyDictionary<string, string> aliasMap)
    {
        if (PlaceholderPatterns.TryParseQualifiedArrayName(source, out var parents, out var prop))
        {
            if (parents.Count == 0)
                return prop + "_i";
            return string.Join(".", parents.Select(p => p + "_i")) + "." + prop + "_i";
        }

        // exp.skills.primary → experiences_i.skills.primary_i
        if (PlaceholderPatterns.TryParseAliasQualifiedSource(source, out var parentAlias, out var relative) &&
            aliasMap.TryGetValue(parentAlias, out var parentPrefix))
        {
            return parentPrefix + "." + relative + "_i";
        }

        // Absolute root path: extra_infos / experience.company_1.industries → …_i
        if (PlaceholderPatterns.TryParseRootArrayPath(source, out var arrayPath, out _))
            return arrayPath + "_i";

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

        if (PlaceholderPatterns.TryParseAliasQualifiedSource(source, out var parentAlias, out var relative) &&
            aliasMap.TryGetValue(parentAlias, out var parentPrefix))
        {
            return parentPrefix + "." + relative;
        }

        if (PlaceholderPatterns.TryParseRootArrayPath(source, out var arrayPath, out _))
            return arrayPath;

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
