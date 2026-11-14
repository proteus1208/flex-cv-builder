using System.Text.Json;
using System.Text.RegularExpressions;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Services;

public sealed class DocxTemplateEngine
{
    private readonly JsonPathResolver _paths = new();

    /// <param name="PropertyName">JSON array property (e.g. experience / detail).</param>
    /// <param name="Alias">Optional iterate name from ArrayStart; use ${alias.field} instead of ${prop_i.field}.</param>
    private sealed record LoopFrame(string PropertyName, string? Alias, JsonElement Item);

    public void Generate(string templatePath, string outputPath, JsonElement root, IssueCollector issues)
    {
        try
        {
            File.Copy(templatePath, outputPath, overwrite: true);
            using var doc = WordprocessingDocument.Open(outputPath, true);
            if (doc.MainDocumentPart?.Document?.Body == null)
            {
                issues.Error("InvalidDocx", "Docx", "Template has no document body.");
                return;
            }

            var stack = new Stack<LoopFrame>();
            ProcessContainer(doc.MainDocumentPart.Document.Body, doc.MainDocumentPart, root, stack, issues);

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                if (header.Header != null)
                    ProcessContainer(header.Header, header, root, new Stack<LoopFrame>(), issues);
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                if (footer.Footer != null)
                    ProcessContainer(footer.Footer, footer, root, new Stack<LoopFrame>(), issues);
            }

            doc.MainDocumentPart.Document.Save();
        }
        catch (Exception ex)
        {
            issues.Error("DocxFailed", "Docx", $"DOCX generation failed: {ex.Message}", detail: ex.ToString());
        }
    }

    private void ProcessContainer(
        OpenXmlElement container,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        // Block loops first so nested aliases (exp / skill) are on the stack, then
        // same-paragraph loops: "Technology Stack: ${ArrayStart ...}${item}, ${ArrayEnd ...}"
        ExpandLoops(container, part, root, loops, issues);
        ExpandInlineLoops(container, part, root, loops, issues);
        ReplacePlaceholders(container, part, root, loops, issues);
    }

    private void ExpandInlineLoops(
        OpenXmlElement container,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        foreach (var para in container.Descendants<Paragraph>().ToList())
            ExpandInlineLoopsInParagraph(para, part, root, loops, issues);
    }

    private void ExpandInlineLoopsInParagraph(
        Paragraph paragraph,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        var guard = 0;
        while (guard++ < 64 &&
               TryFindInlineLoop(
                   RunTextMerger.GetParagraphText(paragraph),
                   out var startMatch,
                   out var endMatch,
                   out var marker))
        {
            var text = RunTextMerger.GetParagraphText(paragraph);
            var innerStart = startMatch.Index + startMatch.Length;
            var innerLength = Math.Max(0, endMatch.Index - innerStart);
            var inner = text.Substring(innerStart, innerLength);
            var replaceStart = startMatch.Index;
            var replaceLength = endMatch.Index + endMatch.Length - startMatch.Index;

            if (!TryResolveArrayItems(
                    root,
                    loops,
                    marker.Source,
                    out var propertyName,
                    out var items,
                    out var coerced,
                    out var missing))
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Cannot resolve inline array '{marker.Source}' — removed.",
                    marker.Source);
                items = Array.Empty<JsonElement>();
                propertyName = marker.Source.Contains('.')
                    ? marker.Source[(marker.Source.LastIndexOf('.') + 1)..]
                    : marker.Source;
            }
            else if (missing)
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Inline array '{marker.Source}' missing — removed.",
                    marker.Source);
            }
            else if (items.Count == 0)
            {
                issues.Warning(
                    "EmptyLoopArray",
                    "Docx",
                    $"Inline array '{marker.Source}' is empty — removed.",
                    marker.Source);
            }
            else if (coerced)
            {
                issues.Info(
                    "LoopObjectCoerced",
                    "Docx",
                    $"Inline array '{marker.Source}' was an object; coerced to a single-item array.",
                    marker.Source);
            }

            // ${skill}${JoinStart}, ${JoinEnd} → join with ", " and no trailing separator
            var itemTemplate = inner;
            var separator = "";
            if (PlaceholderPatterns.TryExtractJoin(inner, out var body, out var sep))
            {
                itemTemplate = body;
                separator = sep;
            }

            var parts = new List<string>();
            foreach (var item in items)
            {
                var childLoops = CloneLoopStack(loops);
                childLoops.Push(new LoopFrame(propertyName, marker.Alias, item));

                var frag = new Paragraph(new Run(new Text(itemTemplate)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
                ExpandInlineLoopsInParagraph(frag, part, root, childLoops, issues);
                ReplacePlaceholdersInParagraph(frag, part, root, childLoops, issues);
                parts.Add(RunTextMerger.GetParagraphText(frag));
            }

            RunTextMerger.ReplaceRange(
                paragraph,
                replaceStart,
                replaceLength,
                string.Join(separator, parts));
        }
    }

    private static bool TryFindInlineLoop(
        string text,
        out Match startMatch,
        out Match endMatch,
        out ArrayMarker marker)
    {
        startMatch = null!;
        endMatch = null!;
        marker = null!;
        if (string.IsNullOrEmpty(text))
            return false;

        Match? bestStart = null;
        Match? bestEnd = null;
        var bestSpan = int.MaxValue;

        var starts = PlaceholderPatterns.ArrayStart().Matches(text).Cast<Match>().ToList();
        var ends = PlaceholderPatterns.ArrayEnd().Matches(text).Cast<Match>().ToList();
        if (starts.Count == 0 || ends.Count == 0)
            return false;

        foreach (var start in starts)
        {
            var source = start.Groups[1].Value;
            var depth = 0;
            Match? foundEnd = null;

            var events = starts.Concat(ends)
                .Where(m => m.Index > start.Index)
                .OrderBy(m => m.Index)
                .ThenBy(m => m.Value.Contains("ArrayEnd", StringComparison.OrdinalIgnoreCase) ? 1 : 0)
                .ToList();

            foreach (var ev in events)
            {
                var isStart = PlaceholderPatterns.ArrayStart().IsMatch(ev.Value);
                var evSource = isStart
                    ? PlaceholderPatterns.ArrayStart().Match(ev.Value).Groups[1].Value
                    : PlaceholderPatterns.ArrayEnd().Match(ev.Value).Groups[1].Value;
                if (!evSource.Equals(source, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (isStart)
                {
                    depth++;
                    continue;
                }

                if (depth == 0)
                {
                    foundEnd = ev;
                    break;
                }

                depth--;
            }

            if (foundEnd is null)
                continue;

            var span = foundEnd.Index + foundEnd.Length - start.Index;
            // Prefer innermost (smallest) complete pair.
            if (span < bestSpan)
            {
                bestSpan = span;
                bestStart = start;
                bestEnd = foundEnd;
            }
        }

        if (bestStart is null || bestEnd is null)
            return false;

        startMatch = bestStart;
        endMatch = bestEnd;
        marker = new ArrayMarker(
            bestStart.Groups[1].Value,
            bestStart.Groups[2].Success ? bestStart.Groups[2].Value : null);
        return true;
    }

    private void ExpandLoops(
        OpenXmlElement container,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        while (TryFindLoop(container, out var startPara, out var endPara, out var source))
        {
            var parent = startPara.Parent;
            if (parent is null || !ReferenceEquals(parent, endPara.Parent))
            {
                issues.Error(
                    "LoopParentMismatch",
                    "Docx",
                    $"ArrayStart/ArrayEnd for '{source}' must share the same parent.",
                    source);
                ClearMarker(startPara);
                continue;
            }

            PlaceholderPatterns.IsArrayStart(RunTextMerger.GetParagraphText(startPara), out ArrayMarker marker);
            var alias = marker.Alias;

            if (!TryResolveArrayItems(root, loops, source, out var propertyName, out var items, out var coerced, out var missing))
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Cannot resolve array '{source}' from the current loop context — section left empty.",
                    source);
                items = Array.Empty<JsonElement>();
                propertyName = source.Contains('.') ? source[(source.LastIndexOf('.') + 1)..] : source;
            }
            else if (missing)
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Array '{source}' missing — section left empty.",
                    source);
            }
            else if (items.Count == 0)
            {
                issues.Warning(
                    "EmptyLoopArray",
                    "Docx",
                    $"Array '{source}' is empty — section left empty.",
                    source);
            }
            else if (coerced)
            {
                issues.Info(
                    "LoopObjectCoerced",
                    "Docx",
                    $"Array '{source}' was an object; coerced to a single-item array.",
                    source);
            }

            var block = startPara.ElementsAfter().TakeWhile(e => !ReferenceEquals(e, endPara)).ToList();
            OpenXmlElement insertAfter = endPara;

            foreach (var item in items)
            {
                var temp = new Body();
                foreach (var el in block)
                    temp.AppendChild(el.CloneNode(true));

                var childLoops = CloneLoopStack(loops);
                childLoops.Push(new LoopFrame(propertyName, alias, item));
                ProcessContainer(temp, part, root, childLoops, issues);

                while (temp.ChildElements.Count > 0)
                {
                    var child = temp.ChildElements[0];
                    child.Remove();
                    insertAfter.InsertAfterSelf(child);
                    insertAfter = child;
                }
            }

            foreach (var el in block)
                el.Remove();
            startPara.Remove();
            endPara.Remove();
        }
    }

    private bool TryResolveArrayItems(
        JsonElement root,
        Stack<LoopFrame> loops,
        string source,
        out string propertyName,
        out IReadOnlyList<JsonElement> items,
        out bool coerced,
        out bool missing)
    {
        propertyName = source;
        items = Array.Empty<JsonElement>();
        coerced = false;
        missing = true;

        // Classic: skills / skills_i.detail
        if (PlaceholderPatterns.TryParseQualifiedArrayName(source, out var parentChain, out propertyName))
        {
            JsonElement owner;
            if (parentChain.Count == 0)
                owner = ResolveDataContext(root, loops);
            else if (!TryResolveFrameChain(loops, parentChain, out owner))
                return false;

            items = _paths.GetArrayItems(owner, propertyName, out coerced, out missing);
            return true;
        }

        // Alias parent: skill.detail
        if (PlaceholderPatterns.TryParseAliasQualifiedSource(source, out var parentAlias, out propertyName))
        {
            if (!TryFindAliasFrame(loops, parentAlias, out var owner))
                return false;

            items = _paths.GetArrayItems(owner, propertyName, out coerced, out missing);
            return true;
        }

        return false;
    }

    private static bool TryResolveFrameChain(
        Stack<LoopFrame> loops,
        IReadOnlyList<string> chain,
        out JsonElement item)
    {
        item = default;
        if (chain.Count == 0 || loops.Count < chain.Count)
            return false;

        var frames = loops.Reverse().ToList(); // outer → inner
        for (var start = frames.Count - chain.Count; start >= 0; start--)
        {
            var matched = true;
            for (var i = 0; i < chain.Count; i++)
            {
                if (!frames[start + i].PropertyName.Equals(chain[i], StringComparison.OrdinalIgnoreCase))
                {
                    matched = false;
                    break;
                }
            }

            if (!matched)
                continue;

            item = frames[start + chain.Count - 1].Item;
            return true;
        }

        return false;
    }

    private static bool TryFindAliasFrame(
        Stack<LoopFrame> loops,
        string alias,
        out JsonElement item)
    {
        item = default;
        foreach (var frame in loops)
        {
            if (frame.Alias != null &&
                frame.Alias.Equals(alias, StringComparison.OrdinalIgnoreCase))
            {
                item = frame.Item;
                return true;
            }
        }

        return false;
    }

    private static bool TryResolveAliasReference(
        Stack<LoopFrame> loops,
        string inner,
        out JsonElement item,
        out string? fieldPath)
    {
        item = default;
        fieldPath = null;
        var s = inner.Trim();
        if (string.IsNullOrEmpty(s))
            return false;

        // Stack enumerates top→bottom (innermost first). Prefer longest alias at each level.
        LoopFrame? best = null;
        foreach (var frame in loops)
        {
            if (frame.Alias is null)
                continue;

            var alias = frame.Alias;
            if (!(s.Equals(alias, StringComparison.OrdinalIgnoreCase) ||
                  s.StartsWith(alias + ".", StringComparison.OrdinalIgnoreCase)))
                continue;

            if (best is null || alias.Length > best.Alias!.Length)
                best = frame;
        }

        if (best?.Alias is null)
            return false;

        item = best.Item;
        if (s.Equals(best.Alias, StringComparison.OrdinalIgnoreCase))
        {
            fieldPath = null;
            return true;
        }

        fieldPath = s[(best.Alias.Length + 1)..];
        return true;
    }

    private static bool TryFindLoop(
        OpenXmlElement container,
        out Paragraph startPara,
        out Paragraph endPara,
        out string source)
    {
        startPara = null!;
        endPara = null!;
        source = "";

        foreach (var para in container.Descendants<Paragraph>())
        {
            var text = RunTextMerger.GetParagraphText(para);
            if (!PlaceholderPatterns.IsArrayStart(text, out ArrayMarker startMarker))
                continue;

            if (para.Parent is null)
                continue;

            var name = startMarker.Source;

            // Same-paragraph loops are handled by ExpandInlineLoops — skip them here.
            if (ParagraphHasInlineEndFor(text, name, startMarker))
                continue;

            var depth = 0;
            foreach (var sibling in para.ElementsAfter())
            {
                if (sibling is not Paragraph p)
                    continue;

                var t = RunTextMerger.GetParagraphText(p);
                if (PlaceholderPatterns.IsArrayStart(t, out ArrayMarker nested) &&
                    nested.Source.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    depth++;
                    continue;
                }

                if (PlaceholderPatterns.IsArrayEnd(t, out ArrayMarker endMarker) &&
                    endMarker.Source.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    if (depth == 0)
                    {
                        startPara = para;
                        endPara = p;
                        source = name;
                        return true;
                    }

                    depth--;
                }
            }
        }

        return false;
    }

    private static bool ParagraphHasInlineEndFor(string text, string source, ArrayMarker startMarker)
    {
        // Re-check using the first ArrayStart for this source in the paragraph.
        if (!TryFindInlineLoop(text, out var startMatch, out _, out var marker))
            return false;
        return marker.Source.Equals(source, StringComparison.OrdinalIgnoreCase) &&
               startMatch.Groups[1].Value.Equals(startMarker.Source, StringComparison.OrdinalIgnoreCase);
    }

    private void ReplacePlaceholders(
        OpenXmlElement container,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        foreach (var paragraph in container.Descendants<Paragraph>().ToList())
            ReplacePlaceholdersInParagraph(paragraph, part, root, loops, issues);
    }

    private void ReplacePlaceholdersInParagraph(
        Paragraph paragraph,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        var text = RunTextMerger.GetParagraphText(paragraph);
        if (!text.Contains("${", StringComparison.Ordinal))
            return;

        // Still has loop markers — leave for loop expanders; don't wipe sibling fields.
        if (PlaceholderPatterns.ArrayStart().IsMatch(text) || PlaceholderPatterns.ArrayEnd().IsMatch(text))
            return;

        RunTextMerger.ReplacePlaceholders(paragraph, part, inner =>
        {
            if (PlaceholderPatterns.TryParseLoopReference(inner, out var loopRef))
            {
                if (!TryResolveFrameChain(loops, loopRef.Chain, out var item))
                {
                    issues.Warning(
                        "OrphanLoopField",
                        "Docx",
                        $"Loop field ${{{inner}}} has no matching active ArrayStart chain — empty.",
                        inner);
                    return string.Empty;
                }

                return FormatItemValue(item, loopRef.FieldPath, inner, issues);
            }

            if (TryResolveAliasReference(loops, inner, out var aliasItem, out var aliasField))
                return FormatItemValue(aliasItem, aliasField, inner, issues);

            if (inner.StartsWith("ArrayStart", StringComparison.OrdinalIgnoreCase) ||
                inner.StartsWith("ArrayEnd", StringComparison.OrdinalIgnoreCase) ||
                PlaceholderPatterns.IsJoinKeyword(inner))
                return string.Empty;

            if (_paths.TryResolveValue(root, inner.Trim(), out var rootValue))
            {
                if (string.IsNullOrEmpty(rootValue))
                {
                    issues.Warning(
                        "EmptyOptionalField",
                        "Docx",
                        $"Field '{inner.Trim()}' is empty.",
                        inner.Trim());
                }

                return rootValue;
            }

            issues.Warning(
                "UnresolvedPlaceholder",
                "Docx",
                $"Unresolved placeholder ${{{inner.Trim()}}} — empty.",
                inner.Trim());
            return string.Empty;
        });
    }

    private string FormatItemValue(
        JsonElement item,
        string? fieldPath,
        string inner,
        IssueCollector issues)
    {
        if (string.IsNullOrEmpty(fieldPath))
        {
            return item.ValueKind switch
            {
                JsonValueKind.String => item.GetString() ?? "",
                JsonValueKind.Number => item.ToString(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                JsonValueKind.Null => "",
                _ => item.ToString()
            };
        }

        if (_paths.TryResolveValue(item, fieldPath, out var value))
            return value;

        issues.Warning(
            "UnresolvedPlaceholder",
            "Docx",
            $"Unresolved ${{{inner}}} — empty.",
            inner);
        return string.Empty;
    }

    private static JsonElement ResolveDataContext(JsonElement root, Stack<LoopFrame> loops)
        => loops.Count > 0 ? loops.Peek().Item : root;

    private static Stack<LoopFrame> CloneLoopStack(Stack<LoopFrame> source)
    {
        var clone = new Stack<LoopFrame>();
        foreach (var frame in source.Reverse())
            clone.Push(frame);
        return clone;
    }

    private static void ClearMarker(Paragraph paragraph)
    {
        var text = RunTextMerger.GetParagraphText(paragraph);
        if (text.Length > 0)
            RunTextMerger.ReplaceRange(paragraph, 0, text.Length, string.Empty);
    }
}
