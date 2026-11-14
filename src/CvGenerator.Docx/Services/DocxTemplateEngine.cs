using System.Text.Json;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Services;

public sealed class DocxTemplateEngine
{
    private readonly JsonPathResolver _paths = new();

    /// <param name="ArrayName">Local array property name (e.g. detail), used for stack matching.</param>
    private sealed record LoopFrame(string ArrayName, JsonElement Item);

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
        ExpandLoops(container, part, root, loops, issues);
        ReplacePlaceholders(container, part, root, loops, issues);
    }

    private void ExpandLoops(
        OpenXmlElement container,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        while (TryFindLoop(container, out var startPara, out var endPara, out var arrayToken))
        {
            var parent = startPara.Parent;
            if (parent is null || !ReferenceEquals(parent, endPara.Parent))
            {
                issues.Error(
                    "LoopParentMismatch",
                    "Docx",
                    $"ArrayStart/ArrayEnd for '{arrayToken}' must share the same parent.",
                    arrayToken);
                ClearMarker(startPara);
                continue;
            }

            if (!TryResolveArrayItems(root, loops, arrayToken, out var localName, out var items, out var coerced, out var missing))
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Cannot resolve array '{arrayToken}' from the current loop context — section left empty.",
                    arrayToken);
                items = Array.Empty<JsonElement>();
                localName = arrayToken;
            }
            else if (missing)
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Array '{arrayToken}' missing — section left empty.",
                    arrayToken);
            }
            else if (items.Count == 0)
            {
                issues.Warning(
                    "EmptyLoopArray",
                    "Docx",
                    $"Array '{arrayToken}' is empty — section left empty.",
                    arrayToken);
            }
            else if (coerced)
            {
                issues.Info(
                    "LoopObjectCoerced",
                    "Docx",
                    $"Array '{arrayToken}' was an object; coerced to a single-item array.",
                    arrayToken);
            }

            var block = startPara.ElementsAfter().TakeWhile(e => !ReferenceEquals(e, endPara)).ToList();
            OpenXmlElement insertAfter = endPara;

            foreach (var item in items)
            {
                var temp = new Body();
                foreach (var el in block)
                    temp.AppendChild(el.CloneNode(true));

                var childLoops = CloneLoopStack(loops);
                childLoops.Push(new LoopFrame(localName, item));
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
        string arrayToken,
        out string localName,
        out IReadOnlyList<JsonElement> items,
        out bool coerced,
        out bool missing)
    {
        localName = arrayToken;
        items = Array.Empty<JsonElement>();
        coerced = false;
        missing = true;

        if (!PlaceholderPatterns.TryParseQualifiedArrayName(arrayToken, out var parentChain, out var propertyName))
            return false;

        localName = propertyName;
        JsonElement owner;

        if (parentChain.Count == 0)
        {
            owner = ResolveDataContext(root, loops);
        }
        else if (!TryResolveFrameChain(loops, parentChain, out owner))
        {
            return false;
        }

        items = _paths.GetArrayItems(owner, propertyName, out coerced, out missing);
        return true;
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
        // Prefer the innermost contiguous match so `${experience_i.title}` still works
        // inside a nested `${ArrayStart highlights}` / `${ArrayStart skills_i.detail}` block.
        for (var start = frames.Count - chain.Count; start >= 0; start--)
        {
            var matched = true;
            for (var i = 0; i < chain.Count; i++)
            {
                if (!frames[start + i].ArrayName.Equals(chain[i], StringComparison.OrdinalIgnoreCase))
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

    private static bool TryFindLoop(
        OpenXmlElement container,
        out Paragraph startPara,
        out Paragraph endPara,
        out string arrayName)
    {
        startPara = null!;
        endPara = null!;
        arrayName = "";

        foreach (var para in container.Descendants<Paragraph>())
        {
            var text = RunTextMerger.GetParagraphText(para);
            if (!PlaceholderPatterns.IsArrayStart(text, out var name))
                continue;

            if (para.Parent is null)
                continue;

            var depth = 0;
            foreach (var sibling in para.ElementsAfter())
            {
                if (sibling is not Paragraph p)
                    continue;

                var t = RunTextMerger.GetParagraphText(p);
                if (PlaceholderPatterns.IsArrayStart(t, out var nested) &&
                    nested.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    depth++;
                    continue;
                }

                if (PlaceholderPatterns.IsArrayEnd(t, out var endName) &&
                    endName.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    if (depth == 0)
                    {
                        startPara = para;
                        endPara = p;
                        arrayName = name;
                        return true;
                    }

                    depth--;
                }
            }
        }

        return false;
    }

    private void ReplacePlaceholders(
        OpenXmlElement container,
        OpenXmlPart part,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        foreach (var paragraph in container.Descendants<Paragraph>().ToList())
        {
            var text = RunTextMerger.GetParagraphText(paragraph);
            if (!text.Contains("${", StringComparison.Ordinal))
                continue;

            if (PlaceholderPatterns.IsArrayStart(text, out _) || PlaceholderPatterns.IsArrayEnd(text, out _))
                continue;

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

                    if (string.IsNullOrEmpty(loopRef.FieldPath))
                    {
                        // Scalar array item (string/number/bool)
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

                    if (_paths.TryResolveValue(item, loopRef.FieldPath, out var value))
                        return value;

                    issues.Warning(
                        "UnresolvedPlaceholder",
                        "Docx",
                        $"Unresolved ${{{inner}}} — empty.",
                        inner);
                    return string.Empty;
                }

                if (inner.StartsWith("ArrayStart", StringComparison.OrdinalIgnoreCase) ||
                    inner.StartsWith("ArrayEnd", StringComparison.OrdinalIgnoreCase))
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
