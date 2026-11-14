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
            ProcessContainer(doc.MainDocumentPart.Document.Body, root, stack, issues);

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                if (header.Header != null)
                    ProcessContainer(header.Header, root, new Stack<LoopFrame>(), issues);
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                if (footer.Footer != null)
                    ProcessContainer(footer.Footer, root, new Stack<LoopFrame>(), issues);
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
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        ExpandLoops(container, root, loops, issues);
        ReplacePlaceholders(container, root, loops, issues);
    }

    private void ExpandLoops(
        OpenXmlElement container,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        // Expand one loop at a time (innermost-capable via recursive ProcessContainer on clones).
        while (TryFindLoop(container, out var startPara, out var endPara, out var arrayName))
        {
            var parent = startPara.Parent;
            if (parent is null || !ReferenceEquals(parent, endPara.Parent))
            {
                issues.Error(
                    "LoopParentMismatch",
                    "Docx",
                    $"ArrayStart/ArrayEnd for '{arrayName}' must share the same parent (cannot span mismatched containers).",
                    arrayName);
                    // Remove start marker text to avoid infinite loop.
                ClearMarker(startPara);
                continue;
            }

            var block = startPara.ElementsAfter().TakeWhile(e => !ReferenceEquals(e, endPara)).ToList();
            var data = ResolveDataContext(root, loops);
            var items = _paths.GetArrayItems(data, arrayName, out var coerced, out var missing);

            if (missing)
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Array '{arrayName}' missing — section left empty.",
                    arrayName);
            }
            else if (items.Count == 0)
            {
                issues.Warning(
                    "EmptyLoopArray",
                    "Docx",
                    $"Array '{arrayName}' is empty — section left empty.",
                    arrayName);
            }
            else if (coerced)
            {
                issues.Info(
                    "LoopObjectCoerced",
                    "Docx",
                    $"Array '{arrayName}' was an object; coerced to a single-item array.",
                    arrayName);
            }

            OpenXmlElement insertAfter = endPara;
            foreach (var item in items)
            {
                // Isolate clones in a temporary body so nested ArrayStart/ArrayEnd + tables/images
                // can be processed with this item as the loop context.
                var temp = new Body();
                foreach (var el in block)
                    temp.AppendChild(el.CloneNode(true));

                var childLoops = CloneLoopStack(loops);
                childLoops.Push(new LoopFrame(arrayName, item));
                ProcessContainer(temp, root, childLoops, issues);

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

    private static bool TryFindLoop(
        OpenXmlElement container,
        out Paragraph startPara,
        out Paragraph endPara,
        out string arrayName)
    {
        startPara = null!;
        endPara = null!;
        arrayName = "";

        // Search paragraphs under this container; start/end must share a parent.
        foreach (var para in container.Descendants<Paragraph>())
        {
            var text = RunTextMerger.GetParagraphText(para);
            if (!PlaceholderPatterns.IsArrayStart(text, out var name))
                continue;

            var parent = para.Parent;
            if (parent is null)
                continue;

            // Match ArrayEnd with nesting depth for the same array name among siblings.
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

            // Unclosed — still return start with no end? Let caller handle via separate scan.
            // Skip and continue searching other starts.
        }

        return false;
    }

    private void ReplacePlaceholders(
        OpenXmlElement container,
        JsonElement root,
        Stack<LoopFrame> loops,
        IssueCollector issues)
    {
        foreach (var paragraph in container.Descendants<Paragraph>().ToList())
        {
            var text = RunTextMerger.GetParagraphText(paragraph);
            if (!text.Contains("${", StringComparison.Ordinal))
                continue;

            // Skip pure marker paragraphs still present (shouldn't after expand).
            if (PlaceholderPatterns.IsArrayStart(text, out _) || PlaceholderPatterns.IsArrayEnd(text, out _))
                continue;

            RunTextMerger.ReplacePlaceholders(paragraph, inner =>
            {
                if (PlaceholderPatterns.IsLoopField(inner, out var arrayName, out var fieldPath))
                {
                    var frame = loops.FirstOrDefault(f =>
                        f.ArrayName.Equals(arrayName, StringComparison.OrdinalIgnoreCase));
                    if (frame is null)
                    {
                        issues.Warning(
                            "OrphanLoopField",
                            "Docx",
                            $"Loop field ${{{inner}}} has no active ArrayStart '{arrayName}' — empty.",
                            inner);
                        return string.Empty;
                    }

                    if (_paths.TryResolveValue(frame.Item, fieldPath, out var value))
                        return value;

                    issues.Warning(
                        "UnresolvedPlaceholder",
                        "Docx",
                        $"Unresolved ${{{inner}}} — empty.",
                        inner);
                    return string.Empty;
                }

                // Ignore leftover ArrayStart/ArrayEnd inners if any.
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

    private JsonElement ResolveDataContext(JsonElement root, Stack<LoopFrame> loops)
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
