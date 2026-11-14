using System.Text.Json;
using System.Text.RegularExpressions;
using CvGenerator.Core.Services;
using CvGenerator.Docx.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CvGenerator.Docx.Services;

public sealed partial class DocxTemplateEngine
{
    private readonly JsonPathResolver _paths = new();

    [GeneratedRegex(@"\$\{\s*start\s+([A-Za-z0-9_.]+)\s*\}", RegexOptions.IgnoreCase)]
    private static partial Regex LoopStartRegex();

    [GeneratedRegex(@"\$\{\s*end\s+([A-Za-z0-9_.]+)\s*\}", RegexOptions.IgnoreCase)]
    private static partial Regex LoopEndRegex();

    [GeneratedRegex(@"\$\{([^{}]+)\}")]
    private static partial Regex PlaceholderRegex();

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

            ProcessContainer(doc.MainDocumentPart.Document.Body, root, issues);

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                if (header.Header != null)
                    ProcessContainer(header.Header, root, issues);
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                if (footer.Footer != null)
                    ProcessContainer(footer.Footer, root, issues);
            }

            doc.MainDocumentPart.Document.Save();
        }
        catch (Exception ex)
        {
            issues.Error("DocxFailed", "Docx", $"DOCX generation failed: {ex.Message}", detail: ex.ToString());
        }
    }

    private void ProcessContainer(OpenXmlElement container, JsonElement root, IssueCollector issues)
    {
        ExpandLoops(container, root, issues);
        ReplaceSimplePlaceholders(container, root, issues);
    }

    private void ExpandLoops(OpenXmlElement container, JsonElement root, IssueCollector issues)
    {
        // Process from outermost by repeatedly finding start markers until none remain.
        while (true)
        {
            var paragraphs = container.Descendants<Paragraph>().ToList();
            int? startIndex = null;
            string? arrayName = null;

            for (var i = 0; i < paragraphs.Count; i++)
            {
                var text = RunTextMerger.GetParagraphText(paragraphs[i]);
                var startMatch = LoopStartRegex().Match(text);
                if (startMatch.Success)
                {
                    startIndex = i;
                    arrayName = startMatch.Groups[1].Value;
                    break;
                }
            }

            if (startIndex is null || arrayName is null)
                break;

            var endIndex = -1;
            for (var i = startIndex.Value + 1; i < paragraphs.Count; i++)
            {
                var text = RunTextMerger.GetParagraphText(paragraphs[i]);
                var endMatch = LoopEndRegex().Match(text);
                if (endMatch.Success &&
                    endMatch.Groups[1].Value.Equals(arrayName, StringComparison.OrdinalIgnoreCase))
                {
                    endIndex = i;
                    break;
                }
            }

            if (endIndex < 0)
            {
                issues.Error(
                    "UnclosedLoop",
                    "Docx",
                    $"Loop start for '{arrayName}' has no matching end marker.",
                    $"${{ start {arrayName} }}");
                // Remove the broken start marker so we don't infinite-loop.
                var broken = paragraphs[startIndex.Value];
                var brokenText = LoopStartRegex().Replace(RunTextMerger.GetParagraphText(broken), string.Empty);
                RunTextMerger.SetParagraphText(broken, brokenText);
                continue;
            }

            var startPara = paragraphs[startIndex.Value];
            var endPara = paragraphs[endIndex];
            var parent = startPara.Parent;
            if (parent == null)
            {
                issues.Error("LoopParentMissing", "Docx", $"Cannot expand loop '{arrayName}' — missing parent.");
                break;
            }

            // Collect block elements between start and end (exclusive), same parent.
            var block = new List<OpenXmlElement>();
            var between = startPara.ElementsAfter().TakeWhile(e => !ReferenceEquals(e, endPara)).ToList();
            foreach (var el in between)
                block.Add(el);

            var items = _paths.GetArrayItems(root, arrayName, out var coerced, out var missing);
            if (missing)
            {
                issues.Warning(
                    "MissingLoopArray",
                    "Docx",
                    $"Loop array '{arrayName}' missing — section left empty.",
                    arrayName);
            }
            else if (items.Count == 0)
            {
                issues.Warning(
                    "EmptyLoopArray",
                    "Docx",
                    $"Loop array '{arrayName}' is empty — section left empty.",
                    arrayName);
            }
            else if (coerced)
            {
                issues.Info(
                    "LoopObjectCoerced",
                    "Docx",
                    $"Loop '{arrayName}' was an object; coerced to a single-item array.",
                    arrayName);
            }

            OpenXmlElement insertAfter = endPara;
            foreach (var item in items)
            {
                foreach (var templateEl in block)
                {
                    var clone = templateEl.CloneNode(true);
                    if (clone is Paragraph p)
                        ApplyLoopItemToParagraph(p, arrayName, item, issues);
                    else
                    {
                        foreach (var para in clone.Descendants<Paragraph>())
                            ApplyLoopItemToParagraph(para, arrayName, item, issues);
                    }

                    insertAfter.InsertAfterSelf(clone);
                    insertAfter = clone;
                }
            }

            // Remove start, block originals, and end.
            foreach (var el in block)
                el.Remove();
            startPara.Remove();
            endPara.Remove();
        }
    }

    private void ApplyLoopItemToParagraph(
        Paragraph paragraph,
        string arrayName,
        JsonElement item,
        IssueCollector issues)
    {
        var text = RunTextMerger.GetParagraphText(paragraph);
        if (!text.Contains("${", StringComparison.Ordinal))
            return;

        var replaced = PlaceholderRegex().Replace(text, match =>
        {
            var inner = match.Groups[1].Value.Trim();
            var prefix = arrayName + ".i.";
            if (inner.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                var field = inner[prefix.Length..];
                if (_paths.TryResolveValue(item, field, out var value))
                    return value;

                issues.Warning(
                    "UnresolvedPlaceholder",
                    "Docx",
                    $"Unresolved loop field ${{{inner}}} — replaced with empty.",
                    inner);
                return string.Empty;
            }

            // Leave non-loop placeholders for the simple pass (if any remain in loop body — unusual).
            return match.Value;
        });

        if (!string.Equals(text, replaced, StringComparison.Ordinal))
            RunTextMerger.SetParagraphText(paragraph, replaced);
    }

    private void ReplaceSimplePlaceholders(OpenXmlElement container, JsonElement root, IssueCollector issues)
    {
        foreach (var paragraph in container.Descendants<Paragraph>().ToList())
        {
            var text = RunTextMerger.GetParagraphText(paragraph);
            if (!text.Contains("${", StringComparison.Ordinal))
                continue;

            // Skip leftover loop markers (should already be gone).
            if (LoopStartRegex().IsMatch(text) || LoopEndRegex().IsMatch(text))
                continue;

            var replaced = PlaceholderRegex().Replace(text, match =>
            {
                var inner = match.Groups[1].Value.Trim();
                if (inner.Contains(".i.", StringComparison.Ordinal))
                {
                    issues.Warning(
                        "OrphanLoopField",
                        "Docx",
                        $"Loop field ${{{inner}}} found outside a loop — replaced with empty.",
                        inner);
                    return string.Empty;
                }

                if (_paths.TryResolveValue(root, inner, out var value))
                {
                    if (string.IsNullOrEmpty(value))
                    {
                        issues.Warning(
                            "EmptyOptionalField",
                            "Docx",
                            $"Field '{inner}' is empty.",
                            inner);
                    }

                    return value;
                }

                issues.Warning(
                    "UnresolvedPlaceholder",
                    "Docx",
                    $"Unresolved placeholder ${{{inner}}} — replaced with empty.",
                    inner);
                return string.Empty;
            });

            if (!string.Equals(text, replaced, StringComparison.Ordinal))
                RunTextMerger.SetParagraphText(paragraph, replaced);

            // Optional: remove paragraph if it became empty after only containing a placeholder.
            if (string.IsNullOrWhiteSpace(replaced))
            {
                // Keep empty paragraphs for spacing fidelity unless they were purely a marker.
            }
        }
    }
}
