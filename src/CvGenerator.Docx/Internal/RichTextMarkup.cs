using System.Text.RegularExpressions;

namespace CvGenerator.Docx.Internal;

internal enum RichSegmentKind
{
    Text,
    Bold,
    Italic,
    Underline,
    Link
}

internal sealed record RichSegment(RichSegmentKind Kind, string Text, string? Url = null);

/// <summary>
/// Lightweight markup inside JSON string values:
///   **bold**  *italic*  __underline__
///   [label](https://example.com)
/// </summary>
internal static partial class RichTextMarkup
{
    // Longest delimiters first so ** wins over * and __ is unambiguous.
    [GeneratedRegex(@"\*\*(.+?)\*\*|__(.+?)__|\*(.+?)\*|\[([^\]]+)\]\(([^)]+)\)")]
    private static partial Regex TokenRegex();

    public static bool ContainsMarkup(string? value)
        => !string.IsNullOrEmpty(value) && TokenRegex().IsMatch(value);

    public static IReadOnlyList<RichSegment> Parse(string value)
    {
        if (string.IsNullOrEmpty(value))
            return Array.Empty<RichSegment>();

        var segments = new List<RichSegment>();
        var index = 0;
        foreach (Match match in TokenRegex().Matches(value))
        {
            if (match.Index > index)
                segments.Add(new RichSegment(RichSegmentKind.Text, value[index..match.Index]));

            if (match.Groups[1].Success)
                segments.Add(new RichSegment(RichSegmentKind.Bold, match.Groups[1].Value));
            else if (match.Groups[2].Success)
                segments.Add(new RichSegment(RichSegmentKind.Underline, match.Groups[2].Value));
            else if (match.Groups[3].Success)
                segments.Add(new RichSegment(RichSegmentKind.Italic, match.Groups[3].Value));
            else
            {
                segments.Add(new RichSegment(
                    RichSegmentKind.Link,
                    match.Groups[4].Value,
                    match.Groups[5].Value.Trim()));
            }

            index = match.Index + match.Length;
        }

        if (index < value.Length)
            segments.Add(new RichSegment(RichSegmentKind.Text, value[index..]));

        return segments;
    }
}
