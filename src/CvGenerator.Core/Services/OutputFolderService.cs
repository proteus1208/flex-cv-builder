using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using CvGenerator.Core.Models;

namespace CvGenerator.Core.Services;

public sealed class OutputFolderService
{
    private static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();
    private static readonly TimeZoneInfo Tokyo = ResolveTokyoTimeZone();
    private readonly JsonPathResolver _paths = new();

    public string BuildPreviewPath(string outputRoot, JsonElement root, DateTimeOffset? instant = null)
    {
        var relative = BuildRelativeOutputPath(root, instant, out _);
        return Path.Combine(outputRoot, relative);
    }

    /// <summary>Backward-compatible overload; <paramref name="date"/> is interpreted as Tokyo wall time when Kind is Unspecified.</summary>
    public string BuildPreviewPath(string outputRoot, JsonElement root, DateTime? date)
        => BuildPreviewPath(outputRoot, root, ToInstant(date));

    public string CreateOutputDirectory(
        string outputRoot,
        JsonElement root,
        IssueCollector issues,
        DateTimeOffset? instant = null)
    {
        var relative = BuildRelativeOutputPath(root, instant, out var usedFallback);

        if (usedFallback)
        {
            issues.Info(
                "FolderFromTimestamp",
                "Folder",
                "No 'directory' in JSON — using Tokyo date folder + time stamp only.",
                relative);
        }

        try
        {
            Directory.CreateDirectory(outputRoot);
            var candidate = Path.Combine(outputRoot, relative);
            var unique = MakeUniqueDirectory(candidate, issues);
            Directory.CreateDirectory(unique);
            return unique;
        }
        catch (Exception ex)
        {
            issues.Error("FolderCreateFailed", "Folder", $"Cannot create output folder: {ex.Message}", detail: ex.ToString());
            return string.Empty;
        }
    }

    public string CreateOutputDirectory(
        string outputRoot,
        JsonElement root,
        IssueCollector issues,
        DateTime? date)
        => CreateOutputDirectory(outputRoot, root, issues, ToInstant(date));

    /// <summary>
    /// Relative output path under the root:
    /// <c>{yyyy-MM-dd}/{HH-mm-ss}_{directory}</c> (Tokyo),
    /// or <c>{yyyy-MM-dd}/{HH-mm-ss}</c> when JSON <c>directory</c> is missing.
    /// </summary>
    public string BuildRelativeOutputPath(JsonElement root, out bool usedFallback)
        => BuildRelativeOutputPath(root, null, out usedFallback);

    public string BuildRelativeOutputPath(JsonElement root, DateTimeOffset? instant, out bool usedFallback)
    {
        var tokyo = ToTokyo(instant ?? DateTimeOffset.UtcNow);
        var dateDir = tokyo.ToString("yyyy-MM-dd");
        var stamp = tokyo.ToString("HH-mm-ss");

        var directory = Sanitize(_paths.GetString(root, "directory") ?? string.Empty);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            usedFallback = false;
            return Path.Combine(dateDir, $"{stamp}_{directory}");
        }

        usedFallback = true;
        return Path.Combine(dateDir, stamp);
    }

    /// <summary>Legacy leaf helper — returns the relative path (date + stamp[+directory]).</summary>
    public string BuildLeafFolderName(JsonElement root, out bool usedFallback)
        => BuildRelativeOutputPath(root, null, out usedFallback);

    public string BuildLeafFolderName(JsonElement root, DateTime now, out bool usedFallback)
        => BuildRelativeOutputPath(root, ToInstant(now), out usedFallback);

    private static DateTimeOffset ToTokyo(DateTimeOffset instant)
        => TimeZoneInfo.ConvertTime(instant, Tokyo);

    private static DateTimeOffset? ToInstant(DateTime? date)
    {
        if (date is null)
            return null;

        var value = date.Value;
        return value.Kind switch
        {
            DateTimeKind.Utc => new DateTimeOffset(value, TimeSpan.Zero),
            DateTimeKind.Local => new DateTimeOffset(value),
            // Tests / callers pass a Tokyo wall-clock time with Unspecified kind.
            _ => new DateTimeOffset(DateTime.SpecifyKind(value, DateTimeKind.Unspecified), Tokyo.GetUtcOffset(value))
        };
    }

    private static TimeZoneInfo ResolveTokyoTimeZone()
    {
        foreach (var id in new[] { "Asia/Tokyo", "Tokyo Standard Time" })
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(id);
            }
            catch (TimeZoneNotFoundException)
            {
            }
            catch (InvalidTimeZoneException)
            {
            }
        }

        // Fixed JST (+09:00, no DST) fallback.
        return TimeZoneInfo.CreateCustomTimeZone(
            "Tokyo Standard Time",
            TimeSpan.FromHours(9),
            "Tokyo Standard Time",
            "Tokyo Standard Time");
    }

    private static string Sanitize(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var sb = new StringBuilder(value.Trim().Length);
        foreach (var ch in value.Trim())
        {
            if (InvalidChars.Contains(ch) || ch is '/' or '\\' or ':')
                sb.Append('-');
            else
                sb.Append(ch);
        }

        var cleaned = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
        cleaned = Regex.Replace(cleaned, @"-+", "-");
        return cleaned.Trim('.', ' ');
    }

    private static string MakeUniqueDirectory(string candidate, IssueCollector issues)
    {
        if (!Directory.Exists(candidate))
            return candidate;

        for (var i = 2; i < 1000; i++)
        {
            var next = $"{candidate}-{i}";
            if (!Directory.Exists(next))
            {
                issues.Info("FolderCollision", "Folder", $"Folder exists; using {Path.GetFileName(next)}.");
                return next;
            }
        }

        var tokyo = ToTokyo(DateTimeOffset.UtcNow);
        var timed = $"{candidate}-{tokyo:HHmmss}";
        issues.Info("FolderCollision", "Folder", $"Folder exists; using {Path.GetFileName(timed)}.");
        return timed;
    }
}
