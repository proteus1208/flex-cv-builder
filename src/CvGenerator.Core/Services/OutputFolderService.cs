using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using CvGenerator.Core.Models;

namespace CvGenerator.Core.Services;

public sealed class OutputFolderService
{
    private static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();
    private readonly JsonPathResolver _paths = new();

    public string BuildPreviewPath(string outputRoot, JsonElement root, DateTime? date = null)
    {
        var now = date ?? DateTime.Now;
        var leaf = BuildLeafFolderName(root, now, out _);
        return Path.Combine(outputRoot, leaf);
    }

    public string CreateOutputDirectory(
        string outputRoot,
        JsonElement root,
        IssueCollector issues,
        DateTime? date = null)
    {
        var now = date ?? DateTime.Now;
        var leaf = BuildLeafFolderName(root, now, out var usedFallback);

        if (usedFallback)
        {
            issues.Info(
                "FolderFromTimestamp",
                "Folder",
                "No 'directory' in JSON — using date/time folder name.",
                leaf);
        }

        try
        {
            Directory.CreateDirectory(outputRoot);
            var candidate = Path.Combine(outputRoot, leaf);
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

    /// <summary>
    /// Leaf folder = JSON <c>directory</c> when present; otherwise current date/time.
    /// </summary>
    public string BuildLeafFolderName(JsonElement root, out bool usedFallback)
        => BuildLeafFolderName(root, DateTime.Now, out usedFallback);

    public string BuildLeafFolderName(JsonElement root, DateTime now, out bool usedFallback)
    {
        var directory = Sanitize(_paths.GetString(root, "directory") ?? string.Empty);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            usedFallback = false;
            return directory;
        }

        usedFallback = true;
        return now.ToString("yyyy-MM-dd_HH-mm-ss");
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

        var timed = $"{candidate}-{DateTime.Now:HHmmss}";
        issues.Info("FolderCollision", "Folder", $"Folder exists; using {Path.GetFileName(timed)}.");
        return timed;
    }
}
