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
        var day = (date ?? DateTime.Now).ToString("yyyy-MM-dd");
        var leaf = BuildLeafFolderName(root, out _);
        return Path.Combine(outputRoot, day, leaf);
    }

    public string CreateOutputDirectory(
        string outputRoot,
        JsonElement root,
        IssueCollector issues,
        DateTime? date = null)
    {
        var day = (date ?? DateTime.Now).ToString("yyyy-MM-dd");
        var leaf = BuildLeafFolderName(root, out var usedFallback);

        if (usedFallback)
        {
            issues.Warning(
                "MissingFolderFields",
                "Folder",
                "Missing name/role/topSkills — using fallback folder name.",
                leaf);
        }

        try
        {
            Directory.CreateDirectory(outputRoot);
            var dayDir = Path.Combine(outputRoot, day);
            Directory.CreateDirectory(dayDir);

            var candidate = Path.Combine(dayDir, leaf);
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

    public string BuildLeafFolderName(JsonElement root, out bool usedFallback)
    {
        usedFallback = false;
        var name = Sanitize(_paths.GetString(root, "name") ?? string.Empty);
        var role = Sanitize(_paths.GetString(root, "role") ?? string.Empty);
        var skills = GetTopSkills(root);

        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(role) && skills.Count == 0)
        {
            usedFallback = true;
            return "Untitled-CV";
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            usedFallback = true;
            name = "Unknown";
        }

        if (string.IsNullOrWhiteSpace(role))
        {
            usedFallback = true;
            role = "Role";
        }

        while (skills.Count < 3)
        {
            usedFallback = true;
            skills.Add("Skill");
        }

        var skillPart = string.Join("-", skills.Take(3).Select(Sanitize));
        return $"{name} - {role} - {skillPart}";
    }

    private List<string> GetTopSkills(JsonElement root)
    {
        var result = new List<string>();

        if (_paths.TryGetElement(root, "topSkills", out var el))
            AddSkills(el, result);
        else if (_paths.TryGetElement(root, "metadata.topStacks", out el))
            AddSkills(el, result);

        return result;
    }

    private static void AddSkills(JsonElement el, List<string> result)
    {
        if (el.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in el.EnumerateArray())
            {
                var s = item.ValueKind == JsonValueKind.String ? item.GetString() : item.ToString();
                if (!string.IsNullOrWhiteSpace(s))
                    result.Add(s!);
            }
        }
        else if (el.ValueKind == JsonValueKind.String)
        {
            foreach (var part in el.GetString()!.Split(new[] { ',', '|' }, StringSplitOptions.RemoveEmptyEntries))
                result.Add(part.Trim());
        }
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
        return cleaned;
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
