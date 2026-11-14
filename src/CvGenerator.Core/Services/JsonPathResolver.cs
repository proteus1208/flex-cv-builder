using System.Text.Json;

namespace CvGenerator.Core.Services;

public sealed class JsonPathResolver
{
    public bool TryGetElement(JsonElement root, string path, out JsonElement element)
    {
        element = default;
        if (string.IsNullOrWhiteSpace(path))
            return false;

        var segments = path.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var current = root;

        foreach (var segment in segments)
        {
            if (current.ValueKind == JsonValueKind.Object &&
                current.TryGetProperty(segment, out var next))
            {
                current = next;
                continue;
            }

            return false;
        }

        element = current;
        return true;
    }

    public string? GetString(JsonElement root, string path)
    {
        if (!TryGetElement(root, path, out var el))
            return null;

        return el.ValueKind switch
        {
            JsonValueKind.String => el.GetString(),
            JsonValueKind.Number => el.ToString(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => null,
            JsonValueKind.Undefined => null,
            _ => el.ToString()
        };
    }

    /// <summary>
    /// Resolves a value for placeholder replacement. Returns false when the path is missing.
    /// </summary>
    public bool TryResolveValue(JsonElement root, string path, out string value)
    {
        value = string.Empty;
        if (!TryGetElement(root, path, out var el))
            return false;

        value = el.ValueKind switch
        {
            JsonValueKind.String => el.GetString() ?? string.Empty,
            JsonValueKind.Number => el.ToString(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => string.Empty,
            JsonValueKind.Array or JsonValueKind.Object => el.GetRawText(),
            _ => string.Empty
        };
        return true;
    }

    /// <summary>
    /// Gets array items for a loop path. Coerces a single object to a one-element list.
    /// </summary>
    public IReadOnlyList<JsonElement> GetArrayItems(
        JsonElement root,
        string arrayPath,
        out bool coercedFromObject,
        out bool missing)
    {
        coercedFromObject = false;
        missing = false;

        if (!TryGetElement(root, arrayPath, out var el))
        {
            missing = true;
            return Array.Empty<JsonElement>();
        }

        if (el.ValueKind == JsonValueKind.Array)
            return el.EnumerateArray().ToList();

        if (el.ValueKind == JsonValueKind.Object)
        {
            coercedFromObject = true;
            return new[] { el };
        }

        missing = true;
        return Array.Empty<JsonElement>();
    }
}
