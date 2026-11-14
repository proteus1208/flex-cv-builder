using System.Text.Json;
using System.Text.Json.Nodes;

namespace CvGenerator.Core.Services;

public sealed class CvJsonService
{
    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = true
    };

    public bool TryParse(string json, IssueCollector issues, out JsonDocument? document)
    {
        document = null;
        if (string.IsNullOrWhiteSpace(json))
        {
            issues.Error("InvalidJson", "ParseJson", "JSON is empty.");
            return false;
        }

        try
        {
            document = JsonDocument.Parse(json);
            return true;
        }
        catch (JsonException ex)
        {
            issues.Error("InvalidJson", "ParseJson", $"Invalid JSON: {ex.Message}", detail: ex.ToString());
            return false;
        }
    }

    public string Format(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return JsonSerializer.Serialize(doc.RootElement, WriteOptions);
    }

    public string WriteResumeJson(JsonElement root, string outputDirectory, IssueCollector issues)
    {
        var path = Path.Combine(outputDirectory, "resume.json");
        try
        {
            var node = JsonNode.Parse(root.GetRawText())?.AsObject()
                       ?? new JsonObject();

            if (node["metadata"] is not JsonObject metadata)
            {
                metadata = new JsonObject();
                node["metadata"] = metadata;
            }

            if (root.TryGetProperty("topSkills", out var skills) && skills.ValueKind == JsonValueKind.Array)
            {
                metadata["topStacks"] = JsonNode.Parse(skills.GetRawText());
            }
            else if (root.TryGetProperty("metadata", out var meta) &&
                     meta.ValueKind == JsonValueKind.Object &&
                     meta.TryGetProperty("topStacks", out var stacks))
            {
                metadata["topStacks"] = JsonNode.Parse(stacks.GetRawText());
            }

            metadata["generatedAt"] = DateTime.UtcNow.ToString("O");

            File.WriteAllText(path, node.ToJsonString(WriteOptions));
            return path;
        }
        catch (Exception ex)
        {
            issues.Error("JsonWriteFailed", "JsonOut", $"Failed to write resume.json: {ex.Message}", detail: ex.ToString());
            return string.Empty;
        }
    }
}
