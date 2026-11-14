using System.Text.Json;

namespace CvGenerator.Core.Services;

public sealed class NotesExportService
{
    private readonly JsonPathResolver _paths = new();

    public string WriteNotes(JsonElement root, string outputDirectory, IssueCollector issues)
    {
        var notesPath = Path.Combine(outputDirectory, "notes.txt");
        try
        {
            var notes = _paths.GetString(root, "notes");
            if (notes is null)
            {
                issues.Warning("MissingNotes", "Notes", "notes field missing — wrote empty notes.txt.", "notes");
                File.WriteAllText(notesPath, string.Empty);
            }
            else
            {
                File.WriteAllText(notesPath, notes);
            }

            return notesPath;
        }
        catch (Exception ex)
        {
            issues.Error("NotesWriteFailed", "Notes", $"Failed to write notes.txt: {ex.Message}", detail: ex.ToString());
            return string.Empty;
        }
    }
}
