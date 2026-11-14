using System.Text.Json;

namespace CvGenerator.Core.Settings;

public sealed class SettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _settingsPath;

    public SettingsStore(string? settingsDirectory = null)
    {
        var dir = settingsDirectory
            ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CvGenerator");
        Directory.CreateDirectory(dir);
        _settingsPath = Path.Combine(dir, "settings.json");
    }

    public string SettingsPath => _settingsPath;

    public AppSettings Load(string defaultOutputRoot, string defaultTemplatesRoot)
    {
        if (!File.Exists(_settingsPath))
        {
            return new AppSettings
            {
                OutputRoot = defaultOutputRoot,
                TemplatesRoot = defaultTemplatesRoot
            };
        }

        try
        {
            var json = File.ReadAllText(_settingsPath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions)
                           ?? new AppSettings();
            if (string.IsNullOrWhiteSpace(settings.OutputRoot))
                settings.OutputRoot = defaultOutputRoot;
            if (string.IsNullOrWhiteSpace(settings.TemplatesRoot))
                settings.TemplatesRoot = defaultTemplatesRoot;
            return settings;
        }
        catch
        {
            return new AppSettings
            {
                OutputRoot = defaultOutputRoot,
                TemplatesRoot = defaultTemplatesRoot
            };
        }
    }

    public void Save(AppSettings settings)
    {
        var dir = Path.GetDirectoryName(_settingsPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(_settingsPath, json);
    }
}
