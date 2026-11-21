namespace CvGenerator.Core.Settings;

public sealed class AppSettings
{
    public string OutputRoot { get; set; } = string.Empty;
    public string? LibreOfficePath { get; set; }
    public string TemplatesRoot { get; set; } = string.Empty;

    /// <summary>Last JSON editor contents, restored on next launch.</summary>
    public string? LastJson { get; set; }
}
