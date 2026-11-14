using System.IO;
using System.Windows;
using CvGenerator.App.Services;
using CvGenerator.Core.Services;
using CvGenerator.Core.Settings;
using CvGenerator.Docx.Services;
using Microsoft.Extensions.DependencyInjection;

namespace CvGenerator.App;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        var baseDir = AppContext.BaseDirectory;
        var defaultOutput = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "..", "Output"));
        var defaultTemplates = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "..", "Templates"));

        // When published, use folders beside the executable.
        var publishedOutput = Path.Combine(baseDir, "Output");
        var publishedTemplates = Path.Combine(baseDir, "Templates");
        if (Directory.Exists(publishedTemplates) || File.Exists(Path.Combine(baseDir, "CvGenerator.exe")))
        {
            defaultOutput = publishedOutput;
            defaultTemplates = publishedTemplates;
        }

        // Prefer repo-relative folders when running from source tree.
        var repoRoot = FindRepoRoot(baseDir);
        if (repoRoot != null)
        {
            defaultOutput = Path.Combine(repoRoot, "Output");
            defaultTemplates = Path.Combine(repoRoot, "Templates");
        }

        Directory.CreateDirectory(defaultOutput);
        Directory.CreateDirectory(defaultTemplates);

        var settingsStore = new SettingsStore();
        var settings = settingsStore.Load(defaultOutput, defaultTemplates);
        Directory.CreateDirectory(settings.OutputRoot);
        Directory.CreateDirectory(settings.TemplatesRoot);

        var services = new ServiceCollection();
        services.AddWpfBlazorWebView();
#if DEBUG
        services.AddBlazorWebViewDeveloperTools();
#endif
        services.AddSingleton(settingsStore);
        services.AddSingleton(settings);
        services.AddSingleton<AppState>();
        services.AddSingleton<CvJsonService>();
        services.AddSingleton<OutputFolderService>();
        services.AddSingleton<TemplateCatalogService>();
        services.AddSingleton<PlaceholderScanner>();
        services.AddSingleton<GenerationService>();

        var provider = services.BuildServiceProvider();
        Resources.Add("services", provider);

        var state = provider.GetRequiredService<AppState>();
        state.Initialize(settings, settingsStore, defaultTemplates);

        InitializeComponent();
    }

    private static string? FindRepoRoot(string start)
    {
        var dir = new DirectoryInfo(start);
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "CvGenerator.sln")) ||
                Directory.Exists(Path.Combine(dir.FullName, "Templates")) &&
                Directory.Exists(Path.Combine(dir.FullName, "src")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        return null;
    }
}
