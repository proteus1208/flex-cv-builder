using System.IO;
using System.Windows;
using CvGenerator.App.Services;
using CvGenerator.Core.Services;
using CvGenerator.Core.Settings;
using CvGenerator.Docx.Services;
using Microsoft.AspNetCore.Components.WebView;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Web.WebView2.Core;

namespace CvGenerator.App;

public partial class MainWindow : Window
{
    private readonly AppState _state;
    private ClipboardHotKeyService? _hotKey;
    private bool _hotKeyBusy;

    public MainWindow()
    {
        var baseDir = AppContext.BaseDirectory;

        // Templates always live beside the running exe (not user-configurable).
        var templatesRoot = Path.GetFullPath(Path.Combine(baseDir, "Templates"));
        Directory.CreateDirectory(templatesRoot);

        // Output root may be changed by the user; default beside exe, or repo Output in dev.
        var defaultOutput = Path.Combine(baseDir, "Output");
        var repoRoot = FindRepoRoot(baseDir);
        if (repoRoot != null)
            defaultOutput = Path.Combine(repoRoot, "Output");
        Directory.CreateDirectory(defaultOutput);

        var settingsStore = new SettingsStore();
        var settings = settingsStore.Load(defaultOutput, templatesRoot);
        // Force exe-relative templates even if an old settings.json pointed elsewhere.
        settings.TemplatesRoot = templatesRoot;
        Directory.CreateDirectory(settings.OutputRoot);

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

        _state = provider.GetRequiredService<AppState>();
        _state.Initialize(settings, settingsStore, templatesRoot);

        InitializeComponent();

        _hotKey = new ClipboardHotKeyService(this);
        _hotKey.Activated += OnClipboardHotKey;
        Closed += (_, _) =>
        {
            _hotKey?.Dispose();
            _hotKey = null;
            _state.Dispose();
        };
    }

    private void OnBlazorWebViewInitialized(object? sender, BlazorWebViewInitializedEventArgs e)
    {
        // Windows dark mode otherwise paints white textarea glyphs on a white editor.
        try
        {
            e.WebView.CoreWebView2.Profile.PreferredColorScheme =
                CoreWebView2PreferredColorScheme.Light;
        }
        catch
        {
            // Older WebView2 runtimes may not expose PreferredColorScheme.
        }
    }

    private async void OnClipboardHotKey()
    {
        if (_hotKeyBusy || _state.IsBusy)
            return;

        _hotKeyBusy = true;
        CursorOrbWindow? orb = null;
        try
        {
            string? clip = null;
            try
            {
                if (Clipboard.ContainsText())
                    clip = Clipboard.GetText();
            }
            catch
            {
                return;
            }

            if (!AppState.LooksLikeJsonObject(clip))
                return;

            orb = new CursorOrbWindow();
            orb.Show();

            var path = await _state.GenerateFromClipboardJsonAsync();
            if (string.IsNullOrWhiteSpace(path))
            {
                // still finish orb animation so UI doesn't stick
            }
        }
        finally
        {
            if (orb != null)
            {
                try { await orb.PlayDisappearAsync(); }
                catch { try { orb.Close(); } catch { /* ignore */ } }
            }

            _hotKeyBusy = false;
        }
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
