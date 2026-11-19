using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Input;

namespace CvGenerator.App.Services;

/// <summary>
/// Global Ctrl+Shift+G → clipboard JSON generate pipeline.
/// </summary>
public sealed class ClipboardHotKeyService : IDisposable
{
    public const int HotKeyId = 0xC6E1;
    private const uint ModControl = 0x0002;
    private const uint ModShift = 0x0004;
    private const int WmHotkey = 0x0312;

    private readonly Window _window;
    private HwndSource? _source;
    private bool _registered;

    public event Action? Activated;

    public ClipboardHotKeyService(Window window)
    {
        _window = window;
        if (_window.IsLoaded)
            Attach();
        else
            _window.Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _window.Loaded -= OnLoaded;
        Attach();
    }

    private void Attach()
    {
        var helper = new WindowInteropHelper(_window);
        if (helper.Handle == IntPtr.Zero)
        {
            _window.SourceInitialized += (_, _) => Attach();
            return;
        }

        _source = HwndSource.FromHwnd(helper.Handle);
        _source?.AddHook(WndProc);

        var vk = (uint)KeyInterop.VirtualKeyFromKey(Key.G);
        _registered = RegisterHotKey(helper.Handle, HotKeyId, ModControl | ModShift, vk);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey && wParam.ToInt32() == HotKeyId)
        {
            handled = true;
            Activated?.Invoke();
        }

        return IntPtr.Zero;
    }

    public void Dispose()
    {
        if (_registered)
        {
            var hwnd = new WindowInteropHelper(_window).Handle;
            if (hwnd != IntPtr.Zero)
                UnregisterHotKey(hwnd, HotKeyId);
            _registered = false;
        }

        _source?.RemoveHook(WndProc);
        _source = null;
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
