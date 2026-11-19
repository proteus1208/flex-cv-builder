using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace CvGenerator.App.Services;

public partial class CursorOrbWindow : Window
{
    private readonly DispatcherTimer _followTimer;
    private Storyboard? _bounceBoard;
    private bool _closing;

    public CursorOrbWindow()
    {
        InitializeComponent();
        _followTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        _followTimer.Tick += (_, _) => FollowCursor();
        SourceInitialized += (_, _) => MakeClickThrough();
        Loaded += (_, _) =>
        {
            FollowCursor();
            _followTimer.Start();
            PlayAppear();
        };
    }

    public void BeginBusyPulse()
    {
        _bounceBoard?.Stop();
        var scaleUp = new DoubleAnimation(1.0, 1.12, TimeSpan.FromSeconds(0.35))
        {
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever,
            EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
        };
        // Full bounce cycle ≈ 0.7s (0.35 up + 0.35 down).
        _bounceBoard = new Storyboard();
        Storyboard.SetTarget(scaleUp, OrbScale);
        Storyboard.SetTargetProperty(scaleUp, new PropertyPath(ScaleTransform.ScaleXProperty));
        _bounceBoard.Children.Add(scaleUp);

        var scaleY = scaleUp.Clone();
        Storyboard.SetTarget(scaleY, OrbScale);
        Storyboard.SetTargetProperty(scaleY, new PropertyPath(ScaleTransform.ScaleYProperty));
        _bounceBoard.Children.Add(scaleY);
        _bounceBoard.Begin();
    }

    public async Task PlayDisappearAsync()
    {
        if (_closing)
            return;
        _closing = true;
        _bounceBoard?.Stop();

        var fade = new DoubleAnimation(Orb.Opacity, 0, TimeSpan.FromSeconds(0.3))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };
        var shrink = new DoubleAnimation(OrbScale.ScaleX, 0.35, TimeSpan.FromSeconds(0.3))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };

        var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        fade.Completed += (_, _) => tcs.TrySetResult();
        Orb.BeginAnimation(OpacityProperty, fade);
        OrbScale.BeginAnimation(ScaleTransform.ScaleXProperty, shrink);
        OrbScale.BeginAnimation(ScaleTransform.ScaleYProperty, shrink.Clone());
        await tcs.Task.WaitAsync(TimeSpan.FromSeconds(0.5));

        _followTimer.Stop();
        Close();
    }

    private void PlayAppear()
    {
        var fade = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(0.3))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        var grow = new DoubleAnimation(0.55, 1.0, TimeSpan.FromSeconds(0.3))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };
        Orb.BeginAnimation(OpacityProperty, fade);
        OrbScale.BeginAnimation(ScaleTransform.ScaleXProperty, grow);
        OrbScale.BeginAnimation(ScaleTransform.ScaleYProperty, grow.Clone());

        fade.Completed += (_, _) => BeginBusyPulse();
    }

    private void FollowCursor()
    {
        if (!GetCursorPos(out var pt))
            return;

        var source = PresentationSource.FromVisual(this);
        var transform = source?.CompositionTarget?.TransformFromDevice;
        Point dip = transform is null
            ? new Point(pt.X, pt.Y)
            : transform.Value.Transform(new Point(pt.X, pt.Y));

        Left = dip.X - (Width / 2);
        Top = dip.Y - (Height / 2);
    }

    private void MakeClickThrough()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        var style = GetWindowLong(hwnd, GwlExstyle);
        SetWindowLong(hwnd, GwlExstyle, style | WsExLayered | WsExTransparent | WsExToolwindow);
    }

    private const int GwlExstyle = -20;
    private const int WsExTransparent = 0x00000020;
    private const int WsExLayered = 0x00080000;
    private const int WsExToolwindow = 0x00000080;

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out PointApi lpPoint);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [StructLayout(LayoutKind.Sequential)]
    private struct PointApi
    {
        public int X;
        public int Y;
    }
}
