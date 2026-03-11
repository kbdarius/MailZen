using System.Windows;
using EmailManage.Services;
using EmailManage.ViewModels;
using System.Text.Json;

namespace EmailManage;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private static readonly string WindowStatePath = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "EmailManage", "window_state.json");

    public MainWindow()
    {
        InitializeComponent();
        try
        {
            DataContext = new MainViewModel();
            RestoreWindowState();
            DiagnosticLogger.Instance.Info("MainWindow initialized with ViewModel.");
        }
        catch (Exception ex)
        {
            // Write to crash log so we can diagnose
            var crashPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EmailManage", "crash.log");
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(crashPath)!);
            System.IO.File.AppendAllText(crashPath, $"\n[{DateTime.UtcNow:O}] MainWindow ctor: {ex}\n");
            MessageBox.Show($"Failed to initialize:\n\n{ex.Message}", "EmailManage Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
            await vm.InitializeAsync();
    }

    private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        SaveWindowState();
    }

    private void Window_Closed(object? sender, EventArgs e)
    {
        if (DataContext is MainViewModel vm)
            vm.Cleanup();
    }

    private void RestoreWindowState()
    {
        try
        {
            if (!System.IO.File.Exists(WindowStatePath)) return;

            var json = System.IO.File.ReadAllText(WindowStatePath);
            var saved = JsonSerializer.Deserialize<SavedWindowState>(json);
            if (saved is null) return;

            var width = Math.Max(MinWidth, saved.Width);
            var height = Math.Max(MinHeight, saved.Height);

            // Keep restored window on a visible screen region, even if monitor setup changed.
            var virtualLeft = SystemParameters.VirtualScreenLeft;
            var virtualTop = SystemParameters.VirtualScreenTop;
            var virtualRight = virtualLeft + SystemParameters.VirtualScreenWidth;
            var virtualBottom = virtualTop + SystemParameters.VirtualScreenHeight;

            var maxLeft = Math.Max(virtualLeft, virtualRight - width);
            var maxTop = Math.Max(virtualTop, virtualBottom - height);

            Left = Math.Clamp(saved.Left, virtualLeft, maxLeft);
            Top = Math.Clamp(saved.Top, virtualTop, maxTop);
            Width = width;
            Height = height;

            WindowStartupLocation = WindowStartupLocation.Manual;

            if (saved.IsMaximized)
            {
                WindowState = WindowState.Maximized;
            }
        }
        catch (Exception ex)
        {
            DiagnosticLogger.Instance.Warn("Could not restore window state: {Message}", ex.Message);
        }
    }

    private void SaveWindowState()
    {
        try
        {
            var bounds = WindowState == WindowState.Normal
                ? new Rect(Left, Top, Width, Height)
                : RestoreBounds;

            var saved = new SavedWindowState
            {
                Left = bounds.Left,
                Top = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height,
                IsMaximized = WindowState == WindowState.Maximized
            };

            var dir = System.IO.Path.GetDirectoryName(WindowStatePath)!;
            if (!System.IO.Directory.Exists(dir))
                System.IO.Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(saved, new JsonSerializerOptions { WriteIndented = true });
            System.IO.File.WriteAllText(WindowStatePath, json);
        }
        catch (Exception ex)
        {
            DiagnosticLogger.Instance.Warn("Could not save window state: {Message}", ex.Message);
        }
    }

    private sealed class SavedWindowState
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public bool IsMaximized { get; set; }
    }
}