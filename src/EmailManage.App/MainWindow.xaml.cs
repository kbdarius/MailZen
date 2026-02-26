using System.Windows;
using EmailManage.Services;
using EmailManage.ViewModels;

namespace EmailManage;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        try
        {
            DataContext = new MainViewModel();
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

    private void Window_Closed(object? sender, EventArgs e)
    {
        if (DataContext is MainViewModel vm)
            vm.Cleanup();
    }
}