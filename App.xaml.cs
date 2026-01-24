using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace aribeth
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "aribeth",
            "app.log");

        private void App_Startup(object sender, StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;

            try
            {
                var window = new MainWindow();
                window.Show();
            }
            catch (Exception ex)
            {
                LogException("Startup exception", ex);
                MessageBox.Show($"Failed to start: {ex.Message}\n\nLog: {LogPath}", "Aribeth",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            LogException("Unhandled UI exception", e.Exception);
            MessageBox.Show($"Unhandled error: {e.Exception.Message}\n\nLog: {LogPath}", "Aribeth",
                MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object? sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogException("Unhandled domain exception", ex);
            }
        }

        private static void LogException(string title, Exception ex)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                var builder = new StringBuilder();
                builder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}");
                builder.AppendLine(ex.ToString());
                builder.AppendLine();
                File.AppendAllText(LogPath, builder.ToString());
            }
            catch
            {
                // Swallow logging failures.
            }
        }
    }

}
