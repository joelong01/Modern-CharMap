using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;

namespace ModernCharMap.WinUI
{
    public partial class App : Application
    {
        private Window? _window;

        public App()
        {
            UnhandledException += OnUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        }

        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            try
            {
                _window = new MainWindow();
                _window.Activate();
                TryInitWindowing(_window);
            }
            catch (Exception ex)
            {
                ShowFatalError(ex);
            }
        }

        private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            ShowFatalError(e.Exception);
        }

        private static void OnDomainUnhandledException(object sender, System.UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
                ShowFatalError(ex);
            else
                ShowFatalError(new Exception(e.ExceptionObject?.ToString() ?? "Unknown error"));
        }

        private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            e.SetObserved();
            ShowFatalError(e.Exception);
        }

        private static void ShowFatalError(Exception ex)
        {
            var msg = $"{ex.GetType().Name}: {ex.Message}\n\n{ex.StackTrace}";

            var inner = ex.InnerException;
            while (inner is not null)
            {
                msg += $"\n\n--- Inner: {inner.GetType().Name} ---\n{inner.Message}\n{inner.StackTrace}";
                inner = inner.InnerException;
            }

            // Write to a log file so the user can copy/paste it.
            string logPath = WriteCrashLog(msg);
            string logNote = !string.IsNullOrEmpty(logPath)
                ? $"\n\nLog written to:\n{logPath}"
                : "";

            string fullMsg = msg + logNote + "\n\nPress Ctrl+C to copy this dialog text.";

            NativeMethods.MessageBoxW(IntPtr.Zero, fullMsg,
                "Modern CharMap - Fatal Error", 0x10 /* MB_ICONERROR */);
        }

        private static string WriteCrashLog(string message)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ModernCharMap", "Logs");
                Directory.CreateDirectory(dir);

                string fileName = $"crash-{DateTime.Now:yyyyMMdd-HHmmss}.log";
                string path = Path.Combine(dir, fileName);
                File.WriteAllText(path, message);
                return path;
            }
            catch
            {
                return "";
            }
        }

        private static void TryInitWindowing(Window window)
        {
            try
            {
                var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
                var id = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hwnd);
                var appWindow = AppWindow.GetFromWindowId(id);
                appWindow.Title = "Modern CharMap";
            }
            catch
            {
                // ignore
            }
        }

        private static class NativeMethods
        {
            [DllImport("user32", CharSet = CharSet.Unicode)]
            public static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);
        }
    }
}
