using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;

namespace ModernCharMap.WinUI
{
    /// <summary>
    /// Application entry point. Handles window creation, global exception handling,
    /// crash logging, and initial window configuration.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Three exception handlers are registered to catch unhandled errors from different sources:
    /// <list type="bullet">
    ///   <item><description>WinUI <c>UnhandledException</c>: XAML framework exceptions.</description></item>
    ///   <item><description><c>AppDomain.UnhandledException</c>: CLR-level unhandled exceptions.</description></item>
    ///   <item><description><c>TaskScheduler.UnobservedTaskException</c>: unobserved async Task failures.</description></item>
    /// </list>
    /// All three delegate to <see cref="ShowFatalError"/> which displays a native
    /// Win32 MessageBox (not XAML, since the UI framework may be in a broken state)
    /// and writes a crash log to <c>%LOCALAPPDATA%\ModernCharMap\Logs\</c>.
    /// </para>
    /// </remarks>
    public partial class App : Application
    {
        private Window? _window;

        /// <summary>
        /// Registers global exception handlers for all unhandled error sources.
        /// </summary>
        public App()
        {
            UnhandledException += OnUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        }

        /// <summary>
        /// Creates and activates the main window. Wrapped in a try/catch so that
        /// startup failures are caught and displayed via <see cref="ShowFatalError"/>.
        /// </summary>
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

        /// <summary>
        /// Handles WinUI XAML framework exceptions. Marks the exception as handled
        /// to prevent process termination, then shows the error dialog.
        /// </summary>
        private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            ShowFatalError(e.Exception);
        }

        /// <summary>
        /// Handles CLR-level unhandled exceptions (e.g. from non-async threads).
        /// </summary>
        private static void OnDomainUnhandledException(object sender, System.UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
                ShowFatalError(ex);
            else
                ShowFatalError(new Exception(e.ExceptionObject?.ToString() ?? "Unknown error"));
        }

        /// <summary>
        /// Handles unobserved Task exceptions (async methods that throw without an await).
        /// Marks the exception as observed to prevent process termination.
        /// </summary>
        private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            e.SetObserved();
            ShowFatalError(e.Exception);
        }

        /// <summary>
        /// Displays a fatal error dialog using a native Win32 MessageBox (not XAML),
        /// including the full exception chain and a crash log file path.
        /// </summary>
        /// <param name="ex">The exception to display.</param>
        /// <remarks>
        /// A native MessageBox is used instead of a XAML dialog because the XAML
        /// framework may be in a broken state when this is called. The dialog text
        /// can be copied via Ctrl+C.
        /// </remarks>
        private static void ShowFatalError(Exception ex)
        {
            var msg = $"{ex.GetType().Name}: {ex.Message}\n\n{ex.StackTrace}";

            // Append inner exception chain
            var inner = ex.InnerException;
            while (inner is not null)
            {
                msg += $"\n\n--- Inner: {inner.GetType().Name} ---\n{inner.Message}\n{inner.StackTrace}";
                inner = inner.InnerException;
            }

            string logPath = WriteCrashLog(msg);
            string logNote = !string.IsNullOrEmpty(logPath)
                ? $"\n\nLog written to:\n{logPath}"
                : "";

            string fullMsg = msg + logNote + "\n\nPress Ctrl+C to copy this dialog text.";

            NativeMethods.MessageBoxW(IntPtr.Zero, fullMsg,
                "Modern CharMap - Fatal Error", 0x10 /* MB_ICONERROR */);
        }

        /// <summary>
        /// Writes the exception details to a timestamped log file in the application's
        /// local data directory.
        /// </summary>
        /// <param name="message">The formatted exception text to write.</param>
        /// <returns>The full path to the log file, or an empty string if writing failed.</returns>
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

        /// <summary>
        /// Attempts to set the window title via the AppWindow API.
        /// Wrapped in a try/catch because this can fail on some Windows versions
        /// or configurations without consequence.
        /// </summary>
        /// <param name="window">The window to configure.</param>
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
                // Non-critical — window still functions without a title
            }
        }

        /// <summary>
        /// Win32 P/Invoke for displaying a native message box.
        /// Used for fatal error dialogs when the XAML framework may be unavailable.
        /// </summary>
        private static class NativeMethods
        {
            /// <summary>
            /// Displays a modal message box with the specified text, caption, and icon.
            /// </summary>
            [DllImport("user32", CharSet = CharSet.Unicode)]
            public static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);
        }
    }
}
