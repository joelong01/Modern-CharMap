using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ModernCharMap.WinUI.Services
{
    public sealed class ClipboardService : IClipboardService
    {
        public static IClipboardService Instance { get; } = new ClipboardService();

        private static readonly uint CF_HTML = NativeMethods.RegisterClipboardFormatW("HTML Format");

        private ClipboardService() { }

        public void SetText(string text)
        {
            SetTextWithFont(text, fontFamily: null);
        }

        public void SetTextWithFont(string text, string? fontFamily)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (!NativeMethods.OpenClipboard(IntPtr.Zero)) return;
            try
            {
                NativeMethods.EmptyClipboard();
                SetClipboardUnicode(text);
                if (!string.IsNullOrEmpty(fontFamily))
                {
                    SetClipboardHtml(text, fontFamily!);
                }
            }
            finally
            {
                NativeMethods.CloseClipboard();
            }
        }

        public string? GetText()
        {
            return null;
        }

        private static void SetClipboardUnicode(string text)
        {
            var byteCount = (text.Length + 1) * 2;
            IntPtr hGlobal = NativeMethods.GlobalAlloc(NativeMethods.GMEM_MOVEABLE, (UIntPtr)byteCount);
            if (hGlobal == IntPtr.Zero) return;

            IntPtr target = NativeMethods.GlobalLock(hGlobal);
            if (target == IntPtr.Zero) return;
            try
            {
                Marshal.Copy(text.ToCharArray(), 0, target, text.Length);
                Marshal.WriteInt16(target, text.Length * 2, 0);
            }
            finally
            {
                NativeMethods.GlobalUnlock(hGlobal);
            }
            NativeMethods.SetClipboardData(NativeMethods.CF_UNICODETEXT, hGlobal);
        }

        private static void SetClipboardHtml(string text, string fontFamily)
        {
            // Encode each character as an HTML numeric entity so the bytes
            // are pure ASCII, which avoids any UTF-8 encoding issues in the
            // CF_HTML header math.
            var entityBuilder = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    int cp = char.ConvertToUtf32(text[i], text[i + 1]);
                    entityBuilder.Append($"&#x{cp:X};");
                    i++;
                }
                else
                {
                    entityBuilder.Append($"&#x{(int)text[i]:X};");
                }
            }
            string entities = entityBuilder.ToString();

            // Build the HTML fragment.
            // Word/rich-text apps read the font-family from the style attribute.
            string htmlBody = $"<span style=\"font-family: '{fontFamily}'; font-size: 24pt;\">{entities}</span>";

            // CF_HTML requires a specific header format with byte offsets.
            // Build with placeholders, then patch the offsets.
            const string header =
                "Version:0.9\r\n" +
                "StartHTML:SSSSSSSSSS\r\n" +
                "EndHTML:EEEEEEEEEE\r\n" +
                "StartFragment:FFFFFFFFFF\r\n" +
                "EndFragment:GGGGGGGGGG\r\n";

            const string htmlStart = "<html><body>\r\n<!--StartFragment-->";
            const string htmlEnd = "<!--EndFragment-->\r\n</body></html>";

            string raw = header + htmlStart + htmlBody + htmlEnd;

            // All content is ASCII, so byte length == character length.
            int startHtml = header.Length;
            int startFragment = startHtml + htmlStart.Length;
            int endFragment = startFragment + htmlBody.Length;
            int endHtml = endFragment + htmlEnd.Length;

            string result = raw
                .Replace("SSSSSSSSSS", startHtml.ToString("D10"))
                .Replace("EEEEEEEEEE", endHtml.ToString("D10"))
                .Replace("FFFFFFFFFF", startFragment.ToString("D10"))
                .Replace("GGGGGGGGGG", endFragment.ToString("D10"));

            byte[] bytes = Encoding.UTF8.GetBytes(result);

            IntPtr hGlobal = NativeMethods.GlobalAlloc(NativeMethods.GMEM_MOVEABLE, (UIntPtr)(bytes.Length + 1));
            if (hGlobal == IntPtr.Zero) return;

            IntPtr target = NativeMethods.GlobalLock(hGlobal);
            if (target == IntPtr.Zero) return;
            try
            {
                Marshal.Copy(bytes, 0, target, bytes.Length);
                Marshal.WriteByte(target, bytes.Length, 0); // null terminator
            }
            finally
            {
                NativeMethods.GlobalUnlock(hGlobal);
            }
            NativeMethods.SetClipboardData(CF_HTML, hGlobal);
        }

        private static class NativeMethods
        {
            public const uint CF_UNICODETEXT = 13;
            public const uint GMEM_MOVEABLE = 0x0002;

            [DllImport("user32", SetLastError = true)]
            public static extern bool OpenClipboard(IntPtr hWndNewOwner);

            [DllImport("user32", SetLastError = true)]
            public static extern bool CloseClipboard();

            [DllImport("user32", SetLastError = true)]
            public static extern bool EmptyClipboard();

            [DllImport("user32", SetLastError = true)]
            public static extern IntPtr SetClipboardData(uint uFormat, IntPtr data);

            [DllImport("user32", SetLastError = true)]
            public static extern uint RegisterClipboardFormatW([MarshalAs(UnmanagedType.LPWStr)] string lpszFormat);

            [DllImport("kernel32", SetLastError = true)]
            public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

            [DllImport("kernel32", SetLastError = true)]
            public static extern IntPtr GlobalLock(IntPtr hMem);

            [DllImport("kernel32", SetLastError = true)]
            public static extern bool GlobalUnlock(IntPtr hMem);
        }
    }
}