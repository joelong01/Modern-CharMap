using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ModernCharMap.WinUI.Services
{
    /// <summary>
    /// Implements clipboard operations using Win32 P/Invoke for direct control
    /// over clipboard formats, including CF_UNICODETEXT and CF_HTML.
    /// </summary>
    /// <remarks>
    /// <para>
    /// WinUI 3 does not provide a built-in clipboard API that supports setting
    /// multiple formats simultaneously. This implementation uses the Win32
    /// clipboard API directly to place both plain text (CF_UNICODETEXT) and
    /// rich HTML (CF_HTML) on the clipboard in a single operation.
    /// </para>
    /// <para>
    /// The CF_HTML format embeds the glyph character as HTML numeric entities
    /// inside a <c>&lt;span&gt;</c> with a <c>font-family</c> style. This allows
    /// rich-text applications (Word, Outlook, OneNote) to paste the character
    /// in the correct font.
    /// </para>
    /// </remarks>
    public sealed class ClipboardService : IClipboardService
    {
        /// <summary>
        /// The process-wide singleton instance of the clipboard service.
        /// </summary>
        public static IClipboardService Instance { get; } = new ClipboardService();

        /// <summary>
        /// The registered clipboard format ID for HTML content.
        /// Obtained at class load time via <c>RegisterClipboardFormatW("HTML Format")</c>.
        /// </summary>
        private static readonly uint CF_HTML = NativeMethods.RegisterClipboardFormatW("HTML Format");

        private ClipboardService() { }

        /// <inheritdoc />
        public void SetText(string text)
        {
            SetTextWithFont(text, fontFamily: null);
        }

        /// <inheritdoc />
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

        /// <inheritdoc />
        public string? GetText()
        {
            return null;
        }

        /// <summary>
        /// Places plain Unicode text on the clipboard as CF_UNICODETEXT.
        /// Allocates a global memory block, copies the UTF-16 string into it
        /// (including a null terminator), and hands ownership to the clipboard.
        /// </summary>
        /// <param name="text">The Unicode text to place on the clipboard.</param>
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
                Marshal.WriteInt16(target, text.Length * 2, 0); // null terminator
            }
            finally
            {
                NativeMethods.GlobalUnlock(hGlobal);
            }
            NativeMethods.SetClipboardData(NativeMethods.CF_UNICODETEXT, hGlobal);
        }

        /// <summary>
        /// Places font-aware HTML on the clipboard as CF_HTML, allowing rich-text
        /// editors to paste the glyph in the specified font.
        /// </summary>
        /// <param name="text">The glyph text (may contain surrogate pairs for SMP characters).</param>
        /// <param name="fontFamily">The font family to embed in the HTML style attribute.</param>
        /// <remarks>
        /// <para>
        /// Each character is encoded as an HTML numeric entity (<c>&amp;#xHEX;</c>) so the
        /// CF_HTML payload is pure ASCII. This avoids UTF-8 multi-byte offset miscalculations
        /// in the CF_HTML header, which requires precise byte-offset fields.
        /// </para>
        /// <para>
        /// The CF_HTML format requires a header with four byte-offset fields
        /// (StartHTML, EndHTML, StartFragment, EndFragment) formatted as 10-digit
        /// zero-padded decimals. The offsets point into the raw UTF-8 byte stream.
        /// </para>
        /// </remarks>
        private static void SetClipboardHtml(string text, string fontFamily)
        {
            // Encode each character as an HTML numeric entity for pure-ASCII output
            var entityBuilder = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    int cp = char.ConvertToUtf32(text[i], text[i + 1]);
                    entityBuilder.Append($"&#x{cp:X};");
                    i++; // skip low surrogate
                }
                else
                {
                    entityBuilder.Append($"&#x{(int)text[i]:X};");
                }
            }
            string entities = entityBuilder.ToString();

            // Build the HTML fragment with font-family styling
            string htmlBody = $"<span style=\"font-family: '{fontFamily}'; font-size: 24pt;\">{entities}</span>";

            // CF_HTML header with placeholder byte-offset fields
            const string header =
                "Version:0.9\r\n" +
                "StartHTML:SSSSSSSSSS\r\n" +
                "EndHTML:EEEEEEEEEE\r\n" +
                "StartFragment:FFFFFFFFFF\r\n" +
                "EndFragment:GGGGGGGGGG\r\n";

            const string htmlStart = "<html><body>\r\n<!--StartFragment-->";
            const string htmlEnd = "<!--EndFragment-->\r\n</body></html>";

            string raw = header + htmlStart + htmlBody + htmlEnd;

            // Compute byte offsets (all content is ASCII, so byte length == char length)
            int startHtml = header.Length;
            int startFragment = startHtml + htmlStart.Length;
            int endFragment = startFragment + htmlBody.Length;
            int endHtml = endFragment + htmlEnd.Length;

            // Patch the placeholders with actual byte offsets
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

        /// <summary>
        /// Win32 P/Invoke declarations for clipboard and global memory operations.
        /// </summary>
        private static class NativeMethods
        {
            /// <summary>Clipboard format identifier for Unicode text.</summary>
            public const uint CF_UNICODETEXT = 13;

            /// <summary>GlobalAlloc flag: allocate moveable memory (required for clipboard).</summary>
            public const uint GMEM_MOVEABLE = 0x0002;

            /// <summary>Opens the clipboard for modification. Must be paired with <see cref="CloseClipboard"/>.</summary>
            [DllImport("user32", SetLastError = true)]
            public static extern bool OpenClipboard(IntPtr hWndNewOwner);

            /// <summary>Closes the clipboard after modification.</summary>
            [DllImport("user32", SetLastError = true)]
            public static extern bool CloseClipboard();

            /// <summary>Clears all data from the clipboard.</summary>
            [DllImport("user32", SetLastError = true)]
            public static extern bool EmptyClipboard();

            /// <summary>Places data on the clipboard in the specified format. Ownership of the memory handle transfers to the system.</summary>
            [DllImport("user32", SetLastError = true)]
            public static extern IntPtr SetClipboardData(uint uFormat, IntPtr data);

            /// <summary>Registers a custom clipboard format by name and returns its format ID.</summary>
            [DllImport("user32", SetLastError = true)]
            public static extern uint RegisterClipboardFormatW([MarshalAs(UnmanagedType.LPWStr)] string lpszFormat);

            /// <summary>Allocates a block of global memory with the specified flags.</summary>
            [DllImport("kernel32", SetLastError = true)]
            public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

            /// <summary>Locks a global memory block and returns a pointer to its data.</summary>
            [DllImport("kernel32", SetLastError = true)]
            public static extern IntPtr GlobalLock(IntPtr hMem);

            /// <summary>Unlocks a previously locked global memory block.</summary>
            [DllImport("kernel32", SetLastError = true)]
            public static extern bool GlobalUnlock(IntPtr hMem);
        }
    }
}
