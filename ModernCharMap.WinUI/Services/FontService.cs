using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using ModernCharMap.WinUI.Services.DirectWrite;

namespace ModernCharMap.WinUI.Services
{
    public sealed class FontService : IFontService, IDisposable
    {
        public static IFontService Instance { get; } = new FontService();

        private readonly IDWriteFactory _factory;
        private IDWriteFontCollection? _fontCollection;
        private IReadOnlyList<string>? _cachedFamilies;
        private FileSystemWatcher? _perUserWatcher;
        private FileSystemWatcher? _systemWatcher;
        private Timer? _debounceTimer;

        public event EventHandler? FontsChanged;

        private static readonly string PerUserFontsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "Windows", "Fonts");

        private static readonly string SystemFontsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

        private const string FontsRegistryKey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

        private FontService()
        {
            _factory = DWrite.CreateFactory();
            StartFontWatchers();
        }

        public IReadOnlyList<string> GetInstalledFontFamilies()
        {
            if (_cachedFamilies is not null)
                return _cachedFamilies;

            EnsureFontCollection();
            var families = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            uint count = _fontCollection!.GetFontFamilyCount();

            for (uint i = 0; i < count; i++)
            {
                try
                {
                    int hr = _fontCollection.GetFontFamily(i, out var family);
                    if (hr != 0) continue;
                    try
                    {
                        hr = family.GetFamilyNames(out var names);
                        if (hr != 0) continue;
                        try
                        {
                            string name = DWriteHelper.GetBestLocalizedString(names);
                            if (!string.IsNullOrEmpty(name))
                                families.Add(name);
                        }
                        finally { Marshal.ReleaseComObject(names); }
                    }
                    finally { Marshal.ReleaseComObject(family); }
                }
                catch
                {
                    // Corrupted font files can throw — ignore
                }
            }

            _cachedFamilies = families.ToList().AsReadOnly();
            return _cachedFamilies;
        }

        public bool IsSymbolFont(string fontFamily)
        {
            EnsureFontCollection();
            int hr = _fontCollection!.FindFamilyName(fontFamily, out uint index, out bool exists);
            if (hr != 0 || !exists) return false;

            hr = _fontCollection.GetFontFamily(index, out var family);
            if (hr != 0) return false;
            try
            {
                hr = family.GetFont(0, out var font);
                if (hr != 0) return false;
                try
                {
                    return font.IsSymbolFont();
                }
                finally { Marshal.ReleaseComObject(font); }
            }
            finally { Marshal.ReleaseComObject(family); }
        }

        public IReadOnlyList<int> GetSupportedCodePoints(string fontFamily)
        {
            // Use DirectWrite IDWriteFontFace1::GetUnicodeRanges for full Unicode
            // support including SMP (emoji, symbols above U+FFFF).
            EnsureFontCollection();
            int hr = _fontCollection!.FindFamilyName(fontFamily, out uint famIndex, out bool exists);
            if (hr != 0 || !exists) return Array.Empty<int>();

            hr = _fontCollection.GetFontFamily(famIndex, out var family);
            if (hr != 0) return Array.Empty<int>();

            try
            {
                hr = family.GetFont(0, out var font);
                if (hr != 0) return Array.Empty<int>();
                try
                {
                    hr = font.CreateFontFace(out var fontFace);
                    if (hr != 0) return Array.Empty<int>();
                    try
                    {
                        // QI for IDWriteFontFace1 (available on Windows 8+)
                        var fontFace1 = (IDWriteFontFace1)fontFace;
                        return ExpandUnicodeRanges(fontFace1);
                    }
                    finally { Marshal.ReleaseComObject(fontFace); }
                }
                finally { Marshal.ReleaseComObject(font); }
            }
            finally { Marshal.ReleaseComObject(family); }
        }

        private static IReadOnlyList<int> ExpandUnicodeRanges(IDWriteFontFace1 fontFace1)
        {
            // First call: get range count
            fontFace1.GetUnicodeRanges(0, IntPtr.Zero, out uint rangeCount);
            if (rangeCount == 0) return Array.Empty<int>();

            // Second call: get actual ranges
            int structSize = Marshal.SizeOf<DWriteUnicodeRange>();
            IntPtr buffer = Marshal.AllocHGlobal(structSize * (int)rangeCount);
            try
            {
                int hr = fontFace1.GetUnicodeRanges(rangeCount, buffer, out _);
                if (hr != 0) return Array.Empty<int>();

                var result = new List<int>();
                for (int i = 0; i < (int)rangeCount; i++)
                {
                    var range = Marshal.PtrToStructure<DWriteUnicodeRange>(buffer + i * structSize);
                    for (uint cp = range.First; cp <= range.Last; cp++)
                    {
                        if (cp < 0x0020) continue;                    // C0 controls
                        if (cp >= 0x007F && cp <= 0x009F) continue;    // DEL + C1 controls
                        if (cp >= 0xD800 && cp <= 0xDFFF) continue;    // surrogates
                        if (cp >= 0xFDD0 && cp <= 0xFDEF) continue;    // noncharacters
                        result.Add((int)cp);
                    }
                }

                return result.AsReadOnly();
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        /// <summary>
        /// Reads glyph names from the font's OpenType 'post' table.
        /// Returns a dictionary mapping Unicode codepoints to glyph names.
        /// Returns an empty dictionary if the font has no meaningful glyph names.
        /// </summary>
        public IReadOnlyDictionary<int, string> GetGlyphNames(string fontFamily)
        {
            var result = new Dictionary<int, string>();

            EnsureFontCollection();
            int hr = _fontCollection!.FindFamilyName(fontFamily, out uint famIndex, out bool exists);
            if (hr != 0 || !exists) return result;

            hr = _fontCollection.GetFontFamily(famIndex, out var family);
            if (hr != 0) return result;

            try
            {
                hr = family.GetFont(0, out var font);
                if (hr != 0) return result;
                try
                {
                    hr = font.CreateFontFace(out var fontFace);
                    if (hr != 0) return result;
                    try
                    {
                        // Read the 'post' table
                        hr = fontFace.TryGetFontTable(DWrite.TAG_POST,
                            out var tableData, out var tableSize,
                            out var tableContext, out var tableExists);

                        if (hr != 0 || !tableExists || tableSize < 34)
                        {
                            if (tableExists) fontFace.ReleaseFontTable(tableContext);
                            return result;
                        }

                        try
                        {
                            var glyphNames = ParsePostTable(tableData, tableSize);
                            if (glyphNames.Count == 0)
                                return result;

                            // Map codepoints to glyph indices, then to names.
                            // Only process BMP characters.
                            var codePoints = GetSupportedCodePoints(fontFamily);
                            const int batchSize = 256;
                            for (int offset = 0; offset < codePoints.Count; offset += batchSize)
                            {
                                int count = Math.Min(batchSize, codePoints.Count - offset);
                                var cpBatch = new uint[count];
                                var giBatch = new ushort[count];

                                for (int j = 0; j < count; j++)
                                    cpBatch[j] = (uint)codePoints[offset + j];

                                hr = fontFace.GetGlyphIndices(cpBatch, (uint)count, giBatch);
                                if (hr != 0) continue;

                                for (int j = 0; j < count; j++)
                                {
                                    if (giBatch[j] != 0 &&
                                        glyphNames.TryGetValue(giBatch[j], out var name))
                                    {
                                        result[codePoints[offset + j]] = FormatGlyphName(name);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            fontFace.ReleaseFontTable(tableContext);
                        }
                    }
                    finally { Marshal.ReleaseComObject(fontFace); }
                }
                finally { Marshal.ReleaseComObject(font); }
            }
            finally { Marshal.ReleaseComObject(family); }

            return result;
        }

        public bool InstallFont(string fontFilePath)
        {
            try
            {
                if (!File.Exists(fontFilePath))
                    return false;

                // Ensure per-user fonts directory exists
                Directory.CreateDirectory(PerUserFontsDir);

                // Copy font file to per-user fonts directory
                string fileName = Path.GetFileName(fontFilePath);
                string destPath = Path.Combine(PerUserFontsDir, fileName);
                File.Copy(fontFilePath, destPath, overwrite: true);

                // Register in HKCU registry
                string fontName = Path.GetFileNameWithoutExtension(fileName);
                string ext = Path.GetExtension(fileName).ToLowerInvariant();
                string registryName = fontName + (ext == ".otf" ? " (OpenType)" : " (TrueType)");

                using (var key = Registry.CurrentUser.OpenSubKey(FontsRegistryKey, writable: true)
                    ?? Registry.CurrentUser.CreateSubKey(FontsRegistryKey))
                {
                    key.SetValue(registryName, destPath, RegistryValueKind.String);
                }

                // Make font available immediately
                NativeMethods.AddFontResourceW(destPath);
                BroadcastFontChange();

                // Start watching the per-user dir if we just created it
                EnsurePerUserWatcher();

                // Notify listeners to refresh
                FontsChanged?.Invoke(this, EventArgs.Empty);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public (bool Success, string Message) UninstallFont(string fontFamily)
        {
            try
            {
                // Look up font in HKCU registry (per-user fonts)
                using var key = Registry.CurrentUser.OpenSubKey(FontsRegistryKey, writable: true);
                if (key is null)
                    return (false, "Cannot access per-user font registry.");

                // Find the registry value matching this font family
                string? matchedValueName = null;
                string? fontFilePath = null;

                foreach (string valueName in key.GetValueNames())
                {
                    string baseName = valueName
                        .Replace(" (TrueType)", "")
                        .Replace(" (OpenType)", "");

                    if (baseName.Equals(fontFamily, StringComparison.OrdinalIgnoreCase) ||
                        valueName.StartsWith(fontFamily, StringComparison.OrdinalIgnoreCase))
                    {
                        matchedValueName = valueName;
                        fontFilePath = key.GetValue(valueName) as string;
                        break;
                    }
                }

                if (matchedValueName is null || fontFilePath is null)
                {
                    // Check if it's a system font (HKLM)
                    using var sysKey = Registry.LocalMachine.OpenSubKey(FontsRegistryKey);
                    if (sysKey is not null)
                    {
                        foreach (string valueName in sysKey.GetValueNames())
                        {
                            string baseName = valueName
                                .Replace(" (TrueType)", "")
                                .Replace(" (OpenType)", "");

                            if (baseName.Equals(fontFamily, StringComparison.OrdinalIgnoreCase) ||
                                valueName.StartsWith(fontFamily, StringComparison.OrdinalIgnoreCase))
                            {
                                return (false, $"'{fontFamily}' is a system font. Removing it requires administrator privileges.");
                            }
                        }
                    }

                    return (false, $"Font '{fontFamily}' not found in per-user fonts.");
                }

                // Remove font resource from current session
                NativeMethods.RemoveFontResourceW(fontFilePath);

                // Delete registry entry
                key.DeleteValue(matchedValueName);

                // Delete font file
                if (File.Exists(fontFilePath))
                {
                    try
                    {
                        File.Delete(fontFilePath);
                    }
                    catch (IOException)
                    {
                        // File may be in use — will be cleaned up on next reboot
                    }
                }

                BroadcastFontChange();
                FontsChanged?.Invoke(this, EventArgs.Empty);
                return (true, $"Font '{fontFamily}' has been uninstalled.");
            }
            catch (Exception ex)
            {
                return (false, $"Failed to uninstall '{fontFamily}': {ex.Message}");
            }
        }

        public void RefreshFontList()
        {
            // Release old collection
            if (_fontCollection is not null)
            {
                try { Marshal.ReleaseComObject(_fontCollection); } catch { }
                _fontCollection = null;
            }
            _cachedFamilies = null;

            // Re-create with checkForUpdates = true
            int hr = _factory.GetSystemFontCollection(out _fontCollection, true);
            if (hr != 0)
                _fontCollection = null;
        }

        private void StartFontWatchers()
        {
            try
            {
                // Watch per-user fonts directory
                if (Directory.Exists(PerUserFontsDir))
                    _perUserWatcher = CreateWatcher(PerUserFontsDir);
            }
            catch { /* best effort */ }

            try
            {
                // Watch system fonts directory
                if (Directory.Exists(SystemFontsDir))
                    _systemWatcher = CreateWatcher(SystemFontsDir);
            }
            catch { /* best effort */ }
        }

        private FileSystemWatcher CreateWatcher(string path)
        {
            var watcher = new FileSystemWatcher(path)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                Filter = "*.*",
                IncludeSubdirectories = false,
                EnableRaisingEvents = true
            };

            watcher.Created += OnFontFileChanged;
            watcher.Deleted += OnFontFileChanged;
            watcher.Renamed += (s, e) => OnFontFileChanged(s, e);

            return watcher;
        }

        private void OnFontFileChanged(object sender, FileSystemEventArgs e)
        {
            // Only care about font files
            var ext = Path.GetExtension(e.FullPath).ToLowerInvariant();
            if (ext != ".ttf" && ext != ".otf" && ext != ".ttc" && ext != ".fon")
                return;

            // Debounce: wait 500ms for rapid changes to settle
            _debounceTimer?.Dispose();
            _debounceTimer = new Timer(_ => FontsChanged?.Invoke(this, EventArgs.Empty),
                null, 500, Timeout.Infinite);
        }

        private void EnsurePerUserWatcher()
        {
            if (_perUserWatcher is null && Directory.Exists(PerUserFontsDir))
            {
                try { _perUserWatcher = CreateWatcher(PerUserFontsDir); }
                catch { /* best effort */ }
            }
        }

        private static void BroadcastFontChange()
        {
            // WM_FONTCHANGE = 0x001D, HWND_BROADCAST = 0xFFFF
            NativeMethods.SendMessageTimeoutW(
                (IntPtr)0xFFFF, 0x001D, IntPtr.Zero, IntPtr.Zero,
                0x0002 /* SMTO_ABORTIFHUNG */, 1000, out _);
        }

        private void EnsureFontCollection()
        {
            if (_fontCollection is not null) return;
            int hr = _factory.GetSystemFontCollection(out _fontCollection, false);
            Marshal.ThrowExceptionForHR(hr);
        }

        /// <summary>
        /// Parses the OpenType 'post' table format 2.0 to extract glyph names.
        /// Returns a map of glyph index -> PostScript glyph name.
        /// </summary>
        private static Dictionary<ushort, string> ParsePostTable(IntPtr data, uint size)
        {
            var result = new Dictionary<ushort, string>();

            uint version = DWriteHelper.ReadUInt32BE(data, 0);
            if (version != 0x00020000) // Only format 2.0 has per-glyph names
                return result;

            ushort numberOfGlyphs = DWriteHelper.ReadUInt16BE(data, 32);
            uint indexArrayEnd = 34u + (uint)numberOfGlyphs * 2u;
            if (size < indexArrayEnd)
                return result;

            // Read glyph name index array
            var glyphNameIndices = new ushort[numberOfGlyphs];
            for (int i = 0; i < numberOfGlyphs; i++)
                glyphNameIndices[i] = DWriteHelper.ReadUInt16BE(data, 34 + i * 2);

            // Read custom Pascal strings that follow the index array
            var customNames = new List<string>();
            int offset = (int)indexArrayEnd;
            while (offset < (int)size)
            {
                byte nameLength = Marshal.ReadByte(data, offset);
                offset++;
                if (offset + nameLength > (int)size) break;

                var nameBytes = new byte[nameLength];
                for (int i = 0; i < nameLength; i++)
                    nameBytes[i] = Marshal.ReadByte(data, offset + i);
                offset += nameLength;

                customNames.Add(Encoding.ASCII.GetString(nameBytes));
            }

            // Build glyph index -> name mapping (only for custom names >= 258)
            for (ushort gi = 0; gi < numberOfGlyphs; gi++)
            {
                ushort nameIndex = glyphNameIndices[gi];
                if (nameIndex >= 258)
                {
                    int customIndex = nameIndex - 258;
                    if (customIndex < customNames.Count)
                    {
                        string name = customNames[customIndex];
                        // Skip generic names like .notdef, .null
                        if (!string.IsNullOrEmpty(name) &&
                            !name.StartsWith('.') &&
                            !name.Equals("space", StringComparison.OrdinalIgnoreCase))
                        {
                            result[gi] = name;
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Converts PostScript glyph names to display-friendly format.
        /// e.g. "building-city" -> "Building City", "uniE900" -> kept as-is
        /// </summary>
        private static string FormatGlyphName(string postScriptName)
        {
            // Skip "uniXXXX" names — they're just codepoint references
            if (postScriptName.StartsWith("uni", StringComparison.OrdinalIgnoreCase) &&
                postScriptName.Length >= 7)
            {
                return postScriptName; // Keep as-is; caller can filter
            }

            // Replace hyphens, underscores, camelCase with spaces
            var sb = new StringBuilder(postScriptName.Length + 4);
            for (int i = 0; i < postScriptName.Length; i++)
            {
                char c = postScriptName[i];
                if (c == '-' || c == '_')
                {
                    sb.Append(' ');
                }
                else if (i > 0 && char.IsUpper(c) && char.IsLower(postScriptName[i - 1]))
                {
                    sb.Append(' ');
                    sb.Append(c);
                }
                else
                {
                    sb.Append(c);
                }
            }

            // Title case each word
            var words = sb.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return string.Join(' ', words.Select(w =>
                char.ToUpper(w[0]) + (w.Length > 1 ? w[1..] : "")));
        }

        private static class NativeMethods
        {
            [DllImport("gdi32", CharSet = CharSet.Unicode)]
            public static extern int AddFontResourceW(string lpFileName);

            [DllImport("gdi32", CharSet = CharSet.Unicode)]
            public static extern int RemoveFontResourceW(string lpFileName);

            [DllImport("user32", CharSet = CharSet.Unicode)]
            public static extern IntPtr SendMessageTimeoutW(
                IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
                uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
        }

        public void Dispose()
        {
            _debounceTimer?.Dispose();
            _perUserWatcher?.Dispose();
            _systemWatcher?.Dispose();
        }
    }
}
