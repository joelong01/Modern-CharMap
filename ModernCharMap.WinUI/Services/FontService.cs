using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using ModernCharMap.WinUI.Services.DirectWrite;

namespace ModernCharMap.WinUI.Services
{
    /// <summary>
    /// Provides font enumeration, codepoint discovery, glyph naming, and per-user
    /// font install/uninstall capabilities using DirectWrite COM interop.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is a singleton service (<see cref="Instance"/>) that wraps a DirectWrite
    /// factory and system font collection. It caches the font family list and
    /// lazily initializes the collection on first use.
    /// </para>
    /// <para>
    /// Font changes (install, uninstall, or external filesystem changes) are detected
    /// via <see cref="FileSystemWatcher"/> instances on the per-user and system font
    /// directories, with a 500ms debounce to coalesce rapid filesystem events.
    /// </para>
    /// <para>
    /// All COM objects obtained from DirectWrite are explicitly released via
    /// <see cref="Marshal.ReleaseComObject"/> in <c>finally</c> blocks to avoid leaking
    /// native resources.
    /// </para>
    /// </remarks>
    public sealed class FontService : IFontService, IDisposable
    {
        /// <summary>
        /// The process-wide singleton instance of the font service.
        /// </summary>
        public static IFontService Instance { get; } = new FontService();

        private IDWriteFactory _factory;
        private IDWriteFontCollection? _fontCollection;
        private IReadOnlyList<string>? _cachedFamilies;
        private FileSystemWatcher? _perUserWatcher;
        private FileSystemWatcher? _systemWatcher;
        private Timer? _debounceTimer;

        /// <inheritdoc />
        public event EventHandler? FontsChanged;

        /// <summary>
        /// Per-user font installation directory.
        /// Fonts placed here and registered in <c>HKCU</c> are available only to the
        /// current user and do not require administrator privileges.
        /// </summary>
        private static readonly string PerUserFontsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "Windows", "Fonts");

        /// <summary>
        /// System-wide font directory (<c>C:\Windows\Fonts</c>).
        /// Monitored for changes but not written to (requires admin).
        /// </summary>
        private static readonly string SystemFontsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

        /// <summary>
        /// Registry key under <c>HKCU</c> (or <c>HKLM</c> for system fonts) that maps
        /// font display names to their file paths.
        /// </summary>
        private const string FontsRegistryKey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

        /// <summary>
        /// Initializes the DirectWrite factory and starts filesystem watchers
        /// for automatic font change detection.
        /// </summary>
        private FontService()
        {
            _factory = DWrite.CreateFactory();
            StartFontWatchers();
        }

        /// <inheritdoc />
        /// <remarks>
        /// Results are cached after the first call. Call <see cref="RefreshFontList"/>
        /// to invalidate the cache after font installs or removals.
        /// Font names are sorted alphabetically (case-insensitive) via a <see cref="SortedSet{T}"/>.
        /// </remarks>
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
                    // Corrupted font files can throw — skip and continue
                }
            }

            _cachedFamilies = families.ToList().AsReadOnly();
            return _cachedFamilies;
        }

        /// <inheritdoc />
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

        /// <inheritdoc />
        /// <remarks>
        /// <para>
        /// Uses DirectWrite <see cref="IDWriteFontFace1.GetUnicodeRanges"/> for full
        /// Unicode support including the Supplementary Multilingual Plane (emoji,
        /// mathematical symbols, historic scripts above U+FFFF).
        /// </para>
        /// <para>
        /// This replaces the earlier GDI-based <c>GetFontUnicodeRanges</c> approach,
        /// which was limited to the Basic Multilingual Plane (U+0000–U+FFFF) because
        /// GDI's <c>WCRANGE</c> structure uses 16-bit <c>WCHAR</c> values.
        /// </para>
        /// </remarks>
        public IReadOnlyList<int> GetSupportedCodePoints(string fontFamily)
        {
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
                        // Cast triggers COM QueryInterface for IDWriteFontFace1 (Windows 8+)
                        var fontFace1 = (IDWriteFontFace1)fontFace;
                        return ExpandUnicodeRanges(fontFace1);
                    }
                    finally { Marshal.ReleaseComObject(fontFace); }
                }
                finally { Marshal.ReleaseComObject(font); }
            }
            finally { Marshal.ReleaseComObject(family); }
        }

        /// <summary>
        /// Expands the Unicode ranges from a font face into individual codepoints,
        /// filtering out control characters, surrogates, and noncharacters.
        /// </summary>
        /// <param name="fontFace1">The DirectWrite font face supporting range enumeration.</param>
        /// <returns>A read-only list of valid Unicode codepoints the font supports.</returns>
        /// <remarks>
        /// Uses the two-call pattern: first call with count=0 to get the number of
        /// ranges, then allocate a native buffer and call again to fill it.
        /// Filtered codepoint ranges:
        /// <list type="bullet">
        ///   <item><description>U+0000–U+001F: C0 control characters</description></item>
        ///   <item><description>U+007F–U+009F: DEL and C1 control characters</description></item>
        ///   <item><description>U+D800–U+DFFF: UTF-16 surrogate pairs (not valid codepoints)</description></item>
        ///   <item><description>U+FDD0–U+FDEF: Unicode noncharacters</description></item>
        /// </list>
        /// </remarks>
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

        /// <inheritdoc />
        /// <remarks>
        /// <para>
        /// Reads the OpenType <c>post</c> (PostScript) table to extract glyph names,
        /// then maps those names back to Unicode codepoints via the font's cmap.
        /// Only <c>post</c> table format 2.0 contains per-glyph names; other formats
        /// return an empty dictionary.
        /// </para>
        /// <para>
        /// Codepoints are processed in batches of 256 for efficient glyph index lookup
        /// via <see cref="IDWriteFontFace.GetGlyphIndices"/>.
        /// </para>
        /// </remarks>
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

                            // Map codepoints to glyph indices, then look up names
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

        /// <inheritdoc />
        /// <remarks>
        /// <para>Installation steps:</para>
        /// <list type="number">
        ///   <item><description>Copy the font file to <c>%LOCALAPPDATA%\Microsoft\Windows\Fonts\</c>.</description></item>
        ///   <item><description>Register the font in <c>HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts</c>.</description></item>
        ///   <item><description>Call <c>AddFontResourceW</c> to make it available in the current session.</description></item>
        ///   <item><description>Broadcast <c>WM_FONTCHANGE</c> so other applications pick up the change.</description></item>
        /// </list>
        /// </remarks>
        public bool InstallFont(string fontFilePath)
        {
            try
            {
                if (!File.Exists(fontFilePath))
                    return false;

                Directory.CreateDirectory(PerUserFontsDir);

                string fileName = Path.GetFileName(fontFilePath);
                string destPath = Path.Combine(PerUserFontsDir, fileName);

                // If overwriting an existing font file, unload the old version from
                // GDI's cache first. Without this, GDI keeps the old file memory-mapped
                // and the new glyphs won't appear until the next login.
                if (File.Exists(destPath))
                {
                    NativeMethods.RemoveFontResourceW(destPath);
                }

                File.Copy(fontFilePath, destPath, overwrite: true);

                // Build registry display name: "FontName (TrueType)" or "(OpenType)"
                string fontName = Path.GetFileNameWithoutExtension(fileName);
                string ext = Path.GetExtension(fileName).ToLowerInvariant();
                string registryName = fontName + (ext == ".otf" ? " (OpenType)" : " (TrueType)");

                using (var key = Registry.CurrentUser.OpenSubKey(FontsRegistryKey, writable: true)
                    ?? Registry.CurrentUser.CreateSubKey(FontsRegistryKey))
                {
                    key.SetValue(registryName, destPath, RegistryValueKind.String);
                }

                NativeMethods.AddFontResourceW(destPath);
                BroadcastFontChange();

                EnsurePerUserWatcher();

                FontsChanged?.Invoke(this, EventArgs.Empty);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <inheritdoc />
        /// <remarks>
        /// <para>
        /// Only per-user fonts (registered under <c>HKCU</c>) can be uninstalled.
        /// If the font is found under <c>HKLM</c> (system-wide), an error message
        /// is returned indicating that administrator privileges are required.
        /// </para>
        /// <para>Uninstallation steps:</para>
        /// <list type="number">
        ///   <item><description>Look up the font file path in the <c>HKCU</c> fonts registry.</description></item>
        ///   <item><description>Call <c>RemoveFontResourceW</c> to unload it from the current session.</description></item>
        ///   <item><description>Delete the registry entry.</description></item>
        ///   <item><description>Delete the font file (best effort — may be locked).</description></item>
        ///   <item><description>Broadcast <c>WM_FONTCHANGE</c>.</description></item>
        /// </list>
        /// </remarks>
        public (bool Success, string Message) UninstallFont(string fontFamily)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(FontsRegistryKey, writable: true);
                if (key is null)
                    return (false, "Cannot access per-user font registry.");

                // Search HKCU for a registry value matching this font family name.
                // The registry value name is the file's display name (e.g. "MyFont (TrueType)")
                // which may differ from the font's internal family name reported by DirectWrite.
                // We try three strategies:
                //   1. Exact match on the base name (value name minus "(TrueType)"/"(OpenType)")
                //   2. Prefix match (value name starts with the font family)
                //   3. Scan all per-user font file paths and use DirectWrite to resolve
                //      which file provides this family name
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

                // Strategy 3: if no name match, check which per-user font file
                // provides this family by looking at file paths in the per-user dir
                if (matchedValueName is null)
                {
                    foreach (string valueName in key.GetValueNames())
                    {
                        string? path = key.GetValue(valueName) as string;
                        if (path is null) continue;

                        // Only consider files in the per-user fonts directory
                        if (!path.StartsWith(PerUserFontsDir, StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Check if this file provides the requested font family
                        if (FontFileProvidesFamilyName(path, fontFamily))
                        {
                            matchedValueName = valueName;
                            fontFilePath = path;
                            break;
                        }
                    }
                }

                if (matchedValueName is null || fontFilePath is null)
                {
                    // Not found in HKCU — check HKLM for system fonts
                    return TryUninstallSystemFont(fontFamily);
                }

                NativeMethods.RemoveFontResourceW(fontFilePath);
                key.DeleteValue(matchedValueName);

                if (File.Exists(fontFilePath))
                {
                    try
                    {
                        File.Delete(fontFilePath);
                    }
                    catch (IOException)
                    {
                        // File may be locked by another process — will be cleaned up on reboot
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

        /// <inheritdoc />
        /// <remarks>
        /// <para>
        /// Destroys the current DirectWrite factory and font collection entirely,
        /// then creates a fresh isolated factory. This is necessary because the
        /// shared factory (<c>DWRITE_FACTORY_TYPE_SHARED</c>) is a process-global
        /// singleton that caches font face data aggressively — even with
        /// <c>checkForUpdates = true</c>, it may serve stale glyph data for a
        /// font file that was overwritten in place.
        /// </para>
        /// <para>
        /// An isolated factory (<c>DWRITE_FACTORY_TYPE_ISOLATED</c>) has its own
        /// independent caches and always reads fresh data from disk.
        /// </para>
        /// </remarks>
        public void RefreshFontList()
        {
            // Release the old collection
            if (_fontCollection is not null)
            {
                try { Marshal.ReleaseComObject(_fontCollection); } catch { }
                _fontCollection = null;
            }

            // Release the old factory — a SHARED factory is a process singleton
            // that caches font face data, so getting a new collection from it
            // won't pick up file content changes.
            try { Marshal.ReleaseComObject(_factory); } catch { }

            // Create a fresh ISOLATED factory with its own caches
            _factory = DWrite.CreateFactory(isolated: true);
            _cachedFamilies = null;

            int hr = _factory.GetSystemFontCollection(out _fontCollection, true);
            if (hr != 0)
                _fontCollection = null;
        }

        /// <summary>
        /// Starts <see cref="FileSystemWatcher"/> instances on the per-user and system
        /// font directories to detect external font installs and removals.
        /// </summary>
        /// <remarks>
        /// Watchers are best-effort: if a directory doesn't exist or access is denied,
        /// the watcher for that directory is simply skipped.
        /// </remarks>
        private void StartFontWatchers()
        {
            try
            {
                if (Directory.Exists(PerUserFontsDir))
                    _perUserWatcher = CreateWatcher(PerUserFontsDir);
            }
            catch { /* best effort */ }

            try
            {
                if (Directory.Exists(SystemFontsDir))
                    _systemWatcher = CreateWatcher(SystemFontsDir);
            }
            catch { /* best effort */ }
        }

        /// <summary>
        /// Creates a <see cref="FileSystemWatcher"/> for a font directory that triggers
        /// <see cref="OnFontFileChanged"/> when font files are created, deleted, or renamed.
        /// </summary>
        /// <param name="path">The directory to monitor.</param>
        /// <returns>A configured and active watcher.</returns>
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

        /// <summary>
        /// Handles filesystem change events for font files with a 500ms debounce
        /// to coalesce rapid events (e.g. multiple files installed at once).
        /// Only responds to font file extensions: <c>.ttf</c>, <c>.otf</c>, <c>.ttc</c>, <c>.fon</c>.
        /// </summary>
        private void OnFontFileChanged(object sender, FileSystemEventArgs e)
        {
            var ext = Path.GetExtension(e.FullPath).ToLowerInvariant();
            if (ext != ".ttf" && ext != ".otf" && ext != ".ttc" && ext != ".fon")
                return;

            // Debounce: reset timer on each event, fire only after 500ms of quiet
            _debounceTimer?.Dispose();
            _debounceTimer = new Timer(_ => FontsChanged?.Invoke(this, EventArgs.Empty),
                null, 500, Timeout.Infinite);
        }

        /// <summary>
        /// Lazily creates a per-user font directory watcher if the directory exists
        /// but no watcher has been created yet (e.g. the directory was just created
        /// during a font install).
        /// </summary>
        private void EnsurePerUserWatcher()
        {
            if (_perUserWatcher is null && Directory.Exists(PerUserFontsDir))
            {
                try { _perUserWatcher = CreateWatcher(PerUserFontsDir); }
                catch { /* best effort */ }
            }
        }

        /// <summary>
        /// Broadcasts <c>WM_FONTCHANGE</c> (0x001D) to all top-level windows so that
        /// other applications refresh their font lists.
        /// Uses <c>SMTO_ABORTIFHUNG</c> with a 1-second timeout to avoid blocking
        /// if a window is unresponsive.
        /// </summary>
        private static void BroadcastFontChange()
        {
            // HWND_BROADCAST = 0xFFFF, WM_FONTCHANGE = 0x001D
            NativeMethods.SendMessageTimeoutW(
                (IntPtr)0xFFFF, 0x001D, IntPtr.Zero, IntPtr.Zero,
                0x0002 /* SMTO_ABORTIFHUNG */, 1000, out _);
        }

        /// <summary>
        /// Lazily initializes the DirectWrite system font collection on first use.
        /// </summary>
        /// <exception cref="COMException">Thrown if DirectWrite fails to enumerate fonts.</exception>
        private void EnsureFontCollection()
        {
            if (_fontCollection is not null) return;
            int hr = _factory.GetSystemFontCollection(out _fontCollection, false);
            Marshal.ThrowExceptionForHR(hr);
        }

        /// <summary>
        /// Parses the OpenType <c>post</c> (PostScript naming) table, format 2.0,
        /// to extract per-glyph PostScript names.
        /// </summary>
        /// <param name="data">Pointer to the raw <c>post</c> table data.</param>
        /// <param name="size">Size of the table in bytes.</param>
        /// <returns>
        /// A dictionary mapping glyph indices to their PostScript names.
        /// Returns empty if the table is not format 2.0 or is malformed.
        /// </returns>
        /// <remarks>
        /// <para>
        /// The <c>post</c> table format 2.0 layout:
        /// <list type="bullet">
        ///   <item><description>Bytes 0–3: Version (0x00020000 for format 2.0)</description></item>
        ///   <item><description>Bytes 32–33: Number of glyphs</description></item>
        ///   <item><description>Bytes 34+: Array of glyph name indices (2 bytes each)</description></item>
        ///   <item><description>Following the array: Pascal strings (length byte + ASCII chars)</description></item>
        /// </list>
        /// Glyph name indices 0–257 refer to the standard Macintosh glyph name set;
        /// indices 258+ refer to the custom names in the Pascal string list.
        /// </para>
        /// </remarks>
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

            // Read the glyph name index array
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
                        // Skip generic/placeholder names
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
        /// Converts a PostScript glyph name into a human-readable display name.
        /// </summary>
        /// <param name="postScriptName">The raw PostScript glyph name (e.g. "building-city").</param>
        /// <returns>
        /// A title-cased, space-separated display name (e.g. "Building City").
        /// Names starting with "uni" followed by hex digits (codepoint references)
        /// are returned as-is.
        /// </returns>
        private static string FormatGlyphName(string postScriptName)
        {
            // Skip "uniXXXX" names — they're just codepoint references, not real names
            if (postScriptName.StartsWith("uni", StringComparison.OrdinalIgnoreCase) &&
                postScriptName.Length >= 7)
            {
                return postScriptName;
            }

            // Replace hyphens, underscores, and camelCase boundaries with spaces
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

            // Title-case each word
            var words = sb.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return string.Join(' ', words.Select(w =>
                char.ToUpper(w[0]) + (w.Length > 1 ? w[1..] : "")));
        }

        /// <summary>
        /// Searches HKLM for a system font matching <paramref name="fontFamily"/> and
        /// attempts to uninstall it. Uses three matching strategies (exact name, prefix,
        /// and OpenType name table inspection), then either uninstalls directly if the
        /// process has admin rights or launches an elevated PowerShell process.
        /// </summary>
        /// <param name="fontFamily">The DirectWrite font family name to uninstall.</param>
        /// <returns>A tuple indicating success and a user-facing message.</returns>
        private (bool Success, string Message) TryUninstallSystemFont(string fontFamily)
        {
            using var sysKey = Registry.LocalMachine.OpenSubKey(FontsRegistryKey);
            if (sysKey is null)
                return (false, $"Font '{fontFamily}' not found in installed fonts.");

            string? sysValueName = null;
            string? sysFilePath = null;

            // Strategy 1 & 2: Exact base-name match or prefix match
            foreach (string valueName in sysKey.GetValueNames())
            {
                string baseName = valueName
                    .Replace(" (TrueType)", "")
                    .Replace(" (OpenType)", "");

                if (baseName.Equals(fontFamily, StringComparison.OrdinalIgnoreCase) ||
                    valueName.StartsWith(fontFamily, StringComparison.OrdinalIgnoreCase))
                {
                    sysValueName = valueName;
                    string? raw = sysKey.GetValue(valueName) as string;
                    if (raw is not null && !Path.IsPathRooted(raw))
                        raw = Path.Combine(SystemFontsDir, raw);
                    sysFilePath = raw;
                    break;
                }
            }

            // Strategy 3: Read the OpenType name table from each system font file
            if (sysValueName is null)
            {
                foreach (string valueName in sysKey.GetValueNames())
                {
                    string? raw = sysKey.GetValue(valueName) as string;
                    if (raw is null) continue;

                    string path = Path.IsPathRooted(raw) ? raw : Path.Combine(SystemFontsDir, raw);
                    if (FontFileProvidesFamilyName(path, fontFamily))
                    {
                        sysValueName = valueName;
                        sysFilePath = path;
                        break;
                    }
                }
            }

            if (sysValueName is null || sysFilePath is null)
                return (false, $"Font '{fontFamily}' not found in installed fonts.");

            // Unload from GDI session cache before removing
            NativeMethods.RemoveFontResourceW(sysFilePath);

            // Try direct uninstall (succeeds if process is running as admin).
            // OpenSubKey with writable:true throws SecurityException on most systems
            // when not elevated, rather than returning null.
            bool directSuccess = false;
            try
            {
                using var sysKeyWritable = Registry.LocalMachine.OpenSubKey(FontsRegistryKey, writable: true);
                if (sysKeyWritable is not null)
                {
                    sysKeyWritable.DeleteValue(sysValueName, throwOnMissingValue: false);
                    if (File.Exists(sysFilePath))
                    {
                        try { File.Delete(sysFilePath); }
                        catch (IOException) { /* file locked — cleaned up on reboot */ }
                    }
                    directSuccess = true;
                }
            }
            catch (System.Security.SecurityException) { /* expected when not admin */ }
            catch (UnauthorizedAccessException) { /* expected when not admin */ }

            if (directSuccess)
            {
                BroadcastFontChange();
                FontsChanged?.Invoke(this, EventArgs.Empty);
                return (true, $"System font '{fontFamily}' has been uninstalled.");
            }

            // Not admin — launch elevated PowerShell to remove registry entry and file
            if (UninstallSystemFontElevated(sysValueName, sysFilePath))
            {
                BroadcastFontChange();
                FontsChanged?.Invoke(this, EventArgs.Empty);
                return (true, $"System font '{fontFamily}' has been uninstalled.");
            }

            // Elevation failed or was cancelled — re-add the font resource
            NativeMethods.AddFontResourceW(sysFilePath);
            return (false, $"Failed to uninstall system font '{fontFamily}'. The operation was cancelled or requires administrator privileges.");
        }

        /// <summary>
        /// Launches an elevated PowerShell process to delete a system font's registry
        /// entry from HKLM and remove the font file from <c>C:\Windows\Fonts</c>.
        /// Shows a UAC consent dialog to the user.
        /// </summary>
        /// <param name="registryValueName">The HKLM registry value name (e.g. "Arial (TrueType)").</param>
        /// <param name="fontFilePath">The full path to the font file.</param>
        /// <returns><c>true</c> if the elevated process completed successfully.</returns>
        private static bool UninstallSystemFontElevated(string registryValueName, string fontFilePath)
        {
            try
            {
                string escapedName = registryValueName.Replace("'", "''");
                string escapedPath = fontFilePath.Replace("'", "''");

                string script =
                    $"Remove-ItemProperty -Path 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts' " +
                    $"-Name '{escapedName}' -Force -ErrorAction Stop; " +
                    $"if (Test-Path '{escapedPath}') {{ Remove-Item -Path '{escapedPath}' -Force -ErrorAction SilentlyContinue }}";

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command \"{script}\"",
                    Verb = "runas",
                    UseShellExecute = true,
                };

                using var proc = Process.Start(psi);
                if (proc is null) return false;
                proc.WaitForExit(15000);
                return proc.ExitCode == 0;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // User declined the UAC prompt
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks whether a font file on disk provides the specified font family name
        /// by reading the OpenType <c>name</c> table directly from the file.
        /// Handles both single fonts (TTF/OTF) and TrueType Collections (TTC).
        /// </summary>
        /// <param name="filePath">Full path to the font file.</param>
        /// <param name="familyName">The DirectWrite family name to look for.</param>
        /// <returns><c>true</c> if the file contains a font with the given family name.</returns>
        private static bool FontFileProvidesFamilyName(string filePath, string familyName)
        {
            try
            {
                if (!File.Exists(filePath)) return false;

                byte[] data = File.ReadAllBytes(filePath);
                if (data.Length < 12) return false;

                uint tag = FontReadBE32(data, 0);

                if (tag == 0x74746366) // 'ttcf' — TrueType Collection
                {
                    if (data.Length < 16) return false;
                    uint numFonts = FontReadBE32(data, 8);
                    for (uint i = 0; i < numFonts && i < 64; i++)
                    {
                        if (data.Length < 12 + (i + 1) * 4) break;
                        uint fontOffset = FontReadBE32(data, (int)(12 + i * 4));
                        if (CheckFontNameTable(data, fontOffset, familyName))
                            return true;
                    }
                    return false;
                }

                // Single font (TTF or OTF)
                return CheckFontNameTable(data, 0, familyName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Parses the OpenType <c>name</c> table at the given offset within a font file
        /// to check whether it contains the specified family name (Name ID 1).
        /// </summary>
        /// <param name="data">The raw font file bytes.</param>
        /// <param name="fontOffset">
        /// Byte offset to the start of this font's offset table (0 for single fonts,
        /// or the TTC directory entry offset for collection fonts).
        /// </param>
        /// <param name="familyName">The family name to match (case-insensitive).</param>
        /// <returns><c>true</c> if the name table contains a matching family name.</returns>
        private static bool CheckFontNameTable(byte[] data, uint fontOffset, string familyName)
        {
            if (data.Length < fontOffset + 12) return false;

            ushort numTables = FontReadBE16(data, (int)(fontOffset + 4));

            // Locate the 'name' table in the table directory
            uint nameTableOffset = 0;
            for (int i = 0; i < numTables; i++)
            {
                int entryOffset = (int)(fontOffset + 12 + i * 16);
                if (entryOffset + 16 > data.Length) break;

                uint tableTag = FontReadBE32(data, entryOffset);
                if (tableTag == 0x6E616D65) // 'name'
                {
                    nameTableOffset = FontReadBE32(data, entryOffset + 8);
                    break;
                }
            }

            if (nameTableOffset == 0 || data.Length < nameTableOffset + 6)
                return false;

            // Parse name table header
            ushort count = FontReadBE16(data, (int)(nameTableOffset + 2));
            ushort stringStorageOffset = FontReadBE16(data, (int)(nameTableOffset + 4));

            for (int i = 0; i < count; i++)
            {
                int recordOffset = (int)(nameTableOffset + 6 + i * 12);
                if (recordOffset + 12 > data.Length) break;

                ushort platformId = FontReadBE16(data, recordOffset);
                ushort encodingId = FontReadBE16(data, recordOffset + 2);
                ushort nameId = FontReadBE16(data, recordOffset + 6);
                ushort length = FontReadBE16(data, recordOffset + 8);
                ushort offset = FontReadBE16(data, recordOffset + 10);

                // Name ID 1 = Font Family; Platform 3 = Windows, Encoding 1 = Unicode BMP
                if (nameId == 1 && platformId == 3 && encodingId == 1)
                {
                    uint strPos = nameTableOffset + (uint)stringStorageOffset + offset;
                    if (strPos + length > data.Length) continue;

                    string name = Encoding.BigEndianUnicode.GetString(data, (int)strPos, length);
                    if (name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Reads a 32-bit big-endian unsigned integer from a byte array.
        /// Used for parsing OpenType font file headers and table directories.
        /// </summary>
        private static uint FontReadBE32(byte[] data, int offset) =>
            (uint)((data[offset] << 24) | (data[offset + 1] << 16) |
                   (data[offset + 2] << 8) | data[offset + 3]);

        /// <summary>
        /// Reads a 16-bit big-endian unsigned integer from a byte array.
        /// Used for parsing OpenType font file name table records.
        /// </summary>
        private static ushort FontReadBE16(byte[] data, int offset) =>
            (ushort)((data[offset] << 8) | data[offset + 1]);

        /// <summary>
        /// Win32 P/Invoke declarations for font installation and broadcasting.
        /// </summary>
        private static class NativeMethods
        {
            /// <summary>Registers a font file so it is available in the current session.</summary>
            [DllImport("gdi32", CharSet = CharSet.Unicode)]
            public static extern int AddFontResourceW(string lpFileName);

            /// <summary>Removes a previously registered font file from the current session.</summary>
            [DllImport("gdi32", CharSet = CharSet.Unicode)]
            public static extern int RemoveFontResourceW(string lpFileName);

            /// <summary>
            /// Sends a message to a window with a timeout, used to broadcast
            /// <c>WM_FONTCHANGE</c> to all top-level windows.
            /// </summary>
            [DllImport("user32", CharSet = CharSet.Unicode)]
            public static extern IntPtr SendMessageTimeoutW(
                IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
                uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
        }

        /// <summary>
        /// Releases filesystem watchers and the debounce timer.
        /// </summary>
        public void Dispose()
        {
            _debounceTimer?.Dispose();
            _perUserWatcher?.Dispose();
            _systemWatcher?.Dispose();
        }
    }
}
