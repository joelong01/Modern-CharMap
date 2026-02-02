using System;
using System.Collections.Generic;

namespace ModernCharMap.WinUI.Services
{
    /// <summary>
    /// Abstracts font enumeration, codepoint discovery, glyph naming, and
    /// per-user font management operations.
    /// </summary>
    /// <remarks>
    /// The default implementation (<see cref="FontService"/>) uses DirectWrite
    /// COM interop for font enumeration and Win32 APIs for font installation.
    /// </remarks>
    public interface IFontService
    {
        /// <summary>
        /// Returns all installed font family names, sorted alphabetically.
        /// Includes both system-wide and per-user installed fonts.
        /// </summary>
        /// <returns>A read-only list of font family names (e.g. "Arial", "Segoe UI Emoji").</returns>
        IReadOnlyList<string> GetInstalledFontFamilies();

        /// <summary>
        /// Returns every Unicode codepoint that has a glyph in the specified font,
        /// spanning the full Unicode range (BMP and Supplementary Planes).
        /// Control characters, surrogates, and noncharacters are excluded.
        /// </summary>
        /// <param name="fontFamily">The font family name to query.</param>
        /// <returns>
        /// A read-only list of codepoints (as <c>int</c>) in ascending order,
        /// or an empty list if the font is not found.
        /// </returns>
        IReadOnlyList<int> GetSupportedCodePoints(string fontFamily);

        /// <summary>
        /// Determines whether the specified font uses a symbol encoding
        /// (e.g. Wingdings, Symbol) rather than standard Unicode mapping.
        /// </summary>
        /// <param name="fontFamily">The font family name to check.</param>
        /// <returns><c>true</c> if the font is a symbol font; otherwise <c>false</c>.</returns>
        bool IsSymbolFont(string fontFamily);

        /// <summary>
        /// Reads per-glyph PostScript names from the font's OpenType <c>post</c> table
        /// and maps them back to Unicode codepoints via the font's glyph index mapping.
        /// </summary>
        /// <param name="fontFamily">The font family name to query.</param>
        /// <returns>
        /// A dictionary mapping Unicode codepoints to human-readable glyph names.
        /// Returns an empty dictionary if the font has no meaningful glyph names
        /// (e.g. the <c>post</c> table is not format 2.0).
        /// </returns>
        IReadOnlyDictionary<int, string> GetGlyphNames(string fontFamily);

        /// <summary>
        /// Installs a font file as a per-user font (no administrator privileges required).
        /// The font is copied to the per-user fonts directory and registered in the
        /// current user's registry hive.
        /// </summary>
        /// <param name="fontFilePath">Full path to the <c>.ttf</c>, <c>.otf</c>, or <c>.ttc</c> file.</param>
        /// <returns><c>true</c> if the font was installed successfully; otherwise <c>false</c>.</returns>
        bool InstallFont(string fontFilePath);

        /// <summary>
        /// Uninstalls a per-user font by removing its registry entry and deleting the
        /// font file. System fonts (installed under HKLM) cannot be removed without
        /// administrator privileges.
        /// </summary>
        /// <param name="fontFamily">The font family name to uninstall.</param>
        /// <returns>
        /// A tuple containing a success flag and a human-readable status message.
        /// On failure, the message explains why (e.g. "system font", "not found").
        /// </returns>
        (bool Success, string Message) UninstallFont(string fontFamily);

        /// <summary>
        /// Invalidates all cached font data and re-enumerates fonts from the system.
        /// Call this after installing or removing fonts to pick up changes.
        /// </summary>
        void RefreshFontList();

        /// <summary>
        /// Raised when the available font list changes due to installation, removal,
        /// or external filesystem changes detected by the font directory watchers.
        /// </summary>
        /// <remarks>
        /// This event may fire on a background thread (from <see cref="System.IO.FileSystemWatcher"/>).
        /// UI subscribers must dispatch to the UI thread before updating controls.
        /// </remarks>
        event EventHandler? FontsChanged;
    }
}
