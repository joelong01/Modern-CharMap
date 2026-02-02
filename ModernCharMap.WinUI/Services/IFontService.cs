using System;
using System.Collections.Generic;

namespace ModernCharMap.WinUI.Services
{
    public interface IFontService
    {
        IReadOnlyList<string> GetInstalledFontFamilies();
        IReadOnlyList<int> GetSupportedCodePoints(string fontFamily);
        bool IsSymbolFont(string fontFamily);
        IReadOnlyDictionary<int, string> GetGlyphNames(string fontFamily);

        /// <summary>
        /// Installs a font file as a per-user font (no admin required).
        /// Copies to %LOCALAPPDATA%\Microsoft\Windows\Fonts and registers in HKCU.
        /// </summary>
        bool InstallFont(string fontFilePath);

        /// <summary>
        /// Uninstalls a per-user font. System fonts (HKLM) cannot be removed without admin.
        /// Returns false with an error message if the font is a system font.
        /// </summary>
        (bool Success, string Message) UninstallFont(string fontFamily);

        /// <summary>
        /// Clears cached font data and re-enumerates from the system.
        /// </summary>
        void RefreshFontList();

        /// <summary>
        /// Raised when the font list changes (install, uninstall, or external change detected).
        /// May fire on a background thread — callers must dispatch to UI.
        /// </summary>
        event EventHandler? FontsChanged;
    }
}
