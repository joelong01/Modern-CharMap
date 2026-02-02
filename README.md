# Modern CharMap

A Windows character map application built with WinUI 3 and .NET 9. Browse, search, and copy Unicode glyphs from any installed font — with per-user font install/uninstall and automatic font change detection.

![Dark themed UI with grouped glyph grid](ModernCharMap.WinUI/Assets/Square150x150Logo.scale-200.png)

## Features

- **Font browsing** — Dropdown lists all installed fonts (system + per-user), enumerated via DirectWrite
- **Font search** — Type-ahead AutoSuggestBox filters fonts by name
- **Grouped glyph display** — Characters grouped by Unicode block (Basic Latin, Greek, CJK, Emoji, etc.) with category headers
- **Glyph metadata** — Shows PostScript glyph names (parsed from OpenType `post` table) when available
- **Copy with font** — Copies the selected character to clipboard with CF_HTML formatting so it pastes into Word/Outlook with the correct font
- **Install fonts** — Browse for `.ttf`/`.otf`/`.ttc` files and install them per-user (no admin required)
- **Uninstall fonts** — Remove per-user fonts; system fonts show a message explaining admin is required
- **Auto-refresh** — FileSystemWatcher monitors both system and per-user font directories; when any application installs or removes a font, the list updates automatically
- **Dark mode** — Full dark theme
- **DPI aware** — PerMonitorV2 for crisp text on mixed-DPI multi-monitor setups

## Requirements

- Windows 10 version 1903 (build 19041) or later
- .NET 9 SDK (for building)
- Visual Studio 2022+ with the **.NET Desktop Development** workload (for building)

The build script can install missing prerequisites automatically.

## Building

All build operations go through `build.ps1`, which runs a cascading pipeline: **doctor** checks prerequisites, **install** fixes anything missing, **build** formats and compiles, **deploy** creates a self-contained install.

```powershell
# Check environment (no changes)
./build.ps1 doctor

# Install missing prerequisites
./build.ps1 install

# Build (default — runs doctor + install + format + compile)
./build.ps1

# Build and deploy to %LOCALAPPDATA%\ModernCharMap with Start Menu shortcut
./build.ps1 deploy
```

Optional parameters:

```powershell
./build.ps1 build -Platform ARM64
./build.ps1 build -Configuration Debug
```

Supported platforms: `x64` (default), `x86`, `ARM64`.

## Deploy

```powershell
./build.ps1 deploy
```

This will:

1. Build a self-contained executable (no .NET runtime required on target)
2. Stop any running instance
3. Copy everything to `%LOCALAPPDATA%\ModernCharMap\`
4. Create a Start Menu shortcut

Launch from Start Menu or run directly:

```
%LOCALAPPDATA%\ModernCharMap\ModernCharMap.WinUI.exe
```

## Usage

### Browsing Fonts

Select a font from the **Font** dropdown. The glyph grid populates with every character the font supports, grouped by Unicode block. Each card shows:

- The rendered character
- Its name (from the font's PostScript glyph names, if available) or hex codepoint
- The hex codepoint (U+XXXX)
- The decimal value

### Searching Fonts

Type in the **Search** box to filter fonts by name. Select a suggestion to switch to that font.

### Copying Characters

Click a glyph card to select it, then click **Copy**. The character is placed on the clipboard with HTML formatting that preserves the font family, so pasting into Word or Outlook renders in the correct font.

### Installing Fonts

Click **Install Font** to browse for a `.ttf`, `.otf`, or `.ttc` file. The font is installed per-user:

- Copied to `%LOCALAPPDATA%\Microsoft\Windows\Fonts\`
- Registered in `HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`
- Made immediately available via `AddFontResourceW` + `WM_FONTCHANGE` broadcast

No administrator privileges are needed.

### Uninstalling Fonts

Select a font in the dropdown and click **Uninstall**. This works for per-user fonts only. System fonts (installed in `C:\Windows\Fonts` with HKLM registry entries) require administrator privileges and cannot be removed from this app.

### Auto-Refresh

The app watches both font directories for changes. If you install or remove a font using any other application (e.g., right-click a `.ttf` > Install), the font list updates automatically within about 500ms.

## Architecture

```
ModernCharMap.WinUI/
  App.xaml(.cs)              Application entry, dark theme, global exception handling
  MainWindow.xaml(.cs)       UI layout, file picker, splash screen
  app.manifest               PerMonitorV2 DPI awareness
  Services/
    IFontService.cs          Font management interface
    FontService.cs           DirectWrite enumeration, GDI ranges, font install/uninstall,
                             FileSystemWatcher, OpenType post table parsing
    IClipboardService.cs     Clipboard interface
    ClipboardService.cs      CF_HTML clipboard for font-aware copy
    UnicodeBlocks.cs         Static Unicode block name lookup (binary search)
    DirectWrite/
      DWriteInterop.cs       COM interop for IDWriteFactory, IDWriteFontCollection,
                             IDWriteFont, IDWriteFontFace, IDWriteLocalizedStrings
  ViewModels/
    CharMapViewModel.cs      Font selection, glyph loading, commands (MVVM Toolkit)
    GlyphItem.cs             Single glyph: codepoint, display, name, block
    GlyphGroup.cs            Named group of glyphs (ObservableCollection)
  Assets/
    app.ico                  Multi-size application icon
    SplashScreen.scale-*.png Splash screens for various DPI scales
```

### Key Technologies

| Component | Technology |
|---|---|
| UI Framework | WinUI 3 (Windows App SDK 1.6) |
| Runtime | .NET 9 |
| MVVM | CommunityToolkit.Mvvm 8.3 |
| Font enumeration | DirectWrite COM interop (P/Invoke) |
| Character ranges | GDI `GetFontUnicodeRanges` (P/Invoke) |
| Glyph names | OpenType `post` table binary parsing |
| Font install/uninstall | Registry (HKCU) + `AddFontResourceW`/`RemoveFontResourceW` |
| Font change broadcast | `SendMessageTimeoutW` with `WM_FONTCHANGE` |
| Auto-detection | `FileSystemWatcher` with 500ms debounce |
| Clipboard | Win32 clipboard API with CF_HTML format |
| Packaging | Unpackaged (`WindowsPackageType=None`) |

### Error Handling

Unhandled exceptions are caught by three handlers (WinUI, AppDomain, TaskScheduler). A crash log is written to `%LOCALAPPDATA%\ModernCharMap\Logs\crash-<timestamp>.log` and a MessageBox displays the error with the log path. Press Ctrl+C on the MessageBox to copy its text.

## License

See [LICENSE](LICENSE) for details.
