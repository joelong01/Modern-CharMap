using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Dispatching;
using ModernCharMap.WinUI.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ModernCharMap.WinUI.ViewModels
{
    /// <summary>
    /// Main view model for the character map window. Manages font selection,
    /// glyph loading with sort/grouping, codepoint navigation, clipboard
    /// operations, and font install/uninstall commands.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Uses CommunityToolkit.Mvvm source generators: <c>[ObservableProperty]</c>
    /// generates property-changed notifications, and <c>[RelayCommand]</c> generates
    /// <see cref="IRelayCommand"/> properties for data binding.
    /// </para>
    /// <para>
    /// The glyph display pipeline:
    /// <list type="number">
    ///   <item><description>User selects a font → <see cref="OnSelectedFontFamilyChanged"/> fires.</description></item>
    ///   <item><description><see cref="LoadGlyphsForFont"/> queries the font service for codepoints and glyph names, caching them in <c>_loadedGlyphs</c>.</description></item>
    ///   <item><description><see cref="RebuildGlyphGroups"/> sorts and groups the cached glyphs according to the current <see cref="SortMode"/>.</description></item>
    ///   <item><description>The grouped results are placed into <see cref="GlyphGroups"/> (for the grouped GridView) and <see cref="Glyphs"/> (flat list for selection).</description></item>
    /// </list>
    /// </para>
    /// </remarks>
    public sealed partial class CharMapViewModel : ObservableObject
    {
        private readonly IFontService _fontService;
        private readonly IClipboardService _clipboardService;

        /// <summary>Complete list of font family names from the font service.</summary>
        private IReadOnlyList<string> _allFontFamilies = Array.Empty<string>();

        /// <summary>UI thread dispatcher for marshalling font-change events.</summary>
        private DispatcherQueue? _dispatcherQueue;

        /// <summary>
        /// Raw glyph items for the currently selected font, cached between sort mode changes.
        /// Populated by <see cref="LoadGlyphsForFont"/> and consumed by <see cref="RebuildGlyphGroups"/>.
        /// </summary>
        private List<GlyphItem> _loadedGlyphs = new();

        /// <summary>The currently selected glyph in the grid.</summary>
        [ObservableProperty]
        private GlyphItem? _selectedGlyph;

        /// <summary>The currently selected font family name (bound to the ComboBox).</summary>
        [ObservableProperty]
        private string? _selectedFontFamily;

        /// <summary>Status bar text displayed at the bottom of the window.</summary>
        [ObservableProperty]
        private string? _statusText;

        /// <summary>User input for codepoint navigation (e.g. "1FA99", "U+0041").</summary>
        [ObservableProperty]
        private string? _codepointInput;

        /// <summary>The active sort/grouping mode for the glyph grid.</summary>
        [ObservableProperty]
        private GlyphSortMode _sortMode = GlyphSortMode.ByBlock;

        /// <summary>The currently selected sort option in the sort ComboBox.</summary>
        [ObservableProperty]
        private SortModeOption _selectedSortOption;

        /// <summary>
        /// All font family names for the font selection ComboBox.
        /// </summary>
        public ObservableCollection<string> FontFamilies { get; } = new();

        /// <summary>
        /// Filtered font names for the search AutoSuggestBox dropdown.
        /// Updated as the user types via <see cref="UpdateSearchSuggestions"/>.
        /// </summary>
        public ObservableCollection<string> SearchSuggestions { get; } = new();

        /// <summary>
        /// Grouped glyph items for the GridView, bound via a <c>CollectionViewSource</c>
        /// with <c>IsSourceGrouped = true</c>.
        /// </summary>
        /// <remarks>
        /// In "By Block" mode, each group represents a Unicode block.
        /// In flat sort modes (By Codepoint, By Name), a single group is used
        /// to avoid toggling <c>IsSourceGrouped</c> at runtime (which is fragile in WinUI 3).
        /// </remarks>
        public ObservableCollection<GlyphGroup> GlyphGroups { get; } = new();

        /// <summary>
        /// Flat list of all glyphs for the current font, used for selection binding
        /// and codepoint lookup.
        /// </summary>
        public ObservableCollection<GlyphItem> Glyphs { get; } = new();

        /// <summary>
        /// Display-friendly sort mode options for the sort ComboBox.
        /// Each option's <see cref="SortModeOption.ToString"/> returns the display name.
        /// </summary>
        public IReadOnlyList<SortModeOption> SortModeOptions { get; } = new[]
        {
            new SortModeOption(GlyphSortMode.ByBlock, "By Block"),
            new SortModeOption(GlyphSortMode.ByCodepoint, "By Codepoint"),
            new SortModeOption(GlyphSortMode.ByName, "By Name"),
        };

        /// <summary>
        /// Raised when the view should scroll to make <see cref="SelectedGlyph"/> visible.
        /// Subscribed by MainWindow to call <c>GridView.ScrollIntoView</c>.
        /// </summary>
        public event EventHandler? ScrollToSelectedRequested;

        /// <summary>
        /// Initializes the view model with the given services and loads the font family list.
        /// </summary>
        /// <param name="fontService">Service for font enumeration and glyph data.</param>
        /// <param name="clipboardService">Service for clipboard copy operations.</param>
        public CharMapViewModel(IFontService fontService, IClipboardService clipboardService)
        {
            _fontService = fontService;
            _clipboardService = clipboardService;
            _selectedSortOption = SortModeOptions[0];

            LoadFontFamilies();
        }

        /// <summary>
        /// Completes initialization by wiring up the dispatcher queue (for UI thread
        /// marshalling) and subscribing to font change notifications.
        /// Must be called from MainWindow after construction.
        /// </summary>
        /// <param name="dispatcherQueue">The UI thread's dispatcher queue.</param>
        public void Initialize(DispatcherQueue dispatcherQueue)
        {
            _dispatcherQueue = dispatcherQueue;
            _fontService.FontsChanged += OnFontsChanged;
        }

        /// <summary>
        /// Handles <see cref="IFontService.FontsChanged"/> by dispatching a font list
        /// reload to the UI thread. The event may fire on a background thread from
        /// <see cref="System.IO.FileSystemWatcher"/>.
        /// </summary>
        private void OnFontsChanged(object? sender, EventArgs e)
        {
            if (_dispatcherQueue is not null)
                _dispatcherQueue.TryEnqueue(ReloadAfterFontChange);
            else
                ReloadAfterFontChange();
        }

        /// <summary>
        /// Refreshes the font list from the system, repopulates the ComboBox,
        /// and restores the previous selection if the font still exists.
        /// </summary>
        private void ReloadAfterFontChange()
        {
            _fontService.RefreshFontList();

            var currentFont = SelectedFontFamily;
            _allFontFamilies = _fontService.GetInstalledFontFamilies();

            FontFamilies.Clear();
            foreach (var name in _allFontFamilies)
                FontFamilies.Add(name);

            if (currentFont is not null && FontFamilies.Contains(currentFont))
                SelectedFontFamily = currentFont;
            else if (FontFamilies.Count > 0)
                SelectedFontFamily = FontFamilies[0];
        }

        /// <summary>
        /// Installs a font file via the font service and updates the status bar
        /// with the result.
        /// </summary>
        /// <param name="filePath">Full path to the font file selected by the user.</param>
        public void InstallFont(string filePath)
        {
            bool success = _fontService.InstallFont(filePath);
            if (success)
                StatusText = $"Installed: {Path.GetFileNameWithoutExtension(filePath)}";
            else
                StatusText = $"Failed to install {Path.GetFileName(filePath)}";
        }

        /// <summary>
        /// Uninstalls the currently selected per-user font and updates the status bar.
        /// </summary>
        [RelayCommand]
        private void UninstallSelectedFont()
        {
            if (string.IsNullOrEmpty(SelectedFontFamily)) return;

            var (success, message) = _fontService.UninstallFont(SelectedFontFamily);
            StatusText = message;
        }

        /// <summary>
        /// Parses the codepoint input, finds the matching glyph in the current font,
        /// selects it, and requests the view to scroll it into view.
        /// </summary>
        /// <remarks>
        /// Accepts hex input in several formats:
        /// <list type="bullet">
        ///   <item><description><c>1FA99</c> — raw hex</description></item>
        ///   <item><description><c>U+1FA99</c> — standard Unicode notation</description></item>
        ///   <item><description><c>0x1FA99</c> — programmer-style hex prefix</description></item>
        /// </list>
        /// Updates <see cref="StatusText"/> with the result or an error message.
        /// </remarks>
        [RelayCommand]
        private void NavigateToCodepoint()
        {
            if (string.IsNullOrWhiteSpace(CodepointInput))
                return;

            // Strip common hex prefixes
            string hex = CodepointInput.Trim();
            if (hex.StartsWith("U+", StringComparison.OrdinalIgnoreCase))
                hex = hex[2..];
            else if (hex.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                hex = hex[2..];

            if (!int.TryParse(hex, NumberStyles.HexNumber, null, out int codePoint))
            {
                StatusText = $"Invalid codepoint: \"{CodepointInput}\"";
                return;
            }

            var match = Glyphs.FirstOrDefault(g => g.CodePoint == codePoint);
            if (match is null)
            {
                StatusText = $"U+{codePoint:X4} not found in {SelectedFontFamily}";
                return;
            }

            SelectedGlyph = match;
            StatusText = $"U+{codePoint:X4}  \u2014  {match.Label}";
            ScrollToSelectedRequested?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Syncs the <see cref="SortMode"/> enum when the user changes the ComboBox selection.
        /// Generated by CommunityToolkit.Mvvm for the <see cref="SelectedSortOption"/> property.
        /// </summary>
        partial void OnSelectedSortOptionChanged(SortModeOption value)
        {
            SortMode = value.Mode;
        }

        /// <summary>
        /// Rebuilds the glyph groups when the sort mode changes (if glyphs are loaded).
        /// Generated by CommunityToolkit.Mvvm for the <see cref="SortMode"/> property.
        /// </summary>
        partial void OnSortModeChanged(GlyphSortMode value)
        {
            if (_loadedGlyphs.Count > 0)
                RebuildGlyphGroups();
        }

        /// <summary>
        /// Queries the font service for all installed font families, populates the
        /// <see cref="FontFamilies"/> collection, and selects the first font.
        /// </summary>
        private void LoadFontFamilies()
        {
            _allFontFamilies = _fontService.GetInstalledFontFamilies();
            FontFamilies.Clear();
            foreach (var name in _allFontFamilies)
                FontFamilies.Add(name);

            if (_allFontFamilies.Count > 0)
            {
                SelectedFontFamily = _allFontFamilies[0];
            }
        }

        /// <summary>
        /// Filters the font family list for the search AutoSuggestBox.
        /// Shows all fonts if the text is empty, otherwise filters by case-insensitive substring match.
        /// </summary>
        /// <param name="text">The current text in the search box.</param>
        public void UpdateSearchSuggestions(string text)
        {
            SearchSuggestions.Clear();
            foreach (var name in _allFontFamilies)
            {
                if (string.IsNullOrEmpty(text) ||
                    name.Contains(text, StringComparison.OrdinalIgnoreCase))
                {
                    SearchSuggestions.Add(name);
                }
            }
        }

        /// <summary>
        /// Selects a font from the search box result. Tries an exact match first,
        /// then falls back to the first suggestion in the filtered list.
        /// </summary>
        /// <param name="fontName">The font name submitted by the user or selected from suggestions.</param>
        public void SelectFontFromSearch(string fontName)
        {
            var match = _allFontFamilies.FirstOrDefault(
                f => f.Equals(fontName, StringComparison.OrdinalIgnoreCase));
            if (match is null && SearchSuggestions.Count > 0)
                match = SearchSuggestions[0];
            if (match is not null)
            {
                SelectedFontFamily = match;
            }
        }

        /// <summary>
        /// Triggers glyph loading when the selected font family changes.
        /// Generated by CommunityToolkit.Mvvm for the <see cref="SelectedFontFamily"/> property.
        /// </summary>
        partial void OnSelectedFontFamilyChanged(string? value)
        {
            LoadGlyphsForFont(value);
        }

        /// <summary>
        /// Queries the font service for all supported codepoints and glyph names
        /// for the specified font, caches them in <see cref="_loadedGlyphs"/>,
        /// and calls <see cref="RebuildGlyphGroups"/> to populate the display collections.
        /// </summary>
        /// <param name="fontFamily">The font family to load, or <c>null</c> to clear.</param>
        private void LoadGlyphsForFont(string? fontFamily)
        {
            GlyphGroups.Clear();
            Glyphs.Clear();
            _loadedGlyphs.Clear();
            SelectedGlyph = null;

            if (string.IsNullOrEmpty(fontFamily))
            {
                StatusText = null;
                return;
            }

            var codePoints = _fontService.GetSupportedCodePoints(fontFamily);

            IReadOnlyDictionary<int, string> glyphNames;
            try
            {
                glyphNames = _fontService.GetGlyphNames(fontFamily);
            }
            catch
            {
                glyphNames = new Dictionary<int, string>();
            }

            foreach (var cp in codePoints)
            {
                glyphNames.TryGetValue(cp, out var name);
                _loadedGlyphs.Add(new GlyphItem(cp, fontFamily, name));
            }

            RebuildGlyphGroups();
        }

        /// <summary>
        /// Sorts and groups the cached <see cref="_loadedGlyphs"/> according to the
        /// current <see cref="SortMode"/>, populating <see cref="GlyphGroups"/> and
        /// <see cref="Glyphs"/>. Preserves the current selection across rebuilds.
        /// </summary>
        /// <remarks>
        /// <list type="table">
        ///   <listheader><term>Sort Mode</term><description>Behavior</description></listheader>
        ///   <item><term>ByBlock</term><description>Groups by Unicode block name, preserving codepoint order within each block.</description></item>
        ///   <item><term>ByCodepoint</term><description>Single flat group sorted by numeric codepoint value.</description></item>
        ///   <item><term>ByName</term><description>Single flat group sorted alphabetically by glyph name; unnamed glyphs appear at the end, sorted by codepoint.</description></item>
        /// </list>
        /// Flat sort modes use a single <see cref="GlyphGroup"/> to keep
        /// <c>CollectionViewSource.IsSourceGrouped</c> always <c>true</c>,
        /// which avoids runtime toggling issues in WinUI 3.
        /// </remarks>
        private void RebuildGlyphGroups()
        {
            var previousCodePoint = SelectedGlyph?.CodePoint;

            GlyphGroups.Clear();
            Glyphs.Clear();

            IEnumerable<GlyphItem> sorted = SortMode switch
            {
                GlyphSortMode.ByCodepoint => _loadedGlyphs.OrderBy(g => g.CodePoint),
                GlyphSortMode.ByName => _loadedGlyphs
                    .OrderBy(g => g.GlyphName is null ? 1 : 0)
                    .ThenBy(g => g.GlyphName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(g => g.CodePoint),
                _ => _loadedGlyphs, // ByBlock: already in codepoint order
            };

            switch (SortMode)
            {
                case GlyphSortMode.ByBlock:
                    {
                        var groups = new Dictionary<string, GlyphGroup>();
                        foreach (var glyph in sorted)
                        {
                            if (!groups.TryGetValue(glyph.BlockName, out var group))
                            {
                                group = new GlyphGroup(glyph.BlockName);
                                groups[glyph.BlockName] = group;
                            }
                            group.Add(glyph);
                            Glyphs.Add(glyph);
                        }
                        foreach (var group in groups.Values)
                            GlyphGroups.Add(group);
                        break;
                    }

                case GlyphSortMode.ByCodepoint:
                    {
                        var group = new GlyphGroup("All Glyphs");
                        foreach (var glyph in sorted)
                        {
                            group.Add(glyph);
                            Glyphs.Add(glyph);
                        }
                        GlyphGroups.Add(group);
                        break;
                    }

                case GlyphSortMode.ByName:
                    {
                        var group = new GlyphGroup("All Glyphs (by name)");
                        foreach (var glyph in sorted)
                        {
                            group.Add(glyph);
                            Glyphs.Add(glyph);
                        }
                        GlyphGroups.Add(group);
                        break;
                    }
            }

            // Restore the previously selected glyph after rebuild
            if (previousCodePoint.HasValue)
                SelectedGlyph = Glyphs.FirstOrDefault(g => g.CodePoint == previousCodePoint.Value);

            UpdateStatusText();
        }

        /// <summary>
        /// Updates the status bar with a summary of the current font, glyph count,
        /// sort mode, and named glyph count.
        /// </summary>
        private void UpdateStatusText()
        {
            if (string.IsNullOrEmpty(SelectedFontFamily))
            {
                StatusText = null;
                return;
            }

            string sortInfo = SortMode switch
            {
                GlyphSortMode.ByCodepoint => "sorted by codepoint",
                GlyphSortMode.ByName => "sorted by name",
                _ => $"{GlyphGroups.Count} blocks"
            };

            int namedCount = _loadedGlyphs.Count(g => g.GlyphName is not null);
            string nameInfo = namedCount > 0 ? $"  ({namedCount} named)" : "";
            StatusText = $"{SelectedFontFamily}  \u2014  {Glyphs.Count} glyphs  \u2014  {sortInfo}{nameInfo}";
        }

        /// <summary>
        /// Copies a glyph's display character to the clipboard with font-family metadata.
        /// Used by both the Copy button command and right-click context action.
        /// </summary>
        /// <param name="glyph">The glyph item to copy.</param>
        public void CopyGlyph(GlyphItem glyph)
        {
            _clipboardService.SetTextWithFont(glyph.Display, glyph.FontFamily);
        }

        /// <summary>
        /// Copies the currently selected glyph to the clipboard.
        /// Bound to the toolbar Copy button.
        /// </summary>
        [RelayCommand]
        private void Copy()
        {
            if (SelectedGlyph is null) return;
            CopyGlyph(SelectedGlyph);
        }
    }
}
