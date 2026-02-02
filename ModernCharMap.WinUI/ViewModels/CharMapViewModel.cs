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
    public sealed partial class CharMapViewModel : ObservableObject
    {
        private readonly IFontService _fontService;
        private readonly IClipboardService _clipboardService;
        private IReadOnlyList<string> _allFontFamilies = Array.Empty<string>();
        private DispatcherQueue? _dispatcherQueue;
        private List<GlyphItem> _loadedGlyphs = new();

        [ObservableProperty]
        private GlyphItem? _selectedGlyph;

        [ObservableProperty]
        private string? _selectedFontFamily;

        [ObservableProperty]
        private string? _statusText;

        [ObservableProperty]
        private string? _codepointInput;

        [ObservableProperty]
        private GlyphSortMode _sortMode = GlyphSortMode.ByBlock;

        [ObservableProperty]
        private SortModeOption _selectedSortOption;

        /// <summary>All font family names for the ComboBox.</summary>
        public ObservableCollection<string> FontFamilies { get; } = new();

        /// <summary>Filtered font names for the search AutoSuggestBox.</summary>
        public ObservableCollection<string> SearchSuggestions { get; } = new();

        /// <summary>Grouped glyph items for display.</summary>
        public ObservableCollection<GlyphGroup> GlyphGroups { get; } = new();

        /// <summary>Flat glyph list (for selection binding).</summary>
        public ObservableCollection<GlyphItem> Glyphs { get; } = new();

        /// <summary>Display-friendly sort mode options for the ComboBox.</summary>
        public IReadOnlyList<SortModeOption> SortModeOptions { get; } = new[]
        {
            new SortModeOption(GlyphSortMode.ByBlock, "By Block"),
            new SortModeOption(GlyphSortMode.ByCodepoint, "By Codepoint"),
            new SortModeOption(GlyphSortMode.ByName, "By Name"),
        };

        /// <summary>
        /// Raised when the view should scroll to make the SelectedGlyph visible.
        /// </summary>
        public event EventHandler? ScrollToSelectedRequested;

        public CharMapViewModel(IFontService fontService, IClipboardService clipboardService)
        {
            _fontService = fontService;
            _clipboardService = clipboardService;
            _selectedSortOption = SortModeOptions[0];

            LoadFontFamilies();
        }

        /// <summary>
        /// Called by MainWindow after construction to wire up dispatcher and font change events.
        /// </summary>
        public void Initialize(DispatcherQueue dispatcherQueue)
        {
            _dispatcherQueue = dispatcherQueue;
            _fontService.FontsChanged += OnFontsChanged;
        }

        private void OnFontsChanged(object? sender, EventArgs e)
        {
            if (_dispatcherQueue is not null)
                _dispatcherQueue.TryEnqueue(ReloadAfterFontChange);
            else
                ReloadAfterFontChange();
        }

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
        /// Called by MainWindow after file picker returns a path.
        /// </summary>
        public void InstallFont(string filePath)
        {
            bool success = _fontService.InstallFont(filePath);
            if (success)
                StatusText = $"Installed: {Path.GetFileNameWithoutExtension(filePath)}";
            else
                StatusText = $"Failed to install {Path.GetFileName(filePath)}";
        }

        [RelayCommand]
        private void UninstallSelectedFont()
        {
            if (string.IsNullOrEmpty(SelectedFontFamily)) return;

            var (success, message) = _fontService.UninstallFont(SelectedFontFamily);
            StatusText = message;
        }

        [RelayCommand]
        private void NavigateToCodepoint()
        {
            if (string.IsNullOrWhiteSpace(CodepointInput))
                return;

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

        partial void OnSelectedSortOptionChanged(SortModeOption value)
        {
            SortMode = value.Mode;
        }

        partial void OnSortModeChanged(GlyphSortMode value)
        {
            if (_loadedGlyphs.Count > 0)
                RebuildGlyphGroups();
        }

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
        /// Called by the AutoSuggestBox when the user types.
        /// </summary>
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
        /// Called when the user picks a font from the search box.
        /// </summary>
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

        partial void OnSelectedFontFamilyChanged(string? value)
        {
            LoadGlyphsForFont(value);
        }

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

            // Restore selection
            if (previousCodePoint.HasValue)
                SelectedGlyph = Glyphs.FirstOrDefault(g => g.CodePoint == previousCodePoint.Value);

            UpdateStatusText();
        }

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

        public void CopyGlyph(GlyphItem glyph)
        {
            _clipboardService.SetTextWithFont(glyph.Display, glyph.FontFamily);
        }

        [RelayCommand]
        private void Copy()
        {
            if (SelectedGlyph is null) return;
            CopyGlyph(SelectedGlyph);
        }
    }
}
