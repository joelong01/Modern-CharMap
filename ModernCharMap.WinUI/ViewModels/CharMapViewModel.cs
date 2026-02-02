using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Dispatching;
using ModernCharMap.WinUI.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

        [ObservableProperty]
        private GlyphItem? _selectedGlyph;

        [ObservableProperty]
        private string? _selectedFontFamily;

        [ObservableProperty]
        private string? _statusText;

        /// <summary>All font family names for the ComboBox.</summary>
        public ObservableCollection<string> FontFamilies { get; } = new();

        /// <summary>Filtered font names for the search AutoSuggestBox.</summary>
        public ObservableCollection<string> SearchSuggestions { get; } = new();

        /// <summary>Grouped glyph items for display.</summary>
        public ObservableCollection<GlyphGroup> GlyphGroups { get; } = new();

        /// <summary>Flat glyph list (for selection binding).</summary>
        public ObservableCollection<GlyphItem> Glyphs { get; } = new();

        public CharMapViewModel(IFontService fontService, IClipboardService clipboardService)
        {
            _fontService = fontService;
            _clipboardService = clipboardService;

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
        /// Filters the suggestion list (case-insensitive).
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
        /// Updates the ComboBox selection.
        /// </summary>
        public void SelectFontFromSearch(string fontName)
        {
            var match = _allFontFamilies.FirstOrDefault(
                f => f.Equals(fontName, StringComparison.OrdinalIgnoreCase));
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
            SelectedGlyph = null;

            if (string.IsNullOrEmpty(fontFamily))
            {
                StatusText = null;
                return;
            }

            var codePoints = _fontService.GetSupportedCodePoints(fontFamily);

            // Try to get glyph names from the font's post table
            IReadOnlyDictionary<int, string> glyphNames;
            try
            {
                glyphNames = _fontService.GetGlyphNames(fontFamily);
            }
            catch
            {
                glyphNames = new Dictionary<int, string>();
            }

            // Group glyphs by Unicode block
            var groups = new Dictionary<string, GlyphGroup>();

            foreach (var cp in codePoints)
            {
                glyphNames.TryGetValue(cp, out var name);
                var glyph = new GlyphItem(cp, fontFamily, name);

                if (!groups.TryGetValue(glyph.BlockName, out var group))
                {
                    group = new GlyphGroup(glyph.BlockName);
                    groups[glyph.BlockName] = group;
                }

                group.Add(glyph);
                Glyphs.Add(glyph);
            }

            // Add groups in the order they appear (by first codepoint)
            foreach (var group in groups.Values)
            {
                GlyphGroups.Add(group);
            }

            int namedCount = glyphNames.Count;
            string nameInfo = namedCount > 0 ? $"  ({namedCount} named)" : "";
            StatusText = $"{fontFamily}  \u2014  {Glyphs.Count} glyphs  \u2014  {groups.Count} blocks{nameInfo}";
        }

        [RelayCommand]
        private void Copy()
        {
            if (SelectedGlyph is null) return;
            _clipboardService.SetTextWithFont(
                SelectedGlyph.Display,
                SelectedGlyph.FontFamily);
        }
    }
}
