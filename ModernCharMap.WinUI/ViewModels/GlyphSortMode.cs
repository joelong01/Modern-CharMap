namespace ModernCharMap.WinUI.ViewModels
{
    /// <summary>
    /// Defines how glyphs are sorted and grouped in the character map grid.
    /// </summary>
    public enum GlyphSortMode
    {
        /// <summary>
        /// Group glyphs by their Unicode block (e.g. "Basic Latin", "Emoticons").
        /// Each block is a separate group header. Glyphs are in codepoint order within each block.
        /// </summary>
        ByBlock,

        /// <summary>
        /// Display all glyphs in a single flat list sorted by numeric codepoint value.
        /// </summary>
        ByCodepoint,

        /// <summary>
        /// Display all glyphs in a single flat list sorted alphabetically by glyph name.
        /// Glyphs without PostScript names are placed at the end, sorted by codepoint.
        /// </summary>
        ByName
    }

    /// <summary>
    /// Pairs a <see cref="GlyphSortMode"/> enum value with a human-readable display name
    /// for use in ComboBox data binding.
    /// </summary>
    /// <param name="Mode">The sort mode enum value.</param>
    /// <param name="DisplayName">The user-facing label (e.g. "By Block").</param>
    public sealed record SortModeOption(GlyphSortMode Mode, string DisplayName)
    {
        /// <summary>
        /// Returns the <see cref="DisplayName"/> so the ComboBox displays it correctly.
        /// </summary>
        public override string ToString() => DisplayName;
    }
}
