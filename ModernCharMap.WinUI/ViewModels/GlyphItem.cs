using ModernCharMap.WinUI.Services;

namespace ModernCharMap.WinUI.ViewModels
{
    /// <summary>
    /// Represents a single Unicode glyph for display in the character map grid.
    /// Immutable after construction — one instance per codepoint per font.
    /// </summary>
    /// <remarks>
    /// Each glyph carries its codepoint, the font family it belongs to, an optional
    /// PostScript name from the font's <c>post</c> table, and a Unicode block name
    /// looked up from <see cref="UnicodeBlocks"/>.
    /// </remarks>
    public sealed class GlyphItem
    {
        /// <summary>
        /// The Unicode codepoint value (e.g. 0x1FA99 for the coin emoji).
        /// Supports the full Unicode range including Supplementary Planes.
        /// </summary>
        public int CodePoint { get; }

        /// <summary>
        /// The font family this glyph was loaded from (e.g. "Segoe UI Emoji").
        /// Used to render the glyph in the correct font in the UI.
        /// </summary>
        public string FontFamily { get; }

        /// <summary>
        /// The displayable string for this codepoint, obtained via
        /// <see cref="char.ConvertFromUtf32"/>. For BMP characters this is a
        /// single <c>char</c>; for SMP characters it is a surrogate pair.
        /// </summary>
        public string Display => char.ConvertFromUtf32(CodePoint);

        /// <summary>
        /// The codepoint in standard Unicode hex notation (e.g. "U+1FA99").
        /// </summary>
        public string CodePointHex => $"U+{CodePoint:X4}";

        /// <summary>
        /// The codepoint as a decimal integer string (e.g. "129689").
        /// </summary>
        public string DecimalString => CodePoint.ToString();

        /// <summary>
        /// The name of the Unicode block this codepoint belongs to
        /// (e.g. "Symbols and Pictographs Extended-A").
        /// </summary>
        public string BlockName { get; }

        /// <summary>
        /// The PostScript glyph name from the font's <c>post</c> table,
        /// or <c>null</c> if the font does not provide per-glyph names.
        /// </summary>
        public string? GlyphName { get; }

        /// <summary>
        /// Display label shown under the glyph: the glyph name if available,
        /// otherwise the hex codepoint notation.
        /// </summary>
        public string Label => GlyphName ?? CodePointHex;

        /// <summary>
        /// Creates a new glyph item for the specified codepoint and font.
        /// </summary>
        /// <param name="codePoint">The Unicode codepoint value.</param>
        /// <param name="fontFamily">The font family name this glyph belongs to.</param>
        /// <param name="glyphName">
        /// Optional PostScript glyph name. Pass <c>null</c> if the font does not
        /// provide names or the name is not meaningful.
        /// </param>
        public GlyphItem(int codePoint, string fontFamily, string? glyphName = null)
        {
            CodePoint = codePoint;
            FontFamily = fontFamily;
            GlyphName = glyphName;
            BlockName = UnicodeBlocks.GetBlockName(codePoint);
        }
    }
}
