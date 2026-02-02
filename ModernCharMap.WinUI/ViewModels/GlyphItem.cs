using ModernCharMap.WinUI.Services;

namespace ModernCharMap.WinUI.ViewModels
{
    public sealed class GlyphItem
    {
        public int CodePoint { get; }
        public string FontFamily { get; }
        public string Display => char.ConvertFromUtf32(CodePoint);
        public string CodePointHex => $"U+{CodePoint:X4}";
        public string DecimalString => CodePoint.ToString();
        public string BlockName { get; }
        public string? GlyphName { get; }

        /// <summary>
        /// Display label: glyph name if available, otherwise the hex codepoint.
        /// </summary>
        public string Label => GlyphName ?? CodePointHex;

        public GlyphItem(int codePoint, string fontFamily, string? glyphName = null)
        {
            CodePoint = codePoint;
            FontFamily = fontFamily;
            GlyphName = glyphName;
            BlockName = UnicodeBlocks.GetBlockName(codePoint);
        }
    }
}
