namespace ModernCharMap.WinUI.ViewModels
{
    public enum GlyphSortMode
    {
        ByBlock,
        ByCodepoint,
        ByName
    }

    public sealed record SortModeOption(GlyphSortMode Mode, string DisplayName)
    {
        public override string ToString() => DisplayName;
    }
}
