using System.Collections.ObjectModel;

namespace ModernCharMap.WinUI.ViewModels
{
    /// <summary>
    /// A named group of glyph items for grouped display in GridView.
    /// </summary>
    public sealed class GlyphGroup : ObservableCollection<GlyphItem>
    {
        public string Name { get; }

        public GlyphGroup(string name) => Name = name;
    }
}
