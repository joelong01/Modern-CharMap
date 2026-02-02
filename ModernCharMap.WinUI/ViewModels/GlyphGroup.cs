using System.Collections.ObjectModel;

namespace ModernCharMap.WinUI.ViewModels
{
    /// <summary>
    /// A named group of <see cref="GlyphItem"/> instances for grouped display
    /// in a WinUI 3 <c>GridView</c> via <c>CollectionViewSource</c>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Inherits from <see cref="ObservableCollection{GlyphItem}"/> so the GridView
    /// can bind directly to the group's items. The <see cref="Name"/> property is
    /// bound to the group header template.
    /// </para>
    /// <para>
    /// In "By Block" sort mode, each group represents a Unicode block
    /// (e.g. "Basic Latin", "Emoticons"). In flat sort modes, a single group
    /// with a descriptive name (e.g. "All Glyphs") is used.
    /// </para>
    /// </remarks>
    public sealed class GlyphGroup : ObservableCollection<GlyphItem>
    {
        /// <summary>
        /// The display name for this group, shown as the group header in the GridView.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Creates a new glyph group with the specified display name.
        /// </summary>
        /// <param name="name">The group header text (e.g. a Unicode block name).</param>
        public GlyphGroup(string name) => Name = name;
    }
}
