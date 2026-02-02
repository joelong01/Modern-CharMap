namespace ModernCharMap.WinUI.Services
{
    /// <summary>
    /// Abstracts clipboard operations for copying text and font-aware rich content.
    /// </summary>
    public interface IClipboardService
    {
        /// <summary>
        /// Copies plain Unicode text to the clipboard.
        /// </summary>
        /// <param name="text">The text to copy.</param>
        void SetText(string text);

        /// <summary>
        /// Copies text to the clipboard in both plain Unicode and CF_HTML formats.
        /// The HTML format includes a <c>font-family</c> style attribute so that
        /// rich-text editors (Word, Outlook, etc.) preserve the intended font
        /// when pasting.
        /// </summary>
        /// <param name="text">The text (glyph character) to copy.</param>
        /// <param name="fontFamily">
        /// The font family name to embed in the HTML markup, or <c>null</c>
        /// to copy plain text only.
        /// </param>
        void SetTextWithFont(string text, string fontFamily);

        /// <summary>
        /// Retrieves plain text from the clipboard.
        /// </summary>
        /// <returns>The clipboard text, or <c>null</c> if no text is available.</returns>
        string? GetText();
    }
}
