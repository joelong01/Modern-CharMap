namespace ModernCharMap.WinUI.Services
{
    public interface IClipboardService
    {
        void SetText(string text);
        void SetTextWithFont(string text, string fontFamily);
        string? GetText();
    }
}