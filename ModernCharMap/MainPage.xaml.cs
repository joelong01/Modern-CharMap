using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using SharpDX.DirectWrite;
using System.Globalization;
using Windows.UI.Popups;
using Windows.UI.Core;
using Windows.ApplicationModel.DataTransfer;
using Windows.System;



// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace ModernCharMap
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {


        List<FontFamilyItem> InstalledFonts { get; set; }
        List<FontFamilyItem> SearchableFonts { get; set; }

        public MainPage()
        {
            this.InitializeComponent();
            //  InitializeGrid();
            InstalledFonts = LoadInstalledFonts();
            foreach (var fontFamilyItem in InstalledFonts)
            {               
                _cmbFontFamily.Items.Add(fontFamilyItem.FontName);
                
            }
            SearchableFonts = LoadInstalledFonts();

            _searchBox.Text = SearchableFonts[0].FontName;
            

            _cmbFontFamily.SelectedIndex = 0;


            Window.Current.CoreWindow.KeyDown += CoreWindow_KeyDown;

        }

        public static Windows.UI.Xaml.Media.FontFamily SharpDxToXamlFontFamily(SharpDX.DirectWrite.FontFamily ff)
        {
            string familyName = ff.FamilyNames.GetString(0);
            return new Windows.UI.Xaml.Media.FontFamily(familyName);
        }

        private void CoreWindow_KeyDown(CoreWindow sender, KeyEventArgs args)
        {

            args.Handled = false;
            if (args.VirtualKey == VirtualKey.Enter)
            {
                LoadFont();
                args.Handled = true;
                return;
            }

            if (args.VirtualKey == VirtualKey.C)
            {
                if (Window.Current.CoreWindow.GetKeyState(VirtualKey.Control) == CoreVirtualKeyStates.Down)
                {
                    if (_toggle.IsOn)
                    {
                        CopyXamlFormatted();
                    }
                    else
                    {
                        CopyHexFormatted();
                    }
                }
                args.Handled = true;
                return;
            }


        }

        private void OnLoadFont(object sender, RoutedEventArgs e)
        {
            LoadFont();
        }

        private async void LoadFont()
        {
            try
            {
                int currentSelectedIndex = _cmbFontFamily.SelectedIndex;
                if (currentSelectedIndex == -1) return;
                SharpDX.DirectWrite.FontFamily fontFamily = InstalledFonts[currentSelectedIndex].FontFamily;
                Windows.UI.Xaml.Media.FontFamily windowsFF = new Windows.UI.Xaml.Media.FontFamily(InstalledFonts[currentSelectedIndex].FontName);

                List<FontItem> list = new List<FontItem>();

                Font font = fontFamily.GetFont(0);
                int fontSize = Convert.ToInt32(_txtDefaultFontSize.Text);

                for (int i = 0; i < 65535; i++)
                {
                    char c = Convert.ToChar(i);
                    if (c == ' ' || c == '\r' || c == '\n' || c == '\t')
                        continue;

                    if (font.HasCharacter(i))
                    {
                        list.Add(new FontItem()
                        {
                            Symbol = c.ToString(),
                            Label = "U+" + i.ToString("X4"),
                            FontFamily = windowsFF,
                            XamlEncoded = "&#x" + i.ToString("X4") + ";",
                            FontSize = fontSize
                        });
                    }
                }

                ResultList.ItemsSource = list;
                ResultList.ScrollIntoView(list[0]);
            }
            catch (Exception e)
            {
                string msg = e.Message;
                MessageDialog errDlg = new MessageDialog(e.Message);
                await errDlg.ShowAsync();
            }

        }




        private void OnCopy(object sender, RoutedEventArgs e)
        {
            CopyHexFormatted();
        }

        private void CopyHexFormatted()
        {
            DataPackage dp = new DataPackage();
            dp.SetText(_txtSelected.Text);
            Clipboard.SetContent(dp);
        }

        public List<FontFamilyItem> LoadInstalledFonts()
        {
            List<FontFamilyItem> retList = new List<FontFamilyItem>();

            using (var factory = new Factory())
            {

                using (var fontCollection = factory.GetSystemFontCollection(false))
                {
                    var familyCount = fontCollection.FontFamilyCount;
                    for (int i = 0; i < familyCount; i++)
                    {
                        try
                        {
                            var fontFamily = fontCollection.GetFontFamily(i);
                            {
                                var familyNames = fontFamily.FamilyNames;

                                int index;

                                if (!familyNames.FindLocaleName(CultureInfo.CurrentCulture.Name, out index))
                                    familyNames.FindLocaleName("en - us", out index);

                                string name = familyNames.GetString(index);
                                FontFamilyItem ffItem = new FontFamilyItem(name, fontFamily, fontFamily.GetFont(0).IsSymbolFont);
                                retList.Add(ffItem);
                            }
                        }
                        catch { }       // Corrupted font files throw an exception – ignore them
                    }
                }
            }

            retList.Sort(delegate (FontFamilyItem item1, FontFamilyItem item2)
            {
                return String.Compare(item1.FontName, item2.FontName, true);
            });

            return retList;
        }

        private void FontChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadFont();
        }

        private void OnListViewKeyDown(object sender, KeyRoutedEventArgs e)
        {

        }

        private void OnSelectFont(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                _txtSelected.Text = ((FontItem)e.AddedItems[0]).Label;
                _txtXaml.Text = ((FontItem)e.AddedItems[0]).XamlEncoded;
                _txtFont.Text = ((FontItem)e.AddedItems[0]).Symbol;
                _txtFont.FontFamily = ((FontItem)e.AddedItems[0]).FontFamily;
            }
            catch
            { }
        }

        private void OnCopyXaml(object sender, RoutedEventArgs e)
        {
            CopyXamlFormatted();
        }

        private void CopyXamlFormatted()
        {
            DataPackage dp = new DataPackage();
            dp.SetText(_txtXaml.Text);
            Clipboard.SetContent(dp);
        }

        private void OnFontSizeLostFocus(object sender, RoutedEventArgs e)
        {
            if (ResultList.Items.Count == 0) return;

            try
            {
                int newSize = Convert.ToInt32(_txtDefaultFontSize.Text);
                if (((FontItem)ResultList.Items[0]).FontSize != newSize)
                {
                    LoadFont();
                }
            }

            catch
            { }
        }

        private void OnCopyChar(object sender, RoutedEventArgs e)
        {
            

            //string s = String.Format($"<font> face = \"{_cmbFontFamily.SelectedItem.ToString()}\"\n</font>\n{_txtFont.Text}");
            string s = String.Format($"{_txtFont.Text}");
            string html = HtmlFormatHelper.CreateHtmlFormat(s);
            DataPackage dp = new DataPackage();                        
            dp.SetHtmlFormat(html);
            Clipboard.SetContent(dp);
        }

        private void Search_OnTextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
        {
            // Only get results when it was a user typing, 
            // otherwise assume the value got filled in by TextMemberPath 
            // or the handler for SuggestionChosen.
            if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            {
                //Set the ItemsSource to be your filtered dataset

                List<FontFamilyItem> filteredList = new List<FontFamilyItem>();
                foreach (var ffI in SearchableFonts)
                {
                    if (ffI.FontName.Contains(sender.Text))
                        filteredList.Add(ffI);
                }

                sender.ItemsSource = filteredList;
            }
        }

        private void Search_OnSuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
        {
            sender.Text = ((FontFamilyItem)args.SelectedItem).FontName;
            // Set sender.Text. You can use args.SelectedItem to build your text string.
        }

        private void Search_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
        {
            if (args.ChosenSuggestion != null)
            {
                // User selected an item from the suggestion list, take an action on it here.
            }
            else
            {
                // Use args.QueryText to determine what to do.
            }
        }
    }

    public class FontItem
    {
        public string Label { get; set; }
        public string Symbol { get; set; }
        public string XamlEncoded { get; set; }
        public int FontSize { get; set; } = 48;
        public Windows.UI.Xaml.Media.FontFamily FontFamily { get; set; }
        
    }

    public class FontFamilyItem
    {
        public FontFamilyItem(string fontName, SharpDX.DirectWrite.FontFamily fontFamily, bool isSymbol)
        {
            FontName = fontName;
            FontFamily = fontFamily;
            IsSymbol = isSymbol;
           
        }

        public string FontName { get; set; }

        public string DisplayFont
        {
            get
            {
                if (IsSymbol)
                {
                    return "Segoe UI";
                }
                else
                {
                    return FontName;
                }
            }
        }
        public SharpDX.DirectWrite.FontFamily FontFamily { get; set; }
        public bool IsSymbol { get; set; } = false;
        
    }
}
