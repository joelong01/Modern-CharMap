using System;
using System.IO;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Media.Imaging;
using ModernCharMap.WinUI.ViewModels;
using ModernCharMap.WinUI.Services;

namespace ModernCharMap.WinUI
{
    public sealed partial class MainWindow : Window
    {
        public CharMapViewModel ViewModel { get; } = new CharMapViewModel(
            FontService.Instance,
            ClipboardService.Instance);

        public CollectionViewSource GroupedGlyphsSource { get; } = new()
        {
            IsSourceGrouped = true
        };

        public MainWindow()
        {
            InitializeComponent();
            GroupedGlyphsSource.Source = ViewModel.GlyphGroups;
            ViewModel.Initialize(DispatcherQueue);
            ViewModel.ScrollToSelectedRequested += OnScrollToSelectedRequested;

            // Set title bar icon
            AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico"));

            LoadSplashImage();
            DismissSplashAfterDelay();
        }

        private void LoadSplashImage()
        {
            // Load splash from file system — ms-appx:/// with scale qualifiers
            // can crash unpackaged WinUI 3 apps with ExecutionEngineException.
            var path = Path.Combine(AppContext.BaseDirectory, "Assets", "SplashScreen.scale-200.png");
            if (File.Exists(path))
            {
                SplashImage.Source = new BitmapImage(new Uri(path));
            }
        }

        private async void DismissSplashAfterDelay()
        {
            await System.Threading.Tasks.Task.Delay(TimeSpan.FromSeconds(2));

            var fadeOut = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = new Duration(TimeSpan.FromMilliseconds(400)),
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
            };

            var fadeIn = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = new Duration(TimeSpan.FromMilliseconds(400)),
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };

            var storyboard = new Storyboard();
            Storyboard.SetTarget(fadeOut, SplashOverlay);
            Storyboard.SetTargetProperty(fadeOut, "Opacity");
            Storyboard.SetTarget(fadeIn, MainContent);
            Storyboard.SetTargetProperty(fadeIn, "Opacity");
            storyboard.Children.Add(fadeOut);
            storyboard.Children.Add(fadeIn);

            storyboard.Completed += (_, _) =>
            {
                SplashOverlay.Visibility = Visibility.Collapsed;
            };

            storyboard.Begin();
        }

        private void OnScrollToSelectedRequested(object? sender, EventArgs e)
        {
            if (ViewModel.SelectedGlyph is not null)
            {
                GlyphGridView.ScrollIntoView(ViewModel.SelectedGlyph, ScrollIntoViewAlignment.Leading);
            }
        }

        private void CodepointTextBox_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
        {
            if (e.Key == Windows.System.VirtualKey.Enter)
            {
                ViewModel.NavigateToCodepointCommand.Execute(null);
                e.Handled = true;
            }
        }

        private async void GlyphCard_RightTapped(object sender, Microsoft.UI.Xaml.Input.RightTappedRoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is ViewModels.GlyphItem glyph)
            {
                ViewModel.CopyGlyph(glyph);

                var flyout = new Microsoft.UI.Xaml.Controls.Flyout
                {
                    Content = new TextBlock
                    {
                        Text = "Copied!",
                        FontSize = 14,
                        Padding = new Thickness(4, 2, 4, 2)
                    },
                    Placement = Microsoft.UI.Xaml.Controls.Primitives.FlyoutPlacementMode.Top
                };
                flyout.ShowAt(element);

                await System.Threading.Tasks.Task.Delay(800);
                flyout.Hide();
            }
        }

        private async void InstallFont_Click(object sender, RoutedEventArgs e)
        {
            var picker = new Windows.Storage.Pickers.FileOpenPicker();
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

            picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.Desktop;
            picker.FileTypeFilter.Add(".ttf");
            picker.FileTypeFilter.Add(".otf");
            picker.FileTypeFilter.Add(".ttc");

            var file = await picker.PickSingleFileAsync();
            if (file is not null)
            {
                ViewModel.InstallFont(file.Path);
            }
        }

        private void FontSearch_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
        {
            if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            {
                ViewModel.UpdateSearchSuggestions(sender.Text);
            }
        }

        private void FontSearch_SuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
        {
            if (args.SelectedItem is string fontName)
            {
                sender.Text = fontName;
            }
        }

        private void FontSearch_QuerySubmitted(AutoSuggestBox sender, AutoSuggestBoxQuerySubmittedEventArgs args)
        {
            if (args.ChosenSuggestion is string fontName)
            {
                ViewModel.SelectFontFromSearch(fontName);
            }
            else if (!string.IsNullOrEmpty(args.QueryText))
            {
                ViewModel.SelectFontFromSearch(args.QueryText);
            }
        }
    }
}
