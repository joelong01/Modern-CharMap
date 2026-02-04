using System;
using System.ComponentModel;
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
    /// <summary>
    /// Code-behind for the main application window. Wires up the view model,
    /// handles UI events that cannot be expressed purely in XAML data binding
    /// (e.g. scroll-into-view, right-click copy flyout, file picker), and
    /// manages the startup splash animation.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        /// <summary>
        /// The main view model, created with singleton service instances.
        /// Exposed as a property for <c>x:Bind</c> access in XAML.
        /// </summary>
        public CharMapViewModel ViewModel { get; } = new CharMapViewModel(
            FontService.Instance,
            ClipboardService.Instance);

        /// <summary>
        /// Provides grouped data binding for the GridView.
        /// <c>IsSourceGrouped</c> is always <c>true</c>; flat sort modes use a
        /// single <see cref="GlyphGroup"/> to avoid runtime toggling issues.
        /// </summary>
        public CollectionViewSource GroupedGlyphsSource { get; } = new()
        {
            IsSourceGrouped = true
        };

        /// <summary>
        /// Initializes the window, binds the grouped glyph source, sets up event
        /// handlers, configures the title bar icon, and starts the splash animation.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            GroupedGlyphsSource.Source = ViewModel.GlyphGroups;
            ViewModel.Initialize(DispatcherQueue);
            ViewModel.ScrollToSelectedRequested += OnScrollToSelectedRequested;
            ViewModel.PropertyChanged += OnViewModelPropertyChanged;
            SortByBlockMenuItem.IsChecked = true;

            AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico"));

            LoadSplashImage();
            DismissSplashAfterDelay();
        }

        /// <summary>
        /// Loads the splash screen image from the filesystem rather than using
        /// <c>ms-appx:///</c> URIs, which can cause <c>ExecutionEngineException</c>
        /// in unpackaged WinUI 3 apps when scale qualifiers are involved.
        /// </summary>
        private void LoadSplashImage()
        {
            var path = Path.Combine(AppContext.BaseDirectory, "Assets", "SplashScreen.scale-200.png");
            if (File.Exists(path))
            {
                SplashImage.Source = new BitmapImage(new Uri(path));
            }
        }

        /// <summary>
        /// Waits 2 seconds, then cross-fades the splash overlay out and the main
        /// content in using a storyboard animation. The splash overlay is collapsed
        /// after the animation completes to remove it from the visual tree.
        /// </summary>
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

        /// <summary>
        /// Scrolls the glyph GridView to bring the selected glyph into view.
        /// Called when the view model raises <see cref="CharMapViewModel.ScrollToSelectedRequested"/>
        /// (e.g. after codepoint navigation).
        /// </summary>
        private void OnScrollToSelectedRequested(object? sender, EventArgs e)
        {
            if (ViewModel.SelectedGlyph is not null)
            {
                GlyphGridView.ScrollIntoView(ViewModel.SelectedGlyph, ScrollIntoViewAlignment.Leading);
            }
        }

        /// <summary>
        /// Executes the codepoint navigation command when the user presses Enter
        /// in the "Go to" text box.
        /// </summary>
        private void CodepointTextBox_KeyDown(object sender, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs e)
        {
            if (e.Key == Windows.System.VirtualKey.Enter)
            {
                ViewModel.NavigateToCodepointCommand.Execute(null);
                e.Handled = true;
            }
        }

        /// <summary>
        /// Handles right-click (or long-press) on a glyph card: copies the glyph
        /// to the clipboard and shows a brief "Copied!" flyout that auto-dismisses
        /// after 800ms.
        /// </summary>
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

        /// <summary>
        /// Opens a file picker for <c>.ttf</c>, <c>.otf</c>, and <c>.ttc</c> font files.
        /// If a file is selected, delegates to the view model's <see cref="CharMapViewModel.InstallFont"/>
        /// method for per-user installation.
        /// </summary>
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

        /// <summary>
        /// Updates the font search suggestions as the user types in the AutoSuggestBox.
        /// Only responds to user-initiated text changes (not programmatic ones).
        /// </summary>
        private void FontSearch_TextChanged(AutoSuggestBox sender, AutoSuggestBoxTextChangedEventArgs args)
        {
            if (args.Reason == AutoSuggestionBoxTextChangeReason.UserInput)
            {
                ViewModel.UpdateSearchSuggestions(sender.Text);
            }
        }

        /// <summary>
        /// Updates the AutoSuggestBox text when the user highlights a suggestion
        /// in the dropdown (before committing).
        /// </summary>
        private void FontSearch_SuggestionChosen(AutoSuggestBox sender, AutoSuggestBoxSuggestionChosenEventArgs args)
        {
            if (args.SelectedItem is string fontName)
            {
                sender.Text = fontName;
            }
        }

        /// <summary>
        /// Commits a font search: selects the chosen suggestion or, if the user
        /// pressed Enter without selecting, delegates to <see cref="CharMapViewModel.SelectFontFromSearch"/>
        /// which falls back to the first matching suggestion.
        /// </summary>
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

        /// <summary>
        /// Keeps the View menu's sort radio items in sync when the sort mode changes
        /// via the CommandBar's Sort ComboBox.
        /// </summary>
        private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CharMapViewModel.SortMode))
            {
                SortByBlockMenuItem.IsChecked = ViewModel.SortMode == GlyphSortMode.ByBlock;
                SortByCodepointMenuItem.IsChecked = ViewModel.SortMode == GlyphSortMode.ByCodepoint;
                SortByNameMenuItem.IsChecked = ViewModel.SortMode == GlyphSortMode.ByName;
            }
        }

        /// <summary>
        /// Sets the sort mode to "By Block" from the View menu.
        /// </summary>
        private void SortByBlock_Click(object sender, RoutedEventArgs e)
            => ViewModel.SetSortMode(GlyphSortMode.ByBlock);

        /// <summary>
        /// Sets the sort mode to "By Codepoint" from the View menu.
        /// </summary>
        private void SortByCodepoint_Click(object sender, RoutedEventArgs e)
            => ViewModel.SetSortMode(GlyphSortMode.ByCodepoint);

        /// <summary>
        /// Sets the sort mode to "By Name" from the View menu.
        /// </summary>
        private void SortByName_Click(object sender, RoutedEventArgs e)
            => ViewModel.SetSortMode(GlyphSortMode.ByName);

        /// <summary>
        /// Moves keyboard focus to the codepoint navigation TextBox.
        /// Invoked from the Edit menu's "Go to Codepoint..." item (Ctrl+G).
        /// </summary>
        private void GoToCodepoint_Click(object sender, RoutedEventArgs e)
            => CodepointTextBox.Focus(FocusState.Keyboard);

        /// <summary>
        /// Closes the application window. Invoked from the File menu's "Exit" item.
        /// </summary>
        private void ExitApp_Click(object sender, RoutedEventArgs e)
            => Close();
    }
}
