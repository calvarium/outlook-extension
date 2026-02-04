using System.Windows;
using System.Windows.Input;
using outlook_extension.UI.Services;
using outlook_extension.UI.ViewModels;
using FormsKeys = System.Windows.Forms.Keys;

namespace outlook_extension.UI.Views
{
    public partial class SettingsWindow : Window
    {
        private readonly SettingsViewModel _viewModel;
        private readonly ThemeService _themeService;
        private readonly SettingsService _settingsService;

        public SettingsWindow(
            SettingsViewModel viewModel,
            ThemeService themeService,
            SettingsService settingsService)
        {
            InitializeComponent();
            _viewModel = viewModel;
            _themeService = themeService;
            _settingsService = settingsService;
            DataContext = _viewModel;
            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _themeService.ApplyTheme(_settingsService.Current.Theme);
            _themeService.WatchSystemTheme(this);
        }

        private void OnShortcutKeyDown(object sender, KeyEventArgs e)
        {
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            var keys = (FormsKeys)KeyInterop.VirtualKeyFromKey(key);

            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                keys |= FormsKeys.Control;
            }

            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                keys |= FormsKeys.Shift;
            }

            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt))
            {
                keys |= FormsKeys.Alt;
            }

            var formatted = ShortcutParser.Format(keys);
            if (!string.IsNullOrWhiteSpace(formatted))
            {
                _viewModel.Shortcut = formatted;
            }

            e.Handled = true;
        }

        private void OnAddFavorite(object sender, RoutedEventArgs e)
        {
            _viewModel.AddFavorite(this);
        }

        private void OnRemoveFavorite(object sender, RoutedEventArgs e)
        {
            _viewModel.RemoveFavorite();
        }

        private void OnMoveFavoriteUp(object sender, RoutedEventArgs e)
        {
            _viewModel.MoveFavorite(-1);
        }

        private void OnMoveFavoriteDown(object sender, RoutedEventArgs e)
        {
            _viewModel.MoveFavorite(1);
        }

        private void OnRefreshCache(object sender, RoutedEventArgs e)
        {
            _viewModel.RefreshCache();
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.TrySave(out var errorMessage))
            {
                MessageBox.Show(errorMessage, "Quick Move", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Close();
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
