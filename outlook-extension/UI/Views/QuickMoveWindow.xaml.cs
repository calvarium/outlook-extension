using System.Windows;
using System.Windows.Input;
using outlook_extension.UI.Services;
using outlook_extension.UI.ViewModels;

namespace outlook_extension.UI.Views
{
    public partial class QuickMoveWindow : Window
    {
        private readonly QuickMoveViewModel _viewModel;
        private readonly ThemeService _themeService;
        private readonly SettingsService _settingsService;

        public QuickMoveWindow(
            QuickMoveViewModel viewModel,
            ThemeService themeService,
            SettingsService settingsService)
        {
            InitializeComponent();
            _viewModel = viewModel;
            _themeService = themeService;
            _settingsService = settingsService;
            DataContext = _viewModel;
            Loaded += OnLoaded;
            PreviewKeyDown += OnPreviewKeyDown;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _themeService.ApplyTheme(_settingsService.Current.Theme);
            _themeService.WatchSystemTheme(this);
            _viewModel.Initialize();
            SearchBar.FocusInput();
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                _viewModel.MoveSelection(1);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Up)
            {
                _viewModel.MoveSelection(-1);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Enter)
            {
                var keepDialogOpen = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
                var moved = _viewModel.ExecuteSelection(keepDialogOpen);
                if (keepDialogOpen)
                {
                    if (moved)
                    {
                        SearchBar.SelectAll();
                        _viewModel.RefreshResults();
                    }
                }
                else if (moved)
                {
                    Close();
                }

                e.Handled = true;
                return;
            }

            if (e.Key == Key.Escape)
            {
                if (!string.IsNullOrWhiteSpace(_viewModel.SearchText))
                {
                    _viewModel.ClearSearch();
                    SearchBar.FocusInput();
                }
                else
                {
                    Close();
                }

                e.Handled = true;
                return;
            }

            if (e.Key == Key.Z && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                _viewModel.UndoLastMove();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Back && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                SearchBar.DeletePreviousWord();
                e.Handled = true;
                return;
            }

            SearchBar.FocusInput();
        }

        private void OnResultsDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_viewModel.ExecuteSelection(false))
            {
                Close();
            }
        }
    }
}
