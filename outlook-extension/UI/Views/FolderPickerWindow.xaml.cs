using System.Windows;
using System.Windows.Input;
using outlook_extension.UI.Services;
using outlook_extension.UI.ViewModels;

namespace outlook_extension.UI.Views
{
    public partial class FolderPickerWindow : Window
    {
        private readonly FolderPickerViewModel _viewModel;
        private readonly ThemeService _themeService;
        private readonly SettingsService _settingsService;

        public FolderPickerWindow(
            FolderPickerViewModel viewModel,
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

        public FolderInfo SelectedFolder { get; private set; }

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
                ConfirmSelection();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
                return;
            }

            SearchBar.FocusInput();
        }

        private void OnConfirmClick(object sender, RoutedEventArgs e)
        {
            ConfirmSelection();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OnResultsDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ConfirmSelection();
        }

        private void ConfirmSelection()
        {
            SelectedFolder = _viewModel.GetSelectedFolder();
            if (SelectedFolder != null)
            {
                DialogResult = true;
                Close();
            }
        }
    }
}
