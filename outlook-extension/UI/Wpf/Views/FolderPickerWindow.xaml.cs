using System.Windows;
using System.Windows.Input;
using outlook_extension.UI.Wpf.ViewModels;
using Wpf.Ui.Controls;

namespace outlook_extension.UI.Wpf.Views
{
    public partial class FolderPickerWindow : Window
    {
        public FolderPickerWindow()
        {
            InitializeComponent();
            Loaded += OnLoaded;
            PreviewKeyDown += OnPreviewKeyDown;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            SearchBar.FocusInput();
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (DataContext is not FolderPickerViewModel viewModel)
            {
                return;
            }

            if (e.Key == Key.Down)
            {
                viewModel.HandleNavigation(1);
                e.Handled = true;
            }
            else if (e.Key == Key.Up)
            {
                viewModel.HandleNavigation(-1);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                viewModel.ConfirmCommand.Execute(null);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                viewModel.HandleEscape();
                e.Handled = true;
            }
        }
    }
}
