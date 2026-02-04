using System.Windows;
using System.Windows.Input;
using outlook_extension.UI.Wpf.ViewModels;

namespace outlook_extension.UI.Wpf.Views
{
    public partial class QuickMoveWindow : Window
    {
        public QuickMoveWindow()
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
            if (DataContext is not QuickMoveViewModel viewModel)
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
                var keepOpen = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
                viewModel.ExecuteSelected(keepOpen);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                viewModel.HandleEscape();
                e.Handled = true;
            }
            else if (e.Key == Key.Z && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                viewModel.UndoLastMove();
                e.Handled = true;
            }
        }
    }
}
