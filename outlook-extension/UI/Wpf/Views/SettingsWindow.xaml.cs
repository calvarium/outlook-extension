using System.Windows;
using System.Windows.Input;
using outlook_extension.UI.Wpf.ViewModels;

namespace outlook_extension.UI.Wpf.Views
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
        }

        private void OnShortcutKeyDown(object sender, KeyEventArgs e)
        {
            if (DataContext is SettingsViewModel viewModel)
            {
                viewModel.UpdateShortcutFromKey(e);
                e.Handled = true;
            }
        }
    }
}
