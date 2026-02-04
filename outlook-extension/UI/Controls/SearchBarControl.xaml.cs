using System.Windows;
using System.Windows.Controls;

namespace outlook_extension.UI.Controls
{
    public partial class SearchBarControl : UserControl
    {
        public static readonly DependencyProperty PlaceholderTextProperty = DependencyProperty.Register(
            nameof(PlaceholderText),
            typeof(string),
            typeof(SearchBarControl),
            new PropertyMetadata("Suchen"));

        public SearchBarControl()
        {
            InitializeComponent();
        }

        public string PlaceholderText
        {
            get => (string)GetValue(PlaceholderTextProperty);
            set => SetValue(PlaceholderTextProperty, value);
        }

        public void FocusInput()
        {
            SearchTextBox.Focus();
        }

        public void SelectAll()
        {
            SearchTextBox.SelectAll();
        }

        public void DeletePreviousWord()
        {
            var text = SearchTextBox.Text ?? string.Empty;
            var caret = SearchTextBox.SelectionStart;
            if (caret <= 0)
            {
                return;
            }

            var deleteFrom = caret - 1;
            while (deleteFrom > 0 && char.IsWhiteSpace(text[deleteFrom]))
            {
                deleteFrom--;
            }

            while (deleteFrom > 0 && !char.IsWhiteSpace(text[deleteFrom - 1]))
            {
                deleteFrom--;
            }

            SearchTextBox.Text = text.Remove(deleteFrom, caret - deleteFrom);
            SearchTextBox.SelectionStart = deleteFrom;
        }
    }
}
