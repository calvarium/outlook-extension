using System.Windows;
using System.Windows.Controls;

namespace outlook_extension.UI.Wpf.Controls
{
    public partial class SearchBarControl : UserControl
    {
        public static readonly DependencyProperty SearchTextProperty = DependencyProperty.Register(
            nameof(SearchText),
            typeof(string),
            typeof(SearchBarControl),
            new FrameworkPropertyMetadata(string.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        public static readonly DependencyProperty PlaceholderTextProperty = DependencyProperty.Register(
            nameof(PlaceholderText),
            typeof(string),
            typeof(SearchBarControl),
            new PropertyMetadata("Suchen..."));

        public SearchBarControl()
        {
            InitializeComponent();
        }

        public string SearchText
        {
            get => (string)GetValue(SearchTextProperty);
            set => SetValue(SearchTextProperty, value);
        }

        public string PlaceholderText
        {
            get => (string)GetValue(PlaceholderTextProperty);
            set => SetValue(PlaceholderTextProperty, value);
        }

        public void FocusInput()
        {
            SearchInput?.Focus();
        }
    }
}
