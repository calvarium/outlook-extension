using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace outlook_extension
{
    internal static class WpfStyles
    {
        public static Brush GlassBackground => new SolidColorBrush(Color.FromArgb(210, 24, 24, 30));
        public static Brush CardBackground => new SolidColorBrush(Color.FromArgb(180, 255, 255, 255));
        public static Brush SubtleBackground => new SolidColorBrush(Color.FromArgb(60, 255, 255, 255));
        public static Brush AccentBackground => new SolidColorBrush(Color.FromArgb(180, 90, 140, 255));
        public static Brush AccentHoverBackground => new SolidColorBrush(Color.FromArgb(210, 120, 170, 255));
        public static Brush TextPrimary => Brushes.White;
        public static Brush TextSecondary => new SolidColorBrush(Color.FromArgb(200, 255, 255, 255));

        public static Button CreateIconButton(string text)
        {
            var button = new Button
            {
                Content = text,
                Width = 34,
                Height = 34,
                Background = SubtleBackground,
                Foreground = TextPrimary,
                BorderThickness = new Thickness(0),
                FontSize = 14,
                Padding = new Thickness(0),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            button.MouseEnter += (sender, args) => button.Background = AccentHoverBackground;
            button.MouseLeave += (sender, args) => button.Background = SubtleBackground;
            button.Template = CreateRoundedButtonTemplate(10);
            return button;
        }

        public static Button CreatePrimaryButton(string text)
        {
            var button = new Button
            {
                Content = text,
                Background = AccentBackground,
                Foreground = TextPrimary,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(16, 6, 16, 6),
                FontWeight = FontWeights.SemiBold,
                MinHeight = 34
            };
            button.MouseEnter += (sender, args) => button.Background = AccentHoverBackground;
            button.MouseLeave += (sender, args) => button.Background = AccentBackground;
            button.Template = CreateRoundedButtonTemplate(12);
            return button;
        }

        public static Button CreateSubtleButton(string text)
        {
            var button = new Button
            {
                Content = text,
                Background = SubtleBackground,
                Foreground = TextPrimary,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(12, 4, 12, 4),
                MinHeight = 30
            };
            button.MouseEnter += (sender, args) => button.Background = AccentHoverBackground;
            button.MouseLeave += (sender, args) => button.Background = SubtleBackground;
            button.Template = CreateRoundedButtonTemplate(10);
            return button;
        }

        public static ControlTemplate CreateRoundedButtonTemplate(double radius)
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(radius));
            border.SetBinding(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetBinding(ContentPresenter.ContentProperty, new TemplateBindingExtension(Button.ContentProperty));
            border.AppendChild(contentPresenter);
            template.VisualTree = border;
            return template;
        }

        public static Border CreateGlassCard(UIElement content)
        {
            return new Border
            {
                Background = CardBackground,
                CornerRadius = new CornerRadius(18),
                Padding = new Thickness(16),
                Child = content
            };
        }

        public static Border CreateInputCard(UIElement content)
        {
            return new Border
            {
                Background = SubtleBackground,
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(10),
                Child = content
            };
        }

        public static TextBox CreateTextBox(string text = "")
        {
            return new TextBox
            {
                Text = text,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = TextPrimary,
                FontSize = 14,
                Padding = new Thickness(4, 2, 4, 2)
            };
        }

        public static ListBox CreateListBox()
        {
            var listBox = new ListBox
            {
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = TextPrimary
            };
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 4, 8, 4)));
            style.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0, 2, 0, 2)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
            style.Setters.Add(new Setter(Control.ForegroundProperty, TextPrimary));
            var selectedTrigger = new Trigger
            {
                Property = ListBoxItem.IsSelectedProperty,
                Value = true
            };
            selectedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, AccentBackground));
            style.Triggers.Add(selectedTrigger);
            listBox.ItemContainerStyle = style;
            return listBox;
        }
    }
}
