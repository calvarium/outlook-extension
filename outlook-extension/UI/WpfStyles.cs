using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace outlook_extension
{
    internal static class WpfStyles
    {
        public static Brush GlassBackground => new SolidColorBrush(Color.FromArgb(235, 20, 24, 36));
        public static Brush CardBackground => new SolidColorBrush(Color.FromArgb(230, 28, 34, 48));
        public static Brush SubtleBackground => new SolidColorBrush(Color.FromArgb(160, 34, 41, 58));
        public static Brush AccentBackground => new SolidColorBrush(Color.FromArgb(220, 78, 120, 214));
        public static Brush AccentHoverBackground => new SolidColorBrush(Color.FromArgb(230, 96, 138, 230));
        public static Brush TextPrimary => new SolidColorBrush(Color.FromArgb(235, 240, 244, 255));
        public static Brush TextSecondary => new SolidColorBrush(Color.FromArgb(200, 200, 210, 228));

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
                Padding = new Thickness(18, 8, 18, 8),
                FontWeight = FontWeights.SemiBold,
                MinHeight = 36,
                MinWidth = 120,
                FontSize = 13,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center
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
                Padding = new Thickness(14, 6, 14, 6),
                MinHeight = 34,
                MinWidth = 110,
                FontSize = 12,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center
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
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(Button.ContentProperty));
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
                Padding = new Thickness(18),
                Child = content
            };
        }

        public static Border CreateInputCard(UIElement content)
        {
            return new Border
            {
                Background = SubtleBackground,
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(12),
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
                Padding = new Thickness(6, 4, 6, 4),
                CaretBrush = TextPrimary
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
            listBox.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            listBox.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Disabled);
            listBox.Resources.Add(typeof(ScrollBar), CreateScrollBarStyle());
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(10, 6, 10, 6)));
            style.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0, 4, 0, 4)));
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

        private static Style CreateScrollBarStyle()
        {
            var style = new Style(typeof(ScrollBar));
            style.Setters.Add(new Setter(Control.WidthProperty, 8.0));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));

            var template = new ControlTemplate(typeof(ScrollBar));
            var grid = new FrameworkElementFactory(typeof(Grid));
            var track = new FrameworkElementFactory(typeof(Track));
            track.Name = "PART_Track";

            var decreaseRepeat = new FrameworkElementFactory(typeof(RepeatButton));
            decreaseRepeat.SetValue(Control.BackgroundProperty, Brushes.Transparent);
            decreaseRepeat.SetValue(Control.BorderThicknessProperty, new Thickness(0));
            decreaseRepeat.SetValue(RepeatButton.CommandProperty, ScrollBar.LineUpCommand);

            var increaseRepeat = new FrameworkElementFactory(typeof(RepeatButton));
            increaseRepeat.SetValue(Control.BackgroundProperty, Brushes.Transparent);
            increaseRepeat.SetValue(Control.BorderThicknessProperty, new Thickness(0));
            increaseRepeat.SetValue(RepeatButton.CommandProperty, ScrollBar.LineDownCommand);

            var thumb = new FrameworkElementFactory(typeof(Thumb));
            thumb.SetValue(Control.BackgroundProperty, new SolidColorBrush(Color.FromArgb(160, 140, 160, 200)));
            thumb.SetValue(Control.MarginProperty, new Thickness(0, 2, 0, 2));
            thumb.SetValue(FrameworkElement.WidthProperty, 8.0);
            thumb.SetValue(FrameworkElement.MinHeightProperty, 24.0);

            track.SetValue(Track.DecreaseRepeatButtonProperty, decreaseRepeat);
            track.SetValue(Track.ThumbProperty, thumb);
            track.SetValue(Track.IncreaseRepeatButtonProperty, increaseRepeat);
            grid.AppendChild(track);
            template.VisualTree = grid;
            style.Setters.Add(new Setter(Control.TemplateProperty, template));
            return style;
        }
    }
}
