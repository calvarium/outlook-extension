using System;
using System.Linq;
using System.Windows;
using Wpf.Ui.Appearance;

namespace outlook_extension.UI.Wpf
{
    public static class WpfUiAppHost
    {
        public static void EnsureInitialized(SettingsService settingsService)
        {
            if (Application.Current == null)
            {
                new Application();
            }

            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/StaticColors.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/Palette.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/Variables.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/Accent.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/Fonts.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/Typography.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/DefaultFocusVisualStyle.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/DefaultContextMenu.xaml");
            AddResourceDictionary("pack://application:,,,/Wpf.Ui;component/Resources/DefaultTextBoxScrollViewerStyle.xaml");
            AddResourceDictionary("pack://application:,,,/outlook-extension;component/UI/Wpf/Resources/Controls.xaml");

            ApplyTheme(settingsService?.Current?.ThemeMode);
        }

        public static void ApplyTheme(string themeMode)
        {
            var normalized = themeMode ?? "System";
            if (string.Equals(normalized, "Light", StringComparison.OrdinalIgnoreCase))
            {
                ThemeManager.Apply(ThemeType.Light);
            }
            else if (string.Equals(normalized, "Dark", StringComparison.OrdinalIgnoreCase))
            {
                ThemeManager.Apply(ThemeType.Dark);
            }
            else
            {
                ThemeManager.Apply(ThemeType.System);
            }
        }

        private static void AddResourceDictionary(string source)
        {
            var existing = Application.Current.Resources.MergedDictionaries
                .FirstOrDefault(dictionary => dictionary.Source != null && dictionary.Source.OriginalString == source);
            if (existing != null)
            {
                return;
            }

            Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary
            {
                Source = new Uri(source, UriKind.Absolute)
            });
        }
    }
}
