using System;
using System.Linq;
using System.Windows;
using Wpf.Ui.Appearance;

namespace outlook_extension.UI.Services
{
    public class ThemeService
    {
        private static bool _subscribed;

        public ThemeService()
        {
            if (_subscribed)
            {
                return;
            }

            ApplicationThemeManager.Changed += (theme, accent) => UpdateThemeResources(theme);
            _subscribed = true;
        }

        public void ApplyTheme(ThemePreference preference)
        {
            var theme = ResolveApplicationTheme(preference);
            if (Application.Current != null && Application.ResourceAssembly == null)
            {
                Application.ResourceAssembly = typeof(ThemeService).Assembly;
            }
            ApplicationThemeManager.Apply(theme);
            UpdateThemeResources(theme);
        }

        public void WatchSystemTheme(Window window)
        {
            if (window == null)
            {
                return;
            }

            SystemThemeWatcher.Watch(window);
        }

        private ApplicationTheme ResolveApplicationTheme(ThemePreference preference)
        {
            if (preference == ThemePreference.Light)
            {
                return ApplicationTheme.Light;
            }

            if (preference == ThemePreference.Dark)
            {
                return ApplicationTheme.Dark;
            }

            var systemTheme = SystemThemeManager.GetCachedSystemTheme();
            switch (systemTheme)
            {
                case SystemTheme.Dark:
                case SystemTheme.Glow:
                case SystemTheme.CapturedMotion:
                case SystemTheme.HCBlack:
                case SystemTheme.HC1:
                case SystemTheme.HC2:
                    return ApplicationTheme.Dark;
                default:
                    return ApplicationTheme.Light;
            }
        }

        private static void UpdateThemeResources(ApplicationTheme theme)
        {
            if (Application.Current == null)
            {
                return;
            }

            var source = theme == ApplicationTheme.Dark
                ? new Uri("UI/Resources/Theme.Dark.xaml", UriKind.Relative)
                : new Uri("UI/Resources/Theme.Light.xaml", UriKind.Relative);

            var merged = Application.Current.Resources.MergedDictionaries;
            var existing = merged.FirstOrDefault(dictionary =>
                dictionary.Source != null &&
                dictionary.Source.OriginalString.IndexOf("Theme.", StringComparison.OrdinalIgnoreCase) >= 0);

            if (existing != null)
            {
                var index = merged.IndexOf(existing);
                merged[index] = new ResourceDictionary { Source = source };
            }
            else
            {
                merged.Add(new ResourceDictionary { Source = source });
            }
        }

    }
}
