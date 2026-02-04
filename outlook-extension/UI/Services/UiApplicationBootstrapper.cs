using System;
using System.Linq;
using System.Windows;

namespace outlook_extension.UI.Services
{
    public static class UiApplicationBootstrapper
    {
        public static void EnsureApplication()
        {
            if (Application.Current == null)
            {
                var application = new Application
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };

                Application.ResourceAssembly = typeof(UiApplicationBootstrapper).Assembly;
                application.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = BuildResourceUri() });
                return;
            }

            if (Application.ResourceAssembly == null)
            {
                Application.ResourceAssembly = typeof(UiApplicationBootstrapper).Assembly;
            }

            var dictionaries = Application.Current.Resources.MergedDictionaries;
            var hasResources = dictionaries.Any(dictionary =>
                dictionary.Source != null &&
                dictionary.Source.OriginalString.IndexOf("AppResources.xaml", StringComparison.OrdinalIgnoreCase) >= 0);

            if (!hasResources)
            {
                dictionaries.Add(new ResourceDictionary { Source = BuildResourceUri() });
            }
        }

        private static Uri BuildResourceUri()
        {
            return new Uri("UI/Resources/AppResources.xaml", UriKind.Relative);
        }
    }
}
