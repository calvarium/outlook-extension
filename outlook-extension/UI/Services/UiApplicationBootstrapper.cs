using System;
using System.Linq;
using System.Windows;

namespace outlook_extension.UI.Services
{
    public static class UiApplicationBootstrapper
    {
        private static readonly Uri AppResourcesUri =
            new Uri("pack://application:,,,/outlook-extension;component/UI/Resources/AppResources.xaml");

        public static void EnsureApplication()
        {
            if (Application.Current == null)
            {
                var application = new Application
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };

                application.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = AppResourcesUri });
                return;
            }

            var dictionaries = Application.Current.Resources.MergedDictionaries;
            var hasResources = dictionaries.Any(dictionary =>
                dictionary.Source != null &&
                dictionary.Source.OriginalString.Contains("AppResources.xaml", StringComparison.OrdinalIgnoreCase));

            if (!hasResources)
            {
                dictionaries.Add(new ResourceDictionary { Source = AppResourcesUri });
            }
        }
    }
}
