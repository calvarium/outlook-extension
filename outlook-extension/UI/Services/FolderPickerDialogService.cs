using System.Windows;
using outlook_extension.UI.ViewModels;
using outlook_extension.UI.Views;

namespace outlook_extension.UI.Services
{
    public class FolderPickerDialogService : IFolderPickerDialogService
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private readonly ThemeService _themeService;
        private readonly SettingsService _settingsService;

        public FolderPickerDialogService(
            FolderService folderService,
            SearchService searchService,
            SettingsService settingsService,
            ThemeService themeService)
        {
            _folderService = folderService;
            _searchService = searchService;
            _settingsService = settingsService;
            _themeService = themeService;
        }

        public FolderInfo PickFolder(Window owner)
        {
            UiApplicationBootstrapper.EnsureApplication();
            var viewModel = new FolderPickerViewModel(_folderService, _searchService);
            var window = new FolderPickerWindow(viewModel, _themeService, _settingsService)
            {
                Owner = owner
            };

            var result = window.ShowDialog();
            return result == true ? window.SelectedFolder : null;
        }
    }
}
