using System.Collections.ObjectModel;
using System.Windows;
using outlook_extension.UI.Services;

namespace outlook_extension.UI.ViewModels
{
    public class SettingsViewModel : ViewModelBase
    {
        private readonly FolderService _folderService;
        private readonly SettingsService _settingsService;
        private readonly HotkeyService _hotkeyService;
        private readonly IFolderPickerDialogService _folderPickerDialogService;
        private readonly ThemeService _themeService;
        private string _shortcut;
        private int _maxRecents;
        private bool _showInboxOnly;
        private bool _includeArchives;
        private ThemePreference _themePreference;
        private FavoriteItemViewModel _selectedFavorite;

        public SettingsViewModel(
            FolderService folderService,
            SettingsService settingsService,
            HotkeyService hotkeyService,
            IFolderPickerDialogService folderPickerDialogService,
            ThemeService themeService)
        {
            _folderService = folderService;
            _settingsService = settingsService;
            _hotkeyService = hotkeyService;
            _folderPickerDialogService = folderPickerDialogService;
            _themeService = themeService;

            Favorites = new ObservableCollection<FavoriteItemViewModel>();
            Shortcut = _settingsService.Current.Shortcut;
            MaxRecents = _settingsService.Current.MaxRecents;
            ShowInboxOnly = _settingsService.Current.ShowInboxOnly;
            IncludeArchives = _settingsService.Current.IncludeArchives;
            ThemePreference = _settingsService.Current.Theme;

            RefreshFavorites();
        }

        public ObservableCollection<FavoriteItemViewModel> Favorites { get; }

        public FavoriteItemViewModel SelectedFavorite
        {
            get => _selectedFavorite;
            set => SetField(ref _selectedFavorite, value);
        }

        public string Shortcut
        {
            get => _shortcut;
            set => SetField(ref _shortcut, value);
        }

        public int MaxRecents
        {
            get => _maxRecents;
            set
            {
                if (SetField(ref _maxRecents, value))
                {
                    _settingsService.Current.MaxRecents = value;
                }
            }
        }

        public bool ShowInboxOnly
        {
            get => _showInboxOnly;
            set
            {
                if (SetField(ref _showInboxOnly, value))
                {
                    _settingsService.Current.ShowInboxOnly = value;
                }
            }
        }

        public bool IncludeArchives
        {
            get => _includeArchives;
            set
            {
                if (SetField(ref _includeArchives, value))
                {
                    _settingsService.Current.IncludeArchives = value;
                }
            }
        }

        public ThemePreference ThemePreference
        {
            get => _themePreference;
            set
            {
                if (SetField(ref _themePreference, value))
                {
                    _settingsService.Current.Theme = value;
                    _themeService.ApplyTheme(value);
                }
            }
        }

        public void RefreshCache()
        {
            _folderService.RefreshCache();
        }

        public void AddFavorite(Window owner)
        {
            var selected = _folderPickerDialogService.PickFolder(owner);
            if (selected == null)
            {
                return;
            }

            _settingsService.AddFavorite(selected);
            RefreshFavorites();
        }

        public void RemoveFavorite()
        {
            if (SelectedFavorite == null)
            {
                return;
            }

            _settingsService.RemoveFavorite(SelectedFavorite.Identifier);
            RefreshFavorites();
        }

        public void MoveFavorite(int offset)
        {
            if (SelectedFavorite == null)
            {
                return;
            }

            var currentIndex = Favorites.IndexOf(SelectedFavorite);
            _settingsService.MoveFavorite(currentIndex, offset);
            RefreshFavorites();

            var newIndex = currentIndex + offset;
            if (newIndex >= 0 && newIndex < Favorites.Count)
            {
                SelectedFavorite = Favorites[newIndex];
            }
        }

        public bool TrySave(out string errorMessage)
        {
            if (!ShortcutParser.TryParse(Shortcut, out _, out _))
            {
                errorMessage = "Der Shortcut ist ungültig. Bitte eine Kombination wie Ctrl+Shift+M wählen.";
                return false;
            }

            _settingsService.Current.Shortcut = Shortcut;
            _settingsService.Save();
            _hotkeyService.RegisterShortcut();
            errorMessage = null;
            return true;
        }

        private void RefreshFavorites()
        {
            Favorites.Clear();
            foreach (var identifier in _settingsService.Current.Favorites)
            {
                Favorites.Add(new FavoriteItemViewModel
                {
                    Identifier = identifier,
                    Label = ResolveFolderLabel(identifier)
                });
            }
        }

        private string ResolveFolderLabel(FolderIdentifier identifier)
        {
            var info = _folderService.GetFolderByIdentifier(identifier);
            return info?.DisplayText ?? $"Unbekannter Ordner ({identifier.EntryId})";
        }
    }
}
