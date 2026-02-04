using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using outlook_extension.UI.Wpf.Infrastructure;
using outlook_extension.UI.Wpf;

namespace outlook_extension.UI.Wpf.ViewModels
{
    public class SettingsViewModel : ViewModelBase
    {
        private readonly FolderService _folderService;
        private readonly SettingsService _settingsService;
        private readonly HotkeyService _hotkeyService;
        private string _shortcut;
        private int _maxRecents;
        private bool _showInboxOnly;
        private bool _includeArchives;
        private string _themeMode;
        private FavoriteItemViewModel _selectedFavorite;

        public SettingsViewModel(FolderService folderService, SettingsService settingsService, HotkeyService hotkeyService)
        {
            _folderService = folderService;
            _settingsService = settingsService;
            _hotkeyService = hotkeyService;

            ThemeOptions = new ObservableCollection<string> { "System", "Light", "Dark" };
            Favorites = new ObservableCollection<FavoriteItemViewModel>();

            Shortcut = settingsService.Current.Shortcut;
            MaxRecents = settingsService.Current.MaxRecents;
            ShowInboxOnly = settingsService.Current.ShowInboxOnly;
            IncludeArchives = settingsService.Current.IncludeArchives;
            ThemeMode = string.IsNullOrWhiteSpace(settingsService.Current.ThemeMode) ? "System" : settingsService.Current.ThemeMode;

            RefreshFavorites();

            AddFavoriteCommand = new RelayCommand(AddFavorite);
            RemoveFavoriteCommand = new RelayCommand(RemoveFavorite, () => SelectedFavorite != null);
            MoveFavoriteUpCommand = new RelayCommand(() => MoveFavorite(-1), () => SelectedFavorite != null);
            MoveFavoriteDownCommand = new RelayCommand(() => MoveFavorite(1), () => SelectedFavorite != null);
            RefreshFoldersCommand = new RelayCommand(_folderService.RefreshCache);
            SaveCommand = new RelayCommand(Save);
            CancelCommand = new RelayCommand(() => CloseRequested?.Invoke());
        }

        public ObservableCollection<FavoriteItemViewModel> Favorites { get; }

        public ObservableCollection<string> ThemeOptions { get; }

        public string Shortcut
        {
            get => _shortcut;
            set => SetProperty(ref _shortcut, value);
        }

        public int MaxRecents
        {
            get => _maxRecents;
            set => SetProperty(ref _maxRecents, value);
        }

        public bool ShowInboxOnly
        {
            get => _showInboxOnly;
            set => SetProperty(ref _showInboxOnly, value);
        }

        public bool IncludeArchives
        {
            get => _includeArchives;
            set => SetProperty(ref _includeArchives, value);
        }

        public string ThemeMode
        {
            get => _themeMode;
            set => SetProperty(ref _themeMode, value);
        }

        public FavoriteItemViewModel SelectedFavorite
        {
            get => _selectedFavorite;
            set
            {
                if (SetProperty(ref _selectedFavorite, value))
                {
                    RemoveFavoriteCommand.RaiseCanExecuteChanged();
                    MoveFavoriteUpCommand.RaiseCanExecuteChanged();
                    MoveFavoriteDownCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public RelayCommand AddFavoriteCommand { get; }

        public RelayCommand RemoveFavoriteCommand { get; }

        public RelayCommand MoveFavoriteUpCommand { get; }

        public RelayCommand MoveFavoriteDownCommand { get; }

        public RelayCommand RefreshFoldersCommand { get; }

        public RelayCommand SaveCommand { get; }

        public RelayCommand CancelCommand { get; }

        public Func<FolderInfo> PickFolder { get; set; }

        public Action CloseRequested { get; set; }

        public void UpdateShortcutFromKey(KeyEventArgs e)
        {
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            var keyValue = (System.Windows.Forms.Keys)System.Windows.Input.KeyInterop.VirtualKeyFromKey(key);
            var modifiers = System.Windows.Forms.Keys.None;

            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                modifiers |= System.Windows.Forms.Keys.Control;
            }

            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                modifiers |= System.Windows.Forms.Keys.Shift;
            }

            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt))
            {
                modifiers |= System.Windows.Forms.Keys.Alt;
            }

            var formatted = ShortcutParser.Format(modifiers | keyValue);
            if (!string.IsNullOrWhiteSpace(formatted))
            {
                Shortcut = formatted;
            }
        }

        private void RefreshFavorites()
        {
            Favorites.Clear();
            foreach (var identifier in _settingsService.Current.Favorites)
            {
                var info = _folderService.GetFolderByIdentifier(identifier);
                Favorites.Add(new FavoriteItemViewModel(identifier, info?.DisplayText ?? $"Unbekannter Ordner ({identifier.EntryId})"));
            }
        }

        private void AddFavorite()
        {
            var picked = PickFolder?.Invoke();
            if (picked == null)
            {
                return;
            }

            _settingsService.AddFavorite(picked);
            RefreshFavorites();
        }

        private void RemoveFavorite()
        {
            if (SelectedFavorite == null)
            {
                return;
            }

            _settingsService.RemoveFavorite(SelectedFavorite.Identifier);
            RefreshFavorites();
        }

        private void MoveFavorite(int offset)
        {
            if (SelectedFavorite == null)
            {
                return;
            }

            var index = Favorites.IndexOf(SelectedFavorite);
            _settingsService.MoveFavorite(index, offset);
            RefreshFavorites();
            var newIndex = Math.Max(0, Math.Min(Favorites.Count - 1, index + offset));
            SelectedFavorite = Favorites.ElementAtOrDefault(newIndex);
        }

        private void Save()
        {
            if (!ShortcutParser.TryParse(Shortcut, out _, out _))
            {
                System.Windows.Forms.MessageBox.Show(
                    "Der Shortcut ist ungültig. Bitte eine Kombination wie Ctrl+Shift+M wählen.",
                    "Quick Move",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            _settingsService.Current.Shortcut = Shortcut;
            _settingsService.Current.MaxRecents = MaxRecents;
            _settingsService.Current.ShowInboxOnly = ShowInboxOnly;
            _settingsService.Current.IncludeArchives = IncludeArchives;
            _settingsService.Current.ThemeMode = ThemeMode;
            _settingsService.Save();
            _hotkeyService.RegisterShortcut();
            WpfUiAppHost.ApplyTheme(ThemeMode);

            CloseRequested?.Invoke();
        }
    }
}
