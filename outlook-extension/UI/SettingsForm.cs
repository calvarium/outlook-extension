using System;
using System.Linq;
using System.Windows.Forms;

namespace outlook_extension
{
    public class SettingsForm : Form
    {
        private readonly FolderService _folderService;
        private readonly SettingsService _settingsService;
        private readonly HotkeyService _hotkeyService;
        private readonly TextBox _shortcutBox;
        private readonly ListBox _favoritesList;
        private readonly NumericUpDown _maxRecentsInput;
        private readonly CheckBox _showInboxOnly;
        private readonly CheckBox _includeArchives;

        public SettingsForm(FolderService folderService, SettingsService settingsService, HotkeyService hotkeyService)
        {
            _folderService = folderService;
            _settingsService = settingsService;
            _hotkeyService = hotkeyService;

            Text = "Quick Move Einstellungen";
            Width = 640;
            Height = 520;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new Padding(12),
                AutoSize = true
            };
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var shortcutLabel = new Label
            {
                Text = "Shortcut",
                Anchor = AnchorStyles.Left,
                AutoSize = true
            };
            _shortcutBox = new TextBox
            {
                ReadOnly = true,
                Text = _settingsService.Current.Shortcut
            };
            _shortcutBox.KeyDown += OnShortcutKeyDown;

            var favoritesLabel = new Label
            {
                Text = "Favoriten",
                Anchor = AnchorStyles.Left,
                AutoSize = true
            };
            _favoritesList = new ListBox
            {
                Height = 160,
                Dock = DockStyle.Fill
            };

            var favoritesButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill
            };
            var addFavorite = new Button { Text = "Hinzufügen" };
            var removeFavorite = new Button { Text = "Entfernen" };
            var moveUp = new Button { Text = "▲" };
            var moveDown = new Button { Text = "▼" };
            addFavorite.Click += OnAddFavorite;
            removeFavorite.Click += OnRemoveFavorite;
            moveUp.Click += (sender, args) => MoveFavorite(-1);
            moveDown.Click += (sender, args) => MoveFavorite(1);
            favoritesButtons.Controls.Add(addFavorite);
            favoritesButtons.Controls.Add(removeFavorite);
            favoritesButtons.Controls.Add(moveUp);
            favoritesButtons.Controls.Add(moveDown);

            var recentsLabel = new Label
            {
                Text = "Anzahl letzte Ziele",
                Anchor = AnchorStyles.Left,
                AutoSize = true
            };
            _maxRecentsInput = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 50,
                Value = _settingsService.Current.MaxRecents
            };

            _showInboxOnly = new CheckBox
            {
                Text = "Nur Unterordner von Posteingang anzeigen",
                Checked = _settingsService.Current.ShowInboxOnly,
                AutoSize = true
            };
            _includeArchives = new CheckBox
            {
                Text = "Archiv/Online-Archive anzeigen",
                Checked = _settingsService.Current.IncludeArchives,
                AutoSize = true
            };

            var refreshButton = new Button
            {
                Text = "Ordnerliste neu laden"
            };
            refreshButton.Click += (sender, args) => _folderService.RefreshCache();

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft
            };
            var saveButton = new Button { Text = "Speichern" };
            var cancelButton = new Button { Text = "Abbrechen" };
            saveButton.Click += OnSave;
            cancelButton.Click += (sender, args) => Close();
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(cancelButton);

            mainPanel.Controls.Add(shortcutLabel, 0, 0);
            mainPanel.Controls.Add(_shortcutBox, 1, 0);
            mainPanel.Controls.Add(favoritesLabel, 0, 1);
            mainPanel.Controls.Add(_favoritesList, 1, 1);
            mainPanel.Controls.Add(new Label(), 0, 2);
            mainPanel.Controls.Add(favoritesButtons, 1, 2);
            mainPanel.Controls.Add(recentsLabel, 0, 3);
            mainPanel.Controls.Add(_maxRecentsInput, 1, 3);
            mainPanel.Controls.Add(new Label(), 0, 4);
            mainPanel.Controls.Add(_showInboxOnly, 1, 4);
            mainPanel.Controls.Add(new Label(), 0, 5);
            mainPanel.Controls.Add(_includeArchives, 1, 5);

            Controls.Add(mainPanel);
            Controls.Add(refreshButton);
            Controls.Add(buttonPanel);

            refreshButton.Dock = DockStyle.Bottom;
            buttonPanel.Padding = new Padding(12);
            refreshButton.Padding = new Padding(12, 6, 12, 6);

            Load += (sender, args) => RefreshFavorites();
        }

        private void OnShortcutKeyDown(object sender, KeyEventArgs e)
        {
            var shortcut = ShortcutParser.Format(e.KeyData);
            if (!string.IsNullOrWhiteSpace(shortcut))
            {
                _shortcutBox.Text = shortcut;
            }

            e.SuppressKeyPress = true;
        }

        private void OnAddFavorite(object sender, EventArgs e)
        {
            using (var picker = new FolderPickerForm(_folderService, new SearchService(_settingsService)))
            {
                if (picker.ShowDialog() == DialogResult.OK && picker.SelectedFolder != null)
                {
                    _settingsService.AddFavorite(picker.SelectedFolder);
                    RefreshFavorites();
                }
            }
        }

        private void OnRemoveFavorite(object sender, EventArgs e)
        {
            var selected = _favoritesList.SelectedItem as FavoriteItem;
            if (selected == null)
            {
                return;
            }

            _settingsService.RemoveFavorite(selected.Identifier);
            RefreshFavorites();
        }

        private void MoveFavorite(int offset)
        {
            var selectedIndex = _favoritesList.SelectedIndex;
            if (selectedIndex < 0)
            {
                return;
            }

            _settingsService.MoveFavorite(selectedIndex, offset);
            RefreshFavorites();
            _favoritesList.SelectedIndex = Math.Max(0, Math.Min(_favoritesList.Items.Count - 1, selectedIndex + offset));
        }

        private void RefreshFavorites()
        {
            var favorites = _settingsService.Current.Favorites
                .Select(identifier => new FavoriteItem
                {
                    Identifier = identifier,
                    Label = ResolveFolderLabel(identifier)
                })
                .ToList();

            _favoritesList.DataSource = favorites;
            _favoritesList.DisplayMember = nameof(FavoriteItem.Label);
        }

        private string ResolveFolderLabel(FolderIdentifier identifier)
        {
            var info = _folderService.GetFolderByIdentifier(identifier);
            return info?.DisplayText ?? $"Unbekannter Ordner ({identifier.EntryId})";
        }

        private void OnSave(object sender, EventArgs e)
        {
            if (!ShortcutParser.TryParse(_shortcutBox.Text, out _, out _))
            {
                MessageBox.Show(
                    "Der Shortcut ist ungültig. Bitte eine Kombination wie Ctrl+Shift+M wählen.",
                    "Quick Move",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            _settingsService.Current.Shortcut = _shortcutBox.Text;
            _settingsService.Current.MaxRecents = (int)_maxRecentsInput.Value;
            _settingsService.Current.ShowInboxOnly = _showInboxOnly.Checked;
            _settingsService.Current.IncludeArchives = _includeArchives.Checked;
            _settingsService.Save();
            _hotkeyService.RegisterShortcut();

            Close();
        }

        private class FavoriteItem
        {
            public FolderIdentifier Identifier { get; set; }

            public string Label { get; set; }
        }
    }
}
