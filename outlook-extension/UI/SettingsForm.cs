using System;
using System.Drawing;
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
            Width = 720;
            Height = 560;
            MaximizeBox = false;
            MinimizeBox = false;
            GlassStyle.ApplyFormStyle(this);

            var rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3
            };
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 64
            };
            var titleStack = new TableLayoutPanel
            {
                Dock = DockStyle.Left,
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 2
            };
            titleStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            titleStack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var titleLabel = new Label
            {
                Text = "Quick Move",
                Font = new Font("Segoe UI Semibold", 16F, FontStyle.Bold, GraphicsUnit.Point),
                ForeColor = Color.White,
                AutoSize = true
            };
            var subtitleLabel = new Label
            {
                Text = "Einstellungen",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(210, 255, 255, 255),
                AutoSize = true
            };
            titleStack.Controls.Add(titleLabel, 0, 0);
            titleStack.Controls.Add(subtitleLabel, 0, 1);
            var closeButton = new Button
            {
                Text = "✕",
                Width = 36,
                Height = 36,
                Dock = DockStyle.Right
            };
            GlassStyle.StyleSubtleButton(closeButton);
            closeButton.Click += (sender, args) => Close();
            headerPanel.Controls.Add(closeButton);
            headerPanel.Controls.Add(titleStack);
            GlassStyle.EnableDrag(headerPanel);

            var contentCard = GlassStyle.CreateGlassCard();

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 9,
                AutoSize = true
            };
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var shortcutLabel = new Label
            {
                Text = "Shortcut",
                Anchor = AnchorStyles.Left,
                AutoSize = true,
                ForeColor = Color.WhiteSmoke
            };
            _shortcutBox = new TextBox
            {
                ReadOnly = true,
                Text = _settingsService.Current.Shortcut
            };
            GlassStyle.StyleTextInput(_shortcutBox);
            _shortcutBox.KeyDown += OnShortcutKeyDown;

            var favoritesLabel = new Label
            {
                Text = "Favoriten",
                Anchor = AnchorStyles.Left,
                AutoSize = true,
                ForeColor = Color.WhiteSmoke
            };
            _favoritesList = new ListBox
            {
                Height = 160,
                Dock = DockStyle.Fill
            };
            GlassStyle.StyleListBox(_favoritesList);
            _favoritesList.DrawItem += OnFavoritesDrawItem;

            var favoritesButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill
            };
            var addFavorite = new Button { Text = "Hinzufügen" };
            var removeFavorite = new Button { Text = "Entfernen" };
            var moveUp = new Button { Text = "▲" };
            var moveDown = new Button { Text = "▼" };
            GlassStyle.StyleSubtleButton(addFavorite);
            GlassStyle.StyleSubtleButton(removeFavorite);
            GlassStyle.StyleSubtleButton(moveUp);
            GlassStyle.StyleSubtleButton(moveDown);
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
                AutoSize = true,
                ForeColor = Color.WhiteSmoke
            };
            _maxRecentsInput = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 50,
                Value = _settingsService.Current.MaxRecents
            };
            GlassStyle.StyleNumericInput(_maxRecentsInput);

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
            GlassStyle.StyleCheckBox(_showInboxOnly);
            GlassStyle.StyleCheckBox(_includeArchives);

            var refreshButton = new Button
            {
                Text = "Ordnerliste neu laden"
            };
            GlassStyle.StyleSubtleButton(refreshButton);
            refreshButton.Click += (sender, args) => _folderService.RefreshCache();

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft
            };
            var saveButton = new Button { Text = "Speichern" };
            var cancelButton = new Button { Text = "Abbrechen" };
            GlassStyle.StyleGlassButton(saveButton);
            GlassStyle.StyleSubtleButton(cancelButton);
            saveButton.Click += OnSave;
            cancelButton.Click += (sender, args) => Close();
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(cancelButton);

            var shortcutPanel = GlassStyle.CreateInputPanel();
            shortcutPanel.MinimumSize = new Size(0, 42);
            shortcutPanel.Controls.Add(_shortcutBox);

            var recentsPanel = GlassStyle.CreateInputPanel();
            recentsPanel.MinimumSize = new Size(0, 42);
            _maxRecentsInput.Dock = DockStyle.Fill;
            recentsPanel.Controls.Add(_maxRecentsInput);

            mainPanel.Controls.Add(shortcutLabel, 0, 0);
            mainPanel.Controls.Add(shortcutPanel, 1, 0);
            mainPanel.Controls.Add(favoritesLabel, 0, 1);
            mainPanel.Controls.Add(_favoritesList, 1, 1);
            mainPanel.Controls.Add(new Label(), 0, 2);
            mainPanel.Controls.Add(favoritesButtons, 1, 2);
            mainPanel.Controls.Add(recentsLabel, 0, 3);
            mainPanel.Controls.Add(recentsPanel, 1, 3);
            mainPanel.Controls.Add(new Label(), 0, 4);
            mainPanel.Controls.Add(_showInboxOnly, 1, 4);
            mainPanel.Controls.Add(new Label(), 0, 5);
            mainPanel.Controls.Add(_includeArchives, 1, 5);
            mainPanel.Controls.Add(new Label(), 0, 6);
            mainPanel.Controls.Add(refreshButton, 1, 6);

            contentCard.Controls.Add(mainPanel);
            rootLayout.Controls.Add(headerPanel, 0, 0);
            rootLayout.Controls.Add(contentCard, 0, 1);
            rootLayout.Controls.Add(buttonPanel, 0, 2);
            Controls.Add(rootLayout);

            buttonPanel.Padding = new Padding(0, 16, 0, 0);
            mainPanel.Padding = new Padding(12);
            mainPanel.SetColumnSpan(_showInboxOnly, 2);
            mainPanel.SetColumnSpan(_includeArchives, 2);
            mainPanel.SetColumnSpan(refreshButton, 2);
            refreshButton.Anchor = AnchorStyles.Left;
            AcceptButton = saveButton;
            CancelButton = cancelButton;

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

        private void OnFavoritesDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            var item = _favoritesList.Items[e.Index];
            var text = item is FavoriteItem favorite ? favorite.Label : item?.ToString();
            var bounds = e.Bounds;
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using (var background = new SolidBrush(selected
                       ? Color.FromArgb(120, 100, 170, 255)
                       : Color.FromArgb(20, 255, 255, 255)))
            {
                e.Graphics.FillRectangle(background, bounds);
            }

            TextRenderer.DrawText(
                e.Graphics,
                text,
                _favoritesList.Font,
                new Rectangle(bounds.X + 10, bounds.Y + 6, bounds.Width - 20, bounds.Height - 12),
                Color.White,
                TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
        }

        private class FavoriteItem
        {
            public FolderIdentifier Identifier { get; set; }

            public string Label { get; set; }
        }
    }
}
