using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace outlook_extension
{
    public class SettingsWindow : Window
    {
        private readonly FolderService _folderService;
        private readonly SettingsService _settingsService;
        private readonly HotkeyService _hotkeyService;
        private readonly TextBox _shortcutBox;
        private readonly ListBox _favoritesList;
        private readonly TextBox _maxRecentsBox;
        private readonly CheckBox _showInboxOnly;
        private readonly CheckBox _includeArchives;

        public SettingsWindow(FolderService folderService, SettingsService settingsService, HotkeyService hotkeyService)
        {
            _folderService = folderService;
            _settingsService = settingsService;
            _hotkeyService = hotkeyService;

            Width = 720;
            Height = 560;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            var rootBorder = new ContinuousCornerBorder
            {
                Background = WpfStyles.GlassBackground,
                CornerRadius = new CornerRadius(CornerTokens.RadiusXL),
                CornerStyle = WpfStyles.DefaultCornerStyle,
                CornerSmoothing = WpfStyles.DefaultCornerSmoothing,
                Padding = new Thickness(24)
            };

            var layout = new Grid();
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _shortcutBox = WpfStyles.CreateTextBox(_settingsService.Current.Shortcut);
            _shortcutBox.IsReadOnly = true;
            _shortcutBox.KeyDown += OnShortcutKeyDown;

            _favoritesList = WpfStyles.CreateListBox();
            _favoritesList.DisplayMemberPath = nameof(FavoriteItem.Label);

            _maxRecentsBox = WpfStyles.CreateTextBox(_settingsService.Current.MaxRecents.ToString());
            _maxRecentsBox.PreviewTextInput += OnRecentsTextInput;
            _maxRecentsBox.LostFocus += (sender, args) => NormalizeRecents();

            _showInboxOnly = new CheckBox
            {
                Content = "Nur Unterordner von Posteingang anzeigen",
                Foreground = WpfStyles.TextPrimary,
                IsChecked = _settingsService.Current.ShowInboxOnly
            };
            _includeArchives = new CheckBox
            {
                Content = "Archiv/Online-Archive anzeigen",
                Foreground = WpfStyles.TextPrimary,
                IsChecked = _settingsService.Current.IncludeArchives
            };

            var header = BuildHeader();
            Grid.SetRow(header, 0);
            layout.Children.Add(header);

            var content = BuildContent();
            if (content is FrameworkElement contentElement)
            {
                contentElement.Margin = new Thickness(0, 16, 0, 16);
            }
            Grid.SetRow(content, 1);
            layout.Children.Add(content);

            var footer = BuildFooter();
            Grid.SetRow(footer, 2);
            layout.Children.Add(footer);

            rootBorder.Child = layout;
            Content = rootBorder;

            Loaded += (sender, args) => RefreshFavorites();
        }

        private UIElement BuildHeader()
        {
            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titleStack = new StackPanel
            {
                Orientation = Orientation.Vertical
            };
            titleStack.Children.Add(new TextBlock
            {
                Text = "Quick Move",
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                Foreground = WpfStyles.TextPrimary
            });
            titleStack.Children.Add(new TextBlock
            {
                Text = "Einstellungen",
                FontSize = 11,
                Foreground = WpfStyles.TextSecondary
            });

            var closeButton = WpfStyles.CreateIconButton("✕");
            closeButton.Click += (sender, args) => Close();
            Grid.SetColumn(closeButton, 1);

            headerGrid.Children.Add(titleStack);
            headerGrid.Children.Add(closeButton);
            headerGrid.MouseLeftButtonDown += (sender, args) =>
            {
                if (args.ButtonState == MouseButtonState.Pressed)
                {
                    DragMove();
                }
            };

            return headerGrid;
        }

        private UIElement BuildContent()
        {
            var grid = new Grid
            {
                Margin = new Thickness(0, 8, 0, 0)
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            for (int i = 0; i < 7; i++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            AddLabel(grid, "Shortcut", 0, 0);
            var shortcutCard = WpfStyles.CreateInputCard(_shortcutBox);
            shortcutCard.MinHeight = 40;
            Grid.SetRow(shortcutCard, 0);
            Grid.SetColumn(shortcutCard, 1);
            grid.Children.Add(shortcutCard);

            AddLabel(grid, "Favoriten", 0, 1);
            var favoritesCard = WpfStyles.CreateGlassCard(_favoritesList);
            favoritesCard.Height = 180;
            Grid.SetRow(favoritesCard, 1);
            Grid.SetColumn(favoritesCard, 1);
            grid.Children.Add(favoritesCard);

            var favoritesButtons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 10, 0, 10)
            };
            var addFavorite = WpfStyles.CreateSubtleButton("Hinzufügen");
            var removeFavorite = WpfStyles.CreateSubtleButton("Entfernen");
            var moveUp = WpfStyles.CreateSubtleButton("▲");
            var moveDown = WpfStyles.CreateSubtleButton("▼");
            moveUp.MinWidth = 34;
            moveDown.MinWidth = 34;
            addFavorite.Click += OnAddFavorite;
            removeFavorite.Click += OnRemoveFavorite;
            moveUp.Click += (sender, args) => MoveFavorite(-1);
            moveDown.Click += (sender, args) => MoveFavorite(1);
            favoritesButtons.Children.Add(addFavorite);
            favoritesButtons.Children.Add(removeFavorite);
            favoritesButtons.Children.Add(moveUp);
            favoritesButtons.Children.Add(moveDown);
            Grid.SetRow(favoritesButtons, 2);
            Grid.SetColumn(favoritesButtons, 1);
            grid.Children.Add(favoritesButtons);

            AddLabel(grid, "Anzahl letzte Ziele", 0, 3);
            var recentsPanel = new Grid();
            recentsPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            recentsPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var upDownPanel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(6, 0, 0, 0)
            };
            var upButton = WpfStyles.CreateSubtleButton("▲");
            var downButton = WpfStyles.CreateSubtleButton("▼");
            upButton.Width = 30;
            downButton.Width = 30;
            upButton.Click += (sender, args) => AdjustRecents(1);
            downButton.Click += (sender, args) => AdjustRecents(-1);
            upDownPanel.Children.Add(upButton);
            upDownPanel.Children.Add(downButton);
            recentsPanel.Children.Add(_maxRecentsBox);
            Grid.SetColumn(upDownPanel, 1);
            recentsPanel.Children.Add(upDownPanel);
            var recentsCard = WpfStyles.CreateInputCard(recentsPanel);
            recentsCard.MinHeight = 40;
            Grid.SetRow(recentsCard, 3);
            Grid.SetColumn(recentsCard, 1);
            grid.Children.Add(recentsCard);

            Grid.SetRow(_showInboxOnly, 4);
            Grid.SetColumn(_showInboxOnly, 1);
            _showInboxOnly.Margin = new Thickness(0, 12, 0, 0);
            grid.Children.Add(_showInboxOnly);

            Grid.SetRow(_includeArchives, 5);
            Grid.SetColumn(_includeArchives, 1);
            _includeArchives.Margin = new Thickness(0, 8, 0, 0);
            grid.Children.Add(_includeArchives);

            var refreshButton = WpfStyles.CreateSubtleButton("Ordnerliste neu laden");
            refreshButton.Click += (sender, args) => _folderService.RefreshCache();
            Grid.SetRow(refreshButton, 6);
            Grid.SetColumn(refreshButton, 1);
            refreshButton.Margin = new Thickness(0, 14, 0, 0);
            grid.Children.Add(refreshButton);

            return WpfStyles.CreateGlassCard(grid);
        }

        private UIElement BuildFooter()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 8, 0, 0)
            };
            var saveButton = WpfStyles.CreatePrimaryButton("Speichern");
            var cancelButton = WpfStyles.CreateSubtleButton("Abbrechen");
            saveButton.Click += OnSave;
            cancelButton.Click += (sender, args) => Close();
            panel.Children.Add(saveButton);
            panel.Children.Add(cancelButton);
            return panel;
        }

        private void AddLabel(Grid grid, string text, int column, int row)
        {
            var label = new TextBlock
            {
                Text = text,
                Foreground = WpfStyles.TextSecondary,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetColumn(label, column);
            Grid.SetRow(label, row);
            grid.Children.Add(label);
        }

        private void OnShortcutKeyDown(object sender, KeyEventArgs e)
        {
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            var virtualKey = KeyInterop.VirtualKeyFromKey(key);
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

            var formatted = ShortcutParser.Format(modifiers | (System.Windows.Forms.Keys)virtualKey);
            if (!string.IsNullOrWhiteSpace(formatted))
            {
                _shortcutBox.Text = formatted;
            }

            e.Handled = true;
        }

        private void OnRecentsTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = e.Text.Any(ch => !char.IsDigit(ch));
        }

        private void AdjustRecents(int delta)
        {
            NormalizeRecents();
            if (!int.TryParse(_maxRecentsBox.Text, out int value))
            {
                value = _settingsService.Current.MaxRecents;
            }

            value = Math.Max(1, Math.Min(50, value + delta));
            _maxRecentsBox.Text = value.ToString();
        }

        private void NormalizeRecents()
        {
            if (!int.TryParse(_maxRecentsBox.Text, out int value))
            {
                value = _settingsService.Current.MaxRecents;
            }

            value = Math.Max(1, Math.Min(50, value));
            _maxRecentsBox.Text = value.ToString();
        }

        private void OnAddFavorite(object sender, RoutedEventArgs e)
        {
            using (var picker = new FolderPickerForm(_folderService, new SearchService(_settingsService)))
            {
                var owner = GetDialogOwner();
                var result = owner == null ? picker.ShowDialog() : picker.ShowDialog(owner);
                if (result == System.Windows.Forms.DialogResult.OK && picker.SelectedFolder != null)
                {
                    _settingsService.AddFavorite(picker.SelectedFolder);
                    RefreshFavorites();
                }
            }
        }

        private void OnRemoveFavorite(object sender, RoutedEventArgs e)
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

            _favoritesList.ItemsSource = favorites;
        }

        private System.Windows.Forms.IWin32Window GetDialogOwner()
        {
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero)
            {
                return null;
            }

            return new Win32Window(helper.Handle);
        }

        private sealed class Win32Window : System.Windows.Forms.IWin32Window
        {
            public Win32Window(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }

        private string ResolveFolderLabel(FolderIdentifier identifier)
        {
            var info = _folderService.GetFolderByIdentifier(identifier);
            return info?.DisplayText ?? $"Unbekannter Ordner ({identifier.EntryId})";
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            if (!ShortcutParser.TryParse(_shortcutBox.Text, out _, out _))
            {
                System.Windows.Forms.MessageBox.Show(
                    "Der Shortcut ist ungültig. Bitte eine Kombination wie Ctrl+Shift+M wählen.",
                    "Quick Move",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            NormalizeRecents();
            _settingsService.Current.Shortcut = _shortcutBox.Text;
            _settingsService.Current.MaxRecents = int.Parse(_maxRecentsBox.Text);
            _settingsService.Current.ShowInboxOnly = _showInboxOnly.IsChecked ?? false;
            _settingsService.Current.IncludeArchives = _includeArchives.IsChecked ?? false;
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
