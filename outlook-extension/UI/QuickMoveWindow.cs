using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace outlook_extension
{
    public class QuickMoveWindow : Window
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private readonly ThisAddIn _addIn;
        private readonly TextBox _searchBox;
        private readonly ListBox _resultsList;
        private List<FolderInfo> _currentResults = new List<FolderInfo>();

        public QuickMoveWindow(FolderService folderService, SearchService searchService, ThisAddIn addIn)
        {
            _folderService = folderService;
            _searchService = searchService;
            _addIn = addIn;

            Width = 640;
            Height = 360;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            var rootBorder = new Border
            {
                Background = WpfStyles.GlassBackground,
                CornerRadius = new CornerRadius(22),
                Padding = new Thickness(22),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    BlurRadius = 18,
                    Opacity = 0.4,
                    ShadowDepth = 0
                }
            };

            var layout = new Grid();
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            _searchBox = WpfStyles.CreateTextBox();
            _searchBox.TextChanged += OnSearchTextChanged;
            _searchBox.PreviewKeyDown += OnSearchBoxKeyDown;
            var searchCard = WpfStyles.CreateInputCard(_searchBox);
            searchCard.Margin = new Thickness(0, 0, 0, 16);
            Grid.SetRow(searchCard, 0);
            layout.Children.Add(searchCard);

            _resultsList = WpfStyles.CreateListBox();
            _resultsList.DisplayMemberPath = nameof(FolderInfo.DisplayText);
            _resultsList.KeyDown += OnResultsKeyDown;
            _resultsList.PreviewTextInput += OnResultsTextInput;
            _resultsList.MouseDoubleClick += (sender, args) => MoveSelectedFolder(false);
            _resultsList.SelectionChanged += (sender, args) => _searchBox.Focus();
            _resultsList.PreviewMouseDown += (sender, args) => _searchBox.Focus();

            var listCard = WpfStyles.CreateGlassCard(_resultsList);
            listCard.MouseLeftButtonDown += (sender, args) =>
            {
                if (_resultsList.Items.Count > 0)
                {
                    _resultsList.SelectedIndex = Math.Max(_resultsList.SelectedIndex, 0);
                    _searchBox.Focus();
                }
            };
            Grid.SetRow(listCard, 1);
            layout.Children.Add(listCard);

            rootBorder.Child = layout;
            rootBorder.MouseLeftButtonDown += (sender, args) =>
            {
                if (args.ButtonState == MouseButtonState.Pressed)
                {
                    DragMove();
                }
            };
            Content = rootBorder;

            Loaded += (sender, args) =>
            {
                _searchBox.Focus();
                UpdateResults();
            };
            Deactivated += (sender, args) => Close();
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateResults();
        }

        private void UpdateResults()
        {
            var folders = _folderService.GetCachedFolders();
            if (folders.Count == 0)
            {
                _resultsList.ItemsSource = null;
                return;
            }

            _currentResults = _searchService.Search(_searchBox.Text, folders);
            _resultsList.ItemsSource = _currentResults;
            if (_currentResults.Count > 0)
            {
                _resultsList.SelectedIndex = 0;
            }

        }

        private void OnSearchBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && e.Key == Key.Back)
            {
                DeletePreviousWord();
                e.Handled = true;
            }
            else if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && e.Key == Key.Z)
            {
                _addIn.UndoLastMove();
                e.Handled = true;
            }
            else if (e.Key == Key.Down && _resultsList.Items.Count > 0)
            {
                _resultsList.SelectedIndex = Math.Min(_resultsList.SelectedIndex + 1, _resultsList.Items.Count - 1);
                _resultsList.ScrollIntoView(_resultsList.SelectedItem);
                e.Handled = true;
            }
            else if (e.Key == Key.Up && _resultsList.Items.Count > 0)
            {
                _resultsList.SelectedIndex = Math.Max(_resultsList.SelectedIndex - 1, 0);
                _resultsList.ScrollIntoView(_resultsList.SelectedItem);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                Close();
            }
            else if (e.Key == Key.Enter)
            {
                MoveSelectedFolder(Keyboard.Modifiers.HasFlag(ModifierKeys.Control));
                e.Handled = true;
            }
        }

        private void OnResultsKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
            }
            else if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && e.Key == Key.Back)
            {
                _searchBox.Focus();
                DeletePreviousWord();
                e.Handled = true;
            }
            else if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && e.Key == Key.Z)
            {
                _addIn.UndoLastMove();
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                MoveSelectedFolder(Keyboard.Modifiers.HasFlag(ModifierKeys.Control));
                e.Handled = true;
            }
        }

        private void OnResultsTextInput(object sender, TextCompositionEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Text))
            {
                return;
            }

            _searchBox.Focus();
            _searchBox.Text += e.Text;
            _searchBox.SelectionStart = _searchBox.Text.Length;
            e.Handled = true;
        }

        private void DeletePreviousWord()
        {
            var text = _searchBox.Text;
            var caret = _searchBox.SelectionStart;
            if (caret == 0)
            {
                return;
            }

            var deleteFrom = caret - 1;
            while (deleteFrom > 0 && char.IsWhiteSpace(text[deleteFrom]))
            {
                deleteFrom--;
            }

            while (deleteFrom > 0 && !char.IsWhiteSpace(text[deleteFrom - 1]))
            {
                deleteFrom--;
            }

            _searchBox.Text = text.Remove(deleteFrom, caret - deleteFrom);
            _searchBox.SelectionStart = deleteFrom;
        }

        private void MoveSelectedFolder(bool keepDialogOpen)
        {
            var selected = _resultsList.SelectedItem as FolderInfo;
            if (selected == null)
            {
                return;
            }

            var moved = _addIn.MoveSelectionToFolder(selected, keepDialogOpen);
            if (moved)
            {
                if (keepDialogOpen)
                {
                    _searchBox.SelectAll();
                    _searchBox.Focus();
                    UpdateResults();
                    return;
                }

                Close();
            }
        }
    }
}
