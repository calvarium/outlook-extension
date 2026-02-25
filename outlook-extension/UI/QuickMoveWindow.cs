using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Threading;

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
        private bool _isClosing;
        private bool _eventsSubscribed;

        private readonly StackPanel _statusBar;
        private readonly TextBlock _statusText;
        private readonly FrameworkElement _spinner;

        public QuickMoveWindow(FolderService folderService, SearchService searchService, ThisAddIn addIn)
        {
            _folderService = folderService;
            _searchService = searchService;
            _addIn = addIn;

            // DO NOT subscribe to folder events here — subscribe after UI is initialized in Loaded to avoid races

            Width = 640;
            Height = 360;
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
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _searchBox = WpfStyles.CreateTextBox();
            _searchBox.TextChanged += OnSearchTextChanged;
            _searchBox.PreviewKeyDown += OnSearchBoxKeyDown;
            var searchCard = WpfStyles.CreateInputCard(_searchBox);
            searchCard.Margin = new Thickness(0, 0, 0, 16);
            Grid.SetRow(searchCard, 0);
            layout.Children.Add(searchCard);

            _resultsList = WpfStyles.CreateListBox();
            _resultsList.KeyDown += OnResultsKeyDown;
            _resultsList.PreviewTextInput += OnResultsTextInput;
            _resultsList.MouseDoubleClick += (sender, args) => MoveSelectedFolder(false);
            _resultsList.SelectionChanged += (sender, args) => _searchBox.Focus();
            _resultsList.PreviewMouseDown += (sender, args) => _searchBox.Focus();

            // set custom item template
            var dataTemplate = new DataTemplate(typeof(FolderInfo));
            var stackFactory = new FrameworkElementFactory(typeof(StackPanel));
            stackFactory.SetValue(StackPanel.OrientationProperty, Orientation.Vertical);
            stackFactory.SetValue(StackPanel.MarginProperty, new Thickness(6));

            var titleFactory = new FrameworkElementFactory(typeof(TextBlock));
            titleFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("DisplayName"));
            titleFactory.SetValue(TextBlock.FontSizeProperty, 14.0);
            titleFactory.SetValue(TextBlock.FontWeightProperty, FontWeights.SemiBold);
            titleFactory.SetValue(TextBlock.ForegroundProperty, WpfStyles.TextPrimary);
            titleFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);
            // uppercase converter
            var converter = new UppercaseConverter();
            titleFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("DisplayName") { Converter = converter });

            var pathFactory = new FrameworkElementFactory(typeof(TextBlock));
            pathFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("FullPath"));
            pathFactory.SetValue(TextBlock.FontSizeProperty, 12.0);
            pathFactory.SetValue(TextBlock.ForegroundProperty, WpfStyles.TextSecondary);
            pathFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);

            stackFactory.AppendChild(titleFactory);
            stackFactory.AppendChild(pathFactory);

            dataTemplate.VisualTree = stackFactory;
            _resultsList.ItemTemplate = dataTemplate;

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

            // status bar with spinner
            _statusBar = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Left };
            _statusText = new TextBlock { Foreground = WpfStyles.TextSecondary, Margin = new Thickness(8, 6, 8, 6) };
            _spinner = CreateSpinner();
            _spinner.Visibility = Visibility.Collapsed;
            _statusBar.Children.Add(_spinner);
            _statusBar.Children.Add(_statusText);
            Grid.SetRow(_statusBar, 2);
            layout.Children.Add(_statusBar);

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
                // subscribe to events only after UI elements exist
                if (!_eventsSubscribed)
                {
                    try
                    {
                        _folderService.CacheUpdated += OnCacheUpdated;
                        _folderService.RefreshingChanged += OnRefreshingChanged;
                        _folderService.ProgressUpdated += OnProgressUpdated;
                        _eventsSubscribed = true;
                    }
                    catch { }
                }

                _searchBox.Focus();
                UpdateResults();
            };

            Closing += (sender, args) => _isClosing = true;
            Deactivated += (sender, args) => CloseOnDeactivate();
            Closed += (sender, args) => UnsubscribeEvents();
        }

        private FrameworkElement CreateSpinner()
        {
            var ellipse = new System.Windows.Shapes.Ellipse
            {
                Width = 16,
                Height = 16,
                Stroke = WpfStyles.TextSecondary,
                StrokeThickness = 2,
                Margin = new Thickness(8, 6, 8, 6)
            };

            var rotate = new System.Windows.Media.RotateTransform();
            ellipse.RenderTransform = rotate;
            ellipse.RenderTransformOrigin = new Point(0.5, 0.5);

            var animation = new System.Windows.Media.Animation.DoubleAnimation(0, 360, new Duration(TimeSpan.FromSeconds(1)))
            {
                RepeatBehavior = System.Windows.Media.Animation.RepeatBehavior.Forever
            };

            rotate.BeginAnimation(System.Windows.Media.RotateTransform.AngleProperty, animation);

            return ellipse;
        }

        private void UnsubscribeEvents()
        {
            if (!_eventsSubscribed) return;
            try
            {
                _folderService.CacheUpdated -= OnCacheUpdated;
                _folderService.RefreshingChanged -= OnRefreshingChanged;
                _folderService.ProgressUpdated -= OnProgressUpdated;
            }
            catch
            {
                // ignore
            }
            finally
            {
                _eventsSubscribed = false;
            }
        }

        private void OnRefreshingChanged(bool isRefreshing)
        {
            // guard in case events fire before UI elements are initialized
            if (_statusText == null || _spinner == null)
            {
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                _statusText.Text = isRefreshing ? "Aktualisiere Ordner…" : string.Empty;
                _spinner.Visibility = isRefreshing ? Visibility.Visible : Visibility.Collapsed;
                if (!isRefreshing)
                {
                    // no-op
                }
            }), DispatcherPriority.Background);
        }

        private void OnProgressUpdated(int processed, int total)
        {
            // guard in case events fire before UI elements are initialized
            if (_statusText == null)
            {
                return;
            }

            try
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        if (_statusText == null) return;
                        if (total > 0)
                        {
                            _statusText.Text = $"Ordner geladen: {processed}/{total}";
                        }
                    }
                    catch
                    {
                        // ignore UI update failures
                    }
                }), DispatcherPriority.Background);
            }
            catch
            {
                // ignore dispatcher exceptions
            }
        }

        private void OnCacheUpdated()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (!_isClosing)
                {
                    UpdateResults();
                }
            }), DispatcherPriority.Background);
        }

        private void CloseOnDeactivate()
        {
            if (_isClosing)
            {
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_isClosing || !IsVisible)
                {
                    return;
                }

                _isClosing = true;
                Close();
            }), DispatcherPriority.Background);
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

                _isClosing = true;
                Close();
            }
        }
    }
}
