using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Animation;
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
        private readonly ProgressBar _progressBar;
        private readonly FrameworkElement _checkIcon;

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
            _progressBar = new ProgressBar { Width = 140, Height = 12, Margin = new Thickness(8, 10, 8, 8), Visibility = Visibility.Collapsed, Foreground = WpfStyles.AccentBackground, Background = WpfStyles.GlassBackground };
            _checkIcon = CreateCheckIcon();
            _checkIcon.Visibility = Visibility.Collapsed;
            _statusBar.Children.Add(_spinner);
            _statusBar.Children.Add(_progressBar);
            _statusBar.Children.Add(_checkIcon);
            _statusBar.Children.Add(_statusText);
            Grid.SetRow(_statusBar, 2);
            layout.Children.Add(_statusBar);

            // keep status bar visible permanently (it will update contents based on refresh state)
            _statusBar.Visibility = Visibility.Visible;

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
                         _folderService.FullRefreshCompleted += OnFullRefreshCompleted;
                         _eventsSubscribed = true;
                     }
                     catch { }
                 }

                 _searchBox.Focus();
                 UpdateResults();

                 // initialize status display from folder service current state
                 try
                 {
                     if (_folderService.IsRefreshing)
                     {
                         OnRefreshingChanged(true);

                         var lastProcessed = _folderService.LastProgressProcessed;
                         var lastTotal = _folderService.LastProgressTotal;
                         if (lastTotal > 0)
                         {
                             OnProgressUpdated(lastProcessed, lastTotal);
                         }
                     }

                     // If cache is empty and no refresh in progress, trigger a refresh and show spinner
                     var folders = _folderService.GetCachedFolders();
                     if ((folders == null || folders.Count == 0) && !_folderService.IsRefreshing)
                     {
                         try
                         {
                             OnRefreshingChanged(true);
                             _folderService.RefreshCache();
                         }
                         catch { }
                     }
                 }
                 catch { }
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
                Stroke = WpfStyles.TextPrimary,
                StrokeThickness = 3,
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

        private FrameworkElement CreateCheckIcon()
        {
            // Use a vector path for the checkmark to avoid font glyph issues
            var path = new System.Windows.Shapes.Path
            {
                Stroke = new SolidColorBrush(Color.FromRgb(88, 196, 110)),
                StrokeThickness = 2.5,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round,
                StrokeLineJoin = PenLineJoin.Round,
                Width = 18,
                Height = 18,
                Margin = new Thickness(8, 4, 8, 4),
                VerticalAlignment = VerticalAlignment.Center,
                Data = Geometry.Parse("M2,10 L7,15 L16,4")
            };

            // Put path in a Viewbox so it scales nicely with layout
            var box = new Viewbox
            {
                Width = 18,
                Height = 18,
                Child = path,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 4, 4, 4)
            };

            return box;
        }

        private void UnsubscribeEvents()
        {
            if (!_eventsSubscribed) return;
            try
            {
                _folderService.CacheUpdated -= OnCacheUpdated;
                _folderService.RefreshingChanged -= OnRefreshingChanged;
                _folderService.ProgressUpdated -= OnProgressUpdated;
                _folderService.FullRefreshCompleted -= OnFullRefreshCompleted;
                _eventsSubscribed = false;
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
            if (_statusText == null || _spinner == null || _progressBar == null || _checkIcon == null || _statusBar == null)
            {
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (isRefreshing)
                {
                    // show running state
                    _statusBar.Visibility = Visibility.Visible;
                    _statusText.Text = "Aktualisiere Ordner…";
                    _spinner.Visibility = Visibility.Visible;
                    _checkIcon.Visibility = Visibility.Collapsed;
                    _progressBar.IsIndeterminate = true;
                    _progressBar.Visibility = Visibility.Visible;
                }
                else
                {
                    // show success state but keep status bar visible
                    _spinner.Visibility = Visibility.Collapsed;
                    _progressBar.IsIndeterminate = false;
                    _progressBar.Visibility = Visibility.Collapsed;
                    _progressBar.Value = 0;

                    _checkIcon.Visibility = Visibility.Visible;
                    _statusText.Text = "Auf dem aktuellen Stand";

                    _statusBar.Visibility = Visibility.Visible;
                }
            }), DispatcherPriority.Background);
        }

        private void OnProgressUpdated(int processed, int total)
        {
            // guard in case events fire before UI elements are initialized
            if (_statusText == null || _progressBar == null || _statusBar == null)
            {
                return;
            }

            // Only update progress UI while a refresh is running
            if (!_folderService.IsRefreshing)
            {
                return;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    if (_statusText == null) return;
                    if (total > 0)
                    {
                        _statusBar.Visibility = Visibility.Visible;
                        _checkIcon.Visibility = Visibility.Collapsed;
                        _progressBar.IsIndeterminate = false;
                        _progressBar.Minimum = 0;
                        _progressBar.Maximum = total;
                        _progressBar.Value = processed;
                        _progressBar.Visibility = Visibility.Visible;
                        _statusText.Text = $"Ordner geladen: {processed}/{total}";
                    }
                }
                catch
                {
                    // ignore UI update failures
                }
            }), DispatcherPriority.Background);
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
        
        private void OnFullRefreshCompleted()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                // show success state only when the entire refresh completed
                _spinner.Visibility = Visibility.Collapsed;
                _progressBar.IsIndeterminate = false;
                _progressBar.Visibility = Visibility.Collapsed;
                _progressBar.Value = 0;

                _checkIcon.Visibility = Visibility.Visible;
                _statusText.Text = "Auf dem aktuellen Stand";
                _statusBar.Visibility = Visibility.Visible;
            }), DispatcherPriority.Background);
        }
    }
}
