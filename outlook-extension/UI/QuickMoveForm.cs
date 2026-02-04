using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace outlook_extension
{
    public class QuickMoveForm : Form
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private readonly ThisAddIn _addIn;
        private readonly TextBox _searchBox;
        private readonly ListBox _resultsList;
        private readonly Label _statusLabel;
        private List<FolderInfo> _currentResults = new List<FolderInfo>();

        public QuickMoveForm(
            FolderService folderService,
            SearchService searchService,
            ThisAddIn addIn)
        {
            _folderService = folderService;
            _searchService = searchService;
            _addIn = addIn;

            Text = "Quick Move";
            Width = 640;
            Height = 460;
            MaximizeBox = false;
            MinimizeBox = false;
            GlassStyle.ApplyFormStyle(this);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 56
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
                Text = "Schnell verschieben, klar & fokussiert",
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

            var searchContainer = GlassStyle.CreateInputPanel();
            searchContainer.MinimumSize = new Size(0, 44);
            _searchBox = new TextBox
            {
                Dock = DockStyle.Fill
            };
            GlassStyle.StyleTextInput(_searchBox);
            searchContainer.Controls.Add(_searchBox);

            _resultsList = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = nameof(FolderInfo.DisplayText)
            };
            GlassStyle.StyleListBox(_resultsList);
            _resultsList.DrawItem += OnResultsDrawItem;

            var resultsCard = GlassStyle.CreateGlassCard(18, 10);
            resultsCard.Controls.Add(_resultsList);

            _statusLabel = new Label
            {
                Dock = DockStyle.Fill,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(210, 255, 255, 255)
            };

            layout.Controls.Add(headerPanel, 0, 0);
            layout.Controls.Add(searchContainer, 0, 1);
            layout.Controls.Add(resultsCard, 0, 2);
            layout.Controls.Add(_statusLabel, 0, 3);

            Controls.Add(layout);

            _searchBox.TextChanged += OnSearchTextChanged;
            _searchBox.KeyDown += OnSearchBoxKeyDown;
            _resultsList.KeyDown += OnResultsKeyDown;
            _resultsList.KeyPress += OnResultsKeyPress;
            _resultsList.DoubleClick += OnResultsDoubleClick;
            Shown += OnShown;
        }

        private void OnShown(object sender, EventArgs e)
        {
            _searchBox.Focus();
            UpdateResults();
        }

        private void OnSearchTextChanged(object sender, EventArgs e)
        {
            UpdateResults();
        }

        private void UpdateResults()
        {
            var folders = _folderService.GetCachedFolders();
            if (folders.Count == 0)
            {
                _statusLabel.Text = "Keine Ordner im Cache. Bitte Ordnerliste aktualisieren.";
                _resultsList.DataSource = null;
                return;
            }

            _currentResults = _searchService.Search(_searchBox.Text, folders);
            _resultsList.DataSource = _currentResults;
            if (_currentResults.Count > 0)
            {
                _resultsList.SelectedIndex = 0;
            }

            _statusLabel.Text = $"{_currentResults.Count} Treffer";
        }

        private void OnSearchBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Back)
            {
                DeletePreviousWord();
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.Z)
            {
                _addIn.UndoLastMove();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down && _resultsList.Items.Count > 0)
            {
                _resultsList.Focus();
                _resultsList.SelectedIndex = Math.Min(1, _resultsList.Items.Count - 1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                MoveSelectedFolder(e.Control);
                e.Handled = true;
            }
        }

        private void OnResultsKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.Back)
            {
                _searchBox.Focus();
                DeletePreviousWord();
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.Z)
            {
                _addIn.UndoLastMove();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                MoveSelectedFolder(e.Control);
                e.Handled = true;
            }
        }

        private void OnResultsKeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            _searchBox.Focus();
            _searchBox.AppendText(e.KeyChar.ToString());
            _searchBox.SelectionStart = _searchBox.Text.Length;
            e.Handled = true;
        }

        private void OnResultsDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            var item = _resultsList.Items[e.Index];
            var text = item is FolderInfo info ? info.DisplayText : item?.ToString();
            var bounds = e.Bounds;
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using (var background = new SolidBrush(selected
                       ? Color.FromArgb(120, 100, 170, 255)
                       : Color.FromArgb(20, 255, 255, 255)))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillRectangle(background, bounds);
            }

            TextRenderer.DrawText(
                e.Graphics,
                text,
                _resultsList.Font,
                new Rectangle(bounds.X + 10, bounds.Y + 6, bounds.Width - 20, bounds.Height - 12),
                Color.White,
                TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter);
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

        private void OnResultsDoubleClick(object sender, EventArgs e)
        {
            MoveSelectedFolder(false);
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
