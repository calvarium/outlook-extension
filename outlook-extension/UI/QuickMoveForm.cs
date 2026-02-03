using System;
using System.Collections.Generic;
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
            Width = 600;
            Height = 420;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;

            _searchBox = new TextBox
            {
                Dock = DockStyle.Top
            };

            _resultsList = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = nameof(FolderInfo.DisplayText)
            };

            _statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            Controls.Add(_resultsList);
            Controls.Add(_searchBox);
            Controls.Add(_statusLabel);

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
            if (e.KeyCode == Keys.Down && _resultsList.Items.Count > 0)
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
