using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace outlook_extension
{
    public class FolderPickerForm : Form
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private readonly TextBox _searchBox;
        private readonly ListBox _resultsList;
        private List<FolderInfo> _currentResults = new List<FolderInfo>();

        public FolderInfo SelectedFolder { get; private set; }

        public FolderPickerForm(FolderService folderService, SearchService searchService)
        {
            _folderService = folderService;
            _searchService = searchService;

            Text = "Favoritenordner wählen";
            Width = 560;
            Height = 380;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;

            _searchBox = new TextBox { Dock = DockStyle.Top };
            _resultsList = new ListBox { Dock = DockStyle.Fill, DisplayMember = nameof(FolderInfo.DisplayText) };

            Controls.Add(_resultsList);
            Controls.Add(_searchBox);

            _searchBox.TextChanged += OnSearchTextChanged;
            _searchBox.KeyDown += OnSearchBoxKeyDown;
            _resultsList.KeyDown += OnResultsKeyDown;
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
            _currentResults = _searchService.Search(_searchBox.Text, folders);
            _resultsList.DataSource = _currentResults;
            if (_currentResults.Count > 0)
            {
                _resultsList.SelectedIndex = 0;
            }
        }

        private void OnSearchBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                ConfirmSelection();
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
                ConfirmSelection();
                e.Handled = true;
            }
        }

        private void OnResultsDoubleClick(object sender, EventArgs e)
        {
            ConfirmSelection();
        }

        private void ConfirmSelection()
        {
            SelectedFolder = _resultsList.SelectedItem as FolderInfo;
            if (SelectedFolder != null)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}
