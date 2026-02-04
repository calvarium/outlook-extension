using System;
using System.Collections.ObjectModel;

namespace outlook_extension.UI.ViewModels
{
    public class QuickMoveViewModel : ViewModelBase
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private readonly Func<FolderInfo, bool, bool> _moveSelection;
        private readonly Action _undoAction;
        private string _searchText;
        private string _statusText;
        private FolderResultItemViewModel _selectedResult;

        public QuickMoveViewModel(
            FolderService folderService,
            SearchService searchService,
            Func<FolderInfo, bool, bool> moveSelection,
            Action undoAction)
        {
            _folderService = folderService;
            _searchService = searchService;
            _moveSelection = moveSelection;
            _undoAction = undoAction;
            Results = new ObservableCollection<FolderResultItemViewModel>();
        }

        public ObservableCollection<FolderResultItemViewModel> Results { get; }

        public FolderResultItemViewModel SelectedResult
        {
            get => _selectedResult;
            set
            {
                if (SetField(ref _selectedResult, value))
                {
                    UpdateSelectionFlags();
                }
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetField(ref _searchText, value))
                {
                    UpdateResults();
                }
            }
        }

        public string StatusText
        {
            get => _statusText;
            set => SetField(ref _statusText, value);
        }

        public void Initialize()
        {
            UpdateResults();
        }

        public void MoveSelection(int offset)
        {
            if (Results.Count == 0)
            {
                return;
            }

            var currentIndex = SelectedResult == null ? -1 : Results.IndexOf(SelectedResult);
            var newIndex = currentIndex + offset;
            if (newIndex < 0)
            {
                newIndex = 0;
            }

            if (newIndex >= Results.Count)
            {
                newIndex = Results.Count - 1;
            }

            SelectedResult = Results[newIndex];
        }

        public void ClearSearch()
        {
            SearchText = string.Empty;
        }

        public bool ExecuteSelection(bool keepDialogOpen)
        {
            if (SelectedResult?.Info == null)
            {
                return false;
            }

            return _moveSelection?.Invoke(SelectedResult.Info, keepDialogOpen) ?? false;
        }

        public void UndoLastMove()
        {
            _undoAction?.Invoke();
        }

        public void RefreshResults()
        {
            UpdateResults();
        }

        private void UpdateResults()
        {
            var folders = _folderService.GetCachedFolders();
            Results.Clear();

            if (folders.Count == 0)
            {
                StatusText = "Keine Ordner im Cache. Bitte Ordnerliste aktualisieren.";
                SelectedResult = null;
                return;
            }

            var matches = _searchService.Search(SearchText ?? string.Empty, folders);
            foreach (var folder in matches)
            {
                Results.Add(new FolderResultItemViewModel(folder));
            }

            StatusText = $"{Results.Count} Treffer";
            SelectedResult = Results.Count > 0 ? Results[0] : null;
        }

        private void UpdateSelectionFlags()
        {
            foreach (var item in Results)
            {
                item.IsSelected = item == SelectedResult;
            }
        }
    }
}
