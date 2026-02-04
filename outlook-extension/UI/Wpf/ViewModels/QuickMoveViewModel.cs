using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using outlook_extension.UI.Wpf.Infrastructure;

namespace outlook_extension.UI.Wpf.ViewModels
{
    public class QuickMoveViewModel : ViewModelBase
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private readonly ThisAddIn _addIn;
        private string _searchText;
        private string _statusText;
        private SearchResultViewModel _selectedResult;

        public QuickMoveViewModel(FolderService folderService, SearchService searchService, ThisAddIn addIn)
        {
            _folderService = folderService;
            _searchService = searchService;
            _addIn = addIn;

            Results = new ObservableCollection<SearchResultViewModel>();
            SelectResultCommand = new RelayCommand<SearchResultViewModel>(SelectResult);
            ExecuteSelectedCommand = new RelayCommand(ExecuteSelected, () => SelectedResult != null);

            UpdateResults();
        }

        public ObservableCollection<SearchResultViewModel> Results { get; }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                {
                    UpdateResults();
                }
            }
        }

        public string StatusText
        {
            get => _statusText;
            private set => SetProperty(ref _statusText, value);
        }

        public SearchResultViewModel SelectedResult
        {
            get => _selectedResult;
            private set
            {
                if (SetProperty(ref _selectedResult, value))
                {
                    ExecuteSelectedCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public RelayCommand<SearchResultViewModel> SelectResultCommand { get; }

        public RelayCommand ExecuteSelectedCommand { get; }

        public Action CloseRequested { get; set; }

        public void HandleNavigation(int delta)
        {
            if (Results.Count == 0)
            {
                return;
            }

            var currentIndex = SelectedResult == null ? -1 : Results.IndexOf(SelectedResult);
            var newIndex = currentIndex + delta;
            if (newIndex < 0)
            {
                newIndex = 0;
            }
            else if (newIndex >= Results.Count)
            {
                newIndex = Results.Count - 1;
            }

            SelectResult(Results[newIndex]);
        }

        public void HandleEscape()
        {
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                SearchText = string.Empty;
                return;
            }

            CloseRequested?.Invoke();
        }

        public void UndoLastMove()
        {
            _addIn.UndoLastMove();
        }

        public void ExecuteSelected(bool keepDialogOpen)
        {
            if (SelectedResult?.Info == null)
            {
                return;
            }

            var moved = _addIn.MoveSelectionToFolder(SelectedResult.Info, keepDialogOpen);
            if (!moved)
            {
                return;
            }

            if (keepDialogOpen)
            {
                SearchText = string.Empty;
                return;
            }

            CloseRequested?.Invoke();
        }

        private void ExecuteSelected()
        {
            ExecuteSelected(false);
        }

        private void SelectResult(SearchResultViewModel result)
        {
            if (result == null)
            {
                return;
            }

            if (SelectedResult != null)
            {
                SelectedResult.IsSelected = false;
            }

            SelectedResult = result;
            SelectedResult.IsSelected = true;
        }

        private void UpdateResults()
        {
            var folders = _folderService.GetCachedFolders();
            if (folders.Count == 0)
            {
                Results.Clear();
                StatusText = "Keine Ordner im Cache. Bitte Ordnerliste aktualisieren.";
                SelectedResult = null;
                return;
            }

            var matches = _searchService.Search(SearchText ?? string.Empty, folders);
            var previous = SelectedResult?.Info;
            Results.Clear();
            foreach (var match in matches)
            {
                Results.Add(new SearchResultViewModel(match));
            }

            StatusText = $"{Results.Count} Treffer";

            if (Results.Count == 0)
            {
                SelectedResult = null;
                return;
            }

            var selected = Results.FirstOrDefault(item => previous != null && item.Info.EntryId == previous.EntryId)
                ?? Results.First();
            SelectResult(selected);
        }
    }
}
