using System.Collections.ObjectModel;

namespace outlook_extension.UI.ViewModels
{
    public class FolderPickerViewModel : ViewModelBase
    {
        private readonly FolderService _folderService;
        private readonly SearchService _searchService;
        private string _searchText;
        private FolderResultItemViewModel _selectedResult;

        public FolderPickerViewModel(FolderService folderService, SearchService searchService)
        {
            _folderService = folderService;
            _searchService = searchService;
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

        public FolderInfo GetSelectedFolder()
        {
            return SelectedResult?.Info;
        }

        private void UpdateResults()
        {
            var folders = _folderService.GetCachedFolders();
            Results.Clear();

            var matches = _searchService.Search(SearchText ?? string.Empty, folders);
            foreach (var folder in matches)
            {
                Results.Add(new FolderResultItemViewModel(folder));
            }

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
