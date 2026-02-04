using outlook_extension.UI.Wpf.Infrastructure;

namespace outlook_extension.UI.Wpf.ViewModels
{
    public class SearchResultViewModel : ViewModelBase
    {
        private bool _isSelected;

        public SearchResultViewModel(FolderInfo info)
        {
            Info = info;
            Title = info?.DisplayName ?? string.Empty;
            Subtitle = info?.FullPath ?? info?.DisplayText ?? string.Empty;
        }

        public FolderInfo Info { get; }

        public string Title { get; }

        public string Subtitle { get; }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}
