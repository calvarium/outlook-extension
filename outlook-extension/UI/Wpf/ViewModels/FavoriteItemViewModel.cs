using outlook_extension.UI.Wpf.Infrastructure;

namespace outlook_extension.UI.Wpf.ViewModels
{
    public class FavoriteItemViewModel : ViewModelBase
    {
        private bool _isSelected;

        public FavoriteItemViewModel(FolderIdentifier identifier, string label)
        {
            Identifier = identifier;
            Label = label;
        }

        public FolderIdentifier Identifier { get; }

        public string Label { get; }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}
