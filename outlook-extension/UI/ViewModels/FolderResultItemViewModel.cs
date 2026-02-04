namespace outlook_extension.UI.ViewModels
{
    public class FolderResultItemViewModel : ViewModelBase
    {
        private bool _isSelected;

        public FolderResultItemViewModel(FolderInfo info)
        {
            Info = info;
        }

        public FolderInfo Info { get; }

        public string DisplayName => Info?.DisplayName ?? string.Empty;

        public string DisplayPath
        {
            get
            {
                if (Info == null)
                {
                    return string.Empty;
                }

                var mailbox = Info.MailboxName ?? string.Empty;
                var path = Info.FolderPath ?? string.Empty;
                return string.IsNullOrWhiteSpace(mailbox)
                    ? path
                    : $"{mailbox} {path}".Trim();
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetField(ref _isSelected, value);
        }
    }
}
