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

        public string DisplayPath\n+        {\n+            get\n+            {\n+                if (Info == null)\n+                {\n+                    return string.Empty;\n+                }\n+\n+                var mailbox = Info.MailboxName ?? string.Empty;\n+                var path = Info.FolderPath ?? string.Empty;\n+                return string.IsNullOrWhiteSpace(mailbox)\n+                    ? path\n+                    : $\"{mailbox} {path}\".Trim();\n+            }\n+        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetField(ref _isSelected, value);
        }
    }
}
