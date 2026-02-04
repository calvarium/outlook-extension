using System.Collections.Generic;
using System.Runtime.Serialization;

namespace outlook_extension
{
    [DataContract]
    public class SettingsModel
    {
        [DataMember(Order = 1)]
        public string Shortcut { get; set; } = "Ctrl+Shift+M";

        [DataMember(Order = 2)]
        public List<FolderIdentifier> Favorites { get; set; } = new List<FolderIdentifier>();

        [DataMember(Order = 3)]
        public List<FolderIdentifier> Recents { get; set; } = new List<FolderIdentifier>();

        [DataMember(Order = 4)]
        public int MaxRecents { get; set; } = 10;

        [DataMember(Order = 5)]
        public bool ShowInboxOnly { get; set; }

        [DataMember(Order = 6)]
        public bool IncludeArchives { get; set; } = true;

        [DataMember(Order = 7)]
        public string LogLevel { get; set; } = "Info";

        [DataMember(Order = 8)]
        public string ThemeMode { get; set; } = "System";
    }
}
