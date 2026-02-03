namespace outlook_extension
{
    public class FolderInfo
    {
        public string EntryId { get; set; }

        public string StoreId { get; set; }

        public string DisplayName { get; set; }

        public string MailboxName { get; set; }

        public string FolderPath { get; set; }

        public string FullPath { get; set; }

        public bool IsUnderInbox { get; set; }

        public FolderIdentifier Identifier => new FolderIdentifier
        {
            EntryId = EntryId,
            StoreId = StoreId
        };

        public string DisplayText => $"{MailboxName} > {FolderPath}";
    }
}
