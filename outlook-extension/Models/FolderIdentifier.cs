using System;
using System.Runtime.Serialization;

namespace outlook_extension
{
    [DataContract]
    public class FolderIdentifier : IEquatable<FolderIdentifier>
    {
        [DataMember(Order = 1)]
        public string EntryId { get; set; }

        [DataMember(Order = 2)]
        public string StoreId { get; set; }

        public bool Equals(FolderIdentifier other)
        {
            if (other == null)
            {
                return false;
            }

            return string.Equals(EntryId, other.EntryId, StringComparison.OrdinalIgnoreCase)
                && string.Equals(StoreId, other.StoreId, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as FolderIdentifier);
        }

        public override int GetHashCode()
        {
            return (EntryId ?? string.Empty).GetHashCode() ^ (StoreId ?? string.Empty).GetHashCode();
        }
    }
}
