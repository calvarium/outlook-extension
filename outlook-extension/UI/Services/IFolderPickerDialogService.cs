using System.Windows;

namespace outlook_extension.UI.Services
{
    public interface IFolderPickerDialogService
    {
        FolderInfo PickFolder(Window owner);
    }
}
