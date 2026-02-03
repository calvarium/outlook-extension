using System;
using Office = Microsoft.Office.Core;

namespace outlook_extension
{
    public class QuickMoveRibbon : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;

        public QuickMoveRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""QuickMoveTab"" label=""Quick Move"">
        <group id=""QuickMoveGroup"" label=""Quick Move"">
          <button id=""QuickMoveOpen"" label=""Quick Move öffnen"" size=""large"" onAction=""OnOpenQuickMove"" />
          <button id=""QuickMoveSettings"" label=""Einstellungen"" onAction=""OnOpenSettings"" />
          <dropDown id=""QuickMoveRecents"" label=""Letzte Ziele"" getItemCount=""GetRecentCount"" getItemLabel=""GetRecentLabel"" onAction=""OnRecentSelected"" />
          <dropDown id=""QuickMoveFavorites"" label=""Favoriten"" getItemCount=""GetFavoriteCount"" getItemLabel=""GetFavoriteLabel"" onAction=""OnFavoriteSelected"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnOpenQuickMove(Office.IRibbonControl control)
        {
            _addIn.OpenQuickMoveDialog();
        }

        public void OnOpenSettings(Office.IRibbonControl control)
        {
            _addIn.OpenSettingsDialog();
        }

        public int GetRecentCount(Office.IRibbonControl control)
        {
            return _addIn.SettingsService.Current.Recents.Count;
        }

        public string GetRecentLabel(Office.IRibbonControl control, int index)
        {
            return ResolveFolderLabel(_addIn.SettingsService.Current.Recents, index);
        }

        public void OnRecentSelected(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var folder = ResolveFolderInfo(_addIn.SettingsService.Current.Recents, selectedIndex);
            if (folder != null)
            {
                _addIn.MoveSelectionToFolder(folder, false);
            }
        }

        public int GetFavoriteCount(Office.IRibbonControl control)
        {
            return _addIn.SettingsService.Current.Favorites.Count;
        }

        public string GetFavoriteLabel(Office.IRibbonControl control, int index)
        {
            return ResolveFolderLabel(_addIn.SettingsService.Current.Favorites, index);
        }

        public void OnFavoriteSelected(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var folder = ResolveFolderInfo(_addIn.SettingsService.Current.Favorites, selectedIndex);
            if (folder != null)
            {
                _addIn.MoveSelectionToFolder(folder, false);
            }
        }

        private string ResolveFolderLabel(System.Collections.Generic.List<FolderIdentifier> list, int index)
        {
            var folder = ResolveFolderInfo(list, index);
            return folder?.DisplayText ?? "Unbekannter Ordner";
        }

        private FolderInfo ResolveFolderInfo(System.Collections.Generic.List<FolderIdentifier> list, int index)
        {
            if (index < 0 || index >= list.Count)
            {
                return null;
            }

            return _addIn.FolderService.GetFolderByIdentifier(list[index]);
        }
    }
}
