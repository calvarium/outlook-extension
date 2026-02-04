using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace outlook_extension
{
    public partial class ThisAddIn
    {
        private FolderService _folderService;
        private SettingsService _settingsService;
        private SearchService _searchService;
        private HotkeyService _hotkeyService;
        private LoggingService _loggingService;
        private Outlook.Stores _stores;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _loggingService = new LoggingService();
            _settingsService = new SettingsService(_loggingService);
            _folderService = new FolderService(Application, _settingsService, _loggingService);
            _searchService = new SearchService(_settingsService);
            _hotkeyService = new HotkeyService(Application, _settingsService, OpenQuickMoveDialog, _loggingService);

            _folderService.InitializeCache();
            _hotkeyService.RegisterShortcut();

            Application.Explorers.NewExplorer += OnNewExplorer;
            _stores = Application.Session.Stores;
            _stores.StoreAdd += OnStoreChanged;
            _stores.BeforeStoreRemove += OnBeforeStoreRemove;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
            Application.Explorers.NewExplorer -= OnNewExplorer;
            if (_stores != null)
            {
                _stores.StoreAdd -= OnStoreChanged;
                _stores.BeforeStoreRemove -= OnBeforeStoreRemove;
                Marshal.ReleaseComObject(_stores);
                _stores = null;
            }

            _hotkeyService?.Dispose();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new QuickMoveRibbon(this);
        }

        public void OpenQuickMoveDialog()
        {
            try
            {
                var dialog = new QuickMoveWindow(_folderService, _searchService, this);
                SetWindowOwner(dialog);
                dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                _loggingService.LogError("QuickMoveDialog", ex);
                System.Windows.Forms.MessageBox.Show(
                    "Der Quick-Move-Dialog konnte nicht geöffnet werden.",
                    "Quick Move",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        internal FolderService FolderService => _folderService;

        internal SettingsService SettingsService => _settingsService;

        public void OpenSettingsDialog()
        {
            var dialog = new SettingsWindow(_folderService, _settingsService, _hotkeyService);
            SetWindowOwner(dialog);
            dialog.ShowDialog();
        }

        private void SetWindowOwner(System.Windows.Window dialog)
        {
            try
            {
                var ownerHandle = GetOutlookWindowHandle();
                if (ownerHandle == IntPtr.Zero)
                {
                    dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
                    return;
                }

                dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
                var helper = new System.Windows.Interop.WindowInteropHelper(dialog)
                {
                    Owner = ownerHandle
                };
            }
            catch
            {
                dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
                // Ignore owner setup failures to avoid blocking the dialog.
            }
        }

        private IntPtr GetOutlookWindowHandle()
        {
            try
            {
                var explorer = Application.ActiveExplorer();
                if (explorer != null)
                {
                    var explorerHandle = new IntPtr(explorer.HWND);
                    if (explorerHandle != IntPtr.Zero)
                    {
                        return explorerHandle;
                    }
                }
            }
            catch
            {
                // Ignore explorer handle errors.
            }

            try
            {
                var inspector = Application.ActiveInspector();
                if (inspector != null)
                {
                    var inspectorHandle = new IntPtr(inspector.HWND);
                    if (inspectorHandle != IntPtr.Zero)
                    {
                        return inspectorHandle;
                    }
                }
            }
            catch
            {
                // Ignore inspector handle errors.
            }

            var processHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (processHandle != IntPtr.Zero)
            {
                return processHandle;
            }

            return GetForegroundWindow();
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public bool MoveSelectionToFolder(FolderInfo targetFolder, bool keepDialogOpen)
        {
            if (targetFolder == null)
            {
                return false;
            }

            Outlook.MAPIFolder folder = null;
            try
            {
                folder = _folderService.ResolveFolder(targetFolder);
                if (folder == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Der Zielordner konnte nicht gefunden werden.",
                        "Quick Move",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                var movedCount = 0;
                var selection = Application.ActiveExplorer()?.Selection;
                if (selection != null && selection.Count > 0)
                {
                    foreach (object item in selection)
                    {
                        var mailItem = item as Outlook.MailItem;
                        if (mailItem == null)
                        {
                            continue;
                        }

                        mailItem.Move(folder);
                        Marshal.ReleaseComObject(mailItem);
                        movedCount++;
                    }
                }
                else
                {
                    var inspector = Application.ActiveInspector();
                    var mailItem = inspector?.CurrentItem as Outlook.MailItem;
                    if (mailItem != null)
                    {
                        mailItem.Move(folder);
                        Marshal.ReleaseComObject(mailItem);
                        movedCount = 1;
                    }
                }

                if (movedCount == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Keine verschiebbaren E-Mails gefunden.",
                        "Quick Move",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                _settingsService.AddRecent(targetFolder);
                _settingsService.Save();
                _searchService.NotifySettingsChanged();
                return true;
            }
            catch (Exception ex)
            {
                _loggingService.LogError("MoveSelectionToFolder", ex);
                System.Windows.Forms.MessageBox.Show(
                    "Beim Verschieben ist ein Fehler aufgetreten.",
                    "Quick Move",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
            }
        }

        public void UndoLastMove()
        {
            try
            {
                var explorer = Application.ActiveExplorer();
                if (explorer != null)
                {
                    explorer.CommandBars.ExecuteMso("Undo");
                    return;
                }

                var inspector = Application.ActiveInspector();
                if (inspector != null)
                {
                    inspector.CommandBars.ExecuteMso("Undo");
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("UndoLastMove", ex);
            }
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            _hotkeyService.RegisterShortcut();
        }

        private void OnStoreChanged(Outlook.Store store)
        {
            _folderService.RefreshCache();
        }

        private void OnBeforeStoreRemove(Outlook.Store store, ref bool cancel)
        {
            _folderService.RefreshCache();
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
