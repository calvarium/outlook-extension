using System;
using System.Collections.Generic;
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
        private System.Threading.Thread _cacheWarmupThread;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _loggingService = new LoggingService();
            _settingsService = new SettingsService(_loggingService);
            _folderService = new FolderService(Application, _settingsService, _loggingService);
            _searchService = new SearchService(_settingsService);
            _hotkeyService = new HotkeyService(Application, _settingsService, OpenQuickMoveDialog, _loggingService);

            StartCacheWarmupThread();

            Application.Explorers.NewExplorer += OnNewExplorer;
            _stores = Application.Session.Stores;
            _stores.StoreAdd += OnStoreChanged;
            _stores.BeforeStoreRemove += OnBeforeStoreRemove;

            _hotkeyService.RegisterShortcut();
            RegisterHotkeyForExplorer(Application.ActiveExplorer());
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
            DisposeCacheWarmupThread();
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

                dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual;
                var helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                helper.EnsureHandle();
                helper.Owner = ownerHandle;
                CenterDialogOnOwner(dialog, ownerHandle);
            }
            catch
            {
                dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
                // Ignore owner setup failures to avoid blocking the dialog.
            }
        }

        private IntPtr GetOutlookWindowHandle()
        {
            var foregroundHandle = GetForegroundWindow();
            if (IsOutlookWindow(foregroundHandle))
            {
                return foregroundHandle;
            }

            var processHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (processHandle != IntPtr.Zero)
            {
                return processHandle;
            }

            return foregroundHandle;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rect rect);

        private static bool IsOutlookWindow(IntPtr windowHandle)
        {
            if (windowHandle == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                GetWindowThreadProcessId(windowHandle, out var processId);
                if (processId == 0)
                {
                    return false;
                }

                var process = System.Diagnostics.Process.GetProcessById((int)processId);
                return string.Equals(process.ProcessName, "OUTLOOK", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static void CenterDialogOnOwner(System.Windows.Window dialog, IntPtr ownerHandle)
        {
            if (ownerHandle == IntPtr.Zero)
            {
                return;
            }

            if (!GetWindowRect(ownerHandle, out var ownerRect))
            {
                return;
            }

            var dialogWidth = dialog.Width;
            var dialogHeight = dialog.Height;
            if (dialogWidth <= 0 || dialogHeight <= 0)
            {
                return;
            }

            var ownerWidth = ownerRect.Right - ownerRect.Left;
            var ownerHeight = ownerRect.Bottom - ownerRect.Top;
            dialog.Left = ownerRect.Left + (ownerWidth - dialogWidth) / 2;
            dialog.Top = ownerRect.Top + (ownerHeight - dialogHeight) / 2;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

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
                    var itemsToMove = CollectMovableItems(selection);

                    movedCount = MoveItems(itemsToMove, folder);
                }
                else
                {
                    var inspector = Application.ActiveInspector();
                    var currentItem = inspector?.CurrentItem;
                    if (TryMoveItem(currentItem, folder))
                    {
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

        private List<object> CollectMovableItems(Outlook.Selection selection)
        {
            var itemsToMove = new List<object>();
            foreach (var selectionItem in selection)
            {
                if (TryAddMovableItem(selectionItem, itemsToMove))
                {
                    continue;
                }

                var conversationHeader = selectionItem as Outlook.ConversationHeader;
                if (conversationHeader != null)
                {
                    AddConversationItems(conversationHeader, itemsToMove);
                    Marshal.ReleaseComObject(conversationHeader);
                    continue;
                }

                if (Marshal.IsComObject(selectionItem))
                {
                    Marshal.ReleaseComObject(selectionItem);
                }
            }

            return itemsToMove;
        }

        private bool TryAddMovableItem(object item, List<object> itemsToMove)
        {
            var mail = item as Outlook.MailItem;
            var meeting = item as Outlook.MeetingItem;

            if (mail != null || meeting != null)
            {
                itemsToMove.Add(item);
                return true;
            }

            return false;
        }

        private void AddConversationItems(Outlook.ConversationHeader conversationHeader, List<object> itemsToMove)
        {
            Outlook.Conversation conversation = null;
            try
            {
                conversation = conversationHeader.GetConversation();
                if (conversation == null)
                {
                    return;
                }

                AddConversationItems(conversation, itemsToMove);
            }
            finally
            {
                if (conversation != null)
                {
                    Marshal.ReleaseComObject(conversation);
                }
            }
        }

        private void AddConversationItems(Outlook.Conversation conversation, List<object> itemsToMove)
        {
            Outlook.SimpleItems rootItems = null;
            try
            {
                rootItems = conversation.GetRootItems();
                if (rootItems == null)
                {
                    return;
                }

                AddConversationItems(conversation, rootItems, itemsToMove);
            }
            finally
            {
                if (rootItems != null)
                {
                    Marshal.ReleaseComObject(rootItems);
                }
            }
        }

        private void AddConversationItems(Outlook.Conversation conversation, Outlook.SimpleItems items, List<object> itemsToMove)
        {
            foreach (var conversationItem in items)
            {
                var added = TryAddMovableItem(conversationItem, itemsToMove);
                Outlook.SimpleItems children = null;
                try
                {
                    children = conversation.GetChildren(conversationItem);
                    if (children != null)
                    {
                        AddConversationItems(conversation, children, itemsToMove);
                    }
                }
                finally
                {
                    if (children != null)
                    {
                        Marshal.ReleaseComObject(children);
                    }
                }

                if (!added && Marshal.IsComObject(conversationItem))
                {
                    Marshal.ReleaseComObject(conversationItem);
                }
            }
        }

        private int MoveItems(List<object> itemsToMove, Outlook.MAPIFolder folder)
        {
            var movedCount = 0;
            foreach (var item in itemsToMove)
            {
                if (TryMoveItem(item, folder))
                {
                    movedCount++;
                }
            }

            return movedCount;
        }

        private bool TryMoveItem(object item, Outlook.MAPIFolder folder)
        {
            if (item == null)
            {
                return false;
            }

            try
            {
                if (item is Outlook.MailItem mailItem)
                {
                    mailItem.Move(folder);
                    return true;
                }

                if (item is Outlook.MeetingItem meetingItem)
                {
                    meetingItem.Move(folder);
                    return true;
                }
            }
            finally
            {
                if (Marshal.IsComObject(item))
                {
                    Marshal.ReleaseComObject(item);
                }
            }

            return false;
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
            RegisterHotkeyForExplorer(explorer);
        }

        private void RegisterHotkeyForExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return;
            }

            ((Outlook.ExplorerEvents_10_Event)explorer).Activate += OnExplorerActivate;
            TryRegisterHotkey(explorer);
        }

        private void OnExplorerActivate()
        {
            TryRegisterHotkey(Application.ActiveExplorer());
        }

        private void TryRegisterHotkey(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return;
            }

            _hotkeyService.RegisterShortcut();
            if (_hotkeyService.IsRegistered)
            {
                ((Outlook.ExplorerEvents_10_Event)explorer).Activate -= OnExplorerActivate;
            }
        }

        private void OnStoreChanged(Outlook.Store store)
        {
            _folderService.RefreshCache();
        }

        private void OnBeforeStoreRemove(Outlook.Store store, ref bool cancel)
        {
            _folderService.RefreshCache();
        }

        private void StartCacheWarmupThread()
        {
            if (_cacheWarmupThread != null || _folderService.WarmupStarted)
            {
                return;
            }

            _cacheWarmupThread = new System.Threading.Thread(() =>
            {
                Outlook.Application warmupApplication = null;
                try
                {
                    warmupApplication = new Outlook.Application();
                    _folderService.RefreshCache(warmupApplication);
                }
                catch (Exception ex)
                {
                    _loggingService.LogError("FolderCacheWarmup", ex);
                }
                finally
                {
                    if (warmupApplication != null)
                    {
                        Marshal.ReleaseComObject(warmupApplication);
                    }
                }
            })
            {
                IsBackground = true
            };
            _cacheWarmupThread.SetApartmentState(System.Threading.ApartmentState.STA);
            _cacheWarmupThread.Start();
        }

        private void DisposeCacheWarmupThread()
        {
            _cacheWarmupThread = null;
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
