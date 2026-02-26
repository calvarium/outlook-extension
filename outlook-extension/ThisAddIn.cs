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
        private Outlook.Stores _stores; // kept for compatibility but not assigned at startup
        private System.Windows.Forms.Timer _startupTimer;
        private System.Threading.Thread _cacheWarmupThread;
        private System.Threading.Timer _periodicRefreshTimer;

        // Track last move operation to allow programmatic Undo
        private List<MoveEntry> _lastMoveEntries = null;
        private readonly object _moveLock = new object();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _loggingService = new LoggingService();
            _settingsService = new SettingsService(_loggingService);
            _folderService = new FolderService(Application, _settingsService, _loggingService);
            _searchService = new SearchService(_settingsService);
            _hotkeyService = new HotkeyService(Application, _settingsService, OpenQuickMoveDialog, _loggingService);

            StartPostStartupTimer();

            Application.Explorers.NewExplorer += OnNewExplorer;

            // Do NOT enumerate Application.Session.Stores here — it's expensive and blocks the UI when many stores exist.
            // Instead use background refresh and a periodic timer.
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
                try
                {
                    Marshal.ReleaseComObject(_stores);
                }
                catch
                {
                    // ignore
                }
                _stores = null;
            }

            _hotkeyService?.Dispose();
            DisposePostStartupTimer();
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

                // start recording this move operation so it can be undone
                lock (_moveLock)
                {
                    _lastMoveEntries = new List<MoveEntry>();
                }

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
                    // clear recorded move if nothing moved
                    lock (_moveLock)
                    {
                        _lastMoveEntries = null;
                    }
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
            var uniqueEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var selectionItem in selection)
            {
                var currentCount = itemsToMove.Count;
                var conversationHeader = selectionItem as Outlook.ConversationHeader;
                if (conversationHeader != null)
                {
                    AddConversationItems(conversationHeader, itemsToMove, uniqueEntryIds);
                    Marshal.ReleaseComObject(conversationHeader);
                    continue;
                }

                var mailItem = selectionItem as Outlook.MailItem;
                if (mailItem != null)
                {
                    AddConversationItemsFromItem(mailItem, itemsToMove, uniqueEntryIds);
                    if (currentCount != itemsToMove.Count)
                        continue;
                }

                if (TryAddMovableItem(selectionItem, itemsToMove, uniqueEntryIds))
                {
                    continue;
                }

                if (Marshal.IsComObject(selectionItem))
                {
                    Marshal.ReleaseComObject(selectionItem);
                }
            }

            return itemsToMove;
        }

        private bool TryAddMovableItem(object item, List<object> itemsToMove, HashSet<string> uniqueEntryIds)
        {
            var mail = item as Outlook.MailItem;
            var meeting = item as Outlook.MeetingItem;

            if (mail != null || meeting != null)
            {
                var entryId = GetEntryId(item);
                if (!string.IsNullOrEmpty(entryId) && !uniqueEntryIds.Add(entryId))
                {
                    return true;
                }

                itemsToMove.Add(item);
                return true;
            }

            return false;
        }

        private void AddConversationItems(Outlook.ConversationHeader conversationHeader, List<object> itemsToMove, HashSet<string> uniqueEntryIds)
        {
            Outlook.Conversation conversation = null;
            Outlook.SimpleItems headerItems = null;
            try
            {
                headerItems = conversationHeader.GetItems();
                if (headerItems != null && headerItems.Count > 0)
                {
                    AddConversationItems(headerItems, itemsToMove, uniqueEntryIds);
                    return;
                }

                conversation = conversationHeader.GetConversation();
                if (conversation == null)
                {
                    return;
                }

                AddConversationItems(conversation, itemsToMove, uniqueEntryIds);
            }
            finally
            {
                if (headerItems != null)
                {
                    Marshal.ReleaseComObject(headerItems);
                }

                if (conversation != null)
                {
                    Marshal.ReleaseComObject(conversation);
                }
            }
        }

        private void AddConversationItemsFromItem(Outlook.MailItem mailItem, List<object> itemsToMove, HashSet<string> uniqueEntryIds)
        {
            Outlook.Conversation conversation = null;
            try
            {
                conversation = mailItem.GetConversation();
                if (conversation == null)
                {
                    return;
                }

                AddConversationItems(conversation, itemsToMove, uniqueEntryIds);
            }
            finally
            {
                if (conversation != null)
                {
                    Marshal.ReleaseComObject(conversation);
                }
            }
        }

        private void AddConversationItems(Outlook.Conversation conversation, List<object> itemsToMove, HashSet<string> uniqueEntryIds)
        {
            Outlook.SimpleItems rootItems = null;
            try
            {
                rootItems = conversation.GetRootItems();
                if (rootItems == null)
                {
                    return;
                }

                AddConversationItems(conversation, rootItems, itemsToMove, uniqueEntryIds);
            }
            finally
            {
                if (rootItems != null)
                {
                    Marshal.ReleaseComObject(rootItems);
                }
            }
        }

        private void AddConversationItems(Outlook.SimpleItems items, List<object> itemsToMove, HashSet<string> uniqueEntryIds)
        {
            if (items == null)
            {
                return;
            }

            foreach (var conversationItem in items)
            {
                if (TryAddMovableItem(conversationItem, itemsToMove, uniqueEntryIds))
                {
                    continue;
                }

                if (Marshal.IsComObject(conversationItem))
                {
                    Marshal.ReleaseComObject(conversationItem);
                }
            }
        }

        private void AddConversationItems(Outlook.Conversation conversation, Outlook.SimpleItems items, List<object> itemsToMove, HashSet<string> uniqueEntryIds)
        {
            if (items == null)
            {
                return;
            }

            foreach (var conversationItem in items)
            {
                var added = TryAddMovableItem(conversationItem, itemsToMove, uniqueEntryIds);
                Outlook.SimpleItems children = null;
                try
                {
                    children = conversation.GetChildren(conversationItem);
                    if (children != null)
                    {
                        AddConversationItems(conversation, children, itemsToMove, uniqueEntryIds);
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

        private string GetEntryId(object item)
        {
            if (item is Outlook.MailItem mailItem)
            {
                return mailItem.EntryID;
            }

            if (item is Outlook.MeetingItem meetingItem)
            {
                return meetingItem.EntryID;
            }

            return null;
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
                    // capture source folder ids
                    string oldFolderEntryId = null;
                    string oldStoreId = null;
                    Outlook.MAPIFolder oldFolder = null;
                    try
                    {
                        oldFolder = mailItem.Parent as Outlook.MAPIFolder;
                        if (oldFolder != null)
                        {
                            try { oldFolderEntryId = oldFolder.EntryID; } catch { }
                            try { oldStoreId = oldFolder.Store?.StoreID; } catch { }
                        }
                    }
                    catch { }
                    finally
                    {
                        try { if (oldFolder != null) Marshal.ReleaseComObject(oldFolder); } catch { }
                    }

                    // capture identifying metadata before move
                    string subject = null;
                    string convId = null;
                    DateTime? received = null;
                    string sender = null;
                    string internetMessageId = null;
                    try
                    {
                        try { subject = mailItem.Subject; } catch { }
                        try { convId = mailItem.ConversationID; } catch { }
                        try { received = mailItem.ReceivedTime; } catch { }
                        try { sender = mailItem.SenderEmailAddress; } catch { }
                        try
                        {
                            // read InternetMessageId via PropertyAccessor (PIA may not expose it)
                            if (Marshal.IsComObject(mailItem))
                            {
                                try
                                {
                                    var pa = mailItem.PropertyAccessor;
                                    try
                                    {
                                        internetMessageId = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") as string;
                                    }
                                    finally
                                    {
                                        try { if (pa != null) Marshal.ReleaseComObject(pa); } catch { }
                                    }
                                }
                                catch { internetMessageId = null; }
                            }
                        }
                        catch { }
                    }
                    catch { }

                    Outlook.MailItem movedItem = null;
                    try
                    {
                        movedItem = mailItem.Move(folder) as Outlook.MailItem;
                        if (movedItem != null)
                        {
                            // try to ensure moved item is persisted and nudged into indexes/search
                            try
                            {
                                movedItem.Save();
                                TryNudgeSearchIndex(movedItem);
                            }
                            catch (Exception exSave)
                            {
                                try { _loggingService.LogError("MoveSaveOrNudge", exSave); } catch { }
                            }
                            string newEntryId = null;
                            string newStoreId = null;
                            string newFolderEntryId = null;
                            Outlook.MAPIFolder newParent = null;
                            try
                            {
                                try { newEntryId = movedItem.EntryID; } catch { }
                                newParent = movedItem.Parent as Outlook.MAPIFolder;
                                if (newParent != null)
                                {
                                    try { newFolderEntryId = newParent.EntryID; } catch { }
                                    try { newStoreId = newParent.Store?.StoreID; } catch { }
                                }
                            }
                            catch { }
                            finally
                            {
                                try { if (newParent != null) Marshal.ReleaseComObject(newParent); } catch { }
                            }

                            // record move for potential undo
                            lock (_moveLock)
                            {
                                if (_lastMoveEntries != null)
                                {
                                    var entry = new MoveEntry
                                    {
                                        OldFolderEntryId = oldFolderEntryId,
                                        OldStoreId = oldStoreId,
                                        NewEntryId = newEntryId,
                                        NewStoreId = newStoreId,
                                        NewFolderEntryId = newFolderEntryId,
                                        Subject = subject,
                                        ConversationId = convId,
                                        ReceivedTime = received,
                                        SenderEmail = sender,
                                        InternetMessageId = internetMessageId
                                    };
                                    _lastMoveEntries.Add(entry);
                                    try { _loggingService.LogInfo($"Recorded move entry: oldFolder={entry.OldFolderEntryId}, oldStore={entry.OldStoreId}, newEntry={entry.NewEntryId}, newStore={entry.NewStoreId}, subject={entry.Subject}"); } catch { }
                                }
                            }

                            return true;
                        }
                    }
                    finally
                    {
                        try { if (movedItem != null) Marshal.ReleaseComObject(movedItem); } catch { }
                    }
                }

                if (item is Outlook.MeetingItem meetingItem)
                {
                    // capture source folder ids
                    string oldFolderEntryId = null;
                    string oldStoreId = null;
                    Outlook.MAPIFolder oldFolder = null;
                    try
                    {
                        oldFolder = meetingItem.Parent as Outlook.MAPIFolder;
                        if (oldFolder != null)
                        {
                            try { oldFolderEntryId = oldFolder.EntryID; } catch { }
                            try { oldStoreId = oldFolder.Store?.StoreID; } catch { }
                        }
                    }
                    catch { }
                    finally
                    {
                        try { if (oldFolder != null) Marshal.ReleaseComObject(oldFolder); } catch { }
                    }

                    // capture identifying metadata before move
                    string subject = null;
                    DateTime? received = null;
                    try
                    {
                        try { subject = meetingItem.Subject; } catch { }
                        try { received = meetingItem.CreationTime; } catch { }
                    }
                    catch { }

                    Outlook.MeetingItem movedItem = null;
                    try
                    {
                        movedItem = meetingItem.Move(folder) as Outlook.MeetingItem;
                        if (movedItem != null)
                        {
                            try
                            {
                                movedItem.Save();
                                TryNudgeSearchIndex(movedItem);
                            }
                            catch (Exception exSave)
                            {
                                try { _loggingService.LogError("MoveSaveOrNudgeMeeting", exSave); } catch { }
                            }
                            string newEntryId = null;
                            string newStoreId = null;
                            string newFolderEntryId = null;
                            Outlook.MAPIFolder newParent = null;
                            try
                            {
                                try { newEntryId = movedItem.EntryID; } catch { }
                                newParent = movedItem.Parent as Outlook.MAPIFolder;
                                if (newParent != null)
                                {
                                    try { newFolderEntryId = newParent.EntryID; } catch { }
                                    try { newStoreId = newParent.Store?.StoreID; } catch { }
                                }
                            }
                            catch { }
                            finally
                            {
                                try { if (newParent != null) Marshal.ReleaseComObject(newParent); } catch { }
                            }

                            // record move for potential undo
                            lock (_moveLock)
                            {
                                if (_lastMoveEntries != null)
                                {
                                    var entry = new MoveEntry
                                    {
                                        OldFolderEntryId = oldFolderEntryId,
                                        OldStoreId = oldStoreId,
                                        NewEntryId = newEntryId,
                                        NewStoreId = newStoreId,
                                        NewFolderEntryId = newFolderEntryId,
                                        Subject = subject,
                                        ReceivedTime = received
                                    };
                                    _lastMoveEntries.Add(entry);
                                    try { _loggingService.LogInfo($"Recorded move entry (meeting): oldFolder={entry.OldFolderEntryId}, newEntry={entry.NewEntryId}, subject={entry.Subject}"); } catch { }
                                }
                            }

                            return true;
                        }
                    }
                    finally
                    {
                        try { if (movedItem != null) Marshal.ReleaseComObject(movedItem); } catch { }
                    }
                }
            }
            finally
            {
                if (Marshal.IsComObject(item))
                {
                    try { Marshal.ReleaseComObject(item); } catch { }
                }
            }

            return false;
        }

        public void UndoLastMove()
        {
            try
            {
                // Try programmatic undo first using recorded move entries
                List<MoveEntry> entries = null;
                lock (_moveLock)
                {
                    if (_lastMoveEntries != null && _lastMoveEntries.Count > 0)
                    {
                        entries = new List<MoveEntry>(_lastMoveEntries);
                        _lastMoveEntries = null; // clear optimistic
                    }
                }

                if (entries != null && entries.Count > 0)
                {
                    var session = Application.Session;
                    var successCount = 0;

                    // Undo in reverse order to reduce potential conflicts
                    for (int i = entries.Count - 1; i >= 0; i--)
                    {
                        var e = entries[i];
                        try
                        {
                            if (string.IsNullOrEmpty(e.NewEntryId))
                            {
                                _loggingService.LogInfo($"Undo: skipping entry with empty NewEntryId (oldFolder={e.OldFolderEntryId})");
                                continue;
                            }

                            object movedObj = null;
                            try
                            {
                                movedObj = session.GetItemFromID(e.NewEntryId, e.NewStoreId);
                            }
                            catch (Exception exGet)
                            {
                                _loggingService.LogError("UndoGetItemFromID", exGet);
                                movedObj = null;
                            }

                            // Fallback: try without store id
                            if (movedObj == null)
                            {
                                try
                                {
                                    movedObj = session.GetItemFromID(e.NewEntryId);
                                    if (movedObj != null)
                                    {
                                        _loggingService.LogInfo($"Undo: located item by EntryID without store param: {e.NewEntryId}");
                                    }
                                }
                                catch (Exception exGet2)
                                {
                                    _loggingService.LogError("UndoGetItemFromIDFallback", exGet2);
                                }
                            }

                            if (movedObj == null)
                            {
                                _loggingService.LogInfo($"Undo: could not locate moved item by id {e.NewEntryId}");
                                // Try to find the item by metadata in the target folder
                                try
                                {
                                    Outlook.MAPIFolder targetFolder = null;
                                    try
                                    {
                                        if (!string.IsNullOrEmpty(e.NewFolderEntryId))
                                        {
                                            try { targetFolder = session.GetFolderFromID(e.NewFolderEntryId, e.NewStoreId); } catch { targetFolder = null; }
                                            if (targetFolder == null)
                                            {
                                                try { targetFolder = session.GetFolderFromID(e.NewFolderEntryId); } catch { targetFolder = null; }
                                            }
                                        }
                                    }
                                    catch { targetFolder = null; }

                                    if (targetFolder != null)
                                    {
                                        try
                                        {
                                            var items = targetFolder.Items;
                                            try
                                            {
                                                foreach (var it in items)
                                                {
                                                    try
                                                    {
                                                        if (it is Outlook.MailItem mailItem)
                                                        {
                                                            if (!string.IsNullOrEmpty(e.Subject) &&
                                                                string.Equals(mailItem.Subject, e.Subject, StringComparison.OrdinalIgnoreCase) &&
                                                                (!e.ReceivedTime.HasValue || mailItem.ReceivedTime == e.ReceivedTime))
                                                            {
                                                                movedObj = it;
                                                                _loggingService.LogInfo($"Undo: located moved mail by metadata in target folder: {e.Subject}");
                                                                break;
                                                            }
                                                        }
                                                        else if (it is Outlook.MeetingItem meetingItem)
                                                        {
                                                            if (!string.IsNullOrEmpty(e.Subject) &&
                                                                string.Equals(meetingItem.Subject, e.Subject, StringComparison.OrdinalIgnoreCase) &&
                                                                (!e.ReceivedTime.HasValue || meetingItem.ReceivedTime == e.ReceivedTime))
                                                            {
                                                                movedObj = it;
                                                                _loggingService.LogInfo($"Undo: located moved meeting by metadata in target folder: {e.Subject}");
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    catch { }

                                                    // release non-matching item
                                                    try { if (it != null && Marshal.IsComObject(it)) Marshal.ReleaseComObject(it); } catch { }
                                                }
                                            }
                                            finally
                                            {
                                                try { if (items != null) Marshal.ReleaseComObject(items); } catch { }
                                            }
                                        }
                                        catch (Exception exSearch)
                                        {
                                            _loggingService.LogError("UndoSearchTargetFolder", exSearch);
                                        }
                                    }

                                    // If not found in target folder, try searching all stores recursively
                                    if (movedObj == null)
                                    {
                                        try
                                        {
                                            var stores = session.Stores;
                                            try
                                            {
                                                foreach (Outlook.Store s in stores)
                                                {
                                                    Outlook.MAPIFolder root = null;
                                                    try
                                                    {
                                                        root = s.GetRootFolder();
                                                        if (root != null)
                                                        {
                                                            // recursive search
                                                            var found = FindItemInFolderRecursive(root, e);
                                                            if (found != null)
                                                            {
                                                                movedObj = found;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    catch { }
                                                    finally { try { if (root != null) Marshal.ReleaseComObject(root); } catch { } }
                                                }
                                            }
                                            finally { try { if (stores != null) Marshal.ReleaseComObject(stores); } catch { } }
                                        }
                                        catch (Exception exAll)
                                        {
                                            _loggingService.LogError("UndoSearchAllStores", exAll);
                                        }
                                    }

                                    try { if (targetFolder != null) Marshal.ReleaseComObject(targetFolder); } catch { }
                                }
                                catch (Exception ex)
                                {
                                    _loggingService.LogError("UndoFindByMetadata", ex);
                                }

                                if (movedObj == null)
                                {
                                    _loggingService.LogInfo($"Undo: relocated item not found for {e.NewEntryId}, skipping");
                                    continue;
                                }
                            }

                            Outlook.MAPIFolder originalFolder = null;
                            try
                            {
                                if (!string.IsNullOrEmpty(e.OldFolderEntryId))
                                {
                                    try { originalFolder = session.GetFolderFromID(e.OldFolderEntryId, e.OldStoreId); } catch (Exception exGetFolder) { _loggingService.LogError("UndoGetFolderFromID", exGetFolder); originalFolder = null; }
                                    if (originalFolder == null)
                                    {
                                        try { originalFolder = session.GetFolderFromID(e.OldFolderEntryId); } catch (Exception exGetFolder2) { _loggingService.LogError("UndoGetFolderFromIDFallback", exGetFolder2); originalFolder = null; }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                _loggingService.LogError("UndoGetFolderWrap", ex);
                            }

                            try
                            {
                                if (originalFolder == null)
                                {
                                    _loggingService.LogInfo($"Undo: original folder not found for item {e.NewEntryId}, skipping");
                                }
                                else
                                {
                                    // Try typed move first
                                    if (movedObj is Outlook.MailItem m)
                                    {
                                        try { m.Move(originalFolder); successCount++; }
                                        catch (Exception exMove) { _loggingService.LogError("UndoMoveMailItem", exMove); }
                                    }
                                    else if (movedObj is Outlook.MeetingItem mt)
                                    {
                                        try { mt.Move(originalFolder); successCount++; }
                                        catch (Exception exMove) { _loggingService.LogError("UndoMoveMeetingItem", exMove); }
                                    }
                                    else
                                    {
                                        // Try dynamic invocation for other COM item types that expose Move
                                        try
                                        {
                                            var ti = movedObj.GetType();
                                            var moveMethod = ti.GetMethod("Move");
                                            if (moveMethod != null)
                                            {
                                                moveMethod.Invoke(movedObj, new object[] { originalFolder });
                                                successCount++;
                                            }
                                            else
                                            {
                                                _loggingService.LogInfo($"Undo: moved object type {ti.Name} has no Move method");
                                            }
                                        }
                                        catch (Exception exDyn)
                                        {
                                            _loggingService.LogError("UndoDynamicMove", exDyn);
                                        }
                                    }
                                }
                            }
                            catch (Exception exInner)
                            {
                                _loggingService.LogError("UndoMoveItem", exInner);
                            }
                            finally
                            {
                                try { if (movedObj != null && Marshal.IsComObject(movedObj)) Marshal.ReleaseComObject(movedObj); } catch { }
                                try { if (originalFolder != null) Marshal.ReleaseComObject(originalFolder); } catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            _loggingService.LogError("UndoMoveLoop", ex);
                        }
                    }

                    // If programmatic undo didn't find or move any items, try built-in Undo as fallback
                    if (successCount == 0)
                    {
                        _loggingService.LogInfo("Undo: programmatic undo made no changes, falling back to ExecuteMso Undo");
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

                    return;
                }

                // Fallback to the built-in Undo command if no programmatic undo available
                var explorer2 = Application.ActiveExplorer();
                if (explorer2 != null)
                {
                    explorer2.CommandBars.ExecuteMso("Undo");
                    return;
                }

                var inspector2 = Application.ActiveInspector();
                if (inspector2 != null)
                {
                    inspector2.CommandBars.ExecuteMso("Undo");
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

        private void StartPostStartupTimer()
        {
            if (_startupTimer != null || _folderService.WarmupStarted || _cacheWarmupThread != null)
            {
                return;
            }

            _startupTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            _startupTimer.Tick += (sender, args) =>
            {
                try
                {
                    _startupTimer.Stop();

                    // Do not access Application.Session.Stores here to avoid blocking the UI.

                    try
                    {
                        _hotkeyService.RegisterShortcut();
                        RegisterHotkeyForExplorer(Application.ActiveExplorer());
                    }
                    catch (Exception ex)
                    {
                        _loggingService.LogError("HotkeyPostStartup", ex);
                    }

                    // Trigger quick verification of persisted cache (non-blocking)
                    try
                    {
                        _folderService.VerifyCacheOnStartup();
                    }
                    catch (Exception ex)
                    {
                        _loggingService.LogError("VerifyCachePostStartup", ex);
                    }

                    // Trigger initial cache refresh in background (FolderService.RefreshCache is non-blocking)
                    try
                    {
                        _folderService.RefreshCache();
                    }
                    catch (Exception ex)
                    {
                        _loggingService.LogError("FolderCachePostStartup", ex);
                    }

                    // Start periodic background refresh every 5 minutes
                    try
                    {
                        _periodicRefreshTimer = new System.Threading.Timer(_ =>
                        {
                            try
                            {
                                _folderService.RefreshCache();
                            }
                            catch (Exception ex)
                            {
                                _loggingService.LogError("FolderCachePeriodic", ex);
                            }
                        }, null, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
                    }
                    catch (Exception ex)
                    {
                        _loggingService.LogError("PeriodicRefreshTimer", ex);
                    }
                }
                finally
                {
                    // no-op
                }
            };

            _startupTimer.Start();
        }

        private void DisposePostStartupTimer()
        {
            if (_startupTimer != null)
            {
                try
                {
                    _startupTimer.Stop();
                    _startupTimer.Dispose();
                }
                catch { }
                finally { _startupTimer = null; }
            }

            if (_periodicRefreshTimer != null)
            {
                try
                {
                    _periodicRefreshTimer.Dispose();
                }
                catch { }
                finally { _periodicRefreshTimer = null; }
            }
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

        // Helper types for undo
        private class MoveEntry
        {
            public string OldFolderEntryId { get; set; }
            public string OldStoreId { get; set; }
            public string NewEntryId { get; set; }
            public string NewStoreId { get; set; }
            // additional metadata to find items if EntryID lookup fails
            public string NewFolderEntryId { get; set; }
            public string Subject { get; set; }
            public string ConversationId { get; set; }
            public DateTime? ReceivedTime { get; set; }
            public string SenderEmail { get; set; }
            public string InternetMessageId { get; set; }
        }

        // helper: recursively search folder for item matching MoveEntry metadata
        private object FindItemInFolderRecursive(Outlook.MAPIFolder folder, MoveEntry criteria)
        {
            if (folder == null) return null;

            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                if (items != null)
                {
                    foreach (var it in items)
                    {
                        try
                        {
                            if (it is Outlook.MailItem mailItem)
                            {
                                if (!string.IsNullOrEmpty(criteria.Subject) && string.Equals(mailItem.Subject, criteria.Subject, StringComparison.OrdinalIgnoreCase) &&
                                    (!criteria.ReceivedTime.HasValue || mailItem.ReceivedTime == criteria.ReceivedTime))
                                {
                                    return it; // DO NOT release, caller will release
                                }
                            }
                            else if (it is Outlook.MeetingItem meetingItem)
                            {
                                if (!string.IsNullOrEmpty(criteria.Subject) && string.Equals(meetingItem.Subject, criteria.Subject, StringComparison.OrdinalIgnoreCase) &&
                                    (!criteria.ReceivedTime.HasValue || meetingItem.ReceivedTime == criteria.ReceivedTime))
                                {
                                    return it;
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            try { if (it != null && Marshal.IsComObject(it)) Marshal.ReleaseComObject(it); } catch { }
                        }
                    }
                }
            }
            catch { }
            finally
            {
                try { if (items != null) Marshal.ReleaseComObject(items); } catch { }
            }

            // search subfolders
            Outlook.Folders subs = null;
            try
            {
                subs = folder.Folders;
                if (subs != null)
                {
                    foreach (Outlook.MAPIFolder sub in subs)
                    {
                        object found = null;
                        try
                        {
                            found = FindItemInFolderRecursive(sub, criteria);
                            if (found != null)
                            {
                                return found; // leave found object for caller
                            }
                        }
                        finally
                        {
                            try { if (found == null && sub != null) Marshal.ReleaseComObject(sub); } catch { }
                        }
                    }
                }
            }
            catch { }
            finally
            {
                try { if (subs != null) Marshal.ReleaseComObject(subs); } catch { }
            }

            return null;
        }

        // Best-effort nudge to make moved item visible to Outlook search/indexing:
        // - ensure Save() is called
        // - try to refresh UI explorers/inspectors which may trigger indexing/on-demand refresh
        private void TryNudgeSearchIndex(object movedObj)
        {
            if (movedObj == null)
                return;

            try
            {
                try
                {
                    if (movedObj is Outlook.MailItem m)
                    {
                        m.Save();
                    }
                    else if (movedObj is Outlook.MeetingItem mt)
                    {
                        mt.Save();
                    }
                    else
                    {
                        // try dynamic Save
                        var mi = movedObj.GetType().GetMethod("Save");
                        mi?.Invoke(movedObj, null);
                    }
                }
                catch (Exception ex)
                {
                    try { _loggingService.LogError("TryNudge_Save", ex); } catch { }
                }

                try
                {
                    var explorer = Application.ActiveExplorer();
                    if (explorer != null)
                    {
                        try { explorer.CommandBars.ExecuteMso("Refresh"); } catch (Exception ex) { try { _loggingService.LogError("TryNudge_ExplorerRefresh", ex); } catch { } }
                    }

                    var inspector = Application.ActiveInspector();
                    if (inspector != null)
                    {
                        try { inspector.CommandBars.ExecuteMso("Refresh"); } catch (Exception ex) { try { _loggingService.LogError("TryNudge_InspectorRefresh", ex); } catch { } }
                    }
                }
                catch (Exception ex)
                {
                    try { _loggingService.LogError("TryNudge_RefreshWrap", ex); } catch { }
                }
            }
            catch { }
        }

    }
}
