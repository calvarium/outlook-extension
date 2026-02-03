using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace outlook_extension
{
    public class FolderService
    {
        private readonly Outlook.Application _application;
        private readonly SettingsService _settingsService;
        private readonly LoggingService _loggingService;
        private readonly List<FolderInfo> _cache = new List<FolderInfo>();
        private readonly object _lock = new object();
        private bool _initialized;

        public FolderService(Outlook.Application application, SettingsService settingsService, LoggingService loggingService)
        {
            _application = application;
            _settingsService = settingsService;
            _loggingService = loggingService;
        }

        public IReadOnlyList<FolderInfo> GetCachedFolders()
        {
            lock (_lock)
            {
                return _cache.ToList();
            }
        }

        public void InitializeCache()
        {
            if (_initialized)
            {
                return;
            }

            RefreshCache();
            _initialized = true;
        }

        public void RefreshCache()
        {
            lock (_lock)
            {
                _cache.Clear();
                try
                {
                    var namespaceSession = _application.Session;
                    var stores = namespaceSession.Stores;
                    foreach (Outlook.Store store in stores)
                    {
                        if (!ShouldIncludeStore(store))
                        {
                            Marshal.ReleaseComObject(store);
                            continue;
                        }

                        var rootFolder = store.GetRootFolder();
                        try
                        {
                            TraverseFolder(rootFolder, store.DisplayName, new Stack<string>());
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(rootFolder);
                            Marshal.ReleaseComObject(store);
                        }
                    }

                    Marshal.ReleaseComObject(stores);
                }
                catch (Exception ex)
                {
                    _loggingService.LogError("FolderCache", ex);
                }
            }
        }

        public Outlook.MAPIFolder ResolveFolder(FolderInfo info)
        {
            if (info == null)
            {
                return null;
            }

            return _application.Session.GetFolderFromID(info.EntryId, info.StoreId);
        }

        public FolderInfo GetFolderByIdentifier(FolderIdentifier identifier)
        {
            if (identifier == null)
            {
                return null;
            }

            lock (_lock)
            {
                return _cache.FirstOrDefault(folder => folder.Identifier.Equals(identifier));
            }
        }

        private void TraverseFolder(Outlook.MAPIFolder folder, string mailboxName, Stack<string> path)
        {
            if (folder == null)
            {
                return;
            }

            path.Push(folder.Name);
            try
            {
                if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    var folderPath = string.Join(" > ", path.Reverse());
                    var info = new FolderInfo
                    {
                        EntryId = folder.EntryID,
                        StoreId = folder.StoreID,
                        DisplayName = folder.Name,
                        MailboxName = mailboxName,
                        FolderPath = folderPath,
                        FullPath = $"{mailboxName} > {folderPath}",
                        IsUnderInbox = folderPath.StartsWith("Posteingang", StringComparison.OrdinalIgnoreCase)
                    };
                    _cache.Add(info);
                }

                var folders = folder.Folders;
                foreach (Outlook.MAPIFolder child in folders)
                {
                    TraverseFolder(child, mailboxName, path);
                    Marshal.ReleaseComObject(child);
                }

                Marshal.ReleaseComObject(folders);
            }
            catch (Exception ex)
            {
                _loggingService.LogError("FolderTraverse", ex);
            }
            finally
            {
                path.Pop();
            }
        }

        private bool ShouldIncludeStore(Outlook.Store store)
        {
            if (store == null)
            {
                return false;
            }

            if (_settingsService.Current.IncludeArchives)
            {
                return true;
            }

            var displayName = store.DisplayName ?? string.Empty;
            var filePath = store.FilePath ?? string.Empty;
            return !(displayName.IndexOf("Archiv", StringComparison.OrdinalIgnoreCase) >= 0
                || displayName.IndexOf("Archive", StringComparison.OrdinalIgnoreCase) >= 0
                || filePath.IndexOf("archive", StringComparison.OrdinalIgnoreCase) >= 0
                || filePath.IndexOf("archiv", StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}
