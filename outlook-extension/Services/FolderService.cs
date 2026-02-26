using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Json;
using System.Threading;
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
        private bool _warmupStarted;
        private volatile bool _isRefreshing;
        private volatile bool _isVerifying;

        private CancellationTokenSource _refreshCts;

        private readonly string _cachePath;

        public event Action CacheUpdated;
        public event Action<bool> RefreshingChanged;
        public event Action<int,int> ProgressUpdated; // processed, total
        public event Action FullRefreshCompleted;

        private int _lastProgressProcessed;
        private int _lastProgressTotal;

        public bool IsRefreshing => _isRefreshing;
        public int LastProgressProcessed => _lastProgressProcessed;
        public int LastProgressTotal => _lastProgressTotal;

        public FolderService(Outlook.Application application, SettingsService settingsService, LoggingService loggingService)
        {
            _application = application;
            _settingsService = settingsService;
            _loggingService = loggingService;

            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "QuickMoveOutlook");
            Directory.CreateDirectory(folder);
            _cachePath = Path.Combine(folder, "folders.json");

            LoadCacheFromDisk();
        }

        private void LoadCacheFromDisk()
        {
            try
            {
                if (!File.Exists(_cachePath))
                {
                    return;
                }

                using (var stream = File.OpenRead(_cachePath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(List<FolderInfo>));
                    var list = (List<FolderInfo>)serializer.ReadObject(stream);
                    if (list != null)
                    {
                        lock (_lock)
                        {
                            _cache.Clear();
                            _cache.AddRange(list);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("LoadFolderCache", ex);
            }
        }

        private void SaveCacheToDisk(List<FolderInfo> list)
        {
            try
            {
                using (var stream = File.Create(_cachePath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(List<FolderInfo>));
                    serializer.WriteObject(stream, list);
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("SaveFolderCache", ex);
            }
        }

        public IReadOnlyList<FolderInfo> GetCachedFolders()
        {
            lock (_lock)
            {
                return _cache.ToList();
            }
        }

        public bool WarmupStarted => _warmupStarted;

        public void InitializeCache()
        {
            if (_initialized)
            {
                return;
            }

            RefreshCache();
            _initialized = true;
        }

        // Non-blocking public refresh: runs cache build on background STA thread
        public void RefreshCache()
        {
            // cancel previous
            try
            {
                _refreshCts?.Cancel();
            }
            catch { }

            _refreshCts = new CancellationTokenSource();
            var token = _refreshCts.Token;

            if (_isRefreshing)
            {
                // let cancellation take effect and continue
            }

            _isRefreshing = true;
            RefreshingChanged?.Invoke(true);

            var thread = new System.Threading.Thread(() =>
            {
                Outlook.Application app = null;
                var isFullRefresh = true;
                var refreshCompleted = false;
                try
                {
                    app = new Outlook.Application();

                    var namespaceSession = app.Session;
                    var stores = namespaceSession.Stores;

                    // collect stores into a list to allow ordering and safe release
                    var storeList = new List<Outlook.Store>();
                    foreach (Outlook.Store s in stores)
                    {
                        storeList.Add(s);
                    }

                    // prioritize default store first if available
                    Outlook.Store defaultStore = null;
                    try
                    {
                        defaultStore = app.Session.DefaultStore;
                    }
                    catch { }

                    var ordered = new List<Outlook.Store>();
                    if (defaultStore != null)
                    {
                        var match = storeList.FirstOrDefault(x => string.Equals(x.StoreID, defaultStore.StoreID, StringComparison.OrdinalIgnoreCase));
                        if (match != null)
                        {
                            ordered.Add(match);
                        }
                    }

                    // add remaining stores
                    ordered.AddRange(storeList.Where(s => ordered.All(o => o.StoreID != s.StoreID)));

                    // total count for progress
                    var total = ordered.Count;
                    var processed = 0;

                    var cumulative = new List<FolderInfo>();

                    for (int storeIndex = 0; storeIndex < ordered.Count; storeIndex++)
                    {
                        var store = ordered[storeIndex];
                        if (token.IsCancellationRequested) break;

                        try
                        {
                            var perStore = BuildCacheForStore(store, token, storeIndex);
                            if (perStore != null && perStore.Count > 0)
                            {
                                cumulative.AddRange(perStore);

                                lock (_lock)
                                {
                                    _cache.Clear();
                                    _cache.AddRange(cumulative);
                                }

                                // persist intermediate state and notify UI
                                SaveCacheToDisk(cumulative);
                                // update last progress for this store
                                _lastProgressProcessed = processed + 1; // processed stores count
                                _lastProgressTotal = total;
                                ProgressUpdated?.Invoke(_lastProgressProcessed, _lastProgressTotal);
                                CacheUpdated?.Invoke();
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            break;
                        }
                        catch (Exception ex)
                        {
                            _loggingService.LogError("FolderCacheStore", ex);
                        }
                        finally
                        {
                            try { if (store != null) Marshal.ReleaseComObject(store); } catch { }
                        }

                        processed++;
                    }

                    try { if (stores != null) Marshal.ReleaseComObject(stores); } catch { }

                    lock (_lock)
                    {
                        _initialized = true;
                    }

                    // final save (redundant if last store persisted)
                    SaveCacheToDisk(cumulative);
                    CacheUpdated?.Invoke();
                    refreshCompleted = !token.IsCancellationRequested && processed >= total;
                }
                catch (OperationCanceledException)
                {
                    // canceled - ignore
                }
                catch (Exception ex)
                {
                    _loggingService.LogError("FolderCache", ex);
                }
                finally
                {
                    try
                    {
                        if (app != null)
                        {
                            Marshal.ReleaseComObject(app);
                        }
                    }
                    catch { }

                    _isRefreshing = false;
                    // notify full refresh completed only if the refresh actually finished all stores
                    try
                    {
                        if (isFullRefresh && refreshCompleted)
                        {
                            FullRefreshCompleted?.Invoke();
                        }
                    }
                    catch { }
                    RefreshingChanged?.Invoke(false);
                }
            })
            {
                IsBackground = true
            };

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
        }

        // Warmup/synchronous refresh when caller provides an Outlook.Application already on STA thread
        public void RefreshCache(Outlook.Application application)
        {
            if (application == null)
            {
                return;
            }

            _warmupStarted = true;
            try
            {
                // similar incremental processing but synchronous
                var namespaceSession = application.Session;
                var stores = namespaceSession.Stores;

                var storeList = new List<Outlook.Store>();
                foreach (Outlook.Store s in stores)
                {
                    storeList.Add(s);
                }

                Outlook.Store defaultStore = null;
                try { defaultStore = application.Session.DefaultStore; } catch { }

                var ordered = new List<Outlook.Store>();
                if (defaultStore != null)
                {
                    var match = storeList.FirstOrDefault(x => string.Equals(x.StoreID, defaultStore.StoreID, StringComparison.OrdinalIgnoreCase));
                    if (match != null) ordered.Add(match);
                }
                ordered.AddRange(storeList.Where(s => ordered.All(o => o.StoreID != s.StoreID)));

                var cumulative = new List<FolderInfo>();
                foreach (var store in ordered)
                {
                    var perStore = BuildCacheForStore(store, CancellationToken.None, 0);
                    if (perStore != null && perStore.Count > 0)
                    {
                        cumulative.AddRange(perStore);
                    }

                    try { if (store != null) Marshal.ReleaseComObject(store); } catch { }
                }

                try { if (stores != null) Marshal.ReleaseComObject(stores); } catch { }

                lock (_lock)
                {
                    _cache.Clear();
                    _cache.AddRange(cumulative);
                    _initialized = true;
                }

                SaveCacheToDisk(cumulative);
                CacheUpdated?.Invoke();
            }
            catch (Exception ex)
            {
                _loggingService.LogError("FolderCache", ex);
            }
        }

        // Quick verification of existing cached entries on startup. Runs on background STA thread and updates cache with present folders.
        public void VerifyCacheOnStartup()
        {
            // cancel any running refresh and start a short verify
            try { _refreshCts?.Cancel(); } catch { }
            _refreshCts = new CancellationTokenSource();
            var token = _refreshCts.Token;

            if (_isRefreshing || _isVerifying)
            {
                // don't start verify if a full refresh or another verify is running
                return;
            }

            _isVerifying = true;

            var thread = new System.Threading.Thread(() =>
            {
                Outlook.Application app = null;
                try
                {
                    app = new Outlook.Application();
                    List<FolderInfo> currentCacheSnapshot;
                    lock (_lock)
                    {
                        currentCacheSnapshot = _cache.ToList();
                    }

                    var updated = new List<FolderInfo>();
                    var processed = 0;

                    foreach (var cached in currentCacheSnapshot)
                    {
                        if (token.IsCancellationRequested) break;

                        try
                        {
                            if (string.IsNullOrWhiteSpace(cached.EntryId) || string.IsNullOrWhiteSpace(cached.StoreId))
                            {
                                continue;
                            }

                            Outlook.MAPIFolder folder = null;
                            try
                            {
                                folder = app.Session.GetFolderFromID(cached.EntryId, cached.StoreId);
                            }
                            catch
                            {
                                folder = null;
                            }

                            if (folder == null)
                            {
                                // folder no longer exists
                                continue;
                            }

                            try
                            {
                                var mailboxName = folder.Store?.DisplayName ?? cached.MailboxName;
                                var pathParts = new Stack<string>();
                                try
                                {
                                    var current = folder as Outlook.MAPIFolder;
                                    while (current != null)
                                    {
                                        pathParts.Push(current.Name);
                                        var parent = current.Parent as Outlook.MAPIFolder;
                                        if (parent == null)
                                            break;
                                        current = parent;
                                    }
                                }
                                catch { }

                                var folderPath = pathParts.Count > 0 ? string.Join(" > ", pathParts) : cached.FolderPath;

                                // Ensure we don't duplicate the mailbox/display name when the root folder
                                // name already equals the mailbox display name (which can happen for some stores).
                                string fullPath;
                                if (string.IsNullOrEmpty(mailboxName))
                                {
                                    fullPath = folderPath;
                                }
                                else if (!string.IsNullOrEmpty(folderPath) && folderPath.StartsWith(mailboxName + " > ", StringComparison.OrdinalIgnoreCase))
                                {
                                    fullPath = folderPath;
                                }
                                else
                                {
                                    fullPath = string.IsNullOrEmpty(folderPath) ? mailboxName : $"{mailboxName} > {folderPath}";
                                }

                                var info = new FolderInfo
                                {
                                    EntryId = cached.EntryId,
                                    StoreId = cached.StoreId,
                                    DisplayName = folder.Name ?? cached.DisplayName,
                                    MailboxName = mailboxName,
                                    FolderPath = folderPath,
                                    FullPath = fullPath,
                                    IsUnderInbox = (folderPath ?? string.Empty).StartsWith("Posteingang", StringComparison.OrdinalIgnoreCase)
                                };

                                updated.Add(info);
                            }
                            finally
                            {
                                try { if (folder != null) Marshal.ReleaseComObject(folder); } catch { }
                            }
                        }
                        catch (Exception exInner)
                        {
                            _loggingService.LogError("VerifyCacheItem", exInner);
                        }

                        processed++;
                        _lastProgressProcessed = processed;
                        _lastProgressTotal = currentCacheSnapshot.Count;
                        ProgressUpdated?.Invoke(_lastProgressProcessed, _lastProgressTotal);
                    }

                    lock (_lock)
                    {
                        var changed = false;
                        if (updated.Count != _cache.Count)
                        {
                            changed = true;
                        }
                        else
                        {
                            for (int i = 0; i < updated.Count; i++)
                            {
                                if (!updated[i].Identifier.Equals(_cache[i].Identifier))
                                {
                                    changed = true;
                                    break;
                                }
                            }
                        }

                        if (changed)
                        {
                            _cache.Clear();
                            _cache.AddRange(updated);
                            SaveCacheToDisk(updated);
                        }
                    }

                    CacheUpdated?.Invoke();
                }
                catch (OperationCanceledException)
                {
                    // ignored
                }
                catch (Exception ex)
                {
                    _loggingService.LogError("VerifyCacheOnStartup", ex);
                }
                finally
                {
                    try { if (app != null) Marshal.ReleaseComObject(app); } catch { }
                    _isVerifying = false;
                    // Do not fire RefreshingChanged here; Verify is not a full refresh
                }
             })
             {
                 IsBackground = true
             };
 
             thread.SetApartmentState(System.Threading.ApartmentState.STA);
             thread.Start();
         }

        private List<FolderInfo> BuildCacheForStore(Outlook.Store store, CancellationToken token, int storeOrder)
        {
            var result = new List<FolderInfo>();

            if (store == null) return result;

            Outlook.MAPIFolder rootFolder = null;
            try
            {
                rootFolder = store.GetRootFolder();
                TraverseFolderForBuild(rootFolder, store.DisplayName, new Stack<string>(), result, token, storeOrder);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _loggingService.LogError("BuildCacheForStore", ex);
            }
            finally
            {
                try { if (rootFolder != null) Marshal.ReleaseComObject(rootFolder); } catch { }
            }

            return result;
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

        private void TraverseFolderForBuild(Outlook.MAPIFolder folder, string mailboxName, Stack<string> path, List<FolderInfo> target, CancellationToken token, int storeOrder)
        {
            if (folder == null || token.IsCancellationRequested)
            {
                return;
            }

            path.Push(folder.Name);
            try
            {
                // Determine whether to include this folder in the index.
                // Include when the folder itself is a mail folder or when it has child folders (so parent folders like "Posteingang" are included).
                Outlook.Folders folders = null;
                bool includeThis = false;
                try
                {
                    includeThis = folder.DefaultItemType == Outlook.OlItemType.olMailItem;
                    // retrieve child folders once
                    try { folders = folder.Folders; } catch { folders = null; }
                    if (!includeThis && folders != null)
                    {
                        try { includeThis = folders.Count > 0; } catch { includeThis = includeThis || false; }
                    }
                }
                catch { }

                if (includeThis)
                {
                    var folderPath = string.Join(" > ", path.Reverse());
                    var info = new FolderInfo
                    {
                        EntryId = folder.EntryID,
                        StoreId = folder.StoreID,
                        DisplayName = folder.Name,
                        MailboxName = mailboxName,
                        FolderPath = folderPath,
                        // Avoid duplicating mailbox name if the folderPath already begins with it
                        FullPath = (!string.IsNullOrEmpty(mailboxName) && !string.IsNullOrEmpty(folderPath) && folderPath.StartsWith(mailboxName + " > ", StringComparison.OrdinalIgnoreCase))
                                    ? folderPath
                                    : (string.IsNullOrEmpty(mailboxName) ? folderPath : $"{mailboxName} > {folderPath}"),
                        IsUnderInbox = folderPath.StartsWith("Posteingang", StringComparison.OrdinalIgnoreCase),
                        StoreOrder = storeOrder
                    };
                    target.Add(info);
                }

                // iterate children (folders may be null if access failed)
                if (folders == null)
                {
                    try { folders = folder.Folders; } catch { folders = null; }
                }

                if (folders != null)
                {
                    foreach (Outlook.MAPIFolder child in folders)
                    {
                        if (token.IsCancellationRequested) break;

                        try
                        {
                            TraverseFolderForBuild(child, mailboxName, path, target, token, storeOrder);
                        }
                        finally
                        {
                            if (child != null) Marshal.ReleaseComObject(child);
                        }
                    }

                    try { if (folders != null) Marshal.ReleaseComObject(folders); } catch { }
                }
                else
                {
                    // nothing to release
                }
             }
             catch (OperationCanceledException)
             {
                 throw;
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

        public void CancelRefresh()
        {
            try
            {
                _refreshCts?.Cancel();
            }
            catch
            {
                // ignore
            }
        }
    }
}
