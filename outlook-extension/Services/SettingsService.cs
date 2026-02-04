using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;

namespace outlook_extension
{
    public class SettingsService
    {
        private readonly string _settingsPath;
        private readonly LoggingService _loggingService;

        public SettingsModel Current { get; private set; }

        public SettingsService(LoggingService loggingService)
        {
            _loggingService = loggingService;
            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "QuickMoveOutlook");
            Directory.CreateDirectory(folder);
            _settingsPath = Path.Combine(folder, "settings.json");
            Load();
        }

        public void Load()
        {
            try
            {
                if (!File.Exists(_settingsPath))
                {
                    Current = new SettingsModel();
                    Save();
                    return;
                }

                using (var stream = File.OpenRead(_settingsPath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(SettingsModel));
                    Current = (SettingsModel)serializer.ReadObject(stream);
                }

                if (string.IsNullOrWhiteSpace(Current.Shortcut))
                {
                    Current.Shortcut = "Alt+Shift+M";
                    Save();
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("SettingsLoad", ex);
                Current = new SettingsModel();
            }
        }

        public void Save()
        {
            try
            {
                using (var stream = File.Create(_settingsPath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(SettingsModel));
                    serializer.WriteObject(stream, Current);
                }
            }
            catch (Exception ex)
            {
                _loggingService.LogError("SettingsSave", ex);
            }
        }

        public void AddRecent(FolderInfo folder)
        {
            if (folder == null)
            {
                return;
            }

            var identifier = folder.Identifier;
            Current.Recents.RemoveAll(item => item.Equals(identifier));
            Current.Recents.Insert(0, identifier);
            Current.Recents = Current.Recents.Take(Current.MaxRecents).ToList();
        }

        public void AddFavorite(FolderInfo folder)
        {
            if (folder == null)
            {
                return;
            }

            var identifier = folder.Identifier;
            if (!Current.Favorites.Any(item => item.Equals(identifier)))
            {
                Current.Favorites.Add(identifier);
            }
        }

        public void RemoveFavorite(FolderIdentifier identifier)
        {
            Current.Favorites.RemoveAll(item => item.Equals(identifier));
        }

        public void MoveFavorite(int index, int offset)
        {
            var newIndex = index + offset;
            if (index < 0 || index >= Current.Favorites.Count || newIndex < 0 || newIndex >= Current.Favorites.Count)
            {
                return;
            }

            var item = Current.Favorites[index];
            Current.Favorites.RemoveAt(index);
            Current.Favorites.Insert(newIndex, item);
        }
    }
}
