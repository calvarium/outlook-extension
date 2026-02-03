using System;
using System.Collections.Generic;
using System.Linq;

namespace outlook_extension
{
    public class SearchService
    {
        private readonly SettingsService _settingsService;

        public SearchService(SettingsService settingsService)
        {
            _settingsService = settingsService;
        }

        public void NotifySettingsChanged()
        {
        }

        public List<FolderInfo> Search(string query, IReadOnlyList<FolderInfo> folders)
        {
            var normalizedQuery = (query ?? string.Empty).Trim();
            var results = new List<(FolderInfo folder, int score)>();

            foreach (var folder in folders)
            {
                if (_settingsService.Current.ShowInboxOnly && !folder.IsUnderInbox)
                {
                    continue;
                }

                var score = ScoreFolder(folder, normalizedQuery);
                if (score >= 0)
                {
                    results.Add((folder, score));
                }
            }

            return results
                .OrderByDescending(item => item.score)
                .ThenBy(item => item.folder.FullPath)
                .Take(50)
                .Select(item => item.folder)
                .ToList();
        }

        private int ScoreFolder(FolderInfo folder, string query)
        {
            var score = 0;
            var favorites = _settingsService.Current.Favorites;
            var recents = _settingsService.Current.Recents;

            if (favorites.Any(item => item.Equals(folder.Identifier)))
            {
                score += 1000;
            }

            var recentIndex = recents.FindIndex(item => item.Equals(folder.Identifier));
            if (recentIndex >= 0)
            {
                score += 500 - recentIndex;
            }

            if (string.IsNullOrWhiteSpace(query))
            {
                return score;
            }

            var fullPath = folder.FullPath ?? string.Empty;
            var folderName = folder.DisplayName ?? string.Empty;

            if (string.Equals(folderName, query, StringComparison.OrdinalIgnoreCase))
            {
                score += 200;
            }

            if (fullPath.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 100;
                return score;
            }

            return -1;
        }
    }
}
