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

        private static bool IsSubsequence(string small, string large)
        {
            if (string.IsNullOrEmpty(small)) return true;
            if (string.IsNullOrEmpty(large)) return false;

            int si = 0, li = 0;
            while (si < small.Length && li < large.Length)
            {
                if (small[si] == large[li]) si++;
                li++;
            }
            return si == small.Length;
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

            // Tokenize query and normalize to lower for fast comparisons
            var tokens = query.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                              .Select(t => t.ToLowerInvariant()).ToArray();

            var displayName = (folder.DisplayName ?? string.Empty).ToLowerInvariant();
            var fullPath = (folder.FullPath ?? string.Empty).ToLowerInvariant();

            // Primary ranking should favor matches in DisplayName
            int tokenMatchSum = 0;

            foreach (var token in tokens)
            {
                int tokenScore = 0;

                if (displayName.Equals(token, StringComparison.OrdinalIgnoreCase))
                {
                    tokenScore += 300; // exact folder name match
                }
                else if (displayName.StartsWith(token, StringComparison.OrdinalIgnoreCase))
                {
                    tokenScore += 200; // prefix match on name
                }
                else if (displayName.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    tokenScore += 150; // substring in name
                }
                else if (IsSubsequence(token, displayName))
                {
                    tokenScore += 90; // fuzzy subsequence in name
                }

                // If not found in name, check path
                if (tokenScore == 0)
                {
                    if (fullPath.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        tokenScore += 80; // substring in path
                    }
                    else if (IsSubsequence(token, fullPath))
                    {
                        tokenScore += 40; // subsequence in path
                    }
                }

                // If token didn't match anywhere, this folder should not be considered relevant
                if (tokenScore == 0)
                {
                    return -1;
                }

                // smaller tokens are less significant, scale slightly by token length
                tokenScore += Math.Min(20, token.Length * 3);

                tokenMatchSum += tokenScore;
            }

            score += tokenMatchSum;

            return score;
        }
    }
}
