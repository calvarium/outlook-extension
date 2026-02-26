using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace outlook_extension
{
    public class SearchService
    {
        private readonly SettingsService _settingsService;

        // common synonyms (can be extended)
        private static readonly Dictionary<string, string[]> _synonyms = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            { "inbox", new[] { "posteingang" } },
            { "posteingang", new[] { "inbox" } },
            { "sent", new[] { "sent items", "gesendete elemente", "gesendet" } },
        };

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
            // Use StoreOrder from FolderInfo so we preserve Outlook mailbox ordering
            var results = new List<(FolderInfo folder, int score, int depth, int storeOrder)>();

            // pre-normalize query tokens and include synonyms
            var queryTokens = TokenizeAndExpand(normalizedQuery).ToArray();

            for (int i = 0; i < folders.Count; i++)
            {
                var folder = folders[i];

                if (_settingsService.Current.ShowInboxOnly && !folder.IsUnderInbox)
                {
                    continue;
                }

                var score = ScoreFolder(folder, queryTokens);
                if (score >= 0)
                {
                    var depth = GetPathSegments(folder.FullPath).Length;
                    var storeOrder = folder?.StoreOrder ?? int.MaxValue;
                    results.Add((folder, score, depth, storeOrder));
                }
            }

            return results
                .OrderByDescending(item => item.score)
                .ThenBy(item => item.storeOrder) // prefer folders in Outlook store order
                .ThenBy(item => item.depth) // prefer shallower folders
                .ThenBy(item => item.folder.FullPath)
                .Take(50)
                .Select(item => item.folder)
                .ToList();
        }

        private static string[] TokenizeAndExpand(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return new string[0];

            // normalize and split on whitespace and punctuation
            var normalized = Normalize(query);
            var tokens = Regex.Split(normalized, "\\s+")
                              .Where(t => !string.IsNullOrWhiteSpace(t))
                              .ToList();

            // expand synonyms for tokens (simple expansion)
            var expanded = new List<string>();
            foreach (var t in tokens)
            {
                expanded.Add(t);
                if (_synonyms.TryGetValue(t, out var syns))
                {
                    foreach (var s in syns)
                    {
                        expanded.Add(Normalize(s));
                    }
                }
            }

            return expanded.Distinct().ToArray();
        }

        private static string[] GetPathSegments(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath)) return new string[0];
            // common separators: '>' (used in this project), backslash, slash, colon
            var parts = fullPath.Split(new[] { '>', '\\', '/', '|', ':' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Select(p => Normalize(p))
                                 .Where(p => !string.IsNullOrWhiteSpace(p))
                                 .ToArray();
            return parts;
        }

        private int ScoreFolder(FolderInfo folder, string[] queryTokens)
        {
            var score = 0;
            var favorites = _settingsService.Current.Favorites;
            var recents = _settings_service_placeholder();

            if (favorites.Any(item => item.Equals(folder.Identifier)))
            {
                score += 1000;
            }

            var recentIndex = recents.FindIndex(item => item.Equals(folder.Identifier));
            if (recentIndex >= 0)
            {
                score += Math.Max(0, 400 - recentIndex);
            }

            if (queryTokens.Length == 0)
            {
                return score;
            }

            var displayName = Normalize(folder.DisplayName ?? string.Empty);
            var fullPath = Normalize(folder.FullPath ?? string.Empty);
            var segments = GetPathSegments(folder.FullPath);
            var lastSegment = segments.Length > 0 ? segments[segments.Length - 1] : displayName;

            int tokenMatchSum = 0;

            foreach (var token in queryTokens)
            {
                if (string.IsNullOrWhiteSpace(token)) continue;

                int tokenScore = 0;

                // 1) Exact match on last segment or display name
                if (string.Equals(token, lastSegment, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(token, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    tokenScore += 1200; // very strong boost
                }
                else if (lastSegment.StartsWith(token, StringComparison.OrdinalIgnoreCase) || displayName.StartsWith(token, StringComparison.OrdinalIgnoreCase))
                {
                    tokenScore += 800; // prefix on last segment
                }
                else
                {
                    // substring checks
                    if (lastSegment.Contains(token)) tokenScore += 600;
                    if (displayName.Contains(token)) tokenScore += 550;
                }

                // Levenshtein fuzzy match against last segment (allow misspellings)
                var fuzzyLast = FuzzyRatio(token, lastSegment);
                if (fuzzyLast > 0.85) tokenScore += 700; // almost exact
                else if (fuzzyLast > 0.7) tokenScore += 450;
                else if (fuzzyLast > 0.5) tokenScore += 220;

                // best match among other segments
                var bestOther = 0.0;
                foreach (var seg in segments)
                {
                    if (seg == lastSegment) continue;
                    var r = FuzzyRatio(token, seg);
                    if (r > bestOther) bestOther = r;
                }

                if (bestOther > 0.85) tokenScore += 300;
                else if (bestOther > 0.7) tokenScore += 180;
                else if (bestOther > 0.5) tokenScore += 90;

                // substring in full path
                if (fullPath.Contains(token)) tokenScore += 120;

                // acronym matching: e.g. 'inb' -> 'inbox', or 'pm' -> 'Posteingang Mails' etc.
                if (IsAcronymMatch(token, segments)) tokenScore += 300;

                // subsequence matching (characters in order)
                if (IsSubsequence(token, lastSegment)) tokenScore += 150;
                else if (IsSubsequence(token, fullPath)) tokenScore += 60;

                // penalize if token is very short and only matches weakly
                if (token.Length <= 2 && tokenScore < 200)
                {
                    tokenScore = Math.Max(0, tokenScore - 80);
                }

                // if token didn't match anywhere reasonably, reject folder
                if (tokenScore <= 0)
                {
                    return -1;
                }

                tokenMatchSum += tokenScore;
            }

            // small bonus for shorter paths (more specific)
            var pathDepth = segments.Length;
            var depthBonus = Math.Max(0, 50 - (pathDepth * 2));

            score += tokenMatchSum + depthBonus;

            return score;
        }

        private static string Normalize(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            var s = input.Trim().ToLowerInvariant();
            s = RemoveDiacritics(s);
            // replace punctuation with spaces, preserve alphanumerics and separators
            s = Regex.Replace(s, "[\\p{P}+\\p{S}]+", " ");
            s = Regex.Replace(s, "\\s+", " ");
            return s.Trim();
        }

        private static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();
            foreach (var ch in normalized)
            {
                var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(ch);
                }
            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
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

        private static bool IsAcronymMatch(string token, string[] segments)
        {
            if (string.IsNullOrWhiteSpace(token) || segments == null || segments.Length == 0) return false;
            // Build acronym from last few segments (up to 5)
            var acronymBuilder = new StringBuilder();
            for (int i = Math.Max(0, segments.Length - 5); i < segments.Length; i++)
            {
                var seg = segments[i];
                var parts = seg.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                {
                    acronymBuilder.Append(p[0]);
                }
            }
            var acronym = acronymBuilder.ToString();
            if (string.IsNullOrEmpty(acronym)) return false;
            return IsSubsequence(token, acronym) || acronym.Contains(token);
        }

        // Returns ratio between 0..1 where 1 means identical
        private static double FuzzyRatio(string a, string b)
        {
            if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return 1.0;
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;

            if (a.Equals(b, StringComparison.OrdinalIgnoreCase)) return 1.0;

            // limit lengths for performance
            var maxLen = Math.Max(a.Length, b.Length);
            var dist = LevenshteinDistance(a, b);
            var ratio = 1.0 - (double)dist / (double)maxLen;
            return Math.Max(0.0, Math.Min(1.0, ratio));
        }

        // classic Levenshtein distance
        private static int LevenshteinDistance(string s, string t)
        {
            if (s == null) s = string.Empty;
            if (t == null) t = string.Empty;

            var n = s.Length;
            var m = t.Length;
            if (n == 0) return m;
            if (m == 0) return n;

            // Use optimized two-row algorithm
            var prev = new int[m + 1];
            var curr = new int[m + 1];

            for (int j = 0; j <= m; j++) prev[j] = j;

            for (int i = 1; i <= n; i++)
            {
                curr[0] = i;
                var si = s[i - 1];
                for (int j = 1; j <= m; j++)
                {
                    var cost = (si == t[j - 1]) ? 0 : 1;
                    var insertion = curr[j - 1] + 1;
                    var deletion = prev[j] + 1;
                    var substitution = prev[j - 1] + cost;
                    var value = insertion;
                    if (deletion < value) value = deletion;
                    if (substitution < value) value = substitution;
                    curr[j] = value;
                }

                // swap
                var tmp = prev;
                prev = curr;
                curr = tmp;
            }

            return prev[m];
        }

        // Placeholder to avoid referencing _settingsService.Current.Recents directly in the editor diff context; actual code uses _settingsService.Current.Recents
        private List<FolderIdentifier> _settings_service_placeholder()
        {
            return _settingsService.Current.Recents;
        }
    }
}
