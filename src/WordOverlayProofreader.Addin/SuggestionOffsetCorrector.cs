using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordOverlayProofreader.Addin
{
    /// <summary>
    /// Corrects suggestion offsets from the backend to match actual Word document positions.
    /// Handles Arabic text normalization and fuzzy matching similar to the TipTap implementation.
    /// </summary>
    public class SuggestionOffsetCorrector
    {
        private const int SEARCH_WINDOW = 15; // chars to search around backend position
        private const double MIN_SIMILARITY = 0.75; // minimum similarity score for fuzzy match
        private const double EXACT_BOOST = 0.3; // bonus for exact matches
        private const double PROXIMITY_WEIGHT = 0.6; // weight for proximity to expected position
        private const double SIMILARITY_WEIGHT = 0.4; // weight for text similarity

        /// <summary>
        /// Corrects all suggestions in the list by finding their actual positions in the document text.
        /// </summary>
        public static List<Suggestion> CorrectSuggestions(List<Suggestion> suggestions, string documentText)
        {
            if (suggestions == null || suggestions.Count == 0 || string.IsNullOrEmpty(documentText))
                return suggestions;

            var correctedSuggestions = new List<Suggestion>();
            var occurrenceTracker = new Dictionary<string, int>();

            foreach (var suggestion in suggestions)
            {
                try
                {
                    var corrected = CorrectSingleSuggestion(suggestion, documentText, occurrenceTracker);
                    if (corrected != null)
                    {
                        correctedSuggestions.Add(corrected);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[OffsetCorrector] Error correcting suggestion '{suggestion.text}': {ex.Message}");
                    // Add original if correction fails
                    correctedSuggestions.Add(suggestion);
                }
            }

            return correctedSuggestions;
        }

        private static Suggestion CorrectSingleSuggestion(Suggestion original, string docText, Dictionary<string, int> occurrenceTracker)
        {
            if (string.IsNullOrEmpty(original.text))
                return original;

            // Normalize the search text for better Arabic matching
            var searchText = ArabicTextNormalizer.Normalize(original.text);
            var normalizedDoc = ArabicTextNormalizer.Normalize(docText);

            // Track occurrence for this text pattern
            var key = searchText.ToLower();
            if (!occurrenceTracker.ContainsKey(key))
                occurrenceTracker[key] = 0;
            occurrenceTracker[key]++;
            var targetOccurrence = occurrenceTracker[key];

            // Calculate center position to search around
            int centerPos = original.from > 0 ? original.from : (original.to > 0 ? original.to : 0);
            
            // Try to find best match near the expected position
            var match = FindBestMatch(normalizedDoc, searchText, centerPos, SEARCH_WINDOW);

            if (match != null)
            {
                // Map normalized positions back to original document positions
                var actualFrom = MapNormalizedToOriginal(docText, normalizedDoc, match.Start);
                var actualTo = MapNormalizedToOriginal(docText, normalizedDoc, match.End);

                Console.WriteLine($"[OffsetCorrector] '{original.text}' corrected: {original.from}-{original.to} → {actualFrom}-{actualTo}");

                return new Suggestion
                {
                    id = original.id,
                    type = original.type,
                    text = original.text,
                    suggestion = original.suggestion,
                    from = actualFrom,
                    to = actualTo,
                    occurence = targetOccurrence,
                    requestId = original.requestId
                };
            }

            // Fallback: try global search with occurrence counting
            var globalMatch = FindNthOccurrence(normalizedDoc, searchText, targetOccurrence);
            if (globalMatch != null)
            {
                var actualFrom = MapNormalizedToOriginal(docText, normalizedDoc, globalMatch.Start);
                var actualTo = MapNormalizedToOriginal(docText, normalizedDoc, globalMatch.End);

                Console.WriteLine($"[OffsetCorrector] '{original.text}' found via global search: {actualFrom}-{actualTo}");

                return new Suggestion
                {
                    id = original.id,
                    type = original.type,
                    text = original.text,
                    suggestion = original.suggestion,
                    from = actualFrom,
                    to = actualTo,
                    occurence = targetOccurrence,
                    requestId = original.requestId
                };
            }

            Console.WriteLine($"[OffsetCorrector] '{original.text}' could not be matched, using original offsets");
            return original;
        }

        private static MatchResult FindBestMatch(string text, string pattern, int centerPos, int window)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(pattern))
                return null;

            // Define search window
            int searchStart = Math.Max(0, centerPos - window);
            int searchEnd = Math.Min(text.Length, centerPos + window + pattern.Length);
            
            if (searchStart >= text.Length || searchEnd <= searchStart)
                return null;

            var candidates = new List<MatchCandidate>();

            // 1. Try exact match first (case-insensitive)
            int exactPos = text.IndexOf(pattern, searchStart, searchEnd - searchStart, StringComparison.OrdinalIgnoreCase);
            if (exactPos >= 0)
            {
                double proximity = 1.0 - Math.Min(1.0, Math.Abs(exactPos - centerPos) / (double)window);
                double score = EXACT_BOOST + PROXIMITY_WEIGHT * proximity + SIMILARITY_WEIGHT * 1.0;
                
                candidates.Add(new MatchCandidate
                {
                    Start = exactPos,
                    End = exactPos + pattern.Length,
                    Score = score,
                    IsExact = true
                });
            }

            // 2. If no exact match, try fuzzy matching
            if (candidates.Count == 0)
            {
                int patternLen = pattern.Length;
                int fuzzPad = Math.Min(3, (int)(patternLen * 0.2)); // Allow ±20% length variation
                int minLen = Math.Max(1, patternLen - fuzzPad);
                int maxLen = Math.Min(searchEnd - searchStart, patternLen + fuzzPad);

                for (int len = minLen; len <= maxLen; len++)
                {
                    for (int i = searchStart; i + len <= searchEnd; i++)
                    {
                        var chunk = text.Substring(i, len);
                        
                        // Quick filter: check first character
                        if (!chunk[0].ToString().Equals(pattern[0].ToString(), StringComparison.OrdinalIgnoreCase))
                            continue;

                        double similarity = CalculateSimilarity(pattern, chunk);
                        if (similarity >= MIN_SIMILARITY)
                        {
                            double proximity = 1.0 - Math.Min(1.0, Math.Abs(i - centerPos) / (double)window);
                            double score = PROXIMITY_WEIGHT * proximity + SIMILARITY_WEIGHT * similarity;

                            candidates.Add(new MatchCandidate
                            {
                                Start = i,
                                End = i + len,
                                Score = score,
                                IsExact = false
                            });
                        }
                    }
                }
            }

            // Pick best candidate
            if (candidates.Count == 0)
                return null;

            var best = candidates.OrderByDescending(c => c.Score)
                                 .ThenByDescending(c => c.IsExact)
                                 .ThenBy(c => Math.Abs((c.Start + c.End) / 2 - centerPos))
                                 .First();

            return new MatchResult { Start = best.Start, End = best.End };
        }

        private static MatchResult FindNthOccurrence(string text, string pattern, int n)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(pattern) || n < 1)
                return null;

            int currentOccurrence = 0;
            int searchPos = 0;

            while (searchPos < text.Length)
            {
                int found = text.IndexOf(pattern, searchPos, StringComparison.OrdinalIgnoreCase);
                if (found < 0)
                    break;

                currentOccurrence++;
                if (currentOccurrence == n)
                {
                    return new MatchResult { Start = found, End = found + pattern.Length };
                }

                searchPos = found + 1;
            }

            return null;
        }

        private static int MapNormalizedToOriginal(string original, string normalized, int normalizedPos)
        {
            // Account for removed diacritics and normalization
            // This is a simplified mapping - for production, you'd need a more sophisticated approach
            
            if (normalizedPos >= normalized.Length)
                return original.Length;
            
            if (normalizedPos <= 0)
                return 0;

            // Simple proportional mapping (works reasonably well for Arabic)
            double ratio = (double)normalizedPos / normalized.Length;
            int mappedPos = (int)(ratio * original.Length);
            
            return Math.Min(mappedPos, original.Length);
        }

        private static double CalculateSimilarity(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1) && string.IsNullOrEmpty(s2))
                return 1.0;
            
            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
                return 0.0;

            int distance = LevenshteinDistance(s1, s2);
            int maxLen = Math.Max(s1.Length, s2.Length);
            return 1.0 - (double)distance / maxLen;
        }

        private static int LevenshteinDistance(string s1, string s2)
        {
            int m = s1.Length;
            int n = s2.Length;
            int[,] dp = new int[m + 1, n + 1];

            for (int i = 0; i <= m; i++)
                dp[i, 0] = i;
            
            for (int j = 0; j <= n; j++)
                dp[0, j] = j;

            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    dp[i, j] = Math.Min(
                        Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1),
                        dp[i - 1, j - 1] + cost
                    );
                }
            }

            return dp[m, n];
        }

        private class MatchCandidate
        {
            public int Start { get; set; }
            public int End { get; set; }
            public double Score { get; set; }
            public bool IsExact { get; set; }
        }

        private class MatchResult
        {
            public int Start { get; set; }
            public int End { get; set; }
        }
    }

    /// <summary>
    /// Normalizes Arabic text by removing diacritics and standardizing character forms.
    /// Similar to the web implementation's Arabic normalization.
    /// </summary>
    public static class ArabicTextNormalizer
    {
        // Arabic diacritics (tashkeel) and tatweel
        private static readonly Regex ArabicDiacritics = new Regex(@"[\u064B-\u0652\u0670\u0640]", RegexOptions.Compiled);

        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // Remove diacritics
            text = ArabicDiacritics.Replace(text, "");

            // Normalize alef forms
            text = text.Replace("أ", "ا")
                       .Replace("إ", "ا")
                       .Replace("آ", "ا");

            // Normalize yaa
            text = text.Replace("ى", "ي");

            // Normalize taa marbuta
            text = text.Replace("ة", "ه");

            return text;
        }
    }
}
