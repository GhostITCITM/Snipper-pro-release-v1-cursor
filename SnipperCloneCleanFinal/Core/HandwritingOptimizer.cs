using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Lightweight handwriting optimization techniques that improve recognition
    /// without requiring heavy computational resources
    /// </summary>
    public static class HandwritingOptimizer
    {
        // Common word dictionary for validation
        private static readonly HashSet<string> CommonWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Top 100 most common English words
            "the", "be", "to", "of", "and", "a", "in", "that", "have", "i",
            "it", "for", "not", "on", "with", "he", "as", "you", "do", "at",
            "this", "but", "his", "by", "from", "they", "we", "say", "her", "she",
            "or", "an", "will", "my", "one", "all", "would", "there", "their", "what",
            "so", "up", "out", "if", "about", "who", "get", "which", "go", "me",
            "when", "make", "can", "like", "time", "no", "just", "him", "know", "take",
            "people", "into", "year", "your", "good", "some", "could", "them", "see", "other",
            "than", "then", "now", "look", "only", "come", "its", "over", "think", "also",
            "back", "after", "use", "two", "how", "our", "work", "first", "well", "way",
            "even", "new", "want", "because", "any", "these", "give", "day", "most", "us"
        };

        // Common handwriting confusion patterns
        private static readonly Dictionary<string, string[]> ConfusionPatterns = new Dictionary<string, string[]>
        {
            { "rn", new[] { "m" } },
            { "m", new[] { "rn" } },
            { "cl", new[] { "d" } },
            { "d", new[] { "cl" } },
            { "h", new[] { "b" } },
            { "b", new[] { "h", "6" } },
            { "o", new[] { "0", "a" } },
            { "0", new[] { "o", "O" } },
            { "l", new[] { "1", "I" } },
            { "1", new[] { "l", "I" } },
            { "I", new[] { "l", "1" } },
            { "5", new[] { "S" } },
            { "S", new[] { "5" } },
            { "8", new[] { "B" } },
            { "B", new[] { "8" } },
            { "g", new[] { "9" } },
            { "9", new[] { "g" } },
            { "u", new[] { "v" } },
            { "v", new[] { "u" } },
            { "n", new[] { "h" } }
        };

        /// <summary>
        /// Optimize handwriting image for better OCR recognition
        /// </summary>
        public static Bitmap OptimizeForHandwriting(Bitmap original)
        {
            // Create a working copy
            var optimized = new Bitmap(original);

            // 1. Normalize size - handwriting OCR works best at specific sizes
            optimized = NormalizeSize(optimized);

            // 2. Remove noise while preserving strokes
            optimized = RemoveNoisePreserveStrokes(optimized);

            // 3. Enhance stroke continuity
            optimized = EnhanceStrokeContinuity(optimized);

            // 4. Normalize stroke thickness
            optimized = NormalizeStrokeThickness(optimized);

            return optimized;
        }

        /// <summary>
        /// Post-process OCR text to fix common handwriting recognition errors
        /// </summary>
        public static string PostProcessHandwritingText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // 1. Fix common character confusions
            text = FixCharacterConfusions(text);

            // 2. Apply word-level corrections
            text = ApplyWordCorrections(text);

            // 3. Fix spacing issues
            text = FixSpacingIssues(text);

            // 4. Clean up obvious errors
            text = CleanObviousErrors(text);

            return text;
        }

        private static Bitmap NormalizeSize(Bitmap image)
        {
            // Target height for optimal OCR (based on Tesseract recommendations)
            const int targetHeight = 100;
            
            if (Math.Abs(image.Height - targetHeight) < 20)
                return image;

            float scale = (float)targetHeight / image.Height;
            int newWidth = (int)(image.Width * scale);
            
            var normalized = new Bitmap(newWidth, targetHeight);
            using (var g = Graphics.FromImage(normalized))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(image, 0, 0, newWidth, targetHeight);
            }
            
            image.Dispose();
            return normalized;
        }

        private static Bitmap RemoveNoisePreserveStrokes(Bitmap image)
        {
            // Convert to grayscale first
            var grayscale = ConvertToGrayscale(image);
            
            // Apply median filter to remove salt-and-pepper noise
            for (int y = 1; y < grayscale.Height - 1; y++)
            {
                for (int x = 1; x < grayscale.Width - 1; x++)
                {
                    var pixels = new List<int>();
                    
                    // Get 3x3 neighborhood
                    for (int dy = -1; dy <= 1; dy++)
                    {
                        for (int dx = -1; dx <= 1; dx++)
                        {
                            var pixel = grayscale.GetPixel(x + dx, y + dy);
                            pixels.Add(pixel.R); // Since it's grayscale, R=G=B
                        }
                    }
                    
                    pixels.Sort();
                    int median = pixels[4]; // Middle value
                    
                    grayscale.SetPixel(x, y, Color.FromArgb(median, median, median));
                }
            }
            
            image.Dispose();
            return grayscale;
        }

        private static Bitmap ConvertToGrayscale(Bitmap image)
        {
            var grayscale = new Bitmap(image.Width, image.Height);
            
            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    var pixel = image.GetPixel(x, y);
                    int gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);
                    grayscale.SetPixel(x, y, Color.FromArgb(gray, gray, gray));
                }
            }
            
            return grayscale;
        }

        private static Bitmap EnhanceStrokeContinuity(Bitmap image)
        {
            // Apply morphological closing to connect broken strokes
            var enhanced = new Bitmap(image);
            
            // Simple dilation followed by erosion
            for (int pass = 0; pass < 1; pass++)
            {
                // Dilation
                var dilated = new Bitmap(enhanced);
                for (int y = 1; y < enhanced.Height - 1; y++)
                {
                    for (int x = 1; x < enhanced.Width - 1; x++)
                    {
                        int minVal = 255;
                        for (int dy = -1; dy <= 1; dy++)
                        {
                            for (int dx = -1; dx <= 1; dx++)
                            {
                                var pixel = enhanced.GetPixel(x + dx, y + dy);
                                minVal = Math.Min(minVal, pixel.R);
                            }
                        }
                        dilated.SetPixel(x, y, Color.FromArgb(minVal, minVal, minVal));
                    }
                }
                
                // Erosion
                for (int y = 1; y < dilated.Height - 1; y++)
                {
                    for (int x = 1; x < dilated.Width - 1; x++)
                    {
                        int maxVal = 0;
                        for (int dy = -1; dy <= 1; dy++)
                        {
                            for (int dx = -1; dx <= 1; dx++)
                            {
                                var pixel = dilated.GetPixel(x + dx, y + dy);
                                maxVal = Math.Max(maxVal, pixel.R);
                            }
                        }
                        enhanced.SetPixel(x, y, Color.FromArgb(maxVal, maxVal, maxVal));
                    }
                }
                dilated.Dispose();
            }
            
            image.Dispose();
            return enhanced;
        }

        private static Bitmap NormalizeStrokeThickness(Bitmap image)
        {
            // Convert to binary first
            int threshold = CalculateOtsuThreshold(image);
            
            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    var pixel = image.GetPixel(x, y);
                    var newColor = pixel.R < threshold ? Color.Black : Color.White;
                    image.SetPixel(x, y, newColor);
                }
            }
            
            return image;
        }

        private static int CalculateOtsuThreshold(Bitmap image)
        {
            int[] histogram = new int[256];
            
            // Build histogram
            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    histogram[image.GetPixel(x, y).R]++;
                }
            }
            
            int total = image.Width * image.Height;
            float sum = 0;
            for (int i = 0; i < 256; i++)
            {
                sum += i * histogram[i];
            }
            
            float sumB = 0;
            int wB = 0;
            float maxVariance = 0;
            int threshold = 0;
            
            for (int i = 0; i < 256; i++)
            {
                wB += histogram[i];
                if (wB == 0) continue;
                
                int wF = total - wB;
                if (wF == 0) break;
                
                sumB += i * histogram[i];
                
                float mB = sumB / wB;
                float mF = (sum - sumB) / wF;
                
                float variance = wB * wF * (mB - mF) * (mB - mF);
                
                if (variance > maxVariance)
                {
                    maxVariance = variance;
                    threshold = i;
                }
            }
            
            return threshold;
        }

        private static string FixCharacterConfusions(string text)
        {
            // Fix isolated character confusions
            var words = text.Split(' ');
            
            for (int i = 0; i < words.Length; i++)
            {
                var word = words[i];
                
                // Skip if it's a valid common word
                if (CommonWords.Contains(word)) continue;
                
                // Try character substitutions
                foreach (var confusion in ConfusionPatterns)
                {
                    if (word.Contains(confusion.Key))
                    {
                        foreach (var replacement in confusion.Value)
                        {
                            var candidate = word.Replace(confusion.Key, replacement);
                            if (CommonWords.Contains(candidate))
                            {
                                words[i] = candidate;
                                break;
                            }
                        }
                    }
                }
            }
            
            return string.Join(" ", words);
        }

        private static string ApplyWordCorrections(string text)
        {
            // Common handwriting word errors
            var corrections = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "tbe", "the" },
                { "tne", "the" },
                { "amd", "and" },
                { "anc", "and" },
                { "witn", "with" },
                { "witb", "with" },
                { "tnat", "that" },
                { "tbat", "that" },
                { "wnat", "what" },
                { "wnich", "which" },
                { "wbich", "which" },
                { "wnen", "when" },
                { "wno", "who" },
                { "tney", "they" },
                { "tnere", "there" },
                { "tneir", "their" },
                { "nave", "have" },
                { "nere", "here" },
                { "tnrough", "through" },
                { "otner", "other" },
                { "mignt", "might" },
                { "rignt", "right" },
                { "lignt", "light" },
                { "fignt", "fight" },
                { "signt", "sight" },
                { "heigbt", "height" },
                { "weigbt", "weight" }
            };
            
            foreach (var correction in corrections)
            {
                text = Regex.Replace(text, @"\b" + correction.Key + @"\b", correction.Value, RegexOptions.IgnoreCase);
            }
            
            return text;
        }

        private static string FixSpacingIssues(string text)
        {
            // Fix missing spaces between words
            text = Regex.Replace(text, @"([a-z])([A-Z])", "$1 $2");
            
            // Fix multiple spaces
            text = Regex.Replace(text, @"\s+", " ");
            
            // Fix spaces before punctuation
            text = Regex.Replace(text, @"\s+([.,!?;:])", "$1");
            
            // Fix missing spaces after punctuation
            text = Regex.Replace(text, @"([.,!?;:])([A-Za-z])", "$1 $2");
            
            return text.Trim();
        }

        private static string CleanObviousErrors(string text)
        {
            // Remove standalone single characters that don't make sense
            text = Regex.Replace(text, @"\b[^aAiI]\b", "", RegexOptions.IgnoreCase);
            
            // Remove repeated characters (e.g., "ttthe" -> "the")
            text = Regex.Replace(text, @"(.)\1{2,}", "$1$1");
            
            // Fix common OCR artifacts
            text = text.Replace("||", "ll");
            text = text.Replace("|", "l");
            text = text.Replace("~", "-");
            
            return text;
        }

        /// <summary>
        /// Merge multiple OCR results intelligently
        /// </summary>
        public static string MergeOCRResults(List<(string text, double confidence)> results)
        {
            if (results == null || results.Count == 0)
                return "";
            
            if (results.Count == 1)
                return results[0].text;
            
            // For each word position, vote on the best candidate
            var allWords = results.Select(r => r.text.Split(' ')).ToList();
            var maxWords = allWords.Max(w => w.Length);
            
            var mergedWords = new List<string>();
            
            for (int pos = 0; pos < maxWords; pos++)
            {
                var candidates = new Dictionary<string, double>();
                
                for (int i = 0; i < results.Count; i++)
                {
                    if (pos < allWords[i].Length)
                    {
                        var word = allWords[i][pos];
                        var confidence = results[i].confidence;
                        
                        if (!candidates.ContainsKey(word))
                            candidates[word] = 0;
                        
                        candidates[word] += confidence;
                        
                        // Bonus for common words
                        if (CommonWords.Contains(word))
                            candidates[word] += 0.2;
                    }
                }
                
                if (candidates.Count > 0)
                {
                    var bestWord = candidates.OrderByDescending(c => c.Value).First().Key;
                    mergedWords.Add(bestWord);
                }
            }
            
            return string.Join(" ", mergedWords);
        }
    }
} 