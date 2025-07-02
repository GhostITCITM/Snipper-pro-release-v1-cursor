using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.RegularExpressions;
using OpenCvSharp;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Simplified TrOCR Engine that uses Tesseract's LSTM mode optimized for handwriting
    /// This provides better handwriting recognition without requiring Python or external models
    /// </summary>
    public class TrOCREngine : IDisposable
    {
        private bool _initialized = false;
        private readonly string _tessdataPath;
        private static readonly HttpClient _httpClient = new HttpClient();

        public TrOCREngine()
        {
            _tessdataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        }

        public async Task<bool> InitializeAsync()
        {
            try
            {
                if (_initialized) return true;

                // Check if we have the required tessdata files
                if (!Directory.Exists(_tessdataPath))
                {
                    Directory.CreateDirectory(_tessdataPath);
                }

                // Download additional handwriting-optimized traineddata if needed
                await EnsureHandwritingModelsAsync();

                _initialized = true;
                Debug.WriteLine("Advanced handwriting OCR initialized");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Handwriting OCR initialization failed: {ex.Message}");
                return false;
            }
        }

        public async Task<(string text, string[] numbers)> RecognizeHandwritingAsync(Bitmap image)
        {
            return await Task.Run(() => RecognizeHandwriting(image));
        }

        public (string text, string[] numbers) RecognizeHandwriting(Bitmap image)
        {
            if (!_initialized)
            {
                var initTask = InitializeAsync();
                initTask.Wait(5000);
                if (!initTask.Result)
                {
                    return ("Handwriting OCR not available", Array.Empty<string>());
                }
            }

            try
            {
                // First, analyze the image to determine best approach
                var imageStats = AnalyzeImage(image);
                
                // Quick check - if image is already high quality, try direct recognition
                if (imageStats.IsHighQuality)
                {
                    var quickResult = RecognizeWithOptimizedTesseract(image);
                    if (!string.IsNullOrWhiteSpace(quickResult.text) && quickResult.confidence > 0.8)
                    {
                        var numbers = ExtractNumbers(quickResult.text);
                        Debug.WriteLine($"Quick recognition succeeded: '{quickResult.text}' (confidence: {quickResult.confidence:F2})");
                        return (PostProcessText(quickResult.text), numbers);
                    }
                }

                // Detect text regions for focused processing
                var textRegions = DetectTextRegions(image);
                if (textRegions.Count > 0 && textRegions.Count < 5)
                {
                    // Process regions separately for better accuracy
                    var regionResults = new List<string>();
                    foreach (var region in textRegions)
                    {
                        using (var regionBitmap = ExtractRegion(image, region))
                        {
                            var regionResult = ProcessSingleRegion(regionBitmap, imageStats);
                            if (!string.IsNullOrWhiteSpace(regionResult))
                            {
                                regionResults.Add(regionResult);
                            }
                        }
                    }
                    
                    if (regionResults.Count > 0)
                    {
                        var combinedText = string.Join(" ", regionResults);
                        var numbers = ExtractNumbers(combinedText);
                        Debug.WriteLine($"Region-based recognition: '{combinedText}'");
                        return (PostProcessText(combinedText), numbers);
                    }
                }

                // Fall back to full image processing with adaptive preprocessing
                var results = new List<(string text, double confidence)>();
                
                // Select preprocessing based on image characteristics
                var preprocessingModes = SelectPreprocessingModes(imageStats);
                
                foreach (var mode in preprocessingModes)
                {
                    using (var preprocessed = PreprocessForHandwriting(image, mode))
                    {
                        if (preprocessed == null) continue;

                        var result = RecognizeWithOptimizedTesseract(preprocessed);
                        if (!string.IsNullOrWhiteSpace(result.text))
                        {
                            results.Add((PostProcessText(result.text), result.confidence));
                            
                            // Early exit if high confidence
                            if (result.confidence > 0.9)
                                break;
                        }
                    }
                }

                // Select the best result
                if (results.Count > 0)
                {
                    var bestResult = results
                        .Where(r => !string.IsNullOrWhiteSpace(r.text))
                        .OrderByDescending(r => r.confidence * GetTextQualityScore(r.text))
                        .FirstOrDefault();

                    if (!string.IsNullOrWhiteSpace(bestResult.text))
                    {
                        var numbers = ExtractNumbers(bestResult.text);
                        Debug.WriteLine($"Adaptive handwriting OCR: '{bestResult.text}' (confidence: {bestResult.confidence:F2})");
                        return (bestResult.text, numbers);
                    }
                }

                return ("", Array.Empty<string>());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Handwriting recognition error: {ex.Message}");
                return ("Recognition error", Array.Empty<string>());
            }
        }

        private (string text, double confidence) RecognizeWithOptimizedTesseract(Bitmap image)
        {
            using (var engine = new Tesseract.TesseractEngine(_tessdataPath, "eng", Tesseract.EngineMode.LstmOnly))
            {
                // Configure for handwriting
                engine.SetVariable("preserve_interword_spaces", "1");
                engine.SetVariable("tessedit_char_blacklist", "");
                engine.SetVariable("tessedit_char_whitelist", "");
                
                // Try different page segmentation modes
                var modes = new[]
                {
                    Tesseract.PageSegMode.SingleBlock,
                    Tesseract.PageSegMode.Auto,
                    Tesseract.PageSegMode.SingleColumn,
                    Tesseract.PageSegMode.RawLine
                };

                string bestText = "";
                double bestConfidence = 0;

                foreach (var mode in modes)
                {
                    try
                    {
                        engine.DefaultPageSegMode = mode;
                        using (var pix = Tesseract.PixConverter.ToPix(image))
                        using (var page = engine.Process(pix))
                        {
                            var text = page.GetText()?.Trim() ?? "";
                            var confidence = page.GetMeanConfidence();

                            if (!string.IsNullOrWhiteSpace(text) && confidence > bestConfidence)
                            {
                                bestText = text;
                                bestConfidence = confidence;
                            }
                        }
                    }
                    catch { }
                }

                return (bestText, bestConfidence);
            }
        }

        private enum PreprocessingMode
        {
            Enhanced,
            Light,
            Adaptive,
            HighContrast
        }

        private Bitmap PreprocessForHandwriting(Bitmap original, PreprocessingMode mode)
        {
            try
            {
                var processed = new Bitmap(original);

                using (var g = Graphics.FromImage(processed))
                {
                    // Apply different preprocessing based on mode
                    switch (mode)
                    {
                        case PreprocessingMode.Enhanced:
                            // Sharpen and enhance contrast
                            ApplySharpening(processed);
                            ApplyContrastEnhancement(processed, 1.5f);
                            break;

                        case PreprocessingMode.Light:
                            // Minimal processing
                            ApplyContrastEnhancement(processed, 1.2f);
                            break;

                        case PreprocessingMode.Adaptive:
                            // Adaptive thresholding
                            ApplyAdaptiveThreshold(processed);
                            break;

                        case PreprocessingMode.HighContrast:
                            // Strong contrast and binarization
                            ApplyContrastEnhancement(processed, 2.0f);
                            ApplyBinarization(processed);
                            break;
                    }
                }

                return processed;
            }
            catch
            {
                return null;
            }
        }

        private void ApplySharpening(Bitmap image)
        {
            // Simple sharpening kernel
            float[,] kernel = {
                { -1, -1, -1 },
                { -1,  9, -1 },
                { -1, -1, -1 }
            };

            ApplyConvolution(image, kernel);
        }

        private void ApplyConvolution(Bitmap image, float[,] kernel)
        {
            var width = image.Width;
            var height = image.Height;
            var clone = (Bitmap)image.Clone();

            BitmapData srcData = clone.LockBits(
                new System.Drawing.Rectangle(0, 0, width, height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format32bppArgb);

            BitmapData destData = image.LockBits(
                new System.Drawing.Rectangle(0, 0, width, height),
                ImageLockMode.WriteOnly,
                PixelFormat.Format32bppArgb);

            int bytes = Math.Abs(srcData.Stride) * height;
            byte[] srcBuffer = new byte[bytes];
            byte[] destBuffer = new byte[bytes];

            System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcBuffer, 0, bytes);

            // Apply convolution
            for (int y = 1; y < height - 1; y++)
            {
                for (int x = 1; x < width - 1; x++)
                {
                    float red = 0, green = 0, blue = 0;

                    for (int ky = -1; ky <= 1; ky++)
                    {
                        for (int kx = -1; kx <= 1; kx++)
                        {
                            int pixelPos = ((y + ky) * srcData.Stride) + ((x + kx) * 4);
                            red += srcBuffer[pixelPos + 2] * kernel[ky + 1, kx + 1];
                            green += srcBuffer[pixelPos + 1] * kernel[ky + 1, kx + 1];
                            blue += srcBuffer[pixelPos] * kernel[ky + 1, kx + 1];
                        }
                    }

                    int destPos = (y * destData.Stride) + (x * 4);
                    destBuffer[destPos] = (byte)Math.Max(0, Math.Min(255, blue));
                    destBuffer[destPos + 1] = (byte)Math.Max(0, Math.Min(255, green));
                    destBuffer[destPos + 2] = (byte)Math.Max(0, Math.Min(255, red));
                    destBuffer[destPos + 3] = srcBuffer[destPos + 3]; // Alpha
                }
            }

            System.Runtime.InteropServices.Marshal.Copy(destBuffer, 0, destData.Scan0, bytes);

            image.UnlockBits(destData);
            clone.UnlockBits(srcData);
            clone.Dispose();
        }

        private void ApplyContrastEnhancement(Bitmap image, float contrast)
        {
            var width = image.Width;
            var height = image.Height;

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    var pixel = image.GetPixel(x, y);
                    
                    int r = (int)((pixel.R - 128) * contrast + 128);
                    int g = (int)((pixel.G - 128) * contrast + 128);
                    int b = (int)((pixel.B - 128) * contrast + 128);

                    r = Math.Max(0, Math.Min(255, r));
                    g = Math.Max(0, Math.Min(255, g));
                    b = Math.Max(0, Math.Min(255, b));

                    image.SetPixel(x, y, Color.FromArgb(r, g, b));
                }
            }
        }

        private void ApplyBinarization(Bitmap image)
        {
            // Simple Otsu's method approximation
            int threshold = CalculateOtsuThreshold(image);

            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    var pixel = image.GetPixel(x, y);
                    var gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);
                    var newColor = gray > threshold ? Color.White : Color.Black;
                    image.SetPixel(x, y, newColor);
                }
            }
        }

        private void ApplyAdaptiveThreshold(Bitmap image)
        {
            // Simplified adaptive thresholding
            var width = image.Width;
            var height = image.Height;
            var clone = (Bitmap)image.Clone();

            int windowSize = 15;
            int halfWindow = windowSize / 2;

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    // Calculate local mean
                    int sum = 0;
                    int count = 0;

                    for (int dy = -halfWindow; dy <= halfWindow; dy++)
                    {
                        for (int dx = -halfWindow; dx <= halfWindow; dx++)
                        {
                            int nx = x + dx;
                            int ny = y + dy;

                            if (nx >= 0 && nx < width && ny >= 0 && ny < height)
                            {
                                var neighborPixel = clone.GetPixel(nx, ny);
                                sum += (int)(neighborPixel.R * 0.299 + neighborPixel.G * 0.587 + neighborPixel.B * 0.114);
                                count++;
                            }
                        }
                    }

                    int localMean = sum / count;
                    var pixel = clone.GetPixel(x, y);
                    var gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);
                    
                    var newColor = gray > (localMean - 10) ? Color.White : Color.Black;
                    image.SetPixel(x, y, newColor);
                }
            }

            clone.Dispose();
        }

        private int CalculateOtsuThreshold(Bitmap image)
        {
            // Histogram
            int[] histogram = new int[256];
            
            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    var pixel = image.GetPixel(x, y);
                    var gray = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);
                    histogram[gray]++;
                }
            }

            // Calculate threshold
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

        private double GetTextQualityScore(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;

            double score = 1.0;

            // Penalize very short text
            if (text.Length < 3) score *= 0.5;

            // Penalize text with too many special characters
            var specialCharRatio = (double)Regex.Matches(text, @"[^a-zA-Z0-9\s]").Count / text.Length;
            if (specialCharRatio > 0.5) score *= 0.5;

            // Reward text with proper word boundaries
            var words = text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length > 1) score *= 1.2;

            // Reward text with alphabetic characters
            var letterRatio = (double)text.Count(char.IsLetter) / text.Length;
            score *= (0.5 + letterRatio * 0.5);

            return score;
        }

        private string[] ExtractNumbers(string text)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<string>();

            var numbers = new List<string>();
            
            // Enhanced patterns for handwritten numbers
            var patterns = new[]
            {
                @"-?\d+\.?\d*",                    // Basic numbers
                @"\$[\d,]+\.?\d*",                 // Currency
                @"\d{1,3}(?:,\d{3})*(?:\.\d+)?",  // Thousands separator
                @"\(\d+\.?\d*\)",                  // Accounting negatives
                @"\d+\s*[/%]\s*",                  // Percentages
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(text, pattern);
                foreach (Match match in matches)
                {
                    var value = match.Value.Trim();
                    if (NumberHelper.TryParseFlexible(value, out _) && !numbers.Contains(value))
                    {
                        numbers.Add(value);
                    }
                }
            }

            return numbers.ToArray();
        }

        private async Task EnsureHandwritingModelsAsync()
        {
            try
            {
                // Check if we need additional language data for better handwriting
                var scriptPath = Path.Combine(_tessdataPath, "script", "Latin.traineddata");
                if (!File.Exists(scriptPath))
                {
                    var scriptDir = Path.GetDirectoryName(scriptPath);
                    if (!Directory.Exists(scriptDir))
                    {
                        Directory.CreateDirectory(scriptDir);
                    }

                    // Note: In production, you might want to download additional traineddata files
                    // For now, we'll rely on the standard eng.traineddata with LSTM
                    Debug.WriteLine("Using standard English LSTM model for handwriting");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not ensure handwriting models: {ex.Message}");
            }
        }

        private string PostProcessText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // Fix common OCR errors in handwriting
            text = text.Replace("rn", "m");
            text = text.Replace("cl", "d");
            text = text.Replace("||", "ll");
            text = text.Replace("|", "l");
            
            // Fix spacing
            text = Regex.Replace(text, @"\s+", " ");
            text = Regex.Replace(text, @"([a-z])([A-Z])", "$1 $2");
            
            // Common word corrections
            text = text.Replace(" tne ", " the ");
            text = text.Replace(" tbe ", " the ");
            text = text.Replace(" amd ", " and ");
            text = text.Replace(" witn ", " with ");
            text = text.Replace(" tnat ", " that ");
            text = text.Replace(" wnat ", " what ");
            
            return text.Trim();
        }

        private List<System.Drawing.Rectangle> DetectTextRegions(Bitmap image)
        {
            var regions = new List<System.Drawing.Rectangle>();
            
            using (var mat = BitmapToMat(image))
            using (var gray = new Mat())
            using (var binary = new Mat())
            {
                Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                Cv2.Threshold(gray, binary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
                
                // Find contours
                Cv2.FindContours(binary, out var contours, out _, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
                
                foreach (var contour in contours)
                {
                    var rect = Cv2.BoundingRect(contour);
                    
                    // Filter based on size and aspect ratio
                    if (rect.Width > 20 && rect.Height > 10 && 
                        rect.Width < image.Width * 0.9 && 
                        rect.Height < image.Height * 0.9)
                    {
                        // Expand region slightly
                        rect.X = Math.Max(0, rect.X - 5);
                        rect.Y = Math.Max(0, rect.Y - 5);
                        rect.Width = Math.Min(image.Width - rect.X, rect.Width + 10);
                        rect.Height = Math.Min(image.Height - rect.Y, rect.Height + 10);
                        
                        regions.Add(new System.Drawing.Rectangle(rect.X, rect.Y, rect.Width, rect.Height));
                    }
                }
            }
            
            // Merge overlapping regions
            return MergeOverlappingRegions(regions);
        }

        private List<System.Drawing.Rectangle> MergeOverlappingRegions(List<System.Drawing.Rectangle> regions)
        {
            if (regions.Count <= 1) return regions;
            
            var merged = new List<System.Drawing.Rectangle>();
            var used = new bool[regions.Count];
            
            for (int i = 0; i < regions.Count; i++)
            {
                if (used[i]) continue;
                
                var current = regions[i];
                used[i] = true;
                
                // Find all overlapping regions
                bool foundOverlap;
                do
                {
                    foundOverlap = false;
                    for (int j = 0; j < regions.Count; j++)
                    {
                        if (used[j]) continue;
                        
                        if (IntersectsWith(current, regions[j]))
                        {
                            current = UnionRectangles(current, regions[j]);
                            used[j] = true;
                            foundOverlap = true;
                        }
                    }
                } while (foundOverlap);
                
                merged.Add(current);
            }
            
            return merged;
        }

        private bool IntersectsWith(System.Drawing.Rectangle a, System.Drawing.Rectangle b)
        {
            return a.IntersectsWith(b);
        }

        private System.Drawing.Rectangle UnionRectangles(System.Drawing.Rectangle a, System.Drawing.Rectangle b)
        {
            int x = Math.Min(a.X, b.X);
            int y = Math.Min(a.Y, b.Y);
            int right = Math.Max(a.Right, b.Right);
            int bottom = Math.Max(a.Bottom, b.Bottom);
            return new System.Drawing.Rectangle(x, y, right - x, bottom - y);
        }

        private Mat BitmapToMat(Bitmap bitmap)
        {
            using (var ms = new MemoryStream())
            {
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                return Cv2.ImDecode(ms.ToArray(), ImreadModes.Color);
            }
        }

        private string ProcessSingleRegion(Bitmap region, ImageStatistics stats)
        {
            // Optimize small regions differently
            if (region.Width < 100 || region.Height < 50)
            {
                // Upscale small regions for better recognition
                using (var upscaled = UpscaleImage(region, 2.0))
                {
                    var result = RecognizeWithOptimizedTesseract(upscaled);
                    return result.confidence > 0.5 ? PostProcessText(result.text) : "";
                }
            }
            
            var result1 = RecognizeWithOptimizedTesseract(region);
            return result1.confidence > 0.6 ? PostProcessText(result1.text) : "";
        }

        private Bitmap ExtractRegion(Bitmap source, System.Drawing.Rectangle region)
        {
            var extracted = new Bitmap(region.Width, region.Height);
            using (var g = Graphics.FromImage(extracted))
            {
                g.DrawImage(source, 0, 0, region, GraphicsUnit.Pixel);
            }
            return extracted;
        }

        private Bitmap UpscaleImage(Bitmap image, double scale)
        {
            int newWidth = (int)(image.Width * scale);
            int newHeight = (int)(image.Height * scale);
            var upscaled = new Bitmap(newWidth, newHeight);
            
            using (var g = Graphics.FromImage(upscaled))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(image, 0, 0, newWidth, newHeight);
            }
            
            return upscaled;
        }

        private class ImageStatistics
        {
            public bool IsHighQuality { get; set; }
            public bool IsLowContrast { get; set; }
            public bool HasNoise { get; set; }
            public bool IsHandwritten { get; set; }
            public double AverageIntensity { get; set; }
            public double ContrastRatio { get; set; }
        }

        private ImageStatistics AnalyzeImage(Bitmap image)
        {
            var stats = new ImageStatistics();
            
            // Quick analysis using sampling
            int sampleSize = 100;
            int stepX = Math.Max(1, image.Width / sampleSize);
            int stepY = Math.Max(1, image.Height / sampleSize);
            
            double sumIntensity = 0;
            double minIntensity = 255;
            double maxIntensity = 0;
            
            for (int y = 0; y < image.Height; y += stepY)
            {
                for (int x = 0; x < image.Width; x += stepX)
                {
                    var pixel = image.GetPixel(x, y);
                    var intensity = pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114;
                    
                    sumIntensity += intensity;
                    minIntensity = Math.Min(minIntensity, intensity);
                    maxIntensity = Math.Max(maxIntensity, intensity);
                }
            }
            
            int sampleCount = (image.Width / stepX) * (image.Height / stepY);
            stats.AverageIntensity = sumIntensity / sampleCount;
            stats.ContrastRatio = (maxIntensity - minIntensity) / 255.0;
            
            stats.IsHighQuality = stats.ContrastRatio > 0.7 && stats.AverageIntensity > 150;
            stats.IsLowContrast = stats.ContrastRatio < 0.5;
            stats.HasNoise = false; // Simplified for performance
            stats.IsHandwritten = true; // We're in handwriting mode
            
            return stats;
        }

        private PreprocessingMode[] SelectPreprocessingModes(ImageStatistics stats)
        {
            var modes = new List<PreprocessingMode>();
            
            if (stats.IsLowContrast)
            {
                modes.Add(PreprocessingMode.Enhanced);
                modes.Add(PreprocessingMode.HighContrast);
            }
            else if (stats.IsHighQuality)
            {
                modes.Add(PreprocessingMode.Light);
            }
            else
            {
                modes.Add(PreprocessingMode.Adaptive);
                modes.Add(PreprocessingMode.Enhanced);
            }
            
            return modes.ToArray();
        }

        public void Dispose()
        {
            // Nothing to dispose in this implementation
        }
    }
} 