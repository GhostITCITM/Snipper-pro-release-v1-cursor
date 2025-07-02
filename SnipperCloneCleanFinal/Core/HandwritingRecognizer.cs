using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Tesseract;
using OpenCvSharp;
using System.Diagnostics;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Specialized handwriting recognition using Tesseract LSTM with enhanced preprocessing
    /// </summary>
    public static class HandwritingRecognizer
    {
        private static readonly string _tessdataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        
        public static (string text, string[] numbers) Recognize(Bitmap bmp)
        {
            try
            {
                // Apply handwriting-specific preprocessing
                using var preprocessed = PreprocessForHandwriting(bmp);
                
                string text = string.Empty;
                double bestConfidence = 0;

                // Try multiple segmentation modes optimized for handwriting
                var handwritingModes = new[]
                {
                    PageSegMode.SingleBlock,      // Best for paragraphs of handwriting
                    PageSegMode.SingleColumn,     // Good for notes
                    PageSegMode.SparseText,       // Good for scattered handwritten text
                    PageSegMode.Auto,             // Let Tesseract decide
                    PageSegMode.RawLine           // For single lines of handwriting
                };

                using (var engine = new TesseractEngine(_tessdataPath, "eng", EngineMode.LstmOnly))
                {
                    // Enable LSTM which is better for handwriting
                    engine.SetVariable("tessedit_pageseg_mode", "6");
                    engine.SetVariable("preserve_interword_spaces", "1");
                    
                    foreach (var mode in handwritingModes)
                    {
                        try
                        {
                            engine.DefaultPageSegMode = mode;
                            using var pix = PixConverter.ToPix(preprocessed);
                            using var page = engine.Process(pix);
                            
                            var confidence = page.GetMeanConfidence();
                            var candidateText = page.GetText()?.Trim() ?? string.Empty;
                            
                            // Keep the best result based on confidence and text length
                            if (!string.IsNullOrWhiteSpace(candidateText) && 
                                (confidence > bestConfidence || 
                                 (confidence == bestConfidence && candidateText.Length > text.Length)))
                            {
                                text = candidateText;
                                bestConfidence = confidence;
                            }
                        }
                        catch
                        {
                            // Continue with next mode
                        }
                    }
                }

                // If LSTM fails, try legacy engine as fallback
                if (string.IsNullOrWhiteSpace(text))
                {
                    using (var engine = new TesseractEngine(_tessdataPath, "eng", EngineMode.TesseractOnly))
                    {
                        engine.DefaultPageSegMode = PageSegMode.Auto;
                        using var pix = PixConverter.ToPix(preprocessed);
                        using var page = engine.Process(pix);
                        text = page.GetText()?.Trim() ?? string.Empty;
                    }
                }

                // Extract numbers with patterns suitable for handwritten text
                var numbers = ExtractHandwrittenNumbers(text);
                
                Debug.WriteLine($"Handwriting recognition result: '{text}' (confidence: {bestConfidence:F2})");
                
                return (text, numbers);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Handwriting recognition error: {ex.Message}");
                return (string.Empty, Array.Empty<string>());
            }
        }

        private static Bitmap PreprocessForHandwriting(Bitmap original)
        {
            using var mat = BitmapToMat(original);
            using var processed = new Mat();
            
            // Convert to grayscale
            Cv2.CvtColor(mat, processed, ColorConversionCodes.BGR2GRAY);
            
            // Apply bilateral filter to reduce noise while preserving edges
            using var filtered = new Mat();
            Cv2.BilateralFilter(processed, filtered, 9, 75, 75);
            
            // Increase contrast using CLAHE (Contrast Limited Adaptive Histogram Equalization)
            using var clahe = Cv2.CreateCLAHE(clipLimit: 2.0, tileGridSize: new OpenCvSharp.Size(8, 8));
            using var enhanced = new Mat();
            clahe.Apply(filtered, enhanced);
            
            // Apply morphological operations to connect broken strokes
            using var kernel = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(2, 2));
            using var morphed = new Mat();
            Cv2.MorphologyEx(enhanced, morphed, MorphTypes.Close, kernel);
            
            // Binarize with Otsu's method (good for handwriting)
            using var binary = new Mat();
            Cv2.Threshold(morphed, binary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
            
            // Slight dilation to make strokes thicker
            using var dilateKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(2, 2));
            using var dilated = new Mat();
            Cv2.Dilate(binary, dilated, dilateKernel, iterations: 1);
            
            // Optional: Deskew if text is slanted
            var deskewed = DeskewImage(dilated);
            
            return MatToBitmap(deskewed);
        }

        private static Mat DeskewImage(Mat image)
        {
            // Find contours
            Cv2.FindContours(image, out var contours, out _, RetrievalModes.List, ContourApproximationModes.ApproxSimple);
            
            if (contours.Length == 0) return image.Clone();
            
            // Get the largest contour (assuming it's the text area)
            var largestContour = contours.OrderByDescending(c => Cv2.ContourArea(c)).FirstOrDefault();
            if (largestContour == null) return image.Clone();
            
            // Get minimum area rectangle
            var rect = Cv2.MinAreaRect(largestContour);
            var angle = rect.Angle;
            
            // Adjust angle for proper orientation
            if (angle < -45)
                angle += 90;
            
            // Only deskew if angle is significant
            if (Math.Abs(angle) < 0.5)
                return image.Clone();
            
            // Rotate image
            var center = new Point2f(image.Width / 2f, image.Height / 2f);
            var rotMatrix = Cv2.GetRotationMatrix2D(center, angle, 1.0);
            var deskewed = new Mat();
            Cv2.WarpAffine(image, deskewed, rotMatrix, image.Size(), InterpolationFlags.Cubic, BorderTypes.Replicate);
            
            return deskewed;
        }

        private static string[] ExtractHandwrittenNumbers(string text)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<string>();
            
            var numbers = new System.Collections.Generic.HashSet<string>();
            
            // Patterns optimized for handwritten numbers (may have inconsistent spacing)
            var patterns = new[]
            {
                @"\$?\s*[\d,]+\.?\d*",                     // Currency with optional space after $
                @"-?\s*\d+\.?\d*",                         // Numbers with optional space after minus
                @"\d{1,3}(?:[,\s]\d{3})*(?:\.\d+)?",      // Thousands with comma or space
                @"\(\s*\d+\.?\d*\s*\)",                   // Accounting negatives with spaces
                @"\d+\s*\.\s*\d+",                        // Decimals with spaces around dot
                @"(?<=\s|^)-?\d[\d\s,\.]*\d(?=\s|$)"     // Numbers with internal spaces
            };
            
            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(text, pattern);
                foreach (Match m in matches)
                {
                    var value = m.Value;
                    
                    // Clean the value - remove internal spaces and standardize
                    value = Regex.Replace(value, @"\s+", "");
                    value = value.Replace("$", "").Replace(",", "").Trim();
                    
                    // Handle accounting format negatives
                    if (value.StartsWith("(") && value.EndsWith(")"))
                    {
                        value = "-" + value.Trim('(', ')');
                    }
                    
                    // Validate and add if it's a valid number
                    if (double.TryParse(value, out _) && !string.IsNullOrWhiteSpace(value))
                    {
                        numbers.Add(value);
                    }
                }
            }
            
            return numbers.ToArray();
        }

        private static Mat BitmapToMat(Bitmap bitmap)
        {
            using var stream = new MemoryStream();
            bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
            return Cv2.ImDecode(stream.ToArray(), ImreadModes.Color);
        }

        private static Bitmap MatToBitmap(Mat mat)
        {
            byte[] bytes;
            Cv2.ImEncode(".png", mat, out bytes);
            using var stream = new MemoryStream(bytes);
            return new Bitmap(stream);
        }
    }
} 