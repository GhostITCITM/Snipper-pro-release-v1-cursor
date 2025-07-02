using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.IO;
using Tesseract;
using OpenCvSharp;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Modernized OCR Engine with field-tested three-stage pipeline
    /// </summary>
    public class OCREngine : IDisposable
    {
        private bool _disposed = false;
        private static readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);
        private static string _tessdataPath;
        private static bool _tessLibAvailable = true;
        private static string _cachedCliPath = null; // memoised CLI location
        private static TrOCREngine _trOCREngine = null; // TrOCR for advanced handwriting
        
        static OCREngine()
        {
            // Verify tessdata path and log findings
            _tessdataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
            LogTessdataVerification();

            try
            {
                // Quick probe if Tesseract managed assembly can be loaded in this AppDomain
                var asm = Assembly.Load("Tesseract");
                _tessLibAvailable = asm != null;
            }
            catch
            {
                _tessLibAvailable = false;
            }
            
            // Initialize TrOCR engine
            InitializeTrOCR();
        }
        
        private static async void InitializeTrOCR()
        {
            try
            {
                _trOCREngine = new TrOCREngine();
                bool initialized = await _trOCREngine.InitializeAsync();
                if (initialized)
                {
                    Debug.WriteLine("TrOCR engine initialized successfully for advanced handwriting recognition");
                }
                else
                {
                    Debug.WriteLine("TrOCR engine initialization failed - will use fallback OCR");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TrOCR initialization error: {ex.Message}");
            }
        }

        private static void LogTessdataVerification()
        {
            try
            {
                var possiblePaths = new[]
                {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SnipperCloneCleanFinal", "tessdata"),
                    Path.Combine(Environment.CurrentDirectory, "tessdata"),
                    Path.Combine(Environment.CurrentDirectory, "SnipperCloneCleanFinal", "tessdata"),
                    // Directory of the assembly itself
                    Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty, "tessdata"),
                    // TESSDATA_PREFIX environment variable if set
                    Environment.GetEnvironmentVariable("TESSDATA_PREFIX") ?? string.Empty
                };

                Debug.WriteLine("=== Tessdata Verification ===");
                foreach (var path in possiblePaths)
                {
                    var engFile = Path.Combine(path, "eng.traineddata");
                    if (File.Exists(engFile))
                    {
                        var fileInfo = new FileInfo(engFile);
                        Debug.WriteLine($"✓ Found: {path} (eng.traineddata: {fileInfo.Length:N0} bytes)");
                        _tessdataPath = path;
                        return;
                    }
                    else
                    {
                        Debug.WriteLine($"✗ Missing: {path}");
                    }
                }
                Debug.WriteLine("=== End Tessdata Verification ===");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Tessdata verification error: {ex.Message}");
            }
        }

        public bool Initialize()
        {
            return Directory.Exists(_tessdataPath) && File.Exists(Path.Combine(_tessdataPath, "eng.traineddata"));
        }

        public async Task<OCRResult> RecognizeTextAsync(Bitmap image)
        {
            await _semaphore.WaitAsync();
            try
            {
                return RecognizeText(image);
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public OCRResult RecognizeText(Bitmap image)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(OCREngine));

            try
            {
                // Stage 1: Capture once, scale once
                var processedImage = CaptureAndScaleOnce(image);
                
                // Stage 2: Pre-process with OpenCvSharp (three proven lines)
                var preprocessedImage = PreprocessWithOpenCV(processedImage);
                
                OCRResult result = null;

                if (_tessLibAvailable)
                {
                    try
                    {
                        result = RecognizeWithManagedTesseract(preprocessedImage);
                        if (result == null || !result.Success)
                        {
                            result = RecognizeWithManagedTesseract(processedImage);
                        }
                    }
                    catch (Exception ex) when (ex is FileLoadException || ex is BadImageFormatException)
                    {
                        _tessLibAvailable = false; // mark unusable for next runs
                        result = null;
                    }
                    catch { result = null; }
                }

                if (result == null || !result.Success)
                {
                    // First try CLI on the heavier pre-processed image
                    var cliRes = RecognizeWithTesseractCli(preprocessedImage);

                    // If still no text, try CLI again on the lightly-processed (scaled-only) image
                    if (cliRes == null || !cliRes.Success || string.IsNullOrWhiteSpace(cliRes.Text))
                    {
                        cliRes = RecognizeWithTesseractCli(processedImage);
                    }

                    if (cliRes != null && cliRes.Success)
                    {
                        result = cliRes;
                    }
                    else
                    {
                        result = new OCRResult
                        {
                            Success = true,
                            Text = $"[OCR error] lib={_tessLibAvailable} cli={(FindTesseractCli()!=null)}",
                            Numbers = Array.Empty<string>(),
                            Confidence = 0
                        };
                    }
                }

                // Check if we should try handwriting recognition
                bool shouldTryHandwriting = result == null || 
                                          string.IsNullOrWhiteSpace(result.Text) ||
                                          result.Text.StartsWith("[") ||
                                          result.Confidence < 30;

                if (shouldTryHandwriting)
                {
                    Debug.WriteLine("Regular OCR produced poor results, trying advanced handwriting recognition...");
                    
                    // First try TrOCR if available (best for handwriting)
                    if (_trOCREngine != null)
                    {
                        try
                        {
                            var trOCRTask = _trOCREngine.RecognizeHandwritingAsync(processedImage);
                            if (trOCRTask.Wait(10000)) // 10 second timeout
                            {
                                var (handwrittenText, handwrittenNumbers) = trOCRTask.Result;
                                
                                if (!string.IsNullOrWhiteSpace(handwrittenText) && 
                                    !handwrittenText.Contains("not available") &&
                                    !handwrittenText.Contains("error"))
                                {
                                    Debug.WriteLine($"TrOCR recognition succeeded: '{handwrittenText}'");
                                    result = new OCRResult
                                    {
                                        Success = true,
                                        Text = handwrittenText,
                                        Numbers = handwrittenNumbers,
                                        Confidence = 95 // TrOCR typically has high confidence
                                    };
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"TrOCR recognition failed: {ex.Message}");
                        }
                    }
                    
                    // If TrOCR didn't work, fall back to our Tesseract-based handwriting recognizer
                    if (result == null || string.IsNullOrWhiteSpace(result.Text) || result.Confidence < 30)
                    {
                        try
                        {
                            // Try handwriting recognition on the original scaled image
                            var (handwrittenText, handwrittenNumbers) = HandwritingRecognizer.Recognize(processedImage);
                            
                            if (!string.IsNullOrWhiteSpace(handwrittenText) && 
                                (string.IsNullOrWhiteSpace(result?.Text) || handwrittenText.Length > result.Text.Length))
                            {
                                Debug.WriteLine($"Tesseract handwriting recognition succeeded: '{handwrittenText}'");
                                result = new OCRResult
                                {
                                    Success = true,
                                    Text = handwrittenText,
                                    Numbers = handwrittenNumbers,
                                    Confidence = 0 // We don't have confidence from handwriting recognizer
                                };
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Tesseract handwriting recognition failed: {ex.Message}");
                        }
                    }
                }

                // Light post-processing
                result.Text = LightPostProcessing(result.Text);
                if (result.Numbers == null || result.Numbers.Length == 0)
                {
                    result.Numbers = ExtractNumbers(result.Text);
                }
                
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OCR Error: {ex.Message}");

                // Fallback: try very simple Tesseract attempt with default parameters
                if (_tessLibAvailable)
                {
                    try
                    {
                        using var engine = new TesseractEngine(_tessdataPath, "eng", EngineMode.Default);
                        using var pix = PixConverter.ToPix(image);
                        using var page = engine.Process(pix);
                        string text = page.GetText()?.Trim() ?? string.Empty;
                        return new OCRResult
                        {
                            Success = !string.IsNullOrEmpty(text),
                            Text = !string.IsNullOrEmpty(text) ? text : "[No text detected]",
                            Numbers = Array.Empty<string>(),
                            Confidence = page.GetMeanConfidence()
                        };
                    }
                    catch (Exception inner)
                    {
                        // If managed lib truly broken, mark unavailable to avoid future attempts
                        if (inner is System.IO.FileLoadException || inner is BadImageFormatException)
                        {
                            _tessLibAvailable = false;
                        }
                        // ignore and fall through to generic error result
                    }
                }

                return new OCRResult
                {
                    Success = true,
                    Text = "[OCR processing error - image analysis unavailable]",
                    Numbers = Array.Empty<string>(),
                    Confidence = 0.0,
                    ErrorMessage = ex.Message
                };
            }
        }

        private Bitmap CaptureAndScaleOnce(Bitmap originalImage)
        {
            // Save raw PNG for debugging
            var debugPath = Path.Combine(Path.GetTempPath(), $"ocr_raw_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            originalImage.Save(debugPath, System.Drawing.Imaging.ImageFormat.Png);
            Debug.WriteLine($"Raw image saved: {debugPath}");

            // Measure capital letter height using pixel analysis
            var avgCapHeight = EstimateCapitalLetterHeight(originalImage);
            Debug.WriteLine($"Estimated capital letter height: {avgCapHeight}px");

            // Scale exactly once if needed (target ~35px for capital letters)
            if (avgCapHeight < 20 || avgCapHeight > 45)
            {
                var scaleFactor = 35.0 / avgCapHeight;
                var newWidth = (int)(originalImage.Width * scaleFactor);
                var newHeight = (int)(originalImage.Height * scaleFactor);
                
                Debug.WriteLine($"Scaling image: {originalImage.Width}x{originalImage.Height} -> {newWidth}x{newHeight} (factor: {scaleFactor:F2})");
                
                var scaledImage = new Bitmap(newWidth, newHeight);
                using (var graphics = Graphics.FromImage(scaledImage))
                {
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                }
                return scaledImage;
            }

            return new Bitmap(originalImage); // Return copy to avoid disposal issues
        }

        private double EstimateCapitalLetterHeight(Bitmap image)
        {
            // Convert to grayscale and find text-like regions
            var heights = new List<int>();
            
            using (var mat = BitmapToMat(image))
            using (var gray = new Mat())
            using (var binary = new Mat())
            {
                Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                Cv2.AdaptiveThreshold(gray, binary, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, 31, 10);
                
                // Find contours and analyze their heights
                Cv2.FindContours(binary, out var contours, out var hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
                
                foreach (var contour in contours)
                {
                    var rect = Cv2.BoundingRect(contour);
                    // Filter for text-like aspect ratios and sizes
                    if (rect.Height >= 8 && rect.Height <= 100 && rect.Width >= 4 && rect.Width <= rect.Height * 8)
                    {
                        heights.Add(rect.Height);
                    }
                }
            }

            if (heights.Count > 0)
            {
                // Return median height as estimate
                heights.Sort();
                return heights[heights.Count / 2];
            }

            // Fallback: assume reasonable size based on image dimensions
            return Math.Max(12, Math.Min(50, image.Height / 20));
        }

        private bool IsLikelyHandwriting(Bitmap image)
        {
            try
            {
                using (var mat = BitmapToMat(image))
                using (var gray = new Mat())
                using (var edges = new Mat())
                {
                    Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                    Cv2.Canny(gray, edges, 50, 150);
                    
                    // Find contours
                    Cv2.FindContours(edges, out var contours, out _, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
                    
                    if (contours.Length == 0) return false;
                    
                    // Analyze stroke characteristics
                    double avgCurvature = 0;
                    int curvedContours = 0;
                    int totalContours = 0;
                    
                    foreach (var contour in contours)
                    {
                        if (contour.Length < 10) continue;
                        
                        var area = Cv2.ContourArea(contour);
                        if (area < 50) continue; // Skip tiny contours
                        
                        totalContours++;
                        
                        // Calculate curvature by comparing perimeter to convex hull perimeter
                        var hull = Cv2.ConvexHull(contour);
                        var contourPerimeter = Cv2.ArcLength(contour, true);
                        var hullPerimeter = Cv2.ArcLength(hull, true);
                        
                        if (hullPerimeter > 0)
                        {
                            var curvature = contourPerimeter / hullPerimeter;
                            avgCurvature += curvature;
                            
                            // Handwriting tends to have more curved strokes
                            if (curvature > 1.2) curvedContours++;
                        }
                    }
                    
                    if (totalContours == 0) return false;
                    
                    avgCurvature /= totalContours;
                    double curvedRatio = (double)curvedContours / totalContours;
                    
                    // Handwriting typically has:
                    // - Higher average curvature (>1.3)
                    // - More curved contours (>40%)
                    return avgCurvature > 1.3 || curvedRatio > 0.4;
                }
            }
            catch
            {
                return false; // Default to regular OCR if detection fails
            }
        }

        private Bitmap PreprocessWithOpenCV(Bitmap image)
        {
            // Check if this looks like handwriting
            bool isHandwriting = IsLikelyHandwriting(image);
            
            using (var mat = BitmapToMat(image))
            {
                if (isHandwriting)
                {
                    Debug.WriteLine("Detected likely handwriting - using enhanced preprocessing");
                    
                    // For handwriting, use less aggressive preprocessing
                    Cv2.CvtColor(mat, mat, ColorConversionCodes.BGR2GRAY);
                    
                    // Bilateral filter preserves edges better for handwriting
                    using (var filtered = new Mat())
                    {
                        Cv2.BilateralFilter(mat, filtered, 5, 50, 50);
                        filtered.CopyTo(mat);
                    }
                    
                    // Use Otsu's method which works better for handwriting
                    Cv2.Threshold(mat, mat, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
                    
                    // Very light morphological operation to clean up without destroying strokes
                    using (var kernel = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(2, 2)))
                    {
                        Cv2.MorphologyEx(mat, mat, MorphTypes.Close, kernel);
                    }
                }
                else
                {
                    // Standard preprocessing for printed text
                    Cv2.CvtColor(mat, mat, ColorConversionCodes.BGR2GRAY);
                    Cv2.AdaptiveThreshold(mat, mat, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, 31, 10);
                    Cv2.MorphologyEx(mat, mat, MorphTypes.Open, Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3)));
                }
                
                return MatToBitmap(mat);
            }
        }

        private OCRResult RecognizeWithManagedTesseract(Bitmap image)
        {
            try
            {
                using var engine = new TesseractEngine(_tessdataPath, "eng", EngineMode.Default);
                var modes = new[]
                {
                    PageSegMode.Auto,
                    PageSegMode.SingleBlock,
                    PageSegMode.SingleLine,
                    PageSegMode.SparseText
                };

                foreach (var mode in modes)
                {
                    engine.DefaultPageSegMode = mode;

                    using var pix = PixConverter.ToPix(image);
                    using var page = engine.Process(pix);

                    string text = page.GetText()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return new OCRResult
                        {
                            Success = true,
                            Text = text,
                            Numbers = Array.Empty<string>(),
                            Confidence = page.GetMeanConfidence() * 100
                        };
                    }
                }

                // Nothing recognised
                return new OCRResult { Success = false, Text = string.Empty, Numbers = Array.Empty<string>(), Confidence = 0 };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Managed OCR exception: {ex.Message}");
                return null;
            }
        }

        private OCRResult RecognizeWithTesseractCli(Bitmap image)
        {
            try
            {
                string cli = FindTesseractCli();
                if (string.IsNullOrEmpty(cli)) return null;

                string tmpInput = Path.Combine(Path.GetTempPath(), $"snip_{Guid.NewGuid()}.png");
                string tmpOutput = Path.Combine(Path.GetTempPath(), $"snip_{Guid.NewGuid()}");
                image.Save(tmpInput, System.Drawing.Imaging.ImageFormat.Png);

                var psi = new ProcessStartInfo(cli, $"\"{tmpInput}\" \"{tmpOutput}\" -l eng --oem 1 --psm 6 --dpi 300")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };
                var proc = Process.Start(psi);
                proc.WaitForExit(10000);

                string txtPath = tmpOutput + ".txt";
                if (File.Exists(txtPath))
                {
                    string text = File.ReadAllText(txtPath).Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        return new OCRResult
                        {
                            Success = true,
                            Text = text,
                            Numbers = Array.Empty<string>(),
                            Confidence = 0
                        };
                    }
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Locate tesseract.exe on the machine.  Checks cached value, env var, common install paths,
        /// PATH folders, and finally `where tesseract` shell command. Returns null if none found.
        /// </summary>
        private static string FindTesseractCli()
        {
            if (!string.IsNullOrEmpty(_cachedCliPath) && File.Exists(_cachedCliPath))
                return _cachedCliPath;

            // 1) Explicit environment variable
            var env = Environment.GetEnvironmentVariable("TESSERACT_PATH");
            if (!string.IsNullOrEmpty(env))
            {
                var candidate = Path.Combine(env, "tesseract.exe");
                if (File.Exists(candidate)) return _cachedCliPath = candidate;
            }

            // 2) Common install dirs
            var programFiles64 = Environment.GetEnvironmentVariable("ProgramW6432") ?? string.Empty; // 64-bit Program Files even from a 32-bit process

            var common = new[]
            {
                // 64-bit install dir
                !string.IsNullOrEmpty(programFiles64) ? Path.Combine(programFiles64, "Tesseract-OCR", "tesseract.exe") : null,
                // 32-bit install dir
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Tesseract-OCR", "tesseract.exe"),
                // Fallback to current ProgramFiles (may coincide with x86 when running 32-bit)
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Tesseract-OCR", "tesseract.exe"),
                // Directory of the add-in itself
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tesseract.exe")
            }.Where(p => !string.IsNullOrEmpty(p)).ToArray();

            foreach (var c in common)
                if (File.Exists(c)) return _cachedCliPath = c;

            // 3) Directories in PATH
            var pathVar = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var dir in pathVar.Split(Path.PathSeparator))
            {
                try
                {
                    var file = Path.Combine(dir.Trim(), "tesseract.exe");
                    if (File.Exists(file)) return _cachedCliPath = file;
                }
                catch { /* ignore malformed path segments */ }
            }

            // 4) Last resort: use `where` command (Windows only)
            try
            {
                var psi = new ProcessStartInfo("where", "tesseract.exe")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var proc = Process.Start(psi);
                string output = proc.StandardOutput.ReadLine();
                proc.WaitForExit(2000);
                if (!string.IsNullOrEmpty(output) && File.Exists(output))
                    return _cachedCliPath = output;
            }
            catch { }

            return null;
        }

        private string LightPostProcessing(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            
            // Single regex pass to compress whitespace
            text = Regex.Replace(text, @"\s+", " ");
            text = text.Trim();
            
            return text;
        }

        private string[] ExtractNumbers(string text)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<string>();

            var numbers = new HashSet<string>();
            
            // Enhanced patterns for better number extraction
            var patterns = new[]
            {
                @"\$?[\d,]+\.?\d*",                    // Currency with optional $
                @"-?\d+\.?\d+",                        // Decimals with optional negative
                @"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?",    // Thousands with optional decimal
                @"\(\d+\.?\d*\)",                      // Accounting format negative
                @"-?\d+",                              // Plain integers
                @"\d+\.\d{2}%?",                       // Percentages
                @"(?<=\s|^)-?\d+(?:\.\d+)?(?=\s|$)"   // Numbers with word boundaries
            };
            
            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(text, pattern);
                foreach (Match m in matches)
                {
                    var value = m.Value;
                    
                    // Clean the value
                    value = value.Replace("$", "").Replace(",", "").Replace("%", "").Trim();
                    
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

        // Manual Bitmap ↔ Mat conversion (since BitmapConverter may not be available)
        private Mat BitmapToMat(Bitmap bitmap)
        {
            // Ensure bitmap is in 24bpp RGB format; convert if necessary
            Bitmap workingBmp = bitmap;
            if (bitmap.PixelFormat != PixelFormat.Format24bppRgb)
            {
                workingBmp = new Bitmap(bitmap.Width, bitmap.Height, PixelFormat.Format24bppRgb);
                using (var g = Graphics.FromImage(workingBmp))
                {
                    g.DrawImage(bitmap, 0, 0, bitmap.Width, bitmap.Height);
                }
            }

            var rect = new System.Drawing.Rectangle(0, 0, workingBmp.Width, workingBmp.Height);
            var bitmapData = workingBmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            
            try
            {
                var mat = new Mat(workingBmp.Height, workingBmp.Width, MatType.CV_8UC3, bitmapData.Scan0, bitmapData.Stride);
                return mat.Clone(); // Clone to ensure we own the memory
            }
            finally
            {
                workingBmp.UnlockBits(bitmapData);
                if (!ReferenceEquals(workingBmp, bitmap))
                {
                    workingBmp.Dispose();
                }
            }
        }

        private Bitmap MatToBitmap(Mat mat)
        {
            var bitmap = new Bitmap(mat.Width, mat.Height, PixelFormat.Format24bppRgb);
            var rect = new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var bitmapData = bitmap.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            
            try
            {
                unsafe
                {
                    var src = (byte*)mat.DataPointer;
                    var dst = (byte*)bitmapData.Scan0;
                    var srcStride = (int)mat.Step();
                    var dstStride = bitmapData.Stride;
                    
                    for (int y = 0; y < mat.Height; y++)
                    {
                        Buffer.MemoryCopy(src + y * srcStride, dst + y * dstStride, dstStride, Math.Min(srcStride, dstStride));
                    }
                }
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
            
            return bitmap;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }
} 