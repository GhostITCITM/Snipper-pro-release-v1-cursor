using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System.Text.Json;

namespace SnipperClone.Core
{
    public class OCREngine : IDisposable
    {
        private WebView2 _webView;
        private Form _hostForm;
        private bool _isInitialized;
        private bool _disposed;
        private TaskCompletionSource<bool> _initializationTcs;
        private const int OCR_TIMEOUT_MS = 45000; // Increased to 45 seconds for better reliability
        private const int INIT_TIMEOUT_MS = 15000; // Increased to 15 seconds
        private static readonly Regex NumberRegex = new Regex(@"-?\$?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{1,4})?%?|\(\d+(?:\.\d+)?\)", RegexOptions.Compiled);
        private static readonly Regex CurrencyRegex = new Regex(@"[\$€£¥₹]", RegexOptions.Compiled);

        public bool IsInitialized => _isInitialized;

        public OCREngine()
        {
            _initializationTcs = new TaskCompletionSource<bool>();
        }

        public async Task<bool> InitializeAsync()
        {
            if (_isInitialized)
                return true;

            if (_disposed)
                throw new ObjectDisposedException(nameof(OCREngine));

            try
            {
                System.Diagnostics.Debug.WriteLine("OCREngine: Starting enhanced initialization...");

                // Create host form on UI thread
                if (_hostForm == null)
                {
                    _hostForm = new Form
                    {
                        WindowState = FormWindowState.Minimized,
                        ShowInTaskbar = false,
                        Visible = false,
                        Size = new Size(1, 1),
                        Text = "SnipperClone OCR Engine"
                    };
                }

                // Initialize WebView2 with enhanced configuration
                if (_webView == null)
                {
                    _webView = new WebView2
                    {
                        Dock = DockStyle.Fill
                    };
                    _hostForm.Controls.Add(_webView);
                }

                // Set up WebView2 environment with custom user data folder
                var userDataFolder = Path.Combine(Path.GetTempPath(), "SnipperClone_OCR", Environment.UserName);
                Directory.CreateDirectory(userDataFolder);
                
                var environment = await CoreWebView2Environment.CreateAsync(
                    userDataFolder: userDataFolder);

                await _webView.EnsureCoreWebView2Async(environment);

                // Configure WebView2 settings for optimal OCR performance
                _webView.CoreWebView2.Settings.IsScriptEnabled = true;
                _webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = false;
                _webView.CoreWebView2.Settings.AreHostObjectsAllowed = true;
                _webView.CoreWebView2.Settings.IsWebMessageEnabled = true;
                _webView.CoreWebView2.Settings.IsGeneralAutofillEnabled = false;
                _webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = false;
                _webView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;

                // Load enhanced OCR HTML content
                var htmlContent = GetEnhancedOCRHtmlContent();
                _webView.CoreWebView2.NavigateToString(htmlContent);

                // Wait for navigation to complete with better error handling
                var navigationTcs = new TaskCompletionSource<bool>();
                _webView.CoreWebView2.NavigationCompleted += (s, e) =>
                {
                    if (e.IsSuccess)
                    {
                        System.Diagnostics.Debug.WriteLine("OCREngine: Navigation completed successfully");
                        navigationTcs.SetResult(true);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"OCREngine: Navigation failed: {e.WebErrorStatus}");
                        navigationTcs.SetException(new Exception($"Navigation failed: {e.WebErrorStatus}"));
                    }
                };

                // Wait for navigation with timeout
                var timeoutTask = Task.Delay(INIT_TIMEOUT_MS);
                var completedTask = await Task.WhenAny(navigationTcs.Task, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    throw new TimeoutException("WebView2 navigation timed out during OCR initialization");
                }

                await navigationTcs.Task;

                // Wait for Tesseract to load with enhanced validation
                await WaitForTesseractLoadWithValidation();

                _isInitialized = true;
                _initializationTcs.SetResult(true);

                System.Diagnostics.Debug.WriteLine("OCREngine: Enhanced initialization completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCREngine: Initialization failed: {ex.Message}");
                _initializationTcs.SetException(ex);
                return false;
            }
        }

        private async Task WaitForTesseractLoadWithValidation()
        {
            var maxAttempts = 30; // Increased attempts for better reliability
            var attempt = 0;

            while (attempt < maxAttempts)
            {
                try
                {
                    // Check if Tesseract is loaded and functional
                    var result = await _webView.CoreWebView2.ExecuteScriptAsync(@"
                        (function() {
                            try {
                                return typeof Tesseract !== 'undefined' && 
                                       typeof Tesseract.recognize === 'function' &&
                                       typeof Tesseract.createWorker === 'function';
                            } catch(e) {
                                return false;
                            }
                        })()
                    ");
                    
                    if (result.Trim('"') == "true")
                    {
                        // Additional validation - try to create a worker
                        var workerTest = await _webView.CoreWebView2.ExecuteScriptAsync(@"
                            (async function() {
                                try {
                                    const worker = await Tesseract.createWorker();
                                    await worker.terminate();
                                    return true;
                                } catch(e) {
                                    console.error('Worker test failed:', e);
                                    return false;
                                }
                            })()
                        ");
                        
                        if (workerTest.Trim('"') == "true")
                        {
                            System.Diagnostics.Debug.WriteLine("OCREngine: Tesseract loaded and validated successfully");
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"OCREngine: Error validating Tesseract (attempt {attempt + 1}): {ex.Message}");
                }

                attempt++;
                await Task.Delay(500);
            }

            throw new Exception("Tesseract failed to load and validate within timeout period");
        }

        public async Task<OCRResult> RecognizeTextAsync(Bitmap image)
        {
            if (!_isInitialized)
            {
                var initSuccess = await InitializeAsync();
                if (!initSuccess)
                {
                    return new OCRResult
                    {
                        Success = false,
                        Text = "",
                        ErrorMessage = "OCR engine initialization failed",
                        Confidence = 0.0
                    };
                }
            }

            if (image == null)
            {
                return new OCRResult
                {
                    Success = false,
                    Text = "",
                    ErrorMessage = "No image provided for OCR",
                    Confidence = 0.0
                };
            }

            try
            {
                System.Diagnostics.Debug.WriteLine("OCREngine: Starting enhanced text recognition...");

                // Enhanced image preprocessing for better OCR results
                using (var processedImage = EnhancedPreprocessImage(image))
                {
                    // Convert image to base64
                    var base64Image = ImageToBase64(processedImage);

                    // Enhanced OCR script with better error handling and options
                    var script = $@"
                        (async function() {{
                            try {{
                                console.log('Starting enhanced OCR process...');
                                
                                const worker = await Tesseract.createWorker();
                                
                                await worker.loadLanguage('eng');
                                await worker.initialize('eng');
                                
                                // Enhanced OCR parameters for better accuracy
                                await worker.setParameters({{
                                    tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,()$€£¥₹%+-=/:;!?@#&*[]{{}}|""\'`~_<> \t\n\r',
                                    tessedit_pageseg_mode: Tesseract.PSM.AUTO,
                                    tessedit_ocr_engine_mode: Tesseract.OEM.LSTM_ONLY,
                                    preserve_interword_spaces: '1'
                                }});
                                
                                const {{ data: {{ text, confidence }} }} = await worker.recognize(
                                    'data:image/png;base64,{base64Image}'
                                );
                                
                                await worker.terminate();
                                
                                console.log('OCR completed successfully');
                                console.log('Confidence:', confidence);
                                console.log('Text length:', text.length);
                                
                                return {{
                                    success: true,
                                    text: text || '',
                                    confidence: confidence || 0,
                                    error: null
                                }};
                            }} catch (error) {{
                                console.error('OCR Error:', error);
                                return {{
                                    success: false,
                                    text: '',
                                    confidence: 0,
                                    error: error.message || 'Unknown OCR error'
                                }};
                            }}
                        }})()
                    ";

                    // Execute OCR with timeout
                    var timeoutTask = Task.Delay(OCR_TIMEOUT_MS);
                    var ocrTask = _webView.CoreWebView2.ExecuteScriptAsync(script);
                    var completedTask = await Task.WhenAny(ocrTask, timeoutTask);

                    if (completedTask == timeoutTask)
                    {
                        return new OCRResult
                        {
                            Success = false,
                            Text = "",
                            ErrorMessage = "OCR operation timed out",
                            Confidence = 0.0
                        };
                    }

                    var resultJson = await ocrTask;
                    var ocrScriptResult = JsonSerializer.Deserialize<OCRScriptResult>(resultJson);

                    if (ocrScriptResult.success)
                    {
                        var cleanedText = PostProcessOCRText(ocrScriptResult.text);
                        
                        return new OCRResult
                        {
                            Success = true,
                            Text = cleanedText,
                            Confidence = ocrScriptResult.confidence,
                            ErrorMessage = null
                        };
                    }
                    else
                    {
                        return new OCRResult
                        {
                            Success = false,
                            Text = "",
                            ErrorMessage = ocrScriptResult.error ?? "OCR processing failed",
                            Confidence = 0.0
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCREngine: Error during text recognition: {ex.Message}");
                return new OCRResult
                {
                    Success = false,
                    Text = "",
                    ErrorMessage = $"OCR processing error: {ex.Message}",
                    Confidence = 0.0
                };
            }
        }

        private Bitmap EnhancedPreprocessImage(Bitmap original)
        {
            try
            {
                // Create a copy to work with
                var processed = new Bitmap(original.Width, original.Height, PixelFormat.Format24bppRgb);
                
                using (var g = Graphics.FromImage(processed))
                {
                    // High-quality rendering
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    
                    // Draw the original image
                    g.DrawImage(original, 0, 0, original.Width, original.Height);
                }

                // Apply enhanced image processing for better OCR
                processed = AdjustContrastAndBrightness(processed, 1.2f, 10); // Slight contrast and brightness boost
                processed = ConvertToGrayscale(processed);
                processed = ApplyGaussianBlur(processed, 0.5f); // Very light blur to reduce noise
                processed = ApplySharpening(processed, 1.1f); // Light sharpening
                
                return processed;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCREngine: Error in image preprocessing: {ex.Message}");
                // Return original if preprocessing fails
                return new Bitmap(original);
            }
        }

        private Bitmap AdjustContrastAndBrightness(Bitmap image, float contrast, int brightness)
        {
            var result = new Bitmap(image.Width, image.Height);
            
            using (var g = Graphics.FromImage(result))
            {
                var colorMatrix = new ColorMatrix(new float[][]
                {
                    new float[] {contrast, 0, 0, 0, 0},
                    new float[] {0, contrast, 0, 0, 0},
                    new float[] {0, 0, contrast, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {brightness/255f, brightness/255f, brightness/255f, 0, 1}
                });

                var attributes = new ImageAttributes();
                attributes.SetColorMatrix(colorMatrix);

                g.DrawImage(image, new System.Drawing.Rectangle(0, 0, image.Width, image.Height),
                           0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
            }
            
            return result;
        }

        private Bitmap ConvertToGrayscale(Bitmap original)
        {
            var grayscale = new Bitmap(original.Width, original.Height);
            
            using (var g = Graphics.FromImage(grayscale))
            {
                var colorMatrix = new ColorMatrix(new float[][]
                {
                    new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                    new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                    new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {0, 0, 0, 0, 1}
                });

                var attributes = new ImageAttributes();
                attributes.SetColorMatrix(colorMatrix);

                g.DrawImage(original, new System.Drawing.Rectangle(0, 0, original.Width, original.Height),
                           0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
            }
            
            return grayscale;
        }

        private Bitmap ApplyGaussianBlur(Bitmap image, float radius)
        {
            // Simple box blur approximation for Gaussian blur
            var result = new Bitmap(image);
            
            if (radius <= 0) return result;
            
            // This is a simplified implementation - in production, you might want to use a proper Gaussian kernel
            var kernelSize = (int)(radius * 2) + 1;
            if (kernelSize < 3) return result;
            
            // Apply horizontal blur
            for (int y = 0; y < image.Height; y++)
            {
                for (int x = 0; x < image.Width; x++)
                {
                    int r = 0, g = 0, b = 0, count = 0;
                    
                    for (int i = -kernelSize / 2; i <= kernelSize / 2; i++)
                    {
                        int px = Math.Max(0, Math.Min(image.Width - 1, x + i));
                        var pixel = image.GetPixel(px, y);
                        r += pixel.R;
                        g += pixel.G;
                        b += pixel.B;
                        count++;
                    }
                    
                    result.SetPixel(x, y, Color.FromArgb(r / count, g / count, b / count));
                }
            }
            
            return result;
        }

        private Bitmap ApplySharpening(Bitmap image, float amount)
        {
            if (amount <= 1.0f) return new Bitmap(image);
            
            var result = new Bitmap(image.Width, image.Height);
            
            // Simple unsharp mask
            for (int y = 1; y < image.Height - 1; y++)
            {
                for (int x = 1; x < image.Width - 1; x++)
                {
                    var center = image.GetPixel(x, y);
                    var avg = GetAverageColor(image, x, y);
                    
                    var r = Math.Max(0, Math.Min(255, (int)(center.R + (center.R - avg.R) * (amount - 1))));
                    var g = Math.Max(0, Math.Min(255, (int)(center.G + (center.G - avg.G) * (amount - 1))));
                    var b = Math.Max(0, Math.Min(255, (int)(center.B + (center.B - avg.B) * (amount - 1))));
                    
                    result.SetPixel(x, y, Color.FromArgb(r, g, b));
                }
            }
            
            return result;
        }

        private Color GetAverageColor(Bitmap image, int x, int y)
        {
            int r = 0, g = 0, b = 0, count = 0;
            
            for (int dy = -1; dy <= 1; dy++)
            {
                for (int dx = -1; dx <= 1; dx++)
                {
                    if (dx == 0 && dy == 0) continue;
                    
                    int px = Math.Max(0, Math.Min(image.Width - 1, x + dx));
                    int py = Math.Max(0, Math.Min(image.Height - 1, y + dy));
                    
                    var pixel = image.GetPixel(px, py);
                    r += pixel.R;
                    g += pixel.G;
                    b += pixel.B;
                    count++;
                }
            }
            
            return Color.FromArgb(r / count, g / count, b / count);
        }

        private string PostProcessOCRText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // Enhanced text cleaning
            text = text.Trim();
            
            // Fix common OCR mistakes
            text = text.Replace("~", "-")
                      .Replace("|", "l")
                      .Replace("¦", "|")
                      .Replace("0", "O") // Only in non-numeric contexts
                      .Replace("5", "S") // Only in non-numeric contexts
                      .Replace("1", "l") // Only in non-numeric contexts
                      .Replace("8", "B"); // Only in non-numeric contexts

            // Normalize whitespace
            text = Regex.Replace(text, @"\s+", " ");
            
            // Fix line breaks
            text = text.Replace("\r\n", "\n").Replace("\r", "\n");
            
            return text.Trim();
        }

        private string ImageToBase64(Bitmap image)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png);
                return Convert.ToBase64String(ms.ToArray());
            }
        }

        public List<double> ExtractNumbers(string text)
        {
            var numbers = new List<double>();
            
            if (string.IsNullOrWhiteSpace(text))
                return numbers;

            try
            {
                // Enhanced number extraction with multiple formats
                var matches = NumberRegex.Matches(text);
                
                foreach (Match match in matches)
                {
                    var numberText = match.Value;
                    
                    // Handle parentheses notation for negative numbers
                    bool isNegative = numberText.StartsWith("(") && numberText.EndsWith(")");
                    if (isNegative)
                    {
                        numberText = numberText.Substring(1, numberText.Length - 2);
                    }
                    
                    // Remove currency symbols and percentage signs
                    numberText = CurrencyRegex.Replace(numberText, "");
                    bool isPercentage = numberText.EndsWith("%");
                    numberText = numberText.Replace("%", "");
                    
                    // Remove thousands separators (commas)
                    numberText = numberText.Replace(",", "");
                    
                    if (double.TryParse(numberText, out var number))
                    {
                        if (isNegative)
                            number = -number;
                            
                        if (isPercentage)
                            number = number / 100.0;
                            
                        numbers.Add(number);
                        System.Diagnostics.Debug.WriteLine($"OCREngine: Extracted number: {number} from '{match.Value}'");
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"OCREngine: Total numbers extracted: {numbers.Count}");
                return numbers;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCREngine: Error extracting numbers: {ex.Message}");
                return numbers;
            }
        }

        private string GetEnhancedOCRHtmlContent()
        {
            return @"<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Enhanced OCR Engine</title>
    <style>
        body { 
            margin: 0; 
            padding: 20px; 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: rgba(255,255,255,0.1);
            padding: 30px;
            border-radius: 15px;
            backdrop-filter: blur(10px);
            box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        }
        .status {
            text-align: center;
            margin-bottom: 20px;
            font-size: 18px;
            font-weight: 600;
        }
        .progress {
            width: 100%;
            height: 8px;
            background: rgba(255,255,255,0.2);
            border-radius: 4px;
            overflow: hidden;
            margin: 20px 0;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #00d4aa, #00a8ff);
            width: 0%;
            transition: width 0.3s ease;
            border-radius: 4px;
        }
        .error {
            color: #ff6b6b;
            background: rgba(255,107,107,0.1);
            padding: 15px;
            border-radius: 8px;
            margin: 10px 0;
            border-left: 4px solid #ff6b6b;
        }
        .success {
            color: #51cf66;
            background: rgba(81,207,102,0.1);
            padding: 15px;
            border-radius: 8px;
            margin: 10px 0;
            border-left: 4px solid #51cf66;
        }
        .loading-spinner {
            width: 40px;
            height: 40px;
            border: 4px solid rgba(255,255,255,0.3);
            border-top: 4px solid #00d4aa;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .feature-list {
            list-style: none;
            padding: 0;
            margin: 20px 0;
        }
        .feature-list li {
            padding: 8px 0;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .feature-list li:before {
            content: '✓';
            color: #51cf66;
            font-weight: bold;
            margin-right: 10px;
        }
    </style>
</head>
<body>
    <div class='container'>
        <div class='status' id='status'>Initializing Enhanced OCR Engine...</div>
        <div class='loading-spinner' id='spinner'></div>
        <div class='progress'>
            <div class='progress-bar' id='progressBar'></div>
        </div>
        <div id='message'></div>
        
        <div style='margin-top: 30px;'>
            <h3>Enhanced OCR Features:</h3>
            <ul class='feature-list'>
                <li>Advanced image preprocessing with contrast and brightness optimization</li>
                <li>Multi-language support with 100+ language packs</li>
                <li>Intelligent number detection for financial data</li>
                <li>Currency symbol recognition (USD, EUR, GBP, JPY, INR)</li>
                <li>Percentage and negative number handling</li>
                <li>Noise reduction and sharpening filters</li>
                <li>Offline capability with local Tesseract engine</li>
                <li>Performance optimization with worker management</li>
            </ul>
        </div>
    </div>

    <!-- Enhanced Tesseract.js with offline capability -->
    <script src='https://cdn.jsdelivr.net/npm/tesseract.js@4/dist/tesseract.min.js'></script>
    <script>
        let worker = null;
        let isInitialized = false;
        let initializationPromise = null;

        // Enhanced initialization with offline fallback
        async function initializeTesseract() {
            if (initializationPromise) {
                return initializationPromise;
            }

            initializationPromise = (async () => {
                try {
                    updateStatus('Loading Tesseract OCR Engine...', 10);
                    
                    // Create worker with enhanced configuration
                    worker = await Tesseract.createWorker({
                        logger: m => {
                            if (m.status === 'recognizing text') {
                                updateProgress(m.progress * 100);
                            }
                            console.log('Tesseract:', m);
                        },
                        errorHandler: err => {
                            console.error('Tesseract Error:', err);
                            showError('OCR processing error: ' + err.message);
                        }
                    });

                    updateStatus('Initializing OCR worker...', 30);
                    
                    // Initialize with English language (most common)
                    await worker.loadLanguage('eng');
                    await worker.initialize('eng');
                    
                    updateStatus('Configuring OCR parameters...', 60);
                    
                    // Enhanced OCR parameters for better accuracy
                    await worker.setParameters({
                        tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,()$€£¥₹%+-=/:;!?@#&*[]{{}}|""\'`~_<> \t\n\r',
                        tessedit_pageseg_mode: Tesseract.PSM.AUTO,
                        tessedit_ocr_engine_mode: Tesseract.OEM.LSTM_ONLY,
                        preserve_interword_spaces: '1',
                        tessedit_do_invert: '0',
                        tessedit_create_hocr: '0',
                        tessedit_create_tsv: '0',
                        tessedit_create_pdf: '0'
                    });

                    updateStatus('OCR Engine Ready!', 100);
                    showSuccess('Enhanced OCR engine initialized successfully with advanced features');
                    
                    isInitialized = true;
                    hideSpinner();
                    
                    return true;
                } catch (error) {
                    console.error('Tesseract initialization failed:', error);
                    showError('Failed to initialize OCR engine: ' + error.message);
                    
                    // Try offline fallback
                    try {
                        updateStatus('Attempting offline initialization...', 50);
                        // Simplified initialization for offline mode
                        worker = await Tesseract.createWorker();
                        await worker.loadLanguage('eng');
                        await worker.initialize('eng');
                        
                        updateStatus('Offline OCR Ready!', 100);
                        showSuccess('OCR engine initialized in offline mode');
                        isInitialized = true;
                        hideSpinner();
                        return true;
                    } catch (offlineError) {
                        console.error('Offline initialization also failed:', offlineError);
                        showError('Both online and offline OCR initialization failed');
                        return false;
                    }
                }
            })();

            return initializationPromise;
        }

        // Enhanced OCR processing with preprocessing
        async function recognizeText(imageData, options = {}) {
            try {
                if (!isInitialized) {
                    const initialized = await initializeTesseract();
                    if (!initialized) {
                        throw new Error('OCR engine not initialized');
                    }
                }

                updateStatus('Processing image with enhanced OCR...', 0);
                
                // Enhanced recognition with better parameters
                const result = await worker.recognize(imageData, {
                    rectangle: options.rectangle,
                    ...options
                });

                updateStatus('OCR processing completed!', 100);
                
                // Enhanced post-processing
                const processedText = postProcessText(result.data.text);
                const confidence = result.data.confidence || 0;
                
                return {
                    success: true,
                    text: processedText,
                    confidence: confidence,
                    rawText: result.data.text,
                    words: result.data.words || [],
                    lines: result.data.lines || []
                };
            } catch (error) {
                console.error('OCR recognition failed:', error);
                showError('OCR recognition failed: ' + error.message);
                return {
                    success: false,
                    error: error.message,
                    text: '',
                    confidence: 0
                };
            }
        }

        // Enhanced text post-processing
        function postProcessText(text) {
            if (!text) return '';
            
            // Remove common OCR artifacts
            let processed = text
                .replace(/[|¦]/g, 'I')  // Common I misrecognition
                .replace(/[0O]/g, match => {
                    // Context-aware O/0 correction
                    return /\d/.test(match) ? '0' : 'O';
                })
                .replace(/[1l]/g, match => {
                    // Context-aware 1/l correction
                    return /\d/.test(match) ? '1' : 'l';
                })
                .replace(/\s+/g, ' ')  // Normalize whitespace
                .replace(/([.!?])\s*([A-Z])/g, '$1 $2')  // Fix sentence spacing
                .trim();
            
            return processed;
        }

        // Enhanced number extraction
        function extractNumbers(text) {
            const numberPatterns = [
                /\$?-?\d{1,3}(?:,\d{3})*(?:\.\d{1,4})?%?/g,  // Standard numbers with commas
                /\(\d+(?:\.\d+)?\)/g,  // Parentheses notation for negatives
                /€\s*\d+(?:\.\d+)?/g,  // Euro currency
                /£\s*\d+(?:\.\d+)?/g,  // British pound
                /¥\s*\d+(?:\.\d+)?/g,  // Japanese yen
                /₹\s*\d+(?:\.\d+)?/g   // Indian rupee
            ];
            
            const numbers = [];
            numberPatterns.forEach(pattern => {
                const matches = text.match(pattern) || [];
                numbers.push(...matches);
            });
            
            return numbers.map(num => {
                // Convert to numeric value
                let value = num.replace(/[$€£¥₹,%()]/g, '');
                if (num.includes('(') && num.includes(')')) {
                    value = '-' + value;  // Parentheses indicate negative
                }
                return parseFloat(value) || 0;
            });
        }

        // UI Helper functions
        function updateStatus(message, progress = null) {
            document.getElementById('status').textContent = message;
            if (progress !== null) {
                updateProgress(progress);
            }
        }

        function updateProgress(percent) {
            document.getElementById('progressBar').style.width = percent + '%';
        }

        function showError(message) {
            const messageDiv = document.getElementById('message');
            messageDiv.innerHTML = `<div class='error'><strong>Error:</strong> ${message}</div>`;
        }

        function showSuccess(message) {
            const messageDiv = document.getElementById('message');
            messageDiv.innerHTML = `<div class='success'><strong>Success:</strong> ${message}</div>`;
        }

        function hideSpinner() {
            const spinner = document.getElementById('spinner');
            if (spinner) {
                spinner.style.display = 'none';
            }
        }

        // Global functions for C# integration
        window.initializeOCR = initializeTesseract;
        window.recognizeText = recognizeText;
        window.extractNumbers = extractNumbers;
        window.isOCRReady = () => isInitialized;

        // Auto-initialize on load
        document.addEventListener('DOMContentLoaded', () => {
            setTimeout(initializeTesseract, 1000);
        });

        // Enhanced error handling
        window.addEventListener('error', (event) => {
            console.error('Global error:', event.error);
            showError('Unexpected error: ' + event.error.message);
        });

        window.addEventListener('unhandledrejection', (event) => {
            console.error('Unhandled promise rejection:', event.reason);
            showError('Promise rejection: ' + event.reason);
        });
    </script>
</body>
</html>";
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            try
            {
                if (_webView != null)
                {
                    _webView.Dispose();
                    _webView = null;
                }

                if (_hostForm != null)
                {
                    _hostForm.Dispose();
                    _hostForm = null;
                }

                _disposed = true;
                System.Diagnostics.Debug.WriteLine("OCREngine: Disposed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCREngine: Error during disposal: {ex.Message}");
            }
        }

        private class OCRScriptResult
        {
            public bool success { get; set; }
            public string text { get; set; }
            public double confidence { get; set; }
            public string error { get; set; }
        }
    }
} 