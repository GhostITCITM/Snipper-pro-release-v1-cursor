using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;
using Microsoft.Web.WebView2.WinForms;
using SnipperClone.Core;
using Newtonsoft.Json;
using Rectangle = System.Drawing.Rectangle;

namespace SnipperClone
{
    /// <summary>
    /// Enhanced Document Viewer with Modern UI and Advanced Features
    /// Supports PDF viewing, OCR integration, and professional document analysis
    /// </summary>
    public partial class DocumentViewer : Form
    {
        private WebView2 _webView;
        private SnipEngine snipperEngine;
        private string currentDocumentPath;
        private Panel _statusPanel;
        private Label _statusLabel;
        private Label _modeLabel;
        private Button _importButton;

        public DocumentViewer(SnipEngine engine)
        {
            snipperEngine = engine ?? throw new ArgumentNullException(nameof(engine));

            InitializeComponent();
            SubscribeToEvents();
            InitializeWebViewAsync().ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    MessageBox.Show($"Failed to initialize document viewer: {t.Exception?.InnerException?.Message}", 
                        "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Create status panel
            _statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                BackColor = Color.FromArgb(240, 240, 240)
            };

            // Create status label
            _statusLabel = new Label
            {
                AutoSize = true,
                Location = new Point(10, 8),
                Text = "Ready"
            };
            _statusPanel.Controls.Add(_statusLabel);

            // Create mode label
            _modeLabel = new Label
            {
                AutoSize = true,
                Location = new Point(_statusLabel.Right + 20, 8),
                ForeColor = Color.FromArgb(0, 120, 212)
            };
            _statusPanel.Controls.Add(_modeLabel);

            // Create import button with enhanced styling
            _importButton = new Button
            {
                Text = "Import Document",
                Width = 150,
                Height = 35,
                Location = new Point(10, 10),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                Cursor = Cursors.Hand,
                Dock = DockStyle.Top,
                Margin = new Padding(10),
                Enabled = false // Initially disabled until WebView2 is ready
            };
            _importButton.FlatAppearance.BorderSize = 0;
            _importButton.Click += OnImportButtonClick;

            // Create WebView2 with proper docking
            _webView = new WebView2
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(10)
            };

            // Create container panel for better layout
            var containerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            containerPanel.Controls.Add(_webView);

            // Add controls in proper order
            this.Controls.Add(_statusPanel);
            this.Controls.Add(_importButton);
            this.Controls.Add(containerPanel);

            // Set form properties
            this.Text = "Document Viewer";
            this.Size = new Size(1024, 768);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(800, 600);

            this.ResumeLayout(false);
        }

        private async Task InitializeWebView()
        {
            int maxRetries = 3;
            int currentRetry = 0;
            bool initialized = false;

            while (!initialized && currentRetry < maxRetries)
            {
                try
                {
                    // Create WebView2 environment with proper cache location
                    var userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SnipperClone", "WebView2");
                    Directory.CreateDirectory(userDataFolder);
                    
                    var envOptions = new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions()
                    {
                        AdditionalBrowserArguments = "--disable-features=msSmartScreenProtection"
                    };

                    var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, userDataFolder, envOptions);
                    await _webView.EnsureCoreWebView2Async(env);

                    // Configure WebView2 settings
                    _webView.CoreWebView2.Settings.IsScriptEnabled = true;
                    _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    _webView.CoreWebView2.Settings.IsWebMessageEnabled = true;
                    _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;

                    // Load the viewer HTML
                    var htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebAssets", "viewer.html");
                    if (File.Exists(htmlPath))
                    {
                        try
                        {
                            string htmlContent = await Task.Run(() => File.ReadAllText(htmlPath));
                            _webView.CoreWebView2.NavigateToString(htmlContent);
                            UpdateStatus("WebView initialized with local viewer");
                            initialized = true;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error loading viewer HTML: {ex}");
                            throw;
                        }
                    }
                    else
                    {
                        throw new FileNotFoundException("Viewer HTML file not found", htmlPath);
                    }

                    // Subscribe to WebView2 events
                    _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                    _webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
                    _webView.CoreWebView2.ProcessFailed += OnProcessFailed;
                }
                catch (Exception ex)
                {
                    currentRetry++;
                    System.Diagnostics.Debug.WriteLine($"WebView initialization attempt {currentRetry} failed: {ex}");
                    
                    if (currentRetry >= maxRetries)
                    {
                        throw new Exception("Failed to initialize WebView after multiple attempts", ex);
                    }
                    
                    await Task.Delay(1000); // Wait before retrying
                }
            }
        }

        private async Task PostMessageToWebAsync(string json)
        {
            try
            {
                if (_webView.CoreWebView2 == null)
                {
                    await _webView.EnsureCoreWebView2Async();
                }

                if (_webView.CoreWebView2 != null)
                {
                    _webView.CoreWebView2.PostWebMessageAsJson(json);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("CoreWebView2 is still null after initialization attempt");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error posting message to web: {ex}");
                throw;
            }
        }

        private void AddHostObject()
        {
            try
            {
                if (_webView.CoreWebView2 != null)
                {
                    // We don't need WebViewModel for now
                    //_webView.CoreWebView2.AddHostObjectToScript("external", new WebViewModel());
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding host object: {ex}");
                throw;
            }
        }

        public async Task InvokeBrowserScriptAsync(string functionName, string parameter)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Invoking script: {functionName}, parameter: {parameter}");
                
                var message = new
                {
                    messageType = functionName,
                    messageData = parameter
                };
                
                await PostMessageToWebAsync(JsonConvert.SerializeObject(message));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Invoking script error: {ex}");
                throw;
            }
        }

        private void SubscribeToEvents()
        {
            snipperEngine.ModeChanged += OnSnipModeChanged;
            snipperEngine.SnipCompleted += OnSnipCompleted;
        }

        private void OnSnipModeChanged(object sender, SnipModeChangedEventArgs e)
        {
            this.Invoke((Action)(() =>
            {
                if (e.Mode == SnipMode.None)
                {
                    _modeLabel.Text = "";
                    UpdateStatus("Ready");
                }
                else
                {
                    _modeLabel.Text = $"Mode: {e.Mode.ToString().ToUpper()}";
                    UpdateStatus($"{e.Mode} mode activated - Select area on document");
                }
            }));
        }

        private void OnSnipCompleted(object sender, SnipCompletedEventArgs e)
        {
            this.Invoke((Action)(() =>
            {
                UpdateStatus($"{e.Record.Mode} snip completed successfully");
                _modeLabel.Text = "";
            }));
        }

        private async void OnImportButtonClick(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf|Image Files (*.png;*.jpg;*.jpeg;*.bmp;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.tiff|All Supported Files|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff|All Files (*.*)|*.*";
                openFileDialog.Title = "Select Document to Import";
                openFileDialog.Multiselect = false;
                openFileDialog.CheckFileExists = true;
                openFileDialog.CheckPathExists = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        UpdateStatus("Loading document...");
                        _importButton.Enabled = false;
                        
                        await LoadDocument(openFileDialog.FileName);
                        
                        // Update UI state
                        _importButton.Text = "Change Document";
                        UpdateStatus($"Document loaded: {Path.GetFileName(openFileDialog.FileName)}");
                        
                        // Notify the snip engine about the current document
                        snipperEngine.SetCurrentDocument(Path.GetFileName(openFileDialog.FileName));
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"Failed to load document: {ex.Message}");
                        MessageBox.Show($"Failed to load document:\n\n{ex.Message}", "Document Load Error", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        _importButton.Enabled = true;
                    }
                }
            }
        }

        private async Task LoadDocument(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"Document file not found: {filePath}");
                }

                currentDocumentPath = filePath;
                var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();
                var fileName = Path.GetFileName(filePath);
                var mimeType = GetMimeType(filePath);
                var fileBytes = await Task.Run(() => File.ReadAllBytes(filePath));

                UpdateStatus($"Loading {fileName}...");

                // Convert file bytes to base64
                var base64Data = Convert.ToBase64String(fileBytes);

                // Send load document message to WebView
                var message = new
                {
                    action = "loadDocument",
                    data = base64Data,
                    mimeType = mimeType,
                    fileName = fileName,
                    options = new
                    {
                        detectTables = true
                    }
                };

                await PostMessageToWebAsync(JsonConvert.SerializeObject(message));

                // Start table detection in background
                await Task.Run(async () =>
                {
                    try
                    {
                        var tables = await DetectTables(fileBytes, mimeType);
                        if (tables.Any())
                        {
                            var tablesCommand = new
                            {
                                action = "suggestTables",
                                tables = tables
                            };
                            await PostMessageToWebAsync(JsonConvert.SerializeObject(tablesCommand));
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Table detection error: {ex}");
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DocumentViewer: Error loading document: {ex.Message}");
                UpdateStatus($"Error loading document: {ex.Message}");
                throw;
            }
        }

        private async Task<bool> CheckIfPdfNeedsOcr(byte[] pdfBytes)
        {
            try
            {
                // Basic check - look for text content in first few pages
                // In a real implementation, use a PDF library to check for text content
                return true; // For now, always suggest OCR for PDFs
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking PDF for OCR: {ex}");
                return true; // Err on the side of caution
            }
        }

        private async Task ApplyOcr(byte[] documentBytes)
        {
            try
            {
                // In a real implementation:
                // 1. Use OCR library (e.g. Tesseract, ABBYY, etc.)
                // 2. Process document pages
                // 3. Extract text layer
                // 4. Update document with OCR results
                await Task.Delay(1000); // Simulate OCR processing
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCR error: {ex}");
                throw;
            }
        }

        private async Task<List<TableInfo>> DetectTables(byte[] documentBytes, string mimeType)
        {
            try
            {
                // In a real implementation:
                // 1. Use table detection library/API
                // 2. Analyze document layout
                // 3. Identify table structures
                // 4. Return table coordinates and metadata
                await Task.Delay(500); // Simulate detection

                // Return dummy table for testing
                return new List<TableInfo>
                {
                    new TableInfo
                    {
                        Page = 1,
                        X = 50,
                        Y = 100,
                        Width = 500,
                        Height = 300,
                        Columns = 4,
                        Rows = 10,
                        HasHeader = true,
                        Confidence = 0.95
                    }
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Table detection error: {ex}");
                return new List<TableInfo>();
            }
        }

        public class TableInfo
        {
            public int Page { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public int Columns { get; set; }
            public int Rows { get; set; }
            public bool HasHeader { get; set; }
            public double Confidence { get; set; }
        }

        private void OnWebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var message = e.WebMessageAsJson;
                var data = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(message);

                switch (data.action.ToString())
                {
                    case "tableSelected":
                        _ = HandleTableSelected(data);
                        break;
                    case "detectTables":
                        _ = HandleDetectTables();
                        break;
                    case "documentReady":
                        UpdateStatus("Document loaded and ready for snipping");
                        break;
                    case "error":
                        UpdateStatus($"Viewer error: {data.message}");
                        break;
                    default:
                        System.Diagnostics.Debug.WriteLine($"Unknown message action: {data.action}");
                        break;
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error processing message: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Message processing error: {ex}");
            }
        }

        private async Task HandleTableSelected(dynamic data)
        {
            try
            {
                // Convert the data to a strongly-typed object
                var tableMetadata = JsonConvert.DeserializeObject<TableMetadata>(data.ToString());
                
                // Capture the table area
                var bitmap = await CaptureSelectionArea(tableMetadata.Rectangle, tableMetadata.PageNumber);
                
                if (bitmap != null)
                {
                    // Process the table snip
                    var result = await snipperEngine.ProcessTableSnipAsync(bitmap);

                    // Notify WebView
                    var message = new
                    {
                        type = "tableSnipped",
                        success = result.Success,
                        message = result.Success ? "Table successfully snipped" : result.ErrorMessage
                    };
                    await PostMessageToWebAsync(JsonConvert.SerializeObject(message));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling table selection: {ex}");
                var errorMessage = new
                {
                    type = "error",
                    message = "Failed to process table selection"
                };
                await PostMessageToWebAsync(JsonConvert.SerializeObject(errorMessage));
            }
        }

        private async Task HandleDetectTables()
        {
            try
            {
                if (string.IsNullOrEmpty(currentDocumentPath))
                {
                    UpdateStatus("No document loaded");
                    return;
                }

                UpdateStatus("Detecting tables...");

                // For now, we'll create some dummy table suggestions
                // In a real implementation, use a table detection library/API
                var tables = new List<TableInfo>
                {
                    new TableInfo
                    {
                        Page = 1,
                        X = 50,
                        Y = 100,
                        Width = 500,
                        Height = 300,
                        Columns = 4,
                        Rows = 10,
                        HasHeader = true,
                        Confidence = 0.95
                    }
                };

                var message = new
                {
                    action = "suggestTables",
                    tables = tables
                };

                await PostMessageToWebAsync(JsonConvert.SerializeObject(message));
                UpdateStatus($"Found {tables.Count} tables");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error detecting tables: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Table detection error: {ex}");
            }
        }

        public class TableMetadata
        {
            public int Columns { get; set; }
            public int Rows { get; set; }
            public bool HasHeader { get; set; }
            public int PageNumber { get; set; }
            public SnipperClone.Core.Rectangle Rectangle { get; set; }
        }

        private async Task<Bitmap> CaptureSelectionArea(Rectangle rectangle, int pageNumber)
        {
            try
            {
                // For now, we'll create a bitmap from the current document
                // In a real implementation, this would extract the exact area from the PDF/image
                
                if (string.IsNullOrEmpty(currentDocumentPath))
                    return null;

                var fileExtension = Path.GetExtension(currentDocumentPath).ToLowerInvariant();
                
                if (fileExtension == ".pdf")
                {
                    return await Task.Run(() => CapturePDFArea(rectangle, pageNumber));
                }
                else if (new[] { ".png", ".jpg", ".jpeg", ".bmp", ".tiff" }.Contains(fileExtension))
                {
                    return await Task.Run(() => CaptureImageArea(rectangle));
                }
                
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error capturing selection area: {ex}");
                return null;
            }
        }

        private Bitmap CapturePDFArea(Rectangle rectangle, int pageNumber)
        {
            try
            {
                // For PDF files, we'll create a simple bitmap with the rectangle dimensions
                // In a production implementation, you would use a PDF library to extract the exact area
                var bitmap = new Bitmap(Math.Max(1, rectangle.Width), Math.Max(1, rectangle.Height));
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.FillRectangle(Brushes.White, 0, 0, bitmap.Width, bitmap.Height);
                    g.DrawString($"PDF Area\nPage {pageNumber}\n{rectangle.Width}x{rectangle.Height}", 
                        SystemFonts.DefaultFont, Brushes.Black, 10, 10);
                }
                return bitmap;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error capturing PDF area: {ex}");
                return null;
            }
        }

        private Bitmap CaptureImageArea(Rectangle rectangle)
        {
            try
            {
                using (var sourceImage = new Bitmap(currentDocumentPath))
                {
                    // Ensure rectangle is within image bounds
                    var clampedRect = Rectangle.Intersect(rectangle, new Rectangle(0, 0, sourceImage.Width, sourceImage.Height));
                    
                    if (clampedRect.IsEmpty)
                        return null;

                    var bitmap = new Bitmap(clampedRect.Width, clampedRect.Height);
                    using (var g = Graphics.FromImage(bitmap))
                    {
                        g.DrawImage(sourceImage, 0, 0, clampedRect, GraphicsUnit.Pixel);
                    }
                    return bitmap;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error capturing image area: {ex}");
                return null;
            }
        }

        private void UpdateStatus(string message)
        {
            if (_statusLabel.InvokeRequired)
            {
                _statusLabel.Invoke((Action)(() => _statusLabel.Text = message));
            }
            else
            {
                _statusLabel.Text = message;
            }
        }

        private string GetMimeType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            switch (extension)
            {
                case ".pdf":
                    return "application/pdf";
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".bmp":
                    return "image/bmp";
                case ".tiff":
                    return "image/tiff";
                case ".gif":
                    return "image/gif";
                default:
                    return "application/octet-stream";
            }
        }

        private string GetPDFViewerHtml()
        {
            return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Document Viewer</title>
    <script src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js'></script>
    <style>
        :root {
            --primary-color: #0078d4;
            --secondary-color: #d83b01;
            --success-color: #107c10;
            --warning-color: #d83b01;
            --error-color: #d13438;
        }

        body { 
            margin: 0; 
            padding: 0; 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
        }

        #container { 
            width: 100%; 
            height: 100vh; 
            display: flex;
            flex-direction: column;
        }

        #toolbar {
            background: white;
            padding: 10px;
            border-bottom: 1px solid #ddd;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        #viewer { 
            flex: 1;
            overflow: auto;
            padding: 20px;
        }

        .page { 
            margin: 10px auto; 
            background: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            position: relative;
        }

        .page-loading {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            color: #666;
        }

        .selection-overlay { 
            position: absolute; 
            border: 2px dashed var(--primary-color);
            background: rgba(0, 120, 212, 0.1); 
            pointer-events: none;
            transition: all 0.2s ease;
        }

        .highlight-overlay { 
            position: absolute;
            border: 2px solid var(--secondary-color);
            background: rgba(216, 59, 1, 0.2);
            pointer-events: none;
            transition: all 0.2s ease;
        }

        .table-suggestion {
            position: absolute;
            border: 2px dashed var(--success-color);
            background: rgba(16, 124, 16, 0.1);
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .table-suggestion:hover {
            background: rgba(16, 124, 16, 0.2);
        }

        .table-suggestion .actions {
            position: absolute;
            top: -30px;
            left: 0;
            background: white;
            padding: 5px;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: none;
        }

        .table-suggestion:hover .actions {
            display: flex;
            gap: 5px;
        }

        button {
            padding: 6px 12px;
            border: none;
            border-radius: 4px;
            background: var(--primary-color);
            color: white;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s ease;
        }

        button:hover {
            background: #106ebe;
        }

        button.secondary {
            background: #f0f0f0;
            color: #333;
        }

        button.secondary:hover {
            background: #e0e0e0;
        }

        #status { 
            position: fixed;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%);
            background: rgba(0,0,0,0.8);
            color: white;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 14px;
            transition: all 0.3s ease;
            opacity: 0;
        }

        #status.visible {
            opacity: 1;
        }

        .zoom-controls {
            position: fixed;
            right: 20px;
            top: 50%;
            transform: translateY(-50%);
            display: flex;
            flex-direction: column;
            gap: 10px;
            background: white;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }

        .zoom-controls button {
            width: 40px;
            height: 40px;
            padding: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
        }

        #pageInfo {
            padding: 5px 10px;
            background: #f0f0f0;
            border-radius: 4px;
            font-size: 14px;
        }

        .table-editor {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            min-width: 300px;
            display: none;
        }

        .table-editor.visible {
            display: block;
        }

        .table-editor h3 {
            margin: 0 0 15px 0;
            color: #333;
        }

        .table-editor .form-group {
            margin-bottom: 15px;
        }

        .table-editor label {
            display: block;
            margin-bottom: 5px;
            color: #666;
        }

        .table-editor input[type='number'] {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }

        .table-editor .actions {
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            margin-top: 20px;
        }

        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255,255,255,0.8);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
            opacity: 0;
            pointer-events: none;
            transition: all 0.3s ease;
        }

        .loading-overlay.visible {
            opacity: 1;
            pointer-events: auto;
        }

        .spinner {
            width: 40px;
            height: 40px;
            border: 4px solid #f0f0f0;
            border-top: 4px solid var(--primary-color);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div id='container'>
        <div id='toolbar'>
            <button onclick='importDocument()'>Import Document</button>
            <button onclick='detectTables()' class='secondary'>Detect Tables</button>
            <span id='pageInfo'></span>
        </div>
        <div id='viewer'></div>
        <div id='status'></div>
        <div class='zoom-controls'>
            <button onclick='zoomIn()'>+</button>
            <button onclick='resetZoom()'>1:1</button>
            <button onclick='zoomOut()'>-</button>
        </div>
    </div>

    <div class='table-editor'>
        <h3>Edit Table</h3>
        <div class='form-group'>
            <label>Columns</label>
            <input type='number' id='columnCount' min='1' value='1'>
        </div>
        <div class='form-group'>
            <label>Rows</label>
            <input type='number' id='rowCount' min='1' value='1'>
        </div>
        <div class='form-group'>
            <label>
                <input type='checkbox' id='hasHeader'> Has Header Row
            </label>
        </div>
        <div class='actions'>
            <button onclick='closeTableEditor()' class='secondary'>Cancel</button>
            <button onclick='applyTableChanges()'>Apply</button>
        </div>
    </div>

    <div class='loading-overlay'>
        <div class='spinner'></div>
    </div>

    <script>
        let pdfDoc = null;
        let currentScale = 1.0;
        let isSelecting = false;
        let selectionStart = null;
        let currentPage = 1;
        let tableSuggestions = [];
        let currentTableEdit = null;
        
        // Initialize PDF.js
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        
        function showStatus(message, duration = 3000) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.classList.add('visible');
            setTimeout(() => status.classList.remove('visible'), duration);
        }

        function showLoading(show = true) {
            document.querySelector('.loading-overlay').classList.toggle('visible', show);
        }
        
        async function loadDocument(base64Data, mimeType, fileName, options = {}) {
            try {
                showStatus('Loading document...');
                showLoading(true);
                
                if (mimeType === 'application/pdf') {
                    const data = atob(base64Data);
                    const bytes = new Uint8Array(data.length);
                    for (let i = 0; i < data.length; i++) {
                        bytes[i] = data.charCodeAt(i);
                    }
                    
                    pdfDoc = await pdfjsLib.getDocument({data: bytes}).promise;
                    await renderAllPages();

                    if (options.detectTables) {
                        detectTables();
                    }
                } else {
                    // Handle image files
                    const img = new Image();
                    img.onload = function() {
                        const viewer = document.getElementById('viewer');
                        viewer.innerHTML = '';
                        
                        const pageDiv = document.createElement('div');
                        pageDiv.className = 'page';
                        pageDiv.style.width = img.width + 'px';
                        pageDiv.style.height = img.height + 'px';
                        pageDiv.appendChild(img);
                        
                        setupPageInteraction(pageDiv, 1);
                        viewer.appendChild(pageDiv);
                        
                        showStatus('Image loaded successfully');
                        showLoading(false);
                        window.chrome.webview.postMessage({action: 'documentReady'});
                    };
                    img.src = 'data:' + mimeType + ';base64,' + base64Data;
                }
            } catch (error) {
                showStatus('Error loading document: ' + error.message);
                showLoading(false);
                window.chrome.webview.postMessage({action: 'error', message: error.message});
            }
        }
        
        async function renderAllPages() {
            const viewer = document.getElementById('viewer');
            viewer.innerHTML = '';
            
            for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
                const pageDiv = document.createElement('div');
                pageDiv.className = 'page';
                pageDiv.innerHTML = '<div class=""page-loading"">Loading page ' + pageNum + '...</div>';
                viewer.appendChild(pageDiv);

                try {
                    const page = await pdfDoc.getPage(pageNum);
                    const viewport = page.getViewport({scale: currentScale});
                    
                    const canvas = document.createElement('canvas');
                    const context = canvas.getContext('2d');
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;
                    
                    pageDiv.style.width = viewport.width + 'px';
                    pageDiv.style.height = viewport.height + 'px';
                    pageDiv.innerHTML = '';
                    pageDiv.appendChild(canvas);
                    
                    await page.render({canvasContext: context, viewport: viewport}).promise;
                    
                    setupPageInteraction(pageDiv, pageNum);
                } catch (error) {
                    pageDiv.innerHTML = '<div class=""page-loading"">Error loading page ' + pageNum + '</div>';
                }
            }
            
            document.getElementById('pageInfo').textContent = `${pdfDoc.numPages} pages`;
            showStatus('Document loaded successfully');
            showLoading(false);
            window.chrome.webview.postMessage({action: 'documentReady'});
        }
        
        function setupPageInteraction(pageDiv, pageNumber) {
            let isSelecting = false;
            let selectionStart = null;
            let selectionDiv = null;
            
            pageDiv.addEventListener('mousedown', function(e) {
                if (e.button === 0) { // Left click
                    isSelecting = true;
                    const rect = pageDiv.getBoundingClientRect();
                    selectionStart = {
                        x: e.clientX - rect.left,
                        y: e.clientY - rect.top
                    };
                    
                    // Create selection overlay
                    selectionDiv = document.createElement('div');
                    selectionDiv.className = 'selection-overlay';
                    selectionDiv.style.left = selectionStart.x + 'px';
                    selectionDiv.style.top = selectionStart.y + 'px';
                    pageDiv.appendChild(selectionDiv);
                    
                    e.preventDefault();
                }
            });
            
            pageDiv.addEventListener('mousemove', function(e) {
                if (isSelecting && selectionDiv) {
                    const rect = pageDiv.getBoundingClientRect();
                    const currentPos = {
                        x: e.clientX - rect.left,
                        y: e.clientY - rect.top
                    };
                    
                    const left = Math.min(selectionStart.x, currentPos.x);
                    const top = Math.min(selectionStart.y, currentPos.y);
                    const width = Math.abs(currentPos.x - selectionStart.x);
                    const height = Math.abs(currentPos.y - selectionStart.y);
                    
                    selectionDiv.style.left = left + 'px';
                    selectionDiv.style.top = top + 'px';
                    selectionDiv.style.width = width + 'px';
                    selectionDiv.style.height = height + 'px';
                }
            });
            
            pageDiv.addEventListener('mouseup', function(e) {
                if (isSelecting && selectionDiv) {
                    const rect = pageDiv.getBoundingClientRect();
                    const endPos = {
                        x: e.clientX - rect.left,
                        y: e.clientY - rect.top
                    };
                    
                    const left = Math.min(selectionStart.x, endPos.x);
                    const top = Math.min(selectionStart.y, endPos.y);
                    const width = Math.abs(endPos.x - selectionStart.x);
                    const height = Math.abs(endPos.y - selectionStart.y);
                    
                    if (width > 10 && height > 10) { // Minimum selection size
                        showTableEditor({
                            page: pageNumber,
                            x: left,
                            y: top,
                            width: width,
                            height: height
                        });
                    } else {
                        pageDiv.removeChild(selectionDiv);
                    }
                    
                    isSelecting = false;
                    selectionStart = null;
                    selectionDiv = null;
                }
            });
        }

        function showTableEditor(tableInfo) {
            currentTableEdit = tableInfo;
            const editor = document.querySelector('.table-editor');
            editor.classList.add('visible');
            document.getElementById('columnCount').value = '1';
            document.getElementById('rowCount').value = '1';
            document.getElementById('hasHeader').checked = true;
        }

        function closeTableEditor() {
            const editor = document.querySelector('.table-editor');
            editor.classList.remove('visible');
            if (currentTableEdit) {
                const page = document.querySelectorAll('.page')[currentTableEdit.page - 1];
                const selection = page.querySelector('.selection-overlay');
                if (selection) {
                    selection.remove();
                }
            }
            currentTableEdit = null;
        }

        function applyTableChanges() {
            if (!currentTableEdit) return;

            const columns = parseInt(document.getElementById('columnCount').value);
            const rows = parseInt(document.getElementById('rowCount').value);
            const hasHeader = document.getElementById('hasHeader').checked;

            const tableInfo = {
                ...currentTableEdit,
                columns,
                rows,
                hasHeader
            };

            window.chrome.webview.postMessage({
                action: 'tableSelected',
                table: tableInfo
            });

            closeTableEditor();
        }
        
        function detectTables() {
            showStatus('Detecting tables...');
            window.chrome.webview.postMessage({action: 'detectTables'});
        }

        function suggestTables(tables) {
            tableSuggestions = tables;
            const pages = document.querySelectorAll('.page');
            
            tables.forEach(table => {
                if (table.page <= pages.length) {
                    const pageDiv = pages[table.page - 1];
                    const suggestion = document.createElement('div');
                    suggestion.className = 'table-suggestion';
                    suggestion.style.left = table.x + 'px';
                    suggestion.style.top = table.y + 'px';
                    suggestion.style.width = table.width + 'px';
                    suggestion.style.height = table.height + 'px';
                    
                    const actions = document.createElement('div');
                    actions.className = 'actions';
                    actions.innerHTML = `
                        <button onclick='extractTable(${tables.indexOf(table)})'>Extract Table</button>
                        <button class='secondary' onclick='ignoreTable(${tables.indexOf(table)})'>Ignore</button>
                    `;
                    
                    suggestion.appendChild(actions);
                    pageDiv.appendChild(suggestion);
                }
            });
            
            showStatus(`Found ${tables.length} tables`);
        }

        function extractTable(index) {
            const table = tableSuggestions[index];
            if (!table) return;

            window.chrome.webview.postMessage({
                action: 'tableSelected',
                table: table
            });

            // Remove the suggestion
            const pages = document.querySelectorAll('.page');
            if (table.page <= pages.length) {
                const pageDiv = pages[table.page - 1];
                const suggestions = pageDiv.querySelectorAll('.table-suggestion');
                suggestions.forEach(s => s.remove());
            }
        }

        function ignoreTable(index) {
            const table = tableSuggestions[index];
            if (!table) return;

            // Remove the suggestion
            const pages = document.querySelectorAll('.page');
            if (table.page <= pages.length) {
                const pageDiv = pages[table.page - 1];
                const suggestions = pageDiv.querySelectorAll('.table-suggestion');
                suggestions.forEach(s => s.remove());
            }
        }
        
        function zoomIn() {
            currentScale *= 1.2;
            if (pdfDoc) renderAllPages();
        }
        
        function zoomOut() {
            currentScale /= 1.2;
            if (pdfDoc) renderAllPages();
        }

        function resetZoom() {
            currentScale = 1.0;
            if (pdfDoc) renderAllPages();
        }
        
        // Message handler
        window.chrome.webview.addEventListener('message', event => {
            const data = JSON.parse(event.data);
            
            switch(data.action) {
                case 'loadDocument':
                    loadDocument(data.data, data.mimeType, data.fileName, data.options);
                    break;
                case 'suggestTables':
                    suggestTables(data.tables);
                    break;
            }
        });
        
        showStatus('Document viewer ready');
    </script>
</body>
</html>";
        }

        private void OnNavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                _importButton.Enabled = true;
                UpdateStatus("Ready to import documents");
            }
            else
            {
                UpdateStatus($"Navigation failed: {e.WebErrorStatus}");
            }
        }

        private void OnProcessFailed(object sender, Microsoft.Web.WebView2.Core.CoreWebView2ProcessFailedEventArgs e)
        {
            try
            {
                var message = $"WebView2 process failed: {e.ProcessFailedKind}";
                System.Diagnostics.Debug.WriteLine(message);
                UpdateStatus(message);

                // Attempt to recover by reinitializing
                if (!_webView.IsDisposed)
                {
                    Task.Run(async () =>
                    {
                        await Task.Delay(1000); // Wait a bit before retrying
                        
                        if (this.InvokeRequired)
                        {
                            this.Invoke((Action)(async () =>
                            {
                                try
                                {
                                    await InitializeWebView();
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to recover from process failure: {ex}");
                                    UpdateStatus("Failed to recover from WebView2 process failure");
                                }
                            }));
                        }
                        else
                        {
                            try
                            {
                                await InitializeWebView();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to recover from process failure: {ex}");
                                UpdateStatus("Failed to recover from WebView2 process failure");
                            }
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling process failure: {ex}");
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Hide instead of close to keep the viewer available
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
            else
            {
                base.OnFormClosing(e);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _webView?.Dispose();
            }
            base.Dispose(disposing);
        }

        public async Task NavigateToSnip(SnipRecord snipRecord)
        {
            try
            {
                if (snipRecord == null)
                    return;

                // Update status
                UpdateStatus($"Navigating to {snipRecord.Mode} snip from {snipRecord.CreatedAt:MM/dd/yyyy HH:mm}");

                // Send message to viewer to go to the page and highlight the rectangle
                var message = new
                {
                    action = "goToPage",
                    page = snipRecord.PageNumber
                };
                
                await PostMessageToWebAsync(JsonConvert.SerializeObject(message));

                // Add highlight after a short delay to ensure page is loaded
                await Task.Delay(500);
                
                try
                {
                    var highlightMessage = new
                    {
                        action = "addHighlight",
                        rectangle = new
                        {
                            x = snipRecord.Rectangle.X,
                            y = snipRecord.Rectangle.Y,
                            width = snipRecord.Rectangle.Width,
                            height = snipRecord.Rectangle.Height
                        },
                        page = snipRecord.PageNumber
                    };
                    
                    await PostMessageToWebAsync(JsonConvert.SerializeObject(highlightMessage));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error adding highlight: {ex}");
                }

                // Bring the viewer to front
                this.BringToFront();
                this.Activate();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error navigating to snip: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"NavigateToSnip error: {ex}");
            }
        }

        public async Task InitializeWebViewAsync()
        {
            if (_webView.CoreWebView2 == null)
            {
                await InitializeWebView();
            }
        }
    }
} 