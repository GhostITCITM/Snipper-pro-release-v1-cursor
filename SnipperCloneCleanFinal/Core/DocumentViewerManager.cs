using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PdfiumViewer;
using SnipperCloneCleanFinal.UI;
using SnipperCloneCleanFinal.Infrastructure;
using System.Linq;
using System.Threading.Tasks;
using Tesseract;

namespace SnipperCloneCleanFinal.Core
{
    public static class DataSnipperExtensions
    {
        public static System.Drawing.Color GetSnipColor(this SnipMode snipMode)
        {
            return snipMode switch
            {
                SnipMode.Text => System.Drawing.Color.Blue,
                SnipMode.Sum => System.Drawing.Color.Purple,
                SnipMode.Table => System.Drawing.Color.Purple,
                SnipMode.Validation => System.Drawing.Color.Green,
                SnipMode.Exception => System.Drawing.Color.Red,
                _ => System.Drawing.Color.Gray
            };
        }
    }

    public static class DocumentViewerManager
    {
        private static DocumentViewerPane _currentViewer;
        private static Dictionary<string, DocumentViewerPane> _viewers = new Dictionary<string, DocumentViewerPane>();
        
        public static DocumentViewerPane GetOrCreateViewer()
        {
            if (_currentViewer == null || _currentViewer.IsDisposed)
            {
                _currentViewer = new DocumentViewerPane();
                Logger.Info("Created new document viewer pane");
            }
            return _currentViewer;
        }
        
        public static void CloseAllViewers()
        {
            foreach (var viewer in _viewers.Values)
            {
                if (!viewer.IsDisposed)
                    viewer.Close();
            }
            _viewers.Clear();
            
            if (_currentViewer != null && !_currentViewer.IsDisposed)
                _currentViewer.Close();
                
            _currentViewer = null;
            Logger.Info("Closed all document viewers");
        }
    }
}

namespace SnipperCloneCleanFinal.UI
{
    public partial class DocumentViewerPane : Form
    {
        private Panel _documentPanel;
        private Panel _toolbarPanel;
        private Label _statusLabel;
        private Button _prevPageBtn;
        private Button _nextPageBtn;
        private Label _pageLabel;
        private PictureBox _documentDisplay;
        private Button _zoomInBtn;
        private Button _zoomOutBtn;
        private Button _fitWidthBtn;
        private Button _clearOverlaysBtn;
        private Button _deleteOverlayBtn;
        private Label _overlayCountLabel;
        
        // Search functionality components
        private TextBox _searchTextBox;
        private Button _searchBtn;
        private Button _nextResultBtn;
        private Button _prevResultBtn;
        private Label _searchResultsLabel;
        private Button _closeSearchBtn;
        
        private string _currentDocumentPath;
        private int _currentPage = 1;
        private int _totalPages = 1;
        private float _zoomLevel = 1.0f;
        private List<Bitmap> _documentPages = new List<Bitmap>();
        
        // Persistent overlay storage - stores overlays per document
        private static Dictionary<string, List<SnipOverlay>> _documentOverlays = new Dictionary<string, List<SnipOverlay>>();
        private List<SnipOverlay> _currentOverlays = new List<SnipOverlay>();
        private SnipOverlay _selectedOverlay = null;
        
        // Search functionality data
        private static Dictionary<string, List<DocumentText>> _documentTexts = new Dictionary<string, List<DocumentText>>();
        private List<SearchResult> _searchResults = new List<SearchResult>();
        private int _currentSearchResultIndex = -1;
        private bool _isSearchMode = false;
        
        private bool _isSnipMode = false;
        private Point _snipStartPoint;
        private Point _snipEndPoint;
        private bool _isDrawing = false;
        private SnipperCloneCleanFinal.Core.SnipMode _currentSnipMode = SnipperCloneCleanFinal.Core.SnipMode.Text;
        
        public event EventHandler<SnipAreaSelectedEventArgs> SnipAreaSelected;
        
        public DocumentViewerPane()
        {
            InitializeComponent();
            SetupViewer();
        }
        
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Save overlays before disposing
                SaveCurrentOverlays();
                
                // Dispose of document pages
                foreach (var page in _documentPages)
                {
                    page?.Dispose();
                }
                _documentPages.Clear();
            }
            base.Dispose(disposing);
        }
        
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Save overlays when form is closing
            SaveCurrentOverlays();
            base.OnFormClosing(e);
        }
        
        private void InitializeComponent()
        {
            // Required for Windows Forms designer support
            this.SuspendLayout();
            this.ResumeLayout(false);
        }
        
        private void SetupViewer()
        {
            this.Text = "Snipper Pro - Document Viewer";
            this.Size = new Size(400, 600);
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            
            // Position on right side of screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(screen.Width - this.Width, 0);
            
            CreateControls();
            SetupEventHandlers();
            
            Logger.Info("Document viewer pane initialized");
        }
        
        private void CreateControls()
        {
            // Toolbar - increased height to accommodate search and controls
            _toolbarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = Color.LightGray
            };
            
            // First row - navigation and zoom
            _prevPageBtn = new Button { Text = "◀", Size = new Size(30, 30), Location = new Point(5, 5) };
            _nextPageBtn = new Button { Text = "▶", Size = new Size(30, 30), Location = new Point(40, 5) };
            _pageLabel = new Label { Text = "Page 1 of 1", Location = new Point(80, 12), Size = new Size(80, 20) };
            _zoomInBtn = new Button { Text = "+", Size = new Size(30, 30), Location = new Point(170, 5) };
            _zoomOutBtn = new Button { Text = "-", Size = new Size(30, 30), Location = new Point(205, 5) };
            _fitWidthBtn = new Button { Text = "Fit", Size = new Size(40, 30), Location = new Point(240, 5) };
            
            // Second row - overlay management
            _overlayCountLabel = new Label { Text = "Overlays: 0", Location = new Point(5, 42), Size = new Size(80, 20) };
            _deleteOverlayBtn = new Button { Text = "Delete", Size = new Size(50, 25), Location = new Point(90, 40), Enabled = false };
            _clearOverlaysBtn = new Button { Text = "Clear All", Size = new Size(60, 25), Location = new Point(145, 40) };
            
            // Third row - search functionality
            var searchLabel = new Label { Text = "Search:", Location = new Point(5, 72), Size = new Size(45, 20) };
            _searchTextBox = new TextBox { Location = new Point(55, 70), Size = new Size(120, 23) };
            _searchBtn = new Button { Text = "🔍", Size = new Size(25, 23), Location = new Point(180, 70) };
            _prevResultBtn = new Button { Text = "◀", Size = new Size(25, 23), Location = new Point(210, 70), Enabled = false };
            _nextResultBtn = new Button { Text = "▶", Size = new Size(25, 23), Location = new Point(240, 70), Enabled = false };
            _searchResultsLabel = new Label { Text = "", Location = new Point(270, 72), Size = new Size(80, 20) };
            _closeSearchBtn = new Button { Text = "✕", Size = new Size(20, 23), Location = new Point(355, 70), Visible = false };
            
            _toolbarPanel.Controls.AddRange(new Control[] { 
                _prevPageBtn, _nextPageBtn, _pageLabel, _zoomInBtn, _zoomOutBtn, _fitWidthBtn,
                _overlayCountLabel, _deleteOverlayBtn, _clearOverlaysBtn,
                searchLabel, _searchTextBox, _searchBtn, _prevResultBtn, _nextResultBtn, _searchResultsLabel, _closeSearchBtn
            });
            
            // Document display
            _documentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                AutoScroll = true
            };
            
            _documentDisplay = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.Zoom,
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            
            _documentPanel.Controls.Add(_documentDisplay);
            
            // Status bar
            _statusLabel = new Label
            {
                Text = "Ready - Open a document to start snipping",
                Dock = DockStyle.Bottom,
                Height = 25,
                BackColor = Color.LightGray,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(5)
            };
            
            this.Controls.Add(_documentPanel);
            this.Controls.Add(_toolbarPanel);
            this.Controls.Add(_statusLabel);
        }
        
        private void SetupEventHandlers()
        {
            _prevPageBtn.Click += (s, e) => NavigatePage(-1);
            _nextPageBtn.Click += (s, e) => NavigatePage(1);
            _zoomInBtn.Click += (s, e) => Zoom(1.2f);
            _zoomOutBtn.Click += (s, e) => Zoom(0.8f);
            _fitWidthBtn.Click += (s, e) => FitToWidth();
            _deleteOverlayBtn.Click += OnDeleteOverlay;
            _clearOverlaysBtn.Click += OnClearAllOverlays;
            
            // Search functionality event handlers
            _searchBtn.Click += OnSearch;
            _searchTextBox.KeyDown += OnSearchTextKeyDown;
            _nextResultBtn.Click += (s, e) => NavigateSearchResult(1);
            _prevResultBtn.Click += (s, e) => NavigateSearchResult(-1);
            _closeSearchBtn.Click += OnCloseSearch;
            
            _documentDisplay.MouseDown += OnDocumentMouseDown;
            _documentDisplay.MouseMove += OnDocumentMouseMove;
            _documentDisplay.MouseUp += OnDocumentMouseUp;
            _documentDisplay.Paint += OnDocumentPaint;
        }
        
        public bool LoadDocument(string filePath)
        {
            try
            {
                // Save current overlays before switching documents
                if (!string.IsNullOrEmpty(_currentDocumentPath))
                {
                    SaveCurrentOverlays();
                }

                _currentDocumentPath = filePath;

                // Dispose previously loaded pages to avoid memory leaks
                foreach (var page in _documentPages)
                {
                    page?.Dispose();
                }
                _documentPages.Clear();

                // Load document pages
                if (IsImageFile(filePath))
                {
                    var image = new Bitmap(filePath);
                    _documentPages.Add(image);
                    _totalPages = 1;
                }
                else if (filePath.ToLower().EndsWith(".pdf"))
                {
                    _documentPages = RenderPdfToBitmaps(filePath);
                    _totalPages = _documentPages.Count;
                }
                else
                {
                    Logger.Error($"Unsupported file format: {filePath}");
                    return false;
                }

                _currentPage = 1;
                
                // Load overlays for this document
                LoadOverlaysForCurrentDocument();
                
                // Extract text for search functionality (async to avoid blocking UI)
                _ = Task.Run(() => ExtractDocumentTextAsync(filePath));
                
                UpdateDisplay();
                UpdatePageLabel();
                UpdateOverlayCount();

                Logger.Info($"Document loaded: {filePath} ({_totalPages} pages, {_currentOverlays.Count} overlays)");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to load document {filePath}: {ex.Message}", ex);
                return false;
            }
        }
        
        public void NavigateToPage(int pageNumber)
        {
            if (pageNumber >= 1 && pageNumber <= _totalPages)
            {
                _currentPage = pageNumber;
                UpdateDisplay();
                UpdatePageLabel();
            }
        }
        
        public void HighlightRegion(Rectangle bounds, Color color)
        {
            var overlay = new SnipOverlay
            {
                Bounds = new SnipperCloneCleanFinal.Core.Rectangle(bounds.X, bounds.Y, bounds.Width, bounds.Height).ToDrawingRectangle(),
                Color = color,
                Page = _currentPage
            };
            
            _currentOverlays.Add(overlay);
            _documentDisplay.Invalidate(); // Trigger repaint
        }
        
        public void SetSnipMode(SnipperCloneCleanFinal.Core.SnipMode snipMode, bool enabled)
        {
            _isSnipMode = enabled;
            _currentSnipMode = snipMode;
            
            if (enabled)
            {
                _statusLabel.Text = $"{snipMode} Snip Mode - Draw rectangle on document";
                _documentDisplay.Cursor = Cursors.Cross;
            }
            else
            {
                _statusLabel.Text = "Snip mode disabled";
                _documentDisplay.Cursor = Cursors.Default;
            }
        }
        
        private bool IsImageFile(string filePath)
        {
            var ext = System.IO.Path.GetExtension(filePath).ToLower();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".tiff";
        }
        
        private void NavigatePage(int direction)
        {
            var newPage = _currentPage + direction;
            if (newPage >= 1 && newPage <= _totalPages)
            {
                _currentPage = newPage;
                UpdateDisplay();
                UpdatePageLabel();
                UpdateOverlayCount(); // Update overlay count when changing pages
            }
        }
        
        private void Zoom(float factor)
        {
            _zoomLevel *= factor;
            UpdateDisplay();
        }
        
        private void FitToWidth()
        {
            if (_documentPages.Count > 0 && _currentPage <= _documentPages.Count)
            {
                var image = _documentPages[_currentPage - 1];
                _zoomLevel = (float)_documentPanel.Width / image.Width;
                UpdateDisplay();
            }
        }

        private List<Bitmap> RenderPdfToBitmaps(string pdfPath)
        {
            var images = new List<Bitmap>();
            try
            {
                using (var document = PdfDocument.Load(pdfPath))
                {
                    var dpiX = 150;
                    var dpiY = 150;
                    for (int i = 0; i < document.PageCount; i++)
                    {
                        var size = document.PageSizes[i];
                        var width = (int)(size.Width * (dpiX / 72.0));
                        var height = (int)(size.Height * (dpiY / 72.0));
                        var rendered = document.Render(i, width, height, dpiX, dpiY, PdfRenderFlags.Annotations);
                        images.Add(new Bitmap(rendered));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to render PDF {pdfPath}: {ex.Message}", ex);
            }
            return images;
        }
        
        private void UpdateDisplay()
        {
            if (_documentPages.Count > 0 && _currentPage <= _documentPages.Count)
            {
                var image = _documentPages[_currentPage - 1];
                var scaledSize = new Size((int)(image.Width * _zoomLevel), (int)(image.Height * _zoomLevel));
                _documentDisplay.Size = scaledSize;
                _documentDisplay.Image = image;
            }
        }
        
        private void UpdatePageLabel()
        {
            _pageLabel.Text = $"Page {_currentPage} of {_totalPages}";
        }
        
        private void OnDocumentMouseDown(object sender, MouseEventArgs e)
        {
            if (_isSnipMode && e.Button == MouseButtons.Left)
            {
                _isDrawing = true;
                _snipStartPoint = e.Location;
                _snipEndPoint = e.Location;
            }
            else if (e.Button == MouseButtons.Left)
            {
                // Check if clicking on an existing overlay for selection
                _selectedOverlay = GetOverlayAtPoint(e.Location);
                _deleteOverlayBtn.Enabled = _selectedOverlay != null;
                _documentDisplay.Invalidate();
            }
        }
        
        private void OnDocumentMouseMove(object sender, MouseEventArgs e)
        {
            if (_isDrawing)
            {
                _snipEndPoint = e.Location;
                _documentDisplay.Invalidate(); // Trigger repaint to show selection rectangle
            }
        }
        
        private void OnDocumentMouseUp(object sender, MouseEventArgs e)
        {
            if (_isDrawing && e.Button == MouseButtons.Left)
            {
                _isDrawing = false;
                
                var rect = GetSelectionRectangle();
                if (rect.Width > 5 && rect.Height > 5) // Minimum selection size
                {
                    // Create persistent overlay immediately
                    var overlay = new SnipOverlay
                    {
                        Id = Guid.NewGuid(),
                        Bounds = rect,
                        Color = SnipperCloneCleanFinal.Core.DataSnipperExtensions.GetSnipColor(_currentSnipMode),
                        Page = _currentPage,
                        SnipMode = _currentSnipMode,
                        DocumentPath = _currentDocumentPath,
                        CreatedAt = DateTime.Now
                    };
                    
                    _currentOverlays.Add(overlay);
                    UpdateOverlayCount();
                    
                    var snipRect = new Rectangle(rect.X, rect.Y, rect.Width, rect.Height);
                    SnipAreaSelected?.Invoke(this, new SnipAreaSelectedEventArgs
                    {
                        SnipMode = _currentSnipMode,
                        Bounds = snipRect,
                        DocumentPath = _currentDocumentPath,
                        PageNumber = _currentPage,
                        SelectedImage = CaptureSelectedArea(rect)
                    });
                }
                
                _documentDisplay.Invalidate();
            }
        }
        
        private void OnDocumentPaint(object sender, PaintEventArgs e)
        {
            // Draw search results highlighting
            if (_isSearchMode && _searchResults.Count > 0)
            {
                foreach (var result in _searchResults)
                {
                    if (result.DocumentPath == _currentDocumentPath && result.PageNumber == _currentPage)
                    {
                        var isCurrentResult = _currentSearchResultIndex >= 0 && 
                                            _searchResults[_currentSearchResultIndex] == result;
                        
                        var highlightColor = isCurrentResult ? Color.Yellow : Color.LightYellow;
                        var borderColor = isCurrentResult ? Color.Orange : Color.Gold;
                        
                        using (var brush = new SolidBrush(Color.FromArgb(80, highlightColor)))
                        using (var pen = new Pen(borderColor, isCurrentResult ? 2 : 1))
                        {
                            e.Graphics.FillRectangle(brush, result.Word.Bounds);
                            e.Graphics.DrawRectangle(pen, result.Word.Bounds);
                        }
                    }
                }
            }
            
            // Draw existing overlays
            foreach (var overlay in _currentOverlays)
            {
                if (overlay.Page == _currentPage)
                {
                    var isSelected = overlay == _selectedOverlay;
                    var penWidth = isSelected ? 3 : 2;
                    var color = isSelected ? Color.Orange : overlay.Color;
                    
                    using (var pen = new Pen(color, penWidth))
                    {
                        e.Graphics.DrawRectangle(pen, overlay.Bounds);
                    }
                    
                    // Draw semi-transparent fill for selected overlay
                    if (isSelected)
                    {
                        using (var brush = new SolidBrush(Color.FromArgb(30, color)))
                        {
                            e.Graphics.FillRectangle(brush, overlay.Bounds);
                        }
                    }
                    
                    // Draw overlay info for selected overlay
                    if (isSelected)
                    {
                        var info = $"{overlay.SnipMode} ({overlay.CreatedAt:HH:mm})";
                        using (var font = new Font("Arial", 8))
                        using (var brush = new SolidBrush(Color.White))
                        using (var bgBrush = new SolidBrush(Color.FromArgb(180, Color.Black)))
                        {
                            var textSize = e.Graphics.MeasureString(info, font);
                            var textRect = new RectangleF(overlay.Bounds.X, overlay.Bounds.Y - textSize.Height - 2, textSize.Width + 4, textSize.Height + 2);
                            e.Graphics.FillRectangle(bgBrush, textRect);
                            e.Graphics.DrawString(info, font, brush, textRect.X + 2, textRect.Y + 1);
                        }
                    }
                }
            }
            
            // Draw current selection
            if (_isDrawing)
            {
                var rect = GetSelectionRectangle();
                var color = SnipperCloneCleanFinal.Core.DataSnipperExtensions.GetSnipColor(_currentSnipMode);
                using (var pen = new Pen(color, 2))
                {
                    e.Graphics.DrawRectangle(pen, rect);
                }
                
                // Draw semi-transparent fill
                using (var brush = new SolidBrush(Color.FromArgb(50, color)))
                {
                    e.Graphics.FillRectangle(brush, rect);
                }
            }
        }
        
        private SnipOverlay GetOverlayAtPoint(Point point)
        {
            // Find overlay at the clicked point (in reverse order to get topmost)
            for (int i = _currentOverlays.Count - 1; i >= 0; i--)
            {
                var overlay = _currentOverlays[i];
                if (overlay.Page == _currentPage && overlay.Bounds.Contains(point))
                {
                    return overlay;
                }
            }
            return null;
        }

        private System.Drawing.Rectangle GetSelectionRectangle()
        {
            return new System.Drawing.Rectangle(
                Math.Min(_snipStartPoint.X, _snipEndPoint.X),
                Math.Min(_snipStartPoint.Y, _snipEndPoint.Y),
                Math.Abs(_snipEndPoint.X - _snipStartPoint.X),
                Math.Abs(_snipEndPoint.Y - _snipStartPoint.Y)
            );
        }
        
        private Bitmap CaptureSelectedArea(System.Drawing.Rectangle rect)
        {
            if (_documentPages.Count > 0 && _currentPage <= _documentPages.Count)
            {
                var sourceImage = _documentPages[_currentPage - 1];
                
                // Scale rectangle back to original image coordinates
                var scaleX = (float)sourceImage.Width / _documentDisplay.Width;
                var scaleY = (float)sourceImage.Height / _documentDisplay.Height;
                
                var scaledRect = new System.Drawing.Rectangle(
                    (int)(rect.X * scaleX),
                    (int)(rect.Y * scaleY),
                    (int)(rect.Width * scaleX),
                    (int)(rect.Height * scaleY)
                );
                
                // Ensure bounds are within image
                scaledRect.Intersect(new System.Drawing.Rectangle(0, 0, sourceImage.Width, sourceImage.Height));
                
                if (scaledRect.Width > 0 && scaledRect.Height > 0)
                {
                    return sourceImage.Clone(scaledRect, sourceImage.PixelFormat);
                }
            }
            
            return null;
        }
        
        private void SaveCurrentOverlays()
        {
            if (!string.IsNullOrEmpty(_currentDocumentPath))
            {
                if (!_documentOverlays.ContainsKey(_currentDocumentPath))
                {
                    _documentOverlays[_currentDocumentPath] = new List<SnipOverlay>();
                }
                else
                {
                    _documentOverlays[_currentDocumentPath].Clear();
                }
                _documentOverlays[_currentDocumentPath].AddRange(_currentOverlays);
            }
        }
        
        private void LoadOverlaysForCurrentDocument()
        {
            _currentOverlays.Clear();
            if (!string.IsNullOrEmpty(_currentDocumentPath) && _documentOverlays.ContainsKey(_currentDocumentPath))
            {
                _currentOverlays.AddRange(_documentOverlays[_currentDocumentPath]);
            }
        }
        
        private void UpdateOverlayCount()
        {
            var currentPageOverlays = _currentOverlays.Count(o => o.Page == _currentPage);
            _overlayCountLabel.Text = $"Overlays: {currentPageOverlays}/{_currentOverlays.Count}";
            _clearOverlaysBtn.Enabled = _currentOverlays.Count > 0;
        }
        
        private void OnDeleteOverlay(object sender, EventArgs e)
        {
            if (_selectedOverlay != null)
            {
                _currentOverlays.Remove(_selectedOverlay);
                _selectedOverlay = null;
                _deleteOverlayBtn.Enabled = false;
                _documentDisplay.Invalidate();
                UpdateOverlayCount();
                SaveCurrentOverlays(); // Persist the change
            }
        }
        
        private void OnClearAllOverlays(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Are you sure you want to clear all overlays for this document?", 
                "Clear Overlays", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                _currentOverlays.Clear();
                _selectedOverlay = null;
                _deleteOverlayBtn.Enabled = false;
                _documentDisplay.Invalidate();
                UpdateOverlayCount();
                SaveCurrentOverlays(); // Persist the change
            }
        }
        
        // Search functionality implementation
        private async Task ExtractDocumentTextAsync(string documentPath)
        {
            try
            {
                if (_documentTexts.ContainsKey(documentPath))
                    return; // Already extracted
                
                var documentTexts = new List<DocumentText>();
                
                using (var engine = new TesseractEngine("./tessdata", "eng", EngineMode.Default))
                {
                    for (int pageIndex = 0; pageIndex < _documentPages.Count; pageIndex++)
                    {
                        try
                        {
                            using (var page = engine.Process(_documentPages[pageIndex]))
                            {
                                var text = page.GetText();
                                var words = new List<TextWord>();
                                
                                using (var iterator = page.GetIterator())
                                {
                                    iterator.Begin();
                                    do
                                    {
                                        if (iterator.TryGetBoundingBox(PageIteratorLevel.Word, out var rect))
                                        {
                                            var word = iterator.GetText(PageIteratorLevel.Word);
                                            if (!string.IsNullOrWhiteSpace(word))
                                            {
                                                words.Add(new TextWord
                                                {
                                                    Text = word.Trim(),
                                                    Bounds = new System.Drawing.Rectangle(rect.X1, rect.Y1, rect.X2 - rect.X1, rect.Y2 - rect.Y1)
                                                });
                                            }
                                        }
                                    } while (iterator.Next(PageIteratorLevel.Word));
                                }
                                
                                documentTexts.Add(new DocumentText
                                {
                                    PageNumber = pageIndex + 1,
                                    FullText = text,
                                    Words = words
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error($"Error extracting text from page {pageIndex + 1}: {ex.Message}", ex);
                        }
                    }
                }
                
                _documentTexts[documentPath] = documentTexts;
                Logger.Info($"Text extraction completed for {documentPath} ({documentTexts.Count} pages)");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error during text extraction for {documentPath}: {ex.Message}", ex);
            }
        }
        
        private void OnSearch(object sender, EventArgs e)
        {
            PerformSearch();
        }
        
        private void OnSearchTextKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                PerformSearch();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                OnCloseSearch(sender, e);
                e.Handled = true;
            }
        }
        
        private void PerformSearch()
        {
            var searchTerm = _searchTextBox.Text.Trim();
            if (string.IsNullOrEmpty(searchTerm))
            {
                ClearSearch();
                return;
            }
            
            _searchResults.Clear();
            _currentSearchResultIndex = -1;
            
            // Search across all loaded documents
            foreach (var docPath in _documentTexts.Keys)
            {
                var documentTexts = _documentTexts[docPath];
                foreach (var pageText in documentTexts)
                {
                    foreach (var word in pageText.Words)
                    {
                        if (word.Text.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _searchResults.Add(new SearchResult
                            {
                                DocumentPath = docPath,
                                PageNumber = pageText.PageNumber,
                                Word = word,
                                SearchTerm = searchTerm
                            });
                        }
                    }
                }
            }
            
            _isSearchMode = _searchResults.Count > 0;
            UpdateSearchUI();
            
            if (_searchResults.Count > 0)
            {
                _currentSearchResultIndex = 0;
                NavigateToSearchResult(_searchResults[0]);
            }
            else
            {
                _statusLabel.Text = $"No results found for '{searchTerm}'";
            }
        }
        
        private void NavigateSearchResult(int direction)
        {
            if (_searchResults.Count == 0) return;
            
            _currentSearchResultIndex += direction;
            
            if (_currentSearchResultIndex >= _searchResults.Count)
                _currentSearchResultIndex = 0;
            else if (_currentSearchResultIndex < 0)
                _currentSearchResultIndex = _searchResults.Count - 1;
            
            NavigateToSearchResult(_searchResults[_currentSearchResultIndex]);
            UpdateSearchUI();
        }
        
        private void NavigateToSearchResult(SearchResult result)
        {
            // Switch to the document if needed
            if (_currentDocumentPath != result.DocumentPath)
            {
                LoadDocument(result.DocumentPath);
            }
            
            // Navigate to the page
            if (_currentPage != result.PageNumber)
            {
                NavigateToPage(result.PageNumber);
            }
            
            // Trigger a repaint to highlight the search result
            _documentDisplay.Invalidate();
            
            _statusLabel.Text = $"Found '{result.SearchTerm}' in {System.IO.Path.GetFileName(result.DocumentPath)} - Page {result.PageNumber}";
        }
        
        private void UpdateSearchUI()
        {
            if (_searchResults.Count > 0)
            {
                _searchResultsLabel.Text = $"{_currentSearchResultIndex + 1}/{_searchResults.Count}";
                _prevResultBtn.Enabled = true;
                _nextResultBtn.Enabled = true;
                _closeSearchBtn.Visible = true;
            }
            else
            {
                _searchResultsLabel.Text = "0/0";
                _prevResultBtn.Enabled = false;
                _nextResultBtn.Enabled = false;
                _closeSearchBtn.Visible = false;
            }
        }
        
        private void OnCloseSearch(object sender, EventArgs e)
        {
            ClearSearch();
        }
        
        private void ClearSearch()
        {
            _searchResults.Clear();
            _currentSearchResultIndex = -1;
            _isSearchMode = false;
            _searchTextBox.Text = "";
            _searchResultsLabel.Text = "";
            _prevResultBtn.Enabled = false;
            _nextResultBtn.Enabled = false;
            _closeSearchBtn.Visible = false;
            _documentDisplay.Invalidate();
            _statusLabel.Text = "Search cleared";
        }
    }
    
    public class SnipOverlay
    {
        public Guid Id { get; set; }
        public System.Drawing.Rectangle Bounds { get; set; }
        public Color Color { get; set; }
        public int Page { get; set; }
        public SnipperCloneCleanFinal.Core.SnipMode SnipMode { get; set; }
        public string DocumentPath { get; set; }
        public DateTime CreatedAt { get; set; }
    }
    
    public class DocumentText
    {
        public int PageNumber { get; set; }
        public string FullText { get; set; }
        public List<TextWord> Words { get; set; } = new List<TextWord>();
    }
    
    public class TextWord
    {
        public string Text { get; set; }
        public System.Drawing.Rectangle Bounds { get; set; }
    }
    
    public class SearchResult
    {
        public string DocumentPath { get; set; }
        public int PageNumber { get; set; }
        public TextWord Word { get; set; }
        public string SearchTerm { get; set; }
    }

    public class SnipAreaSelectedEventArgs : EventArgs
    {
        public SnipperCloneCleanFinal.Core.SnipMode SnipMode { get; set; }
        public Rectangle Bounds { get; set; }
        public string DocumentPath { get; set; }
        public int PageNumber { get; set; }
        public Bitmap SelectedImage { get; set; }
        public string ExtractedText { get; set; } = "";
        public string[] ExtractedNumbers { get; set; } = new string[0];
        public bool Success { get; set; } = true;
    }
}
