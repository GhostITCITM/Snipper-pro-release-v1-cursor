using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SnipperCloneCleanFinal.UI;
using SnipperCloneCleanFinal.Infrastructure;

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
        
        private string _currentDocumentPath;
        private int _currentPage = 1;
        private int _totalPages = 1;
        private float _zoomLevel = 1.0f;
        private List<Bitmap> _documentPages = new List<Bitmap>();
        private List<SnipOverlay> _overlays = new List<SnipOverlay>();
        
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
            // Toolbar
            _toolbarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.LightGray
            };
            
            _prevPageBtn = new Button { Text = "◀", Size = new Size(30, 30), Location = new Point(5, 5) };
            _nextPageBtn = new Button { Text = "▶", Size = new Size(30, 30), Location = new Point(40, 5) };
            _pageLabel = new Label { Text = "Page 1 of 1", Location = new Point(80, 12), Size = new Size(80, 20) };
            _zoomInBtn = new Button { Text = "+", Size = new Size(30, 30), Location = new Point(170, 5) };
            _zoomOutBtn = new Button { Text = "-", Size = new Size(30, 30), Location = new Point(205, 5) };
            _fitWidthBtn = new Button { Text = "Fit", Size = new Size(40, 30), Location = new Point(240, 5) };
            
            _toolbarPanel.Controls.AddRange(new Control[] { 
                _prevPageBtn, _nextPageBtn, _pageLabel, _zoomInBtn, _zoomOutBtn, _fitWidthBtn 
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
            
            _documentDisplay.MouseDown += OnDocumentMouseDown;
            _documentDisplay.MouseMove += OnDocumentMouseMove;
            _documentDisplay.MouseUp += OnDocumentMouseUp;
            _documentDisplay.Paint += OnDocumentPaint;
        }
        
        public bool LoadDocument(string filePath)
        {
            try
            {
                _currentDocumentPath = filePath;

                // Dispose previously loaded pages to avoid memory leaks
                foreach (var page in _documentPages)
                {
                    page?.Dispose();
                }
                _documentPages.Clear();
                
                // For now, handle image files - PDF support would need additional library
                if (IsImageFile(filePath))
                {
                    var image = new Bitmap(filePath);
                    _documentPages.Add(image);
                    _totalPages = 1;
                }
                else
                {
                    _statusLabel.Text = "PDF support coming soon - please use image files for now";
                    return false;
                }
                
                _currentPage = 1;
                UpdateDisplay();
                UpdatePageLabel();
                _statusLabel.Text = $"Loaded: {System.IO.Path.GetFileName(filePath)}";
                
                return true;
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Error loading document: {ex.Message}";
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
            
            _overlays.Add(overlay);
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
            // Draw existing overlays
            foreach (var overlay in _overlays)
            {
                if (overlay.Page == _currentPage)
                {
                    using (var pen = new Pen(overlay.Color, 2))
                    {
                        e.Graphics.DrawRectangle(pen, overlay.Bounds);
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
            }
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
    }
    
    public class SnipOverlay
    {
        public System.Drawing.Rectangle Bounds { get; set; }
        public Color Color { get; set; }
        public int Page { get; set; }
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