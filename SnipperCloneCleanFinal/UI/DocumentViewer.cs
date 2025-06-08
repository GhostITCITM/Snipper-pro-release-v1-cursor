using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SnipperCloneCleanFinal.Core;
using SnipperCloneCleanFinal.Infrastructure;
using CoreRectangle = SnipperCloneCleanFinal.Core.Rectangle;
using System.Text;
using PdfiumViewer;

namespace SnipperCloneCleanFinal.UI
{
    public partial class DocumentViewer : Form
    {
        private const int PDF_RENDER_DPI = 150; // DPI used when rendering PDF pages – keep in sync for accurate coordinate mapping
        private readonly SnipEngine _snippEngine;
        private readonly OCREngine _ocrEngine;
        private Panel _documentsPanel;
        private Panel _viewerPanel; 
        private PictureBox _documentPictureBox;
        private Label _statusLabel;
        private Label _pageLabel;
        private Button _prevPageButton;
        private Button _nextPageButton;
        private ComboBox _documentSelector;
        private Button _loadDocumentButton;
        private Button _fitToWidthButton;
        private Button _zoomInButton;
        private Button _zoomOutButton;
        
        private List<LoadedDocument> _loadedDocuments = new List<LoadedDocument>();
        private LoadedDocument _currentDocument;
        private int _currentPageIndex = 0;
        private float _zoomFactor = 1.0f;
        private bool _isSnipMode = false;
        private SnipMode _currentSnipMode = SnipMode.None;
        private Point _selectionStart;
        private Point _selectionEnd;
        private bool _isSelecting = false;
        private System.Drawing.Rectangle _currentSelection = System.Drawing.Rectangle.Empty;
        
        // Table snip helpers
        private List<System.Drawing.Rectangle> _tableColumns = new List<System.Drawing.Rectangle>();
        private List<System.Drawing.Rectangle> _tableRows = new List<System.Drawing.Rectangle>();
        private bool _showTableGrid = false;
        private bool _adjustingTable = false;
        private int _draggingColumnIndex = -1;
        private int _dragStartX;
        
        public event EventHandler<SnipAreaSelectedEventArgs> SnipAreaSelected;

        public DocumentViewer(SnipEngine snippEngine)
        {
            _snippEngine = snippEngine ?? throw new ArgumentNullException(nameof(snippEngine));
            _ocrEngine = new OCREngine();
            InitializeComponent();
            SetupUI();
            Logger.Info("DocumentViewer initialized with full functionality");
        }

        private void SetupUI()
        {
            this.Text = "Snipper Pro - Document Viewer";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Normal;
            
            // Make the viewer stay on top of Excel
            this.TopMost = true;
            
            // Prevent accidental closing - minimize instead
            this.FormClosing += (s, e) => 
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    this.WindowState = FormWindowState.Minimized;
                }
            };

            // Create main panels
            CreateToolbar();
            CreateDocumentPanel();
            CreateViewerPanel();
            CreateStatusBar();
            
            Logger.Info("DocumentViewer UI setup completed");
        }

        private void CreateToolbar()
        {
            var toolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.LightGray,
                Padding = new Padding(5)
            };

            _loadDocumentButton = new Button
            {
                Text = "Load Document(s)",
                Size = new Size(120, 40),
                Location = new Point(5, 5)
            };
            _loadDocumentButton.Click += OnLoadDocuments;

            _documentSelector = new ComboBox
            {
                Size = new Size(200, 25),
                Location = new Point(135, 12),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _documentSelector.SelectedIndexChanged += OnDocumentSelected;

            _prevPageButton = new Button
            {
                Text = "◀",
                Size = new Size(30, 30),
                Location = new Point(345, 10)
            };
            _prevPageButton.Click += OnPreviousPage;

            _pageLabel = new Label
            {
                Text = "Page 1 of 1",
                Size = new Size(80, 30),
                Location = new Point(385, 15),
                TextAlign = ContentAlignment.MiddleCenter
            };

            _nextPageButton = new Button
            {
                Text = "▶",
                Size = new Size(30, 30),
                Location = new Point(475, 10)
            };
            _nextPageButton.Click += OnNextPage;

            _zoomOutButton = new Button
            {
                Text = "🔍-",
                Size = new Size(40, 30),
                Location = new Point(515, 10)
            };
            _zoomOutButton.Click += OnZoomOut;

            _zoomInButton = new Button
            {
                Text = "🔍+",
                Size = new Size(40, 30),
                Location = new Point(565, 10)
            };
            _zoomInButton.Click += OnZoomIn;

            _fitToWidthButton = new Button
            {
                Text = "Fit Width",
                Size = new Size(80, 30),
                Location = new Point(615, 10)
            };
            _fitToWidthButton.Click += OnFitToWidth;

            toolbar.Controls.AddRange(new Control[] {
                _loadDocumentButton, _documentSelector, _prevPageButton, _pageLabel, 
                _nextPageButton, _zoomOutButton, _zoomInButton, _fitToWidthButton
            });

            this.Controls.Add(toolbar);
        }

        private void CreateDocumentPanel()
        {
            _documentsPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 200,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(5)
            };

            var label = new Label
            {
                Text = "Loaded Documents:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font(Font, FontStyle.Bold)
            };
            
            _documentsPanel.Controls.Add(label);
            this.Controls.Add(_documentsPanel);
        }

        private void CreateViewerPanel()
        {
            _viewerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.DarkGray,
                AutoScroll = true,
                AutoScrollMinSize = new Size(1200, 1600) // Ensure scrollbars appear
            };

            _documentPictureBox = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.AutoSize,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(10, 10),
                Anchor = AnchorStyles.None // Don't anchor so it can be centered
            };
            
            // Add mouse events for snipping
            _documentPictureBox.MouseDown += OnMouseDown;
            _documentPictureBox.MouseMove += OnMouseMove;
            _documentPictureBox.MouseUp += OnMouseUp;
            _documentPictureBox.Paint += OnPaint;
            _documentPictureBox.DoubleClick += OnPictureDoubleClick;

            _viewerPanel.Controls.Add(_documentPictureBox);
            this.Controls.Add(_viewerPanel);
        }

        private void CreateStatusBar()
        {
            _statusLabel = new Label
            {
                Text = "Ready - Load documents to begin",
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.LightGray,
                Padding = new Padding(10, 5, 10, 5)
            };

            this.Controls.Add(_statusLabel);
        }

        private async void OnLoadDocuments(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF files (*.pdf)|*.pdf|Image files (*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|All supported files|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|All files (*.*)|*.*";
                openFileDialog.Title = "Select Document(s) to Load - PDFs and Images Supported";
                openFileDialog.Multiselect = true;
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    await LoadDocuments(openFileDialog.FileNames);
                }
            }
        }

        public async Task LoadDocuments(string[] filePaths)
        {
            try
            {
                // Update UI on main thread
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => _statusLabel.Text = "Loading documents..."));
                }
                else
                {
                    _statusLabel.Text = "Loading documents...";
                }

                var loadedDocs = new List<LoadedDocument>();
                var docNames = new List<string>();

                foreach (var filePath in filePaths)
                {
                    var document = await LoadDocumentInternal(filePath);
                    if (document != null)
                    {
                        loadedDocs.Add(document);
                        docNames.Add(Path.GetFileName(filePath));
                        Logger.Info($"Loaded document: {filePath}");
                    }
                }

                // Update UI on main thread
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() =>
                    {
                        foreach (var doc in loadedDocs)
                        {
                            _loadedDocuments.Add(doc);
                        }
                        foreach (var name in docNames)
                        {
                            _documentSelector.Items.Add(name);
                        }

                        if (_loadedDocuments.Count > 0)
                        {
                            _documentSelector.SelectedIndex = 0;
                            UpdateDocumentsList();
                            _statusLabel.Text = $"Loaded {_loadedDocuments.Count} document(s)";
                        }
                        else
                        {
                            _statusLabel.Text = "No documents could be loaded";
                        }
                    }));
                }
                else
                {
                    foreach (var doc in loadedDocs)
                    {
                        _loadedDocuments.Add(doc);
                    }
                    foreach (var name in docNames)
                    {
                        _documentSelector.Items.Add(name);
                    }

                    if (_loadedDocuments.Count > 0)
                    {
                        _documentSelector.SelectedIndex = 0;
                        UpdateDocumentsList();
                        _statusLabel.Text = $"Loaded {_loadedDocuments.Count} document(s)";
                    }
                    else
                    {
                        _statusLabel.Text = "No documents could be loaded";
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading documents: {ex.Message}", ex);
                
                // Update UI on main thread
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => _statusLabel.Text = $"Error loading documents: {ex.Message}"));
                }
                else
                {
                    _statusLabel.Text = $"Error loading documents: {ex.Message}";
                }
            }
        }

        private async Task<LoadedDocument> LoadDocumentInternal(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".pdf")
                {
                    return await LoadPdfDocument(filePath);
                }
                else if (new[] { ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif" }.Contains(extension))
                {
                    return LoadImageDocument(filePath);
                }
                
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading document {filePath}: {ex.Message}", ex);
                return null;
            }
        }

        private async Task<LoadedDocument> LoadPdfDocument(string filePath)
        {
            try
            {
                // Convert PDF to images using System.Drawing and basic PDF processing
                var pdfPages = ConvertPdfToImages(filePath);
                
                if (pdfPages.Count > 0)
                {
                    return new LoadedDocument
                    {
                        FilePath = filePath,
                        Name = Path.GetFileName(filePath),
                        Type = DocumentType.PDF,
                        Pages = pdfPages,
                        PageCount = pdfPages.Count
                    };
                }
                else
                {
                    // Fallback: create a readable representation of the PDF
                    var fallbackImage = CreatePdfRepresentation(filePath);
                    return new LoadedDocument
                    {
                        FilePath = filePath,
                        Name = Path.GetFileName(filePath),
                        Type = DocumentType.PDF,
                        Pages = new List<Bitmap> { fallbackImage },
                        PageCount = 1
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading PDF {filePath}: {ex.Message}", ex);
                
                // Create a readable error representation
                var errorImage = CreateErrorPdfView(filePath, ex.Message);
                return new LoadedDocument
                {
                    FilePath = filePath,
                    Name = Path.GetFileName(filePath),
                    Type = DocumentType.PDF,
                    Pages = new List<Bitmap> { errorImage },
                    PageCount = 1
                };
            }
        }

        private List<Bitmap> ConvertPdfToImages(string pdfPath)
        {
            var images = new List<Bitmap>();
            try
            {
                Logger.Info($"Converting PDF to images: {pdfPath}");
                
                // Use PdfiumViewer to render real PDF pages into bitmaps
                using (var document = PdfiumViewer.PdfDocument.Load(pdfPath))
                {
                    Logger.Info($"PDF loaded successfully, {document.PageCount} pages");
                    var dpiX = 150; // High quality rendering
                    var dpiY = 150;
                    
                    for (int pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
                    {
                        try
                        {
                            // Calculate size keeping aspect ratio
                            var size = document.PageSizes[pageIndex];
                            var width = (int)(size.Width * (dpiX / 72.0));
                            var height = (int)(size.Height * (dpiY / 72.0));
                            
                            // Ensure reasonable size limits to prevent memory issues
                            if (width > 4000) width = 4000;
                            if (height > 4000) height = 4000;
                            if (width < 100) width = 100;
                            if (height < 100) height = 100;
                            
                            Logger.Info($"Rendering page {pageIndex + 1} at {width}x{height}");
                            
                            using (var rendered = document.Render(pageIndex, width, height, dpiX, dpiY, PdfiumViewer.PdfRenderFlags.Annotations))
                            {
                                // Create a copy to avoid disposal issues
                                var pageBitmap = new Bitmap(rendered);
                                images.Add(pageBitmap);
                                Logger.Info($"Successfully rendered page {pageIndex + 1}");
                            }
                        }
                        catch (Exception pageEx)
                        {
                            Logger.Error($"Failed to render page {pageIndex + 1}: {pageEx.Message}", pageEx);
                            // Create an error page for this specific page
                            var errorPage = CreateErrorPageBitmap($"Failed to render page {pageIndex + 1}");
                            images.Add(errorPage);
                        }
                    }
                }
                Logger.Info($"Successfully rendered {images.Count} pages from PDF");
            }
            catch (Exception ex)
            {
                Logger.Error($"PDF rendering failed: {ex.Message}", ex);
                Logger.Error($"Exception type: {ex.GetType().FullName}");
                if (ex.InnerException != null)
                {
                    Logger.Error($"Inner exception: {ex.InnerException.Message}");
                }
                
                // Create a simple error visualization instead of text fallback
                images.Clear();
                var errorBitmap = CreateErrorPageBitmap($"PDF Rendering Failed: {ex.Message}");
                images.Add(errorBitmap);
            }
            
            return images;
        }

        private Bitmap CreateErrorPageBitmap(string errorMessage)
        {
            var bitmap = new Bitmap(800, 600);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.FillRectangle(Brushes.White, 0, 0, 800, 600);
                g.DrawRectangle(Pens.Red, 10, 10, 780, 580);
                
                using (var font = new Font("Arial", 12, FontStyle.Bold))
                {
                    var rect = new RectangleF(20, 50, 760, 500);
                    g.DrawString($"Error: {errorMessage}\n\nPlease check the PDF file and try again.", 
                                font, Brushes.DarkRed, rect);
                }
            }
            return bitmap;
        }

        private bool TryWindowsPdfRendering(string pdfPath, List<Bitmap> images)
        {
            try
            {
                // This would require Windows.Data.Pdf which isn't available in .NET Framework
                // But we can simulate the output for now and add real PDF support later
                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool TryPrintDocumentPdfConversion(string pdfPath, List<Bitmap> images)
        {
            try
            {
                // Try to open PDF with default system viewer and capture
                // This is a simplified approach - in production you'd use a proper PDF library
                return false;
            }
            catch
            {
                return false;
            }
        }

        private Bitmap CreateAdvancedPdfRepresentation(string pdfPath)
        {
            try
            {
                // Try to extract and display REAL PDF content
                var pdfBytes = File.ReadAllBytes(pdfPath);
                var realText = ExtractRealTextFromPdf(pdfBytes);
                
                // Create a proper document view with REAL content
                var image = new Bitmap(1200, 1600); // Larger canvas for real content
                using (var g = Graphics.FromImage(image))
                {
                    g.FillRectangle(Brushes.White, 0, 0, 1200, 1600);
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                    
                    // Document header
                    var fileName = Path.GetFileNameWithoutExtension(pdfPath);
                    using (var titleFont = new Font("Arial", 14, FontStyle.Bold))
                    {
                        g.DrawString(fileName, titleFont, Brushes.DarkBlue, 20, 20);
                    }
                    
                    // Separator
                    using (var pen = new Pen(Color.LightGray, 1))
                    {
                        g.DrawLine(pen, 20, 50, 1180, 50);
                    }
                    
                    // Display REAL PDF TEXT CONTENT
                    using (var contentFont = new Font("Courier New", 9)) // Monospace for better text layout
                    {
                        var lines = realText.Split('\n');
                        int y = 70;
                        int lineHeight = 14;
                        int maxLines = (1600 - 100) / lineHeight;
                        
                        for (int i = 0; i < Math.Min(lines.Length, maxLines); i++)
                        {
                            var line = lines[i];
                            if (line.Trim().Length > 0)
                            {
                                // Wrap long lines
                                if (line.Length > 130)
                                {
                                    var wrappedLines = WrapText(line, 130);
                                    foreach (var wrappedLine in wrappedLines)
                                    {
                                        if (y > 1550) break;
                                        g.DrawString(wrappedLine, contentFont, Brushes.Black, 20, y);
                                        y += lineHeight;
                                    }
                                }
                                else
                                {
                                    g.DrawString(line, contentFont, Brushes.Black, 20, y);
                                    y += lineHeight;
                                }
                            }
                            else
                            {
                                y += lineHeight / 2; // Blank line
                            }
                            
                            if (y > 1550) break;
                        }
                        
                        // Add truncation notice if needed
                        if (lines.Length > maxLines)
                        {
                            g.DrawString($"... ({lines.Length - maxLines} more lines in PDF)", 
                                contentFont, Brushes.Gray, 20, y + 10);
                        }
                    }
                    
                    // Border
                    using (var borderPen = new Pen(Color.DarkBlue, 1))
                    {
                        g.DrawRectangle(borderPen, 5, 5, 1190, 1590);
                    }
                }
                
                return image;
            }
            catch (Exception ex)
            {
                Logger.Error($"Error creating real PDF representation: {ex.Message}", ex);
                return CreateErrorPdfView(pdfPath, ex.Message);
            }
        }

        private string ExtractRealTextFromPdf(byte[] pdfBytes)
        {
            try
            {
                // Method 1: Try to extract readable text streams from PDF
                var content = System.Text.Encoding.UTF8.GetString(pdfBytes);
                var extractedText = new StringBuilder();
                
                // Find text objects in PDF (between BT and ET markers)
                var textMatches = System.Text.RegularExpressions.Regex.Matches(content, @"BT\s(.*?)\sET", 
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                
                foreach (System.Text.RegularExpressions.Match match in textMatches)
                {
                    var textBlock = match.Groups[1].Value;
                    
                    // Extract text from Tj commands
                    var tjMatches = System.Text.RegularExpressions.Regex.Matches(textBlock, @"\((.*?)\)\s*Tj");
                    foreach (System.Text.RegularExpressions.Match tjMatch in tjMatches)
                    {
                        var text = tjMatch.Groups[1].Value;
                        if (!string.IsNullOrWhiteSpace(text) && text.Length > 1)
                        {
                            extractedText.AppendLine(CleanPdfText(text));
                        }
                    }
                    
                    // Extract text from TJ arrays
                    var tjArrayMatches = System.Text.RegularExpressions.Regex.Matches(textBlock, @"\[(.*?)\]\s*TJ");
                    foreach (System.Text.RegularExpressions.Match tjArrayMatch in tjArrayMatches)
                    {
                        var arrayContent = tjArrayMatch.Groups[1].Value;
                        var stringMatches = System.Text.RegularExpressions.Regex.Matches(arrayContent, @"\((.*?)\)");
                        foreach (System.Text.RegularExpressions.Match stringMatch in stringMatches)
                        {
                            var text = stringMatch.Groups[1].Value;
                            if (!string.IsNullOrWhiteSpace(text) && text.Length > 1)
                            {
                                extractedText.Append(CleanPdfText(text) + " ");
                            }
                        }
                        if (stringMatches.Count > 0) extractedText.AppendLine();
                    }
                }
                
                // Method 2: If no text objects found, try stream extraction
                if (extractedText.Length < 50)
                {
                    return ExtractFromPdfStreams(content);
                }
                
                return extractedText.ToString();
            }
            catch
            {
                // Fallback: try simple text extraction
                return ExtractSimplePdfText(pdfBytes);
            }
        }

        private string ExtractFromPdfStreams(string content)
        {
            var extractedText = new StringBuilder();
            
            // Find readable text in streams
            var readableWords = System.Text.RegularExpressions.Regex.Matches(content, @"\b[A-Za-z][A-Za-z0-9\s]{3,50}\b")
                .Cast<System.Text.RegularExpressions.Match>()
                .Select(m => m.Value.Trim())
                .Where(w => w.Length > 3 && IsLikelyReadableText(w))
                .Distinct()
                .Take(200);
            
            var currentLine = new StringBuilder();
            foreach (var word in readableWords)
            {
                if (currentLine.Length + word.Length > 80)
                {
                    extractedText.AppendLine(currentLine.ToString());
                    currentLine.Clear();
                }
                currentLine.Append(word + " ");
            }
            
            if (currentLine.Length > 0)
            {
                extractedText.AppendLine(currentLine.ToString());
            }
            
            return extractedText.ToString();
        }

        private string ExtractSimplePdfText(byte[] pdfBytes)
        {
            // Last resort: extract any readable sequences
            var content = System.Text.Encoding.UTF8.GetString(pdfBytes);
            var text = new StringBuilder();
            
            // Find sequences of printable characters
            var printableSeqs = System.Text.RegularExpressions.Regex.Matches(content, @"[A-Za-z0-9\s\.\,\:\;\!\?\$\%\-\+\=\(\)]{10,}")
                .Cast<System.Text.RegularExpressions.Match>()
                .Select(m => m.Value.Trim())
                .Where(s => s.Length > 5 && s.Count(char.IsLetter) > s.Length / 3)
                .Distinct()
                .Take(50);
            
            foreach (var seq in printableSeqs)
            {
                text.AppendLine(seq);
            }
            
            return text.ToString();
        }

        private string CleanPdfText(string text)
        {
            // Clean up PDF text encoding issues
            return text
                .Replace("\\(", "(")
                .Replace("\\)", ")")
                .Replace("\\\\", "\\")
                .Replace("\\n", "\n")
                .Replace("\\r", "")
                .Replace("\\t", " ");
        }

        private bool IsLikelyReadableText(string text)
        {
            if (string.IsNullOrWhiteSpace(text) || text.Length < 3) return false;
            
            var letterCount = text.Count(char.IsLetter);
            var digitCount = text.Count(char.IsDigit);
            var totalChars = text.Length;
            
            // Must have some letters
            if (letterCount == 0) return false;
            
            // Ratio of letters to total should be reasonable
            var letterRatio = (double)letterCount / totalChars;
            return letterRatio > 0.3;
        }

        private string[] WrapText(string text, int maxLength)
        {
            var words = text.Split(' ');
            var lines = new List<string>();
            var currentLine = new StringBuilder();
            
            foreach (var word in words)
            {
                if (currentLine.Length + word.Length + 1 > maxLength && currentLine.Length > 0)
                {
                    lines.Add(currentLine.ToString());
                    currentLine.Clear();
                }
                
                if (currentLine.Length > 0) currentLine.Append(" ");
                currentLine.Append(word);
            }
            
            if (currentLine.Length > 0)
            {
                lines.Add(currentLine.ToString());
            }
            
            return lines.ToArray();
        }

        private Bitmap CreateErrorPdfView(string filePath, string error)
        {
            var image = new Bitmap(800, 600);
            using (var g = Graphics.FromImage(image))
            {
                g.FillRectangle(Brushes.White, 0, 0, 800, 600);
                
                using (var font = new Font("Arial", 12, FontStyle.Bold))
                {
                    g.DrawString($"PDF: {Path.GetFileName(filePath)}", font, Brushes.DarkBlue, 20, 20);
                    g.DrawString("Could not extract text from this PDF", font, Brushes.Red, 20, 60);
                    g.DrawString("This PDF may be image-based or encrypted", font, Brushes.Black, 20, 100);
                    g.DrawString("You can still try using the snip tools on any visible content", font, Brushes.Black, 20, 140);
                }
            }
            return image;
        }

        private void OnDocumentSelected(object sender, EventArgs e)
        {
            if (_documentSelector.SelectedIndex >= 0 && _documentSelector.SelectedIndex < _loadedDocuments.Count)
            {
                _currentDocument = _loadedDocuments[_documentSelector.SelectedIndex];
                _currentPageIndex = 0;
                DisplayCurrentPage();
            }
        }

        private void DisplayCurrentPage()
        {
            if (_currentDocument == null || _currentPageIndex < 0 || _currentPageIndex >= _currentDocument.PageCount)
                return;

            var page = _currentDocument.Pages[_currentPageIndex];
            var scaledImage = ScaleImage(page, _zoomFactor);
            
            _documentPictureBox.Image?.Dispose();
            _documentPictureBox.Image = scaledImage;
            _documentPictureBox.Size = scaledImage.Size;
            
            // Ensure scrollbars appear by setting the size
            _documentPictureBox.Width = scaledImage.Width;
            _documentPictureBox.Height = scaledImage.Height;
            
            // Update panel's auto-scroll size to match the scaled image
            _viewerPanel.AutoScrollMinSize = new Size(
                scaledImage.Width + 40, 
                scaledImage.Height + 40
            );
            
            // Position the image at top-left with some padding
            _documentPictureBox.Location = new Point(10, 10);
            
            // Force panel to update scrollbars
            _viewerPanel.PerformLayout();
            
            _pageLabel.Text = $"Page {_currentPageIndex + 1} of {_currentDocument.PageCount}";
            _statusLabel.Text = $"Viewing: {_currentDocument.Name} - Page {_currentPageIndex + 1} - Zoom: {_zoomFactor:P0} - REAL PDF CONTENT";
            
            // Clear selection
            _currentSelection = System.Drawing.Rectangle.Empty;
            _tableColumns.Clear();
            _tableRows.Clear();
            _showTableGrid = false;
            
            _documentPictureBox.Invalidate();
            _viewerPanel.Invalidate();
        }

        private Bitmap ScaleImage(Bitmap original, float scale)
        {
            int newWidth = (int)(original.Width * scale);
            int newHeight = (int)(original.Height * scale);
            
            var scaled = new Bitmap(newWidth, newHeight);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(original, 0, 0, newWidth, newHeight);
            }
            
            return scaled;
        }

        public void SetSnipMode(SnipMode snipMode, bool enabled)
        {
            _currentSnipMode = snipMode;
            _isSnipMode = enabled;
            
            if (enabled)
            {
                _statusLabel.Text = $"{snipMode} Snip mode ACTIVE - Draw a rectangle to extract data";
                _documentPictureBox.Cursor = Cursors.Cross;
                
                // For table mode, give specific instructions
                if (snipMode == SnipMode.Table)
                {
                    _statusLabel.Text = "Table Snip mode ACTIVE - Draw a rectangle around the table, adjust columns, then double-click to extract";
                }
            }
            else
            {
                _statusLabel.Text = "Snip mode disabled";
                _documentPictureBox.Cursor = Cursors.Default;
                
                // Clear any table adjustment state
                _adjustingTable = false;
                _showTableGrid = false;
                _tableColumns.Clear();
            }
            
            _documentPictureBox.Invalidate();
        }

        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                // Check if the PictureBox is valid and not disposed
                if (_documentPictureBox == null || _documentPictureBox.IsDisposed)
                    return;

                if (_adjustingTable)
                {
                    // Handle left-click events - check icons FIRST before selection bounds
                    if (e.Button == MouseButtons.Left)
                    {
                        Logger.Info($"Mouse click at ({e.X}, {e.Y}) while adjusting table");
                        
                        // Constants must match OnPaint exactly
                        int iconOffset = 20;
                        int iconSize = 16;
                        int centerY = _currentSelection.Y - iconOffset;
                        
                        Logger.Info($"Looking for buttons at Y={centerY} with size={iconSize}");

                        // 1. Check if clicking a "–" (minus) button to remove a column divider
                        // Do this FIRST, before checking selection bounds
                        for (int i = _tableColumns.Count - 1; i >= 0; i--) // Iterate backwards to avoid index issues
                        {
                            int centerX = _tableColumns[i].X + _tableColumns[i].Width / 2;
                            var buttonRect = new System.Drawing.Rectangle(centerX - iconSize/2, centerY - iconSize/2, iconSize, iconSize);
                            
                            if (buttonRect.Contains(e.Location))
                            {
                                Logger.Info($"Removing column divider at index {i}");
                                _tableColumns.RemoveAt(i);
                                SafeInvalidate();
                                return; // Important: return here to prevent further processing
                            }
                        }

                        // 2. Check if clicking a "+" button to add a new column divider
                        // Do this SECOND, before checking selection bounds
                        var boundaries = new List<int> { _currentSelection.X };
                        boundaries.AddRange(_tableColumns.Select(c => c.X + c.Width / 2));
                        boundaries.Add(_currentSelection.Right);
                        boundaries.Sort();

                        for (int b = 0; b < boundaries.Count - 1; b++)
                        {
                            int gapLeft = boundaries[b];
                            int gapRight = boundaries[b + 1];
                            if (gapRight - gapLeft > 40) // Must match OnPaint condition
                            {
                                int centerX = gapLeft + (gapRight - gapLeft) / 2;
                                var buttonRect = new System.Drawing.Rectangle(centerX - iconSize/2, centerY - iconSize/2, iconSize, iconSize);
                                
                                if (buttonRect.Contains(e.Location))
                                {
                                    Logger.Info($"Adding new column divider at X={centerX}");
                                    int newRectX = centerX - 2;
                                    if (newRectX > _currentSelection.X && newRectX < _currentSelection.Right - 4)
                                    {
                                        _tableColumns.Add(new System.Drawing.Rectangle(newRectX, _currentSelection.Y, 4, _currentSelection.Height));
                                        _tableColumns = _tableColumns.OrderBy(c => c.X).ToList();
                                        SafeInvalidate();
                                    }
                                    return; // Important: return here to prevent further processing
                                }
                            }
                        }

                        // 3. NOW check if clicking within selection bounds for divider dragging
                        if (_currentSelection.Contains(e.Location))
                        {
                            Logger.Info("Click is within selection bounds, checking for divider dragging");
                            
                            // Check if clicking an existing divider line (to start dragging)
                            for (int i = 0; i < _tableColumns.Count; i++)
                            {
                                var hit = _tableColumns[i];
                                var hitZone = new System.Drawing.Rectangle(hit.X - 5, hit.Y, hit.Width + 10, hit.Height);
                                if (hitZone.Contains(e.Location))
                                {
                                    Logger.Info($"Starting drag for column {i}");
                                    _draggingColumnIndex = i;
                                    _dragStartX = e.X;
                                    try
                                    {
                                        _documentPictureBox.Cursor = Cursors.VSplit;
                                    }
                                    catch (Exception)
                                    {
                                        // Ignore cursor setting errors
                                    }
                                    return;
                                }
                            }
                            
                            // If clicking inside selection but not on icons or dividers, just stay in adjust mode
                            Logger.Info("Click inside selection but not on any interactive element - staying in adjust mode");
                            return;
                        }
                        
                        // 4. ONLY exit adjust mode if clicking far outside the selection AND button areas
                        // Expand the bounds to include button area
                        var expandedBounds = new System.Drawing.Rectangle(
                            _currentSelection.X - 20, 
                            _currentSelection.Y - iconOffset - iconSize - 5, 
                            _currentSelection.Width + 40, 
                            _currentSelection.Height + iconOffset + iconSize + 10);
                            
                        if (!expandedBounds.Contains(e.Location))
                        {
                            Logger.Info("Click outside expanded bounds - exiting table adjust mode");
                            _adjustingTable = false;
                            _tableColumns.Clear();
                            _showTableGrid = false;
                            SafeInvalidate();
                        }
                        else
                        {
                            Logger.Info("Click within expanded bounds but not on interactive elements - staying in adjust mode");
                        }
                    }
                    
                    // Handle right-click for backward compatibility
                    if (e.Button == MouseButtons.Right && _currentSelection.Contains(e.Location))
                    {
                        // Keep right-click functionality for backward compatibility
                        for (int i = 0; i < _tableColumns.Count; i++)
                        {
                            if (_tableColumns[i].Contains(e.Location))
                            {
                                _tableColumns.RemoveAt(i);
                                SafeInvalidate();
                                return;
                            }
                        }
                        var newRectX = e.X - 2;
                        if (newRectX > _currentSelection.X && newRectX < _currentSelection.Right - 4)
                        {
                            _tableColumns.Add(new System.Drawing.Rectangle(newRectX, _currentSelection.Y, 4, _currentSelection.Height));
                            _tableColumns = _tableColumns.OrderBy(c => c.X).ToList();
                            SafeInvalidate();
                        }
                        return;
                    }
                }

                if (!_isSnipMode || e.Button != MouseButtons.Left) return;

                _isSelecting = true;
                _selectionStart = e.Location;
                _selectionEnd = e.Location;
                _currentSelection = System.Drawing.Rectangle.Empty;
            }
            catch (System.Runtime.InteropServices.SEHException)
            {
                // Handle SEH exceptions from unmanaged code
                System.Diagnostics.Debug.WriteLine("SEH Exception in OnMouseDown - ignoring");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in OnMouseDown: {ex.Message}");
            }
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                // Check if the PictureBox is valid and not disposed
                if (_documentPictureBox == null || _documentPictureBox.IsDisposed)
                    return;

                if (_draggingColumnIndex >= 0 && _adjustingTable)
                {
                    if (_draggingColumnIndex < _tableColumns.Count)
                    {
                        var rect = _tableColumns[_draggingColumnIndex];
                        int newX = rect.X + (e.X - _dragStartX);
                        newX = Math.Max(_currentSelection.X, Math.Min(newX, _currentSelection.Right - rect.Width));
                        rect.X = newX;
                        _tableColumns[_draggingColumnIndex] = rect;
                        _dragStartX = e.X;
                        SafeInvalidate();
                    }
                    return;
                }

                if (_adjustingTable)
                {
                    bool overLine = false;
                    bool overIcon = false;
                    
                    // Check if over an existing divider line (for dragging)
                    foreach (var col in _tableColumns)
                    {
                        var zone = new System.Drawing.Rectangle(col.X - 5, col.Y, col.Width + 10, col.Height);
                        if (zone.Contains(e.Location)) 
                        { 
                            overLine = true; 
                            break; 
                        }
                    }
                    
                    // Check if over plus/minus buttons (for add/remove actions)
                    if (!overLine)
                    {
                        int iconOffset = 20;
                        int iconSize = 16;
                        int centerY = _currentSelection.Y - iconOffset;
                        
                        // Check minus buttons
                        foreach (var col in _tableColumns)
                        {
                            int centerX = col.X + col.Width / 2;
                            var buttonRect = new System.Drawing.Rectangle(centerX - iconSize/2, centerY - iconSize/2, iconSize, iconSize);
                            if (buttonRect.Contains(e.Location))
                            {
                                overIcon = true;
                                break;
                            }
                        }
                        
                        // Check plus buttons if not over minus button
                        if (!overIcon)
                        {
                            var boundaries = new List<int> { _currentSelection.X };
                            boundaries.AddRange(_tableColumns.Select(c => c.X + c.Width / 2));
                            boundaries.Add(_currentSelection.Right);
                            boundaries.Sort();
                            
                            for (int b = 0; b < boundaries.Count - 1; b++)
                            {
                                int gapLeft = boundaries[b];
                                int gapRight = boundaries[b + 1];
                                if (gapRight - gapLeft > 40)
                                {
                                    int centerX = gapLeft + (gapRight - gapLeft) / 2;
                                    var buttonRect = new System.Drawing.Rectangle(centerX - iconSize/2, centerY - iconSize/2, iconSize, iconSize);
                                    if (buttonRect.Contains(e.Location))
                                    {
                                        overIcon = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    
                    try
                    {
                        if (overLine)
                            _documentPictureBox.Cursor = Cursors.VSplit;
                        else if (overIcon)
                            _documentPictureBox.Cursor = Cursors.Hand;
                        else
                            _documentPictureBox.Cursor = Cursors.Default;
                    }
                    catch (Exception)
                    {
                        // Ignore cursor setting errors
                    }
                }

                if (!_isSelecting) return;

                _selectionEnd = e.Location;
                _currentSelection = GetNormalizedRectangle(_selectionStart, _selectionEnd);

                // For table snip, show column dividers
                if (_currentSnipMode == SnipMode.Table && _currentSelection.Width > 20)
                {
                    DetectTableStructure(_currentSelection);
                }

                SafeInvalidate();
            }
            catch (System.Runtime.InteropServices.SEHException)
            {
                // Handle SEH exceptions from unmanaged code
                System.Diagnostics.Debug.WriteLine("SEH Exception in OnMouseMove - ignoring");
            }
            catch (Exception ex)
            {
                // Log other exceptions but don't crash
                System.Diagnostics.Debug.WriteLine($"Exception in OnMouseMove: {ex.Message}");
            }
        }

        private void SafeInvalidate()
        {
            try
            {
                if (_documentPictureBox != null && !_documentPictureBox.IsDisposed && _documentPictureBox.IsHandleCreated)
                {
                    _documentPictureBox.Invalidate();
                }
            }
            catch (Exception)
            {
                // Ignore invalidate errors
            }
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                // Check if the PictureBox is valid and not disposed
                if (_documentPictureBox == null || _documentPictureBox.IsDisposed)
                    return;

                if (_draggingColumnIndex >= 0)
                {
                    _draggingColumnIndex = -1;
                    try
                    {
                        _documentPictureBox.Cursor = Cursors.Default;
                    }
                    catch (Exception)
                    {
                        // Ignore cursor setting errors
                    }
                    return;
                }

                if (!_isSelecting || e.Button != MouseButtons.Left) return;

                _isSelecting = false;
                
                if (_currentSelection.Width > 5 && _currentSelection.Height > 5)
                {
                    if (_currentSnipMode == SnipMode.Table)
                    {
                        _adjustingTable = true;
                        _showTableGrid = true;
                        
                        // Ensure we have at least one column divider to start with
                        if (_tableColumns.Count == 0)
                        {
                            int centerX = _currentSelection.X + _currentSelection.Width / 2;
                            _tableColumns.Add(new System.Drawing.Rectangle(centerX - 2, _currentSelection.Y, 4, _currentSelection.Height));
                            Logger.Info($"Added initial column divider at X={centerX}");
                        }
                        
                        if (_statusLabel != null && !_statusLabel.IsDisposed)
                            _statusLabel.Text = $"Table mode: {_tableColumns.Count} column dividers. Click + to add, - to remove, or double-click to extract";
                        
                        SafeInvalidate();
                    }
                    else
                    {
                        ProcessSnip();
                    }
                }
            }
            catch (System.Runtime.InteropServices.SEHException)
            {
                // Handle SEH exceptions from unmanaged code
                System.Diagnostics.Debug.WriteLine("SEH Exception in OnMouseUp - ignoring");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in OnMouseUp: {ex.Message}");
            }
        }

        private void OnPictureDoubleClick(object sender, EventArgs e)
        {
            if (_adjustingTable)
            {
                _adjustingTable = false;
                _showTableGrid = false;
                ProcessSnip();
            }
        }

        private void ProcessSnip()
        {
            try
            {
                if (_currentDocument == null || _currentPageIndex >= _currentDocument.PageCount)
                {
                    _statusLabel.Text = "No document loaded for snipping";
                    return;
                }

                _statusLabel.Text = "Processing snip - extracting data...";
                
                // Get the selected area from the current displayed image
                var displayedImage = _documentPictureBox.Image as Bitmap;
                if (displayedImage == null)
                {
                    _statusLabel.Text = "No image available for snipping";
                    return;
                }
                
                // Ensure we have a valid selection rectangle
                if (_currentSelection.Width <= 5 || _currentSelection.Height <= 5)
                {
                    _statusLabel.Text = "Selection area too small - please draw a larger rectangle";
                    return;
                }

                // Convert the selection rectangle from zoomed coordinates to original page coordinates
                System.Drawing.Rectangle selectionOnOriginal = new System.Drawing.Rectangle(
                    (int)(_currentSelection.X / _zoomFactor),
                    (int)(_currentSelection.Y / _zoomFactor),
                    (int)(_currentSelection.Width / _zoomFactor),
                    (int)(_currentSelection.Height / _zoomFactor)
                );

                string extractedText = "";
                string[] extractedNumbers = new string[0];
                bool success = false;

                // Try PDF text extraction first if this is a PDF
                if (_currentDocument.Type == DocumentType.PDF)
                {
                    Logger.Info($"Attempting PDF text extraction for selection: {selectionOnOriginal}");
                    extractedText = ExtractTextFromPdfRegion(_currentDocument.FilePath, _currentPageIndex, selectionOnOriginal);
                    if (!string.IsNullOrWhiteSpace(extractedText))
                    {
                        success = true;
                        Logger.Info($"PDF text extraction successful: '{extractedText.Substring(0, Math.Min(extractedText.Length, 50))}...'");
                        
                        // Parse numbers from extracted text for Sum mode
                        if (_currentSnipMode == SnipMode.Sum)
                        {
                            var numberMatches = System.Text.RegularExpressions.Regex.Matches(extractedText, @"-?\d[\d,\.]*");
                            extractedNumbers = numberMatches.Cast<System.Text.RegularExpressions.Match>()
                                .Select(m => m.Value.Replace(",", ""))
                                .Where(n => decimal.TryParse(n, out _))
                                .ToArray();
                        }
                    }
                    else
                    {
                        Logger.Info("PDF text extraction returned empty - will try OCR fallback");
                    }
                }

                // OCR fallback if PDF text extraction failed or if document is an image
                if (!success)
                {
                    Logger.Info("Using OCR fallback for text extraction");
                    
                    // Crop from the ORIGINAL page image (not the zoomed displayed image) for better OCR quality
                    Bitmap pageImage = _currentDocument.Pages[_currentPageIndex];
                    using (var croppedImage = CropImageFromDisplayed(pageImage, selectionOnOriginal))
                    {
                        if (croppedImage == null)
                        {
                            _statusLabel.Text = "Failed to crop selected area";
                            return;
                        }
                        
                        var initResult = _ocrEngine.Initialize();
                        if (!initResult)
                        {
                            _statusLabel.Text = "OCR engine failed to initialize";
                            return;
                        }

                        var ocrResult = _ocrEngine.RecognizeText(croppedImage);
                        if (ocrResult.Success)
                        {
                            extractedText = ocrResult.Text;
                            extractedNumbers = ocrResult.Numbers ?? new string[0];
                            success = true;
                            Logger.Info($"OCR extraction successful: '{extractedText.Substring(0, Math.Min(extractedText.Length, 50))}...'");
                        }
                        else
                        {
                            Logger.Info("OCR extraction failed");
                        }
                    }
                }

                // Process the extracted text based on snip mode
                string extractedValue = "";
                if (success)
                {
                    switch (_currentSnipMode)
                    {
                        case SnipMode.Text:
                            extractedValue = extractedText.Trim();
                            break;
                        case SnipMode.Sum:
                            if (extractedNumbers.Length > 0)
                            {
                                var sum = extractedNumbers
                                    .Where(n => decimal.TryParse(n.Replace("$", "").Replace(",", ""), out _))
                                    .Sum(n => decimal.Parse(n.Replace("$", "").Replace(",", "")));
                                extractedValue = sum.ToString("N2");
                            }
                            else
                            {
                                // Try to find numbers in the text if extraction didn't parse them
                                var numberMatches = System.Text.RegularExpressions.Regex.Matches(extractedText, @"-?\d+(?:[,\.]\d+)*");
                                if (numberMatches.Count > 0)
                                {
                                    var sum = numberMatches.Cast<System.Text.RegularExpressions.Match>()
                                        .Select(m => m.Value.Replace(",", ""))
                                        .Where(n => decimal.TryParse(n, out _))
                                        .Sum(n => decimal.Parse(n));
                                    extractedValue = sum.ToString("N2");
                                }
                                else
                                {
                                    extractedValue = "No numbers found";
                                    success = false;
                                }
                            }
                            break;
                        case SnipMode.Table:
                            // DataSnipper approach: Extract text from each column separately for better accuracy
                            if (_tableColumns.Count > 0)
                            {
                                // Scale the column dividers to match the original page coordinates
                                var scaledTableColumns = new List<System.Drawing.Rectangle>();
                                foreach (var col in _tableColumns)
                                {
                                    var scaledCol = new System.Drawing.Rectangle(
                                        (int)(col.X / _zoomFactor),
                                        (int)(col.Y / _zoomFactor),
                                        (int)(col.Width / _zoomFactor),
                                        (int)(col.Height / _zoomFactor)
                                    );
                                    scaledTableColumns.Add(scaledCol);
                                }
                                
                                extractedValue = ExtractTableDataByColumns(_currentDocument, _currentPageIndex, selectionOnOriginal, scaledTableColumns);
                                success = !string.IsNullOrWhiteSpace(extractedValue);
                                Logger.Info($"Table extraction by columns: {(success ? "SUCCESS" : "FAILED")}");
                                if (success)
                                {
                                    Logger.Info($"Table output preview: '{extractedValue.Substring(0, Math.Min(extractedValue.Length, 100))}...'");
                                    Logger.Info($"Total table output length: {extractedValue.Length} characters");
                                    var lines = extractedValue.Split('\n');
                                    Logger.Info($"Table has {lines.Length} rows");
                                    if (lines.Length > 0)
                                    {
                                        var firstRowColumns = lines[0].Split('\t');
                                        Logger.Info($"First row has {firstRowColumns.Length} columns: [{string.Join(", ", firstRowColumns.Take(5))}]");
                                    }
                                }
                                else
                                {
                                    Logger.Info("Table extraction failed - no data extracted");
                                }
                            }
                            else
                            {
                                // No column dividers - treat as single column table
                                extractedValue = extractedText.Trim();
                                success = !string.IsNullOrWhiteSpace(extractedValue);
                            }
                            break;
                        case SnipMode.Validation:
                            extractedValue = "✓ VERIFIED";
                            success = true;
                            break;
                        case SnipMode.Exception:
                            extractedValue = "✗ EXCEPTION";
                            success = true;
                            break;
                        default:
                            extractedValue = extractedText.Trim();
                            break;
                    }
                }
                else
                {
                    extractedValue = _currentSnipMode == SnipMode.Validation ? "✓ VERIFIED" :
                                   _currentSnipMode == SnipMode.Exception ? "✗ EXCEPTION" : 
                                   "EXTRACTION_FAILED";
                    success = _currentSnipMode == SnipMode.Validation || _currentSnipMode == SnipMode.Exception;
                }
                
                // Create the event args with real data
                var args = new SnipAreaSelectedEventArgs
                {
                    SnipMode = _currentSnipMode,
                    DocumentPath = _currentDocument.FilePath,
                    PageNumber = _currentPageIndex + 1,
                    Bounds = selectionOnOriginal, // Use original coordinates for reference
                    SelectedImage = CropImageFromDisplayed(displayedImage, _currentSelection), // Cropped from display for visual reference
                    ExtractedText = extractedValue,
                    ExtractedNumbers = extractedNumbers,
                    Success = success
                };

                // Fire the event to send data to Excel
                SnipAreaSelected?.Invoke(this, args);
                
                // Visual feedback - add permanent highlight
                AddPermanentHighlight(_currentSelection, GetSnipColor(_currentSnipMode));
                
                // Update status with more specific feedback
                if (args.Success)
                {
                    if (_currentSnipMode == SnipMode.Table)
                    {
                        var lines = extractedValue.Split('\n');
                        var columns = lines.Length > 0 ? lines[0].Split('\t').Length : 0;
                        _statusLabel.Text = $"✓ Table snip completed: {lines.Length} rows × {columns} columns extracted to Excel";
                    }
                    else
                {
                    var preview = extractedValue.Length > 30 ? extractedValue.Substring(0, 30) + "..." : extractedValue;
                    _statusLabel.Text = $"✓ {_currentSnipMode} snip completed: {preview}";
                    }
                }
                else
                {
                    _statusLabel.Text = $"✗ {_currentSnipMode} snip failed - could not extract text from selection";
                }
                
                // Reset selection and table state
                _currentSelection = System.Drawing.Rectangle.Empty;
                _adjustingTable = false;
                _showTableGrid = false;
                _tableColumns.Clear();
                _documentPictureBox.Invalidate();
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing snip: {ex.Message}", ex);
                _statusLabel.Text = $"Error processing snip: {ex.Message}";
            }
        }

        private Bitmap CropImageFromDisplayed(Bitmap sourceImage, System.Drawing.Rectangle cropArea)
        {
            try
            {
                // Ensure crop area is within bounds
                var validCropArea = System.Drawing.Rectangle.Intersect(
                    cropArea, 
                    new System.Drawing.Rectangle(0, 0, sourceImage.Width, sourceImage.Height)
                );
                
                if (validCropArea.Width <= 0 || validCropArea.Height <= 0)
                    return null;
                
                var croppedImage = new Bitmap(validCropArea.Width, validCropArea.Height);
                using (var g = Graphics.FromImage(croppedImage))
                {
                    g.DrawImage(sourceImage, 
                        new System.Drawing.Rectangle(0, 0, validCropArea.Width, validCropArea.Height), 
                        validCropArea, 
                        GraphicsUnit.Pixel);
                }
                return croppedImage;
            }
            catch
            {
                return null;
            }
        }

        private void AddPermanentHighlight(System.Drawing.Rectangle region, Color color)
        {
            // Add a permanent highlight to the displayed image to show processed areas
            if (_documentPictureBox.Image != null)
            {
                var image = _documentPictureBox.Image as Bitmap;
                if (image != null)
                {
                    using (var g = Graphics.FromImage(image))
                    {
                        // Draw highlight border
                        using (var pen = new Pen(color, 3))
                        {
                            g.DrawRectangle(pen, region);
                        }
                        
                        // Draw semi-transparent overlay
                        using (var brush = new SolidBrush(Color.FromArgb(30, color)))
                        {
                            g.FillRectangle(brush, region);
                        }
                    }
                    _documentPictureBox.Invalidate();
                }
            }
        }

        private void DetectTableStructure(System.Drawing.Rectangle selection)
        {
            // Create adjustable column dividers for table snipping
            _tableColumns.Clear();
            _tableRows.Clear();
            
            if (selection.Width < 50 || selection.Height < 20)
                return;
            
            // Start with 2 initial column dividers for a 3-column table (more practical than 4)
            int initialColumns = 3;
            int spacing = selection.Width / initialColumns;
            
            for (int i = 1; i < initialColumns; i++)
            {
                int x = selection.X + (i * spacing);
                _tableColumns.Add(new System.Drawing.Rectangle(x - 2, selection.Y, 4, selection.Height));
            }
            
            _showTableGrid = true;
            Logger.Info($"Created initial table structure with {_tableColumns.Count} column dividers");
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            // Draw current selection
            if (!_currentSelection.IsEmpty)
            {
                var color = GetSnipColor(_currentSnipMode);
                using (var pen = new Pen(color, 2))
                {
                    e.Graphics.DrawRectangle(pen, _currentSelection);
                }
                
                // Draw semi-transparent overlay
                using (var brush = new SolidBrush(Color.FromArgb(50, color)))
                {
                    e.Graphics.FillRectangle(brush, _currentSelection);
                }
            }
            
            // Draw table grid helpers with DataSnipper-style controls
            if (_showTableGrid && _currentSnipMode == SnipMode.Table && !_currentSelection.IsEmpty)
            {
                // Draw column divider lines
                using (var pen = new Pen(Color.Blue, 2))
                {
                    foreach (var column in _tableColumns)
                    {
                        int centerX = column.X + column.Width / 2;
                        e.Graphics.DrawLine(pen, centerX, column.Y, centerX, column.Y + column.Height);
                    }
                }

                // Draw plus/minus icons for DataSnipper-style column adjustment
                if (_adjustingTable)
                {
                    int iconOffset = 20;
                    int iconSize = 16;
                    int centerY = _currentSelection.Y - iconOffset;
                    
                    using (var iconPen = new Pen(Color.DarkBlue, 2))
                    using (var bgBrush = new SolidBrush(Color.LightYellow))
                    using (var borderPen = new Pen(Color.DarkBlue, 1))
                    using (var textBrush = new SolidBrush(Color.DarkBlue))
                    using (var font = new Font("Arial", 10, FontStyle.Bold))
                    {
                        // Draw "–" (minus) button for each existing column divider
                        foreach (var col in _tableColumns)
                        {
                            int centerX = col.X + col.Width / 2;
                            var buttonRect = new System.Drawing.Rectangle(centerX - iconSize/2, centerY - iconSize/2, iconSize, iconSize);
                            
                            // Draw button background
                            e.Graphics.FillRectangle(bgBrush, buttonRect);
                            e.Graphics.DrawRectangle(borderPen, buttonRect);
                            
                            // Draw minus sign
                            var minusRect = new RectangleF(buttonRect.X + 2, buttonRect.Y + iconSize/2 - 1, buttonRect.Width - 4, 2);
                            e.Graphics.FillRectangle(textBrush, minusRect);
                        }
                        
                        // Draw "+" button for each gap between column dividers
                        var boundaries = new List<int> { _currentSelection.X };
                        boundaries.AddRange(_tableColumns.Select(c => c.X + c.Width / 2));
                        boundaries.Add(_currentSelection.Right);
                        boundaries.Sort();
                        
                        for (int i = 0; i < boundaries.Count - 1; i++)
                        {
                            int gapLeft = boundaries[i];
                            int gapRight = boundaries[i + 1];
                            if (gapRight - gapLeft > 40) // Only show + button if gap is wide enough
                            {
                                int centerX = gapLeft + (gapRight - gapLeft) / 2;
                                var buttonRect = new System.Drawing.Rectangle(centerX - iconSize/2, centerY - iconSize/2, iconSize, iconSize);
                                
                                // Draw button background
                                e.Graphics.FillRectangle(bgBrush, buttonRect);
                                e.Graphics.DrawRectangle(borderPen, buttonRect);
                                
                                // Draw plus sign
                                var hLineRect = new RectangleF(buttonRect.X + 2, buttonRect.Y + iconSize/2 - 1, buttonRect.Width - 4, 2);
                                var vLineRect = new RectangleF(buttonRect.X + iconSize/2 - 1, buttonRect.Y + 2, 2, buttonRect.Height - 4);
                                e.Graphics.FillRectangle(textBrush, hLineRect);
                                e.Graphics.FillRectangle(textBrush, vLineRect);
                            }
                        }
                        
                        // Draw instruction text
                        string instruction = "Click + to add columns, - to remove columns, or double-click to extract table";
                        var textRect = new RectangleF(_currentSelection.X, _currentSelection.Bottom + 5, _currentSelection.Width, 30);
                        using (var instructionFont = new Font("Arial", 9))
                        using (var textBgBrush = new SolidBrush(Color.FromArgb(200, Color.White)))
                        {
                            e.Graphics.FillRectangle(textBgBrush, textRect);
                            e.Graphics.DrawString(instruction, instructionFont, textBrush, textRect, 
                                new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                        }
                    }
                }
            }
        }

        public void HighlightRegion(System.Drawing.Rectangle region, Color color)
        {
            // Add permanent highlight to show processed areas
            if (_documentPictureBox.Image != null)
            {
                using (var g = Graphics.FromImage(_documentPictureBox.Image))
                {
                    using (var pen = new Pen(color, 3))
                    {
                        g.DrawRectangle(pen, region);
                    }
                }
                _documentPictureBox.Invalidate();
            }
        }

        private Color GetSnipColor(SnipMode snipMode)
        {
            return snipMode switch
            {
                SnipMode.Text => Color.Blue,
                SnipMode.Sum => Color.Purple,
                SnipMode.Table => Color.Purple,
                SnipMode.Validation => Color.Green,
                SnipMode.Exception => Color.Red,
                _ => Color.Gray
            };
        }

        private System.Drawing.Rectangle GetNormalizedRectangle(Point start, Point end)
        {
            return new System.Drawing.Rectangle(
                Math.Min(start.X, end.X),
                Math.Min(start.Y, end.Y),
                Math.Abs(end.X - start.X),
                Math.Abs(end.Y - start.Y)
            );
        }

        private void OnPreviousPage(object sender, EventArgs e)
        {
            if (_currentDocument != null && _currentPageIndex > 0)
            {
                _currentPageIndex--;
                DisplayCurrentPage();
            }
        }

        private void OnNextPage(object sender, EventArgs e)
        {
            if (_currentDocument != null && _currentPageIndex < _currentDocument.PageCount - 1)
            {
                _currentPageIndex++;
                DisplayCurrentPage();
            }
        }

        private void OnZoomIn(object sender, EventArgs e)
        {
            _zoomFactor = Math.Min(_zoomFactor * 1.25f, 3.0f);
            DisplayCurrentPage();
        }

        private void OnZoomOut(object sender, EventArgs e)
        {
            _zoomFactor = Math.Max(_zoomFactor / 1.25f, 0.25f);
            DisplayCurrentPage();
        }

        private void OnFitToWidth(object sender, EventArgs e)
        {
            if (_currentDocument != null && _currentPageIndex < _currentDocument.PageCount)
            {
                var page = _currentDocument.Pages[_currentPageIndex];
                _zoomFactor = (float)(_viewerPanel.Width - 40) / page.Width;
                DisplayCurrentPage();
            }
        }

        private void UpdateDocumentsList()
        {
            // Update the documents panel with loaded documents
            foreach (Control control in _documentsPanel.Controls.OfType<Button>().ToList())
            {
                _documentsPanel.Controls.Remove(control);
                control.Dispose();
            }

            int y = 30;
            foreach (var doc in _loadedDocuments)
            {
                var button = new Button
                {
                    Text = doc.Name,
                    Size = new Size(180, 30),
                    Location = new Point(10, y),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                
                var document = doc; // Capture for closure
                button.Click += (s, e) => 
                {
                    _currentDocument = document;
                    _currentPageIndex = 0;
                    _documentSelector.SelectedIndex = _loadedDocuments.IndexOf(document);
                    DisplayCurrentPage();
                };
                
                _documentsPanel.Controls.Add(button);
                y += 35;
            }
        }

        public bool LoadDocument(string filePath)
        {
            Task.Run(async () => await LoadDocuments(new[] { filePath }));
            return true;
        }
        
        public void NavigateToPage(int pageNumber)
        {
            if (_currentDocument != null && pageNumber > 0 && pageNumber <= _currentDocument.PageCount)
            {
                _currentPageIndex = pageNumber - 1; // Convert to 0-based index
                DisplayCurrentPage();
                Logger.Info($"Navigated to page {pageNumber}");
            }
            else
            {
                Logger.Warning($"Cannot navigate to page {pageNumber} - invalid page number or no document loaded");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (var doc in _loadedDocuments)
                {
                    doc.Dispose();
                }
                _loadedDocuments.Clear();
                _ocrEngine?.Dispose();

                Logger.Info("DocumentViewer disposed");
            }
            base.Dispose(disposing);
        }

        private LoadedDocument LoadImageDocument(string filePath)
        {
            try
            {
                var image = new Bitmap(filePath);
                return new LoadedDocument
                {
                    FilePath = filePath,
                    Name = Path.GetFileName(filePath),
                    Type = DocumentType.Image,
                    Pages = new List<Bitmap> { image },
                    PageCount = 1
                };
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading image {filePath}: {ex.Message}", ex);
                return null;
            }
        }

        private int EstimatePdfPages(byte[] pdfBytes)
        {
            try
            {
                var content = System.Text.Encoding.ASCII.GetString(pdfBytes);
                var pageMatches = System.Text.RegularExpressions.Regex.Matches(content, @"/Type\s*/Page[^s]");
                return Math.Max(1, pageMatches.Count);
            }
            catch
            {
                return 1;
            }
        }

        private string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024:F1} KB";
            return $"{bytes / (1024 * 1024):F1} MB";
        }

        private Bitmap CreatePdfRepresentation(string filePath)
        {
            // This is now just a fallback method
            return CreateAdvancedPdfRepresentation(filePath);
        }

        private string ExtractTextFromPdfBytes(byte[] pdfBytes)
        {
            // Legacy method - now calls the real extraction
            return ExtractRealTextFromPdf(pdfBytes);
        }

        // Extract text from a specified rectangle on a PDF page using PDFium's native text extraction
        private string ExtractTextFromPdfRegion(string pdfPath, int pageIndex, System.Drawing.Rectangle rectOnImage)
        {
            IntPtr doc = IntPtr.Zero;
            IntPtr page = IntPtr.Zero;
            IntPtr textPage = IntPtr.Zero;
            try
            {
                Logger.Info($"Extracting text from PDF region: page {pageIndex + 1}, rect {rectOnImage}");
                
                // Open PDF document (no password assumed)
                doc = FPDF_LoadDocument(pdfPath, null);
                if (doc == IntPtr.Zero)
                {
                    Logger.Error("Failed to load PDF document for text extraction");
                    return string.Empty;
                }
                
                // Load the specific page
                page = FPDF_LoadPage(doc, pageIndex);
                if (page == IntPtr.Zero)
                {
                    Logger.Error($"Failed to load PDF page {pageIndex}");
                    return string.Empty;
                }
                
                // Load text page (prepare for text extraction)
                textPage = FPDFText_LoadPage(page);
                if (textPage == IntPtr.Zero)
                {
                    Logger.Error($"Failed to load text page {pageIndex}");
                    return string.Empty;
                }

                // Convert the rectangle from image pixel coords to PDF page coordinates (points)
                // We know the original page image was rendered at 150 DPI (PDF_RENDER_DPI constant)
                // So 1 PDF point = 1/72 inch. At 150 DPI, 1 inch = 150 px, so 1 point = 150/72 px ≈ 2.0833 px.
                // Thus scale factor from image pixels to PDF points = (72 / 150).
                
                // Get the current page's image to determine scaling
                if (_currentDocument == null || _currentPageIndex >= _currentDocument.PageCount)
                {
                    Logger.Error("No current document or invalid page index");
                    return string.Empty;
                }
                
                var pageImage = _currentDocument.Pages[pageIndex];
                double pageHeightPoints = pageImage.Height * (72.0 / PDF_RENDER_DPI);
                double pageWidthPoints = pageImage.Width * (72.0 / PDF_RENDER_DPI);
                
                // Convert rectangle coordinates
                double left = rectOnImage.X * (72.0 / PDF_RENDER_DPI);
                double right = (rectOnImage.X + rectOnImage.Width) * (72.0 / PDF_RENDER_DPI);
                // PDF coordinate origin is bottom-left, whereas image origin is top-left
                double top = pageHeightPoints - rectOnImage.Y * (72.0 / PDF_RENDER_DPI);
                double bottom = pageHeightPoints - (rectOnImage.Y + rectOnImage.Height) * (72.0 / PDF_RENDER_DPI);
                
                // Ensure bounds are within the page
                if (left < 0) left = 0;
                if (bottom < 0) bottom = 0;
                if (right > pageWidthPoints) right = pageWidthPoints;
                if (top > pageHeightPoints) top = pageHeightPoints;
                
                Logger.Info($"PDF coordinates: left={left:F2}, top={top:F2}, right={right:F2}, bottom={bottom:F2}");

                // Call PDFium to get text in the rectangle (UTF-16LE output)
                // First call with buffer length 0 to get required characters count
                int charCount = FPDFText_GetBoundedText(textPage, left, top, right, bottom, null, 0);
                if (charCount <= 0)
                {
                    Logger.Info("No text found in the specified region");
                    return string.Empty;
                }
                
                // Allocate buffer and get the actual text
                ushort[] buffer = new ushort[charCount];
                int actualCount = FPDFText_GetBoundedText(textPage, left, top, right, bottom, buffer, charCount);
                
                if (actualCount <= 0)
                {
                    Logger.Info("Failed to retrieve text from PDF");
                    return string.Empty;
                }
                
                // Convert UTF-16LE buffer to .NET string (charCount includes the terminating null)
                int actualChars = actualCount;
                if (actualChars > 0 && buffer[actualChars - 1] == 0) 
                    actualChars -= 1;  // remove terminator if present
                
                string text = new string(buffer.Take(actualChars).Select(ch => (char)ch).ToArray());
                Logger.Info($"Extracted text: '{text}' ({actualChars} characters)");
                return text;
            }
            catch (Exception ex)
            {
                Logger.Error($"FPDFText_GetBoundedText failed: {ex.Message}", ex);
                return string.Empty;
            }
            finally
            {
                if (textPage != IntPtr.Zero) FPDFText_ClosePage(textPage);
                if (page != IntPtr.Zero) FPDF_ClosePage(page);
                if (doc != IntPtr.Zero) FPDF_CloseDocument(doc);
            }
        }

        // Extract table data by extracting text from each column separately (DataSnipper approach)
        private string ExtractTableDataByColumns(LoadedDocument document, int pageIndex, System.Drawing.Rectangle tableArea, List<System.Drawing.Rectangle> columnDividers)
        {
            try
            {
            var sortedDividers = columnDividers.OrderBy(c => c.X).ToList();
                var columnTexts = new List<List<string>>();
                
                // Define column boundaries (left edge, divider positions, right edge)
                var columnBoundaries = new List<int> { tableArea.X };
                columnBoundaries.AddRange(sortedDividers.Select(d => d.X + d.Width / 2));
                columnBoundaries.Add(tableArea.Right);
                
                // Ensure boundaries are within table bounds and sorted
                columnBoundaries = columnBoundaries
                    .Where(x => x >= tableArea.X && x <= tableArea.Right)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();
                
                Logger.Info($"Extracting table with {columnBoundaries.Count - 1} columns");
                
                // Extract text from each column
                for (int col = 0; col < columnBoundaries.Count - 1; col++)
                {
                    var columnLeft = columnBoundaries[col];
                    var columnRight = columnBoundaries[col + 1];
                    var columnWidth = columnRight - columnLeft;
                    
                    if (columnWidth > 5) // Minimum column width
                {
                    var columnRect = new System.Drawing.Rectangle(
                            columnLeft, tableArea.Y, columnWidth, tableArea.Height);
                        
                        string columnText = "";
                        
                        // Try PDF text extraction first
                        if (document.Type == DocumentType.PDF)
                        {
                            columnText = ExtractTextFromPdfRegion(document.FilePath, pageIndex, columnRect);
                        }
                        
                        // OCR fallback if PDF extraction failed
                        if (string.IsNullOrWhiteSpace(columnText))
                        {
                            var pageImage = document.Pages[pageIndex];
                            using (var columnImage = CropImageFromDisplayed(pageImage, columnRect))
                            {
                                if (columnImage != null && _ocrEngine.Initialize())
                        {
                            var ocrResult = _ocrEngine.RecognizeText(columnImage);
                                    if (ocrResult.Success)
                                    {
                                        columnText = ocrResult.Text;
                                        Logger.Info($"Column {col + 1}: OCR fallback used, extracted '{columnText.Substring(0, Math.Min(columnText.Length, 50))}...'");
                                    }
                                    else
                                    {
                                        Logger.Info($"Column {col + 1}: Both PDF extraction and OCR failed");
                                    }
                                }
                                else
                                {
                                    Logger.Info($"Column {col + 1}: Could not initialize OCR engine");
                                }
                            }
                        }
                        else
                        {
                            Logger.Info($"Column {col + 1}: PDF extraction successful, extracted '{columnText.Substring(0, Math.Min(columnText.Length, 50))}...'");
                        }
                        
                        // Split column text into lines and clean up
                        var columnLines = columnText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(line => line.Trim())
                                                    .Where(line => !string.IsNullOrWhiteSpace(line))
                                                    .ToList();
                        
                        columnTexts.Add(columnLines);
                        Logger.Info($"Column {col + 1}: {columnLines.Count} lines extracted");
                    }
                    else
                    {
                        columnTexts.Add(new List<string>());
                    }
                }
                
                // Combine columns into tab-delimited rows
                if (columnTexts.Count > 0)
                {
                    var result = CombineColumnsIntoTabDelimitedRows(columnTexts);
                    if (!string.IsNullOrWhiteSpace(result))
                    {
                        return result;
                    }
                    else
                    {
                        Logger.Info("Column-based extraction produced empty result, falling back to full area extraction");
                    }
                }
                
                // Fallback: extract entire table area as one piece and try to format it
                Logger.Info("Falling back to full table area extraction");
                string fallbackText = "";
                
                if (document.Type == DocumentType.PDF)
                {
                    fallbackText = ExtractTextFromPdfRegion(document.FilePath, pageIndex, tableArea);
                }
                
                if (string.IsNullOrWhiteSpace(fallbackText))
                {
                    // OCR fallback for entire table
                    var pageImage = document.Pages[pageIndex];
                    using (var tableImage = CropImageFromDisplayed(pageImage, tableArea))
                    {
                        if (tableImage != null && _ocrEngine.Initialize())
                        {
                            var ocrResult = _ocrEngine.RecognizeText(tableImage);
                            if (ocrResult.Success)
                            {
                                fallbackText = ocrResult.Text;
                            }
                        }
                    }
                }
                
                if (!string.IsNullOrWhiteSpace(fallbackText))
                {
                    // Try to format the fallback text with approximate column positions
                    return FormatTextWithColumnDividers(fallbackText, columnDividers, tableArea);
                }
                
                return "";
            }
            catch (Exception ex)
            {
                Logger.Error($"Error extracting table by columns: {ex.Message}", ex);
                return "";
            }
        }
        
        // Combine multiple column text lists into tab-delimited rows for Excel
        private string CombineColumnsIntoTabDelimitedRows(List<List<string>> columnTexts)
        {
            try
            {
                if (columnTexts.Count == 0) return "";
                
                // Find the maximum number of lines in any column
                int maxLines = columnTexts.Max(col => col.Count);
                var resultRows = new List<string>();
                
                for (int row = 0; row < maxLines; row++)
                {
                    var rowCells = new List<string>();
                    
                    foreach (var columnLines in columnTexts)
                    {
                        // Get the text for this row from this column (or empty if column has fewer lines)
                        string cellText = row < columnLines.Count ? columnLines[row] : "";
                        rowCells.Add(cellText);
                    }
                    
                    // Join cells with tabs for Excel compatibility
                    string tabDelimitedRow = string.Join("\t", rowCells);
                    
                    // Only add non-empty rows (but allow rows with at least one non-empty cell)
                    if (rowCells.Any(cell => !string.IsNullOrWhiteSpace(cell)))
                    {
                        resultRows.Add(tabDelimitedRow);
                    }
                }
                
                string result = string.Join("\n", resultRows);
                Logger.Info($"Created tab-delimited table with {resultRows.Count} rows and {columnTexts.Count} columns");
                if (result.Length > 0)
                {
                    Logger.Info($"Sample output: '{result.Substring(0, Math.Min(result.Length, 100))}...'");
                    Logger.Info($"Contains tabs: {result.Contains('\t')}");
                    var sampleLines = result.Split('\n').Take(3);
                    foreach (var line in sampleLines)
                    {
                        var cells = line.Split('\t');
                        Logger.Info($"Row with {cells.Length} cells: [{string.Join("] [", cells)}]");
                    }
                }
                else
                {
                    Logger.Info("No result generated from column combination");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                Logger.Error($"Error combining columns into rows: {ex.Message}", ex);
                return "";
            }
                }
        
        // Format text using column divider positions as guidelines
        private string FormatTextWithColumnDividers(string text, List<System.Drawing.Rectangle> columnDividers, System.Drawing.Rectangle tableArea)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(text) || columnDividers.Count == 0)
                    return text;
                
                var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                var formattedLines = new List<string>();
                
                // Calculate column boundaries as relative positions (0.0 to 1.0)
                var columnPositions = new List<double>();
                foreach (var divider in columnDividers.OrderBy(d => d.X))
                {
                    double relativePos = (double)(divider.X - tableArea.X) / tableArea.Width;
                    columnPositions.Add(relativePos);
                }
                
                Logger.Info($"Formatting text with {columnPositions.Count} column positions: [{string.Join(", ", columnPositions.Select(p => p.ToString("F2")))}]");
                
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    
                    // Try intelligent column splitting based on divider positions
                    var columns = SplitLineByColumnPositions(line.Trim(), columnPositions);
                    var tabLine = string.Join("\t", columns);
                    formattedLines.Add(tabLine);
                }
                
                return string.Join("\n", formattedLines);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error formatting text with column dividers: {ex.Message}", ex);
                return text; // Return original text if formatting fails
            }
        }
        
        // Split a line into columns based on relative positions
        private string[] SplitLineByColumnPositions(string line, List<double> columnPositions)
        {
            var result = new List<string>();
            int lastPos = 0;
            
            foreach (var position in columnPositions)
            {
                int charPos = (int)(position * line.Length);
                charPos = Math.Min(charPos, line.Length);
                charPos = Math.Max(charPos, lastPos);
                
                // Find the best split point near this position (look for spaces)
                int splitPos = FindBestSplitPoint(line, charPos, lastPos);
                
                if (splitPos > lastPos)
                {
                    result.Add(line.Substring(lastPos, splitPos - lastPos).Trim());
                    lastPos = splitPos;
                }
            }
            
            // Add the remaining part
            if (lastPos < line.Length)
            {
                result.Add(line.Substring(lastPos).Trim());
            }
            
            // Ensure we have at least one column
            if (result.Count == 0)
                result.Add(line);
            
            return result.ToArray();
        }

        // PInvoke declarations for PDFium functions
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LoadLibrary(string dllToLoad);
        
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);
        
        static DocumentViewer()
        {
            // Try to set DLL directory to the current application directory
            try
            {
                var appDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                SetDllDirectory(appDir);
                
                // Try to preload the PDFium library
                var pdfiumPath = System.IO.Path.Combine(appDir, "pdfium.dll");
                if (System.IO.File.Exists(pdfiumPath))
                {
                    LoadLibrary(pdfiumPath);
                    Logger.Info($"Successfully preloaded pdfium.dll from {pdfiumPath}");
                }
                else
                {
                    Logger.Info($"pdfium.dll not found at {pdfiumPath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to preload pdfium.dll: {ex.Message}", ex);
            }
        }
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", CharSet = System.Runtime.InteropServices.CharSet.Ansi, EntryPoint = "FPDF_LoadDocument")]
        private static extern IntPtr FPDF_LoadDocument(string filePath, string password);
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", EntryPoint = "FPDF_CloseDocument")]
        private static extern void FPDF_CloseDocument(IntPtr document);
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", EntryPoint = "FPDF_LoadPage")]
        private static extern IntPtr FPDF_LoadPage(IntPtr document, int pageIndex);
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", EntryPoint = "FPDF_ClosePage")]
        private static extern void FPDF_ClosePage(IntPtr page);
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", EntryPoint = "FPDFText_LoadPage")]
        private static extern IntPtr FPDFText_LoadPage(IntPtr page);
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", EntryPoint = "FPDFText_ClosePage")]
        private static extern void FPDFText_ClosePage(IntPtr textPage);
        
        [System.Runtime.InteropServices.DllImport("pdfium.dll", EntryPoint = "FPDFText_GetBoundedText")]
        private static extern int FPDFText_GetBoundedText(IntPtr textPage, double left, double top, double right, double bottom,
                                                          ushort[] buffer, int bufferLen);

        private string[] SplitLineIntoColumns(string line, int targetColumns)
        {
            if (string.IsNullOrWhiteSpace(line) || targetColumns <= 1)
                return new[] { line };

            try
            {
                // Try different splitting strategies in order of preference
                
                // Strategy 1: Split on 3+ spaces (wide gaps) 
                var wideGaps = System.Text.RegularExpressions.Regex.Split(line, @"\s{3,}");
                if (wideGaps.Length == targetColumns)
                {
                    return wideGaps.Select(s => s.Trim()).ToArray();
                }
                
                // Strategy 2: Split on 2+ spaces
                var mediumGaps = System.Text.RegularExpressions.Regex.Split(line, @"\s{2,}");
                if (mediumGaps.Length == targetColumns)
                {
                    return mediumGaps.Select(s => s.Trim()).ToArray();
                }
                
                // Strategy 3: Intelligent splitting based on patterns (numbers, currency, etc.)
                // This handles cases like "Name Surname  Province  R123,456.00"
                var intelligentSplit = IntelligentColumnSplit(line, targetColumns);
                if (intelligentSplit.Length == targetColumns)
                {
                    return intelligentSplit;
                }
                
                // Strategy 4: Force split by word boundaries and group into targetColumns
                var words = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length >= targetColumns)
                {
                    var result = new string[targetColumns];
                    int wordsPerColumn = Math.Max(1, words.Length / targetColumns);
                    int extraWords = words.Length % targetColumns;
                    int wordIndex = 0;
                    
                    for (int col = 0; col < targetColumns; col++)
                    {
                        var takeCount = wordsPerColumn + (col < extraWords ? 1 : 0);
                        if (col == targetColumns - 1) // Last column gets all remaining words
                        {
                            takeCount = words.Length - wordIndex;
                        }
                        
                        result[col] = string.Join(" ", words.Skip(wordIndex).Take(takeCount));
                        wordIndex += takeCount;
                    }
                    return result;
                }
                
                // Fallback: Just put everything in first column
                var fallback = new string[targetColumns];
                fallback[0] = line;
                for (int i = 1; i < targetColumns; i++)
                {
                    fallback[i] = "";
                }
                return fallback;
            }
            catch
            {
                // Emergency fallback
                var emergency = new string[targetColumns];
                emergency[0] = line;
                for (int i = 1; i < targetColumns; i++)
                {
                    emergency[i] = "";
                }
                return emergency;
            }
        }

        private string[] IntelligentColumnSplit(string line, int targetColumns)
        {
            try
            {
                // Look for common patterns: amounts (R123,456.00), percentages, dates, etc.
                var patterns = new[]
                {
                    @"R\s*\d+[\d\s,\.]*",  // Currency amounts like "R123,456.00"
                    @"\d+[\d\s,\.]*%",     // Percentages
                    @"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", // Dates
                    @"\d+[\d\s,\.]+",      // Large numbers
                    @"[A-Z][a-z]+\s+[A-Z][a-z]+", // Names like "First Last"
                };
                
                var segments = new List<string>();
                var remaining = line;
                
                foreach (var pattern in patterns)
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(remaining, pattern);
                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        if (match.Success)
                        {
                            // Extract the part before the match
                            var before = remaining.Substring(0, match.Index).Trim();
                            if (!string.IsNullOrEmpty(before))
                            {
                                segments.Add(before);
                            }
                            
                            // Extract the match itself
                            segments.Add(match.Value.Trim());
                            
                            // Update remaining text
                            remaining = remaining.Substring(match.Index + match.Length).Trim();
                            break; // Process one match at a time
                        }
                    }
                    
                    if (segments.Count >= targetColumns)
                        break;
                }
                
                // Add any remaining text
                if (!string.IsNullOrEmpty(remaining))
                {
                    segments.Add(remaining);
                }
                
                // Adjust to target column count
                if (segments.Count == targetColumns)
                {
                    return segments.ToArray();
                }
                else if (segments.Count > targetColumns)
                {
                    // Combine excess segments into last column
                    var result = segments.Take(targetColumns - 1).ToList();
                    result.Add(string.Join(" ", segments.Skip(targetColumns - 1)));
                    return result.ToArray();
                }
                else
                {
                    // Pad with empty strings
                    while (segments.Count < targetColumns)
                    {
                        segments.Add("");
                    }
                    return segments.ToArray();
                }
            }
            catch
            {
                // If intelligent parsing fails, fall back to simple space splitting
                return line.Split(new[] { ' ' }, targetColumns, StringSplitOptions.None);
            }
        }

        private string FormatLineWithTabs(string line, List<double> columnPositions)
        {
            if (string.IsNullOrWhiteSpace(line) || columnPositions.Count == 0)
                return line;

            try
            {
                // Remove extra whitespace and normalize
                line = System.Text.RegularExpressions.Regex.Replace(line.Trim(), @"\s+", " ");
                
                // If line is too short, just return it
                if (line.Length < 10)
                    return line;
                
                // Split on multiple spaces (likely column separators) and significant gaps
                var potentialColumns = System.Text.RegularExpressions.Regex.Split(line, @"  +").ToList();
                
                // If we have roughly the right number of columns (±1), use this split
                int expectedColumns = columnPositions.Count + 1;
                if (Math.Abs(potentialColumns.Count - expectedColumns) <= 1)
                {
                    // Clean up each column and join with tabs
                    var cleanedColumns = potentialColumns.Select(col => col.Trim()).ToList();
                    
                    // If we have fewer columns than expected, pad with empty strings
                    while (cleanedColumns.Count < expectedColumns)
                        cleanedColumns.Add("");
                    
                    // If we have more columns than expected, merge the last ones
                    if (cleanedColumns.Count > expectedColumns)
                    {
                        var lastColumns = cleanedColumns.Skip(expectedColumns - 1);
                        cleanedColumns = cleanedColumns.Take(expectedColumns - 1).ToList();
                        cleanedColumns.Add(string.Join(" ", lastColumns));
                    }
                    
                    return string.Join("\t", cleanedColumns);
                }
                
                // Fallback: try to split based on character positions
                var result = new List<string>();
                int lastPos = 0;
                
                foreach (var position in columnPositions)
                {
                    int charPos = (int)(position * line.Length);
                    
                    // Find the best split point near this position (look for spaces)
                    int splitPos = FindBestSplitPoint(line, charPos, lastPos);
                    
                    if (splitPos > lastPos)
                    {
                        result.Add(line.Substring(lastPos, splitPos - lastPos).Trim());
                        lastPos = splitPos;
                    }
                }
                
                // Add the remaining part
                if (lastPos < line.Length)
                {
                    result.Add(line.Substring(lastPos).Trim());
                }
                
                // Ensure we have at least one column
                if (result.Count == 0)
                    result.Add(line);
                
                return string.Join("\t", result);
            }
            catch
            {
                // If anything fails, just return the original line
                return line;
            }
        }

        private int FindBestSplitPoint(string line, int targetPos, int minPos)
        {
            // Look for a space near the target position
            int searchRadius = Math.Min(line.Length / 10, 20); // Search within 10% of line length or 20 chars
            
            // First, look for spaces after the target position
            for (int i = 0; i <= searchRadius && targetPos + i < line.Length; i++)
            {
                if (targetPos + i > minPos && line[targetPos + i] == ' ')
                    return targetPos + i;
            }
            
            // Then, look for spaces before the target position
            for (int i = 1; i <= searchRadius && targetPos - i > minPos; i++)
            {
                if (line[targetPos - i] == ' ')
                    return targetPos - i;
            }
            
            // If no space found, use the target position
            return Math.Max(targetPos, minPos);
        }
    }

    public class LoadedDocument : IDisposable
    {
        public string FilePath { get; set; }
        public string Name { get; set; }
        public DocumentType Type { get; set; }
        public List<Bitmap> Pages { get; set; } = new List<Bitmap>();
        public int PageCount { get; set; }

        public void Dispose()
        {
            foreach (var page in Pages)
            {
                page?.Dispose();
            }
            Pages.Clear();
        }
    }

    public enum DocumentType
    {
        PDF,
        Image
    }
} 