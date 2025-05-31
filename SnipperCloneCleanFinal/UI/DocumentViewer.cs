using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using SnipperCloneCleanFinal.Core;
using SnipperCloneCleanFinal.Infrastructure;

namespace SnipperCloneCleanFinal.UI
{
    public partial class DocumentViewer : Form
    {
        private readonly SnipEngine _snippEngine;
        private bool _isAnnotationMode = false;
        private Panel _contentPanel;
        private Label _statusLabel;

        public DocumentViewer(SnipEngine snippEngine)
        {
            _snippEngine = snippEngine ?? throw new ArgumentNullException(nameof(snippEngine));
            InitializeComponent();
            SetupSimpleUI();
        }

        private void SetupSimpleUI()
        {
            // Create a simple UI without WebView2
            this.Text = "Snipper Pro - Document Viewer";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            _statusLabel = new Label
            {
                Text = "Document viewer ready. Use the SNIPPER PRO ribbon to select snip modes.",
                Dock = DockStyle.Bottom,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.LightGray,
                Padding = new Padding(10, 5, 10, 5)
            };

            this.Controls.Add(_contentPanel);
            this.Controls.Add(_statusLabel);

            // Add a simple click handler for testing
            _contentPanel.Click += OnContentPanelClick;

            Logger.Info("DocumentViewer: Simple UI initialized successfully");
        }

        private async void OnContentPanelClick(object sender, EventArgs e)
        {
            try
            {
                // Simulate a selection for testing
                var rectangle = new Core.Rectangle(10, 10, 100, 50);
                
                // Create a sample bitmap
                var bitmap = new Bitmap(100, 50);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.FillRectangle(Brushes.White, 0, 0, 100, 50);
                    g.DrawString("Test Text", SystemFonts.DefaultFont, Brushes.Black, 10, 10);
                }

                // Process the snip
                var result = await _snippEngine.ProcessSnipAsync(bitmap, 1, rectangle);
                
                if (result.Success)
                {
                    _statusLabel.Text = $"Snip processed successfully: {result.Value}";
                    Logger.Info($"DocumentViewer: Snip processed successfully: {result.Value}");
                }
                else
                {
                    _statusLabel.Text = $"Snip processing failed: {result.ErrorMessage}";
                    Logger.Error($"DocumentViewer: Snip processing failed: {result.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"DocumentViewer: Error handling click: {ex.Message}", ex);
                _statusLabel.Text = $"Error: {ex.Message}";
            }
        }

        public void ToggleAnnotateMode()
        {
            _isAnnotationMode = !_isAnnotationMode;
            _statusLabel.Text = $"Annotation mode: {(_isAnnotationMode ? "ON" : "OFF")}";
            Logger.Info($"DocumentViewer: Annotation mode toggled to {_isAnnotationMode}");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    Logger.Info("DocumentViewer: Disposing");
                }
                catch (Exception ex)
                {
                    Logger.Error($"DocumentViewer: Error during disposal: {ex.Message}", ex);
                }
            }
            base.Dispose(disposing);
        }
    }
} 