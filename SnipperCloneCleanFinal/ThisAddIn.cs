using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using SnipperCloneCleanFinal.Core;
using SnipperCloneCleanFinal.Infrastructure;
using SnipperCloneCleanFinal.UI;
using Extensibility;

namespace SnipperCloneCleanFinal
{
    [ComVisible(true)]
    [Guid("D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111")]
    [ProgId("SnipperPro.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ThisAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private Excel.Application _application;
        private static readonly object _lockObject = new object();
        private SnipEngine _snippEngine;
        private DocumentViewer _documentViewer;
        private SnipMode _currentSnipMode = SnipMode.Text;
        
        // Static instance for access from other classes
        public static ThisAddIn Instance { get; private set; }

        public Excel.Application Application => _application;
        public DocumentViewer DocumentViewer => _documentViewer;
        public SnipManager Snips { get; } = new SnipManager();

        #region IDTExtensibility2 Implementation
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                lock (_lockObject)
                {
                    _application = (Excel.Application)application;
                    Instance = this; // Set static instance for access from other classes

                    var user = Environment.GetEnvironmentVariable("SNIPPER_USER") ?? Environment.UserName;
                    var pass = Environment.GetEnvironmentVariable("SNIPPER_PASS") ?? "snipper";
                    if (!AuthManager.Authenticate(user, pass))
                    {
                        MessageBox.Show("Authentication failed", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    _snippEngine = new SnipEngine(_application);
                    
                    // Register Excel defined names for DS formulas - critical for validation/exception snips
                    RegisterDSNames();
                    
                    // Safely try to register event handlers
                    try
                    {
                        if (_application != null)
                        {
                            _application.SheetBeforeDoubleClick += OnCellDoubleClick;
                            _application.SheetChange += OnSheetChange;
                            _application.WorkbookBeforeSave += OnWorkbookBeforeSave;
                            _application.WorkbookOpen += OnWorkbookOpen;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Snipper Pro: Could not register event handlers: {ex.Message}");
                        // Continue without event handlers
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: OnConnection successful via IDTExtensibility2 (Mode: {connectMode})");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro OnConnection Error: {ex.Message}");
                MessageBox.Show($"OnConnection Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            try
            {
                lock (_lockObject)
                {
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: OnDisconnection (Mode: {disconnectMode})");
                    if (_application != null)
                    {
                        _application.SheetBeforeDoubleClick -= OnCellDoubleClick;
                        _application.SheetChange -= OnSheetChange;
                        _application.WorkbookBeforeSave -= OnWorkbookBeforeSave;
                        _application.WorkbookOpen -= OnWorkbookOpen;
                        Marshal.ReleaseComObject(_application);
                        _application = null;
                    }
                    Instance = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro OnDisconnection Error: {ex.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
        #endregion

        #region IRibbonExtensibility Implementation
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: GetCustomUI called with ribbonID: {ribbonID}");
                string ribbonXml = GetRibbonXml();
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Loaded Ribbon XML: {ribbonXml}");
                return ribbonXml;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro GetCustomUI Error: {ex.Message}");
                MessageBox.Show($"GetCustomUI Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return GetFallbackRibbonXml();
            }
        }

        private string GetRibbonXml()
        {
            try
            {
                string assetsPath = Path.Combine(GetAddInDirectory(), "Assets", "SnipperRibbon.xml");
                
                if (File.Exists(assetsPath))
                {
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Loading ribbon from {assetsPath}");
                    return File.ReadAllText(assetsPath);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Ribbon file not found at {assetsPath}, using fallback");
                    return GetFallbackRibbonXml();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Error loading ribbon XML: {ex.Message}");
                return GetFallbackRibbonXml();
            }
        }

        private string GetAddInDirectory()
        {
            try
            {
                string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
            catch
            {
                return System.IO.Directory.GetCurrentDirectory();
            }
        }
        #endregion

        #region Ribbon Callbacks
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            try
            {
                lock (_lockObject)
                {
                    _ribbon = ribbon;
                    System.Diagnostics.Debug.WriteLine("Snipper Pro: OnRibbonLoad called successfully");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro OnRibbonLoad Error: {ex.Message}");
                MessageBox.Show($"OnRibbonLoad Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTextSnip(IRibbonControl control)
        {
            ExecuteCallback("OnTextSnip", () =>
            {
                SetSnipMode(SnipMode.Text);
                Logger.Info("Text Snip mode activated");
            });
        }

        public void OnSumSnip(IRibbonControl control)
        {
            ExecuteCallback("OnSumSnip", () =>
            {
                SetSnipMode(SnipMode.Sum);
                Logger.Info("Sum Snip mode activated");
            });
        }

        public void OnTableSnip(IRibbonControl control)
        {
            ExecuteCallback("OnTableSnip", () =>
            {
                SetSnipMode(SnipMode.Table);
                Logger.Info("Table Snip mode activated");
            });
        }

        public void OnValidationSnip(IRibbonControl control)
        {
            ExecuteCallback("OnValidationSnip", () =>
            {
                SetSnipMode(SnipMode.Validation);
                Logger.Info("Validation Snip mode activated");
            });
        }

        public void OnExceptionSnip(IRibbonControl control)
        {
            ExecuteCallback("OnExceptionSnip", () =>
            {
                SetSnipMode(SnipMode.Exception);
                Logger.Info("Exception Snip mode activated");
            });
        }

        public async void OnOpenViewer(IRibbonControl control)
        {
            ExecuteCallback("OnOpenViewer", async () =>
            {
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "PDF files (*.pdf)|*.pdf|Image files (*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|All supported files|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|All files (*.*)|*.*";
                    openFileDialog.Title = "Select Document(s) to Load - PDFs and Images Supported";
                    openFileDialog.Multiselect = true;
                    
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        if (_documentViewer == null || _documentViewer.IsDisposed)
                        {
                            InitializeDocumentViewer();
                        }
                        
                        // Load all selected files
                        _documentViewer.Show();
                        _documentViewer.BringToFront();
                        
                        // Call LoadDocuments properly with await
                        try
                        {
                            await _documentViewer.LoadDocuments(openFileDialog.FileNames);
                            Logger.Info($"Loaded {openFileDialog.FileNames.Length} document(s)");
                        }
                        catch (Exception ex)
                        {
                            Logger.Error($"Error loading documents: {ex.Message}", ex);
                            MessageBox.Show($"Error loading documents: {ex.Message}", 
                                "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            });
        }

        public void OnMarkupSnip(IRibbonControl control)
        {
            ExecuteCallback("OnMarkupSnip", () =>
            {
                if (_documentViewer != null && !_documentViewer.IsDisposed)
                {
                    _documentViewer.Show();
                    _documentViewer.BringToFront();
                    MessageBox.Show("Document viewer is now active. Use the snip buttons to start marking up documents.", 
                        "Snipper Pro - Markup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Please open a document first using the 'Open Viewer' button.", 
                        "Snipper Pro - Markup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            });
        }

        // Icon generation methods for ribbon buttons
        public stdole.IPictureDisp GetTextSnipIcon(IRibbonControl control)
        {
            return CreateColoredRectangleIcon(System.Drawing.Color.Blue);
        }

        public stdole.IPictureDisp GetSumSnipIcon(IRibbonControl control)
        {
            return CreateColoredRectangleIcon(System.Drawing.Color.Purple);
        }

        public stdole.IPictureDisp GetTableSnipIcon(IRibbonControl control)
        {
            return CreateColoredRectangleIcon(System.Drawing.Color.Purple);
        }

        public stdole.IPictureDisp GetValidationSnipIcon(IRibbonControl control)
        {
            return CreateColoredRectangleIcon(System.Drawing.Color.Green);
        }

        public stdole.IPictureDisp GetExceptionSnipIcon(IRibbonControl control)
        {
            return CreateColoredRectangleIcon(System.Drawing.Color.Red);
        }

        private stdole.IPictureDisp CreateColoredRectangleIcon(System.Drawing.Color color)
        {
            try
            {
                // Create a larger, more professional bitmap for the ribbon icon (48x48 pixels)
                using (var bitmap = new System.Drawing.Bitmap(48, 48, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                    {
                        // Set high quality rendering for professional appearance
                        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        
                        // Clear with transparent background
                        graphics.Clear(System.Drawing.Color.Transparent);
                        
                        // Create a DataSnipper-style colored square with rounded corners and gradient
                        var rect = new System.Drawing.Rectangle(6, 6, 36, 36);
                        
                        // Create gradient brush for 3D effect like DataSnipper
                        using (var gradientBrush = new System.Drawing.Drawing2D.LinearGradientBrush(
                            rect, 
                            System.Drawing.Color.FromArgb(255, color.R + Math.Min(30, 255 - color.R), color.G + Math.Min(30, 255 - color.G), color.B + Math.Min(30, 255 - color.B)), 
                            System.Drawing.Color.FromArgb(255, Math.Max(0, color.R - 40), Math.Max(0, color.G - 40), Math.Max(0, color.B - 40)),
                            System.Drawing.Drawing2D.LinearGradientMode.Vertical))
                        {
                            // Create rounded rectangle path
                            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
                            {
                                int radius = 6; // Rounded corner radius
                                path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
                                path.AddArc(rect.X + rect.Width - radius, rect.Y, radius, radius, 270, 90);
                                path.AddArc(rect.X + rect.Width - radius, rect.Y + rect.Height - radius, radius, radius, 0, 90);
                                path.AddArc(rect.X, rect.Y + rect.Height - radius, radius, radius, 90, 90);
                                path.CloseAllFigures();
                                
                                // Fill with gradient
                                graphics.FillPath(gradientBrush, path);
                                
                                // Add subtle border with darker shade
                                using (var borderPen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(180, Math.Max(0, color.R - 60), Math.Max(0, color.G - 60), Math.Max(0, color.B - 60)), 1.5f))
                                {
                                    graphics.DrawPath(borderPen, path);
                                }
                                
                                // Add subtle inner highlight for depth (DataSnipper style)
                                using (var highlightPen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(80, 255, 255, 255), 1))
                                {
                                    var innerRect = new System.Drawing.Rectangle(rect.X + 2, rect.Y + 2, rect.Width - 4, rect.Height - 4);
                                    using (var innerPath = new System.Drawing.Drawing2D.GraphicsPath())
                                    {
                                        int innerRadius = 4;
                                        innerPath.AddArc(innerRect.X, innerRect.Y, innerRadius, innerRadius, 180, 90);
                                        innerPath.AddArc(innerRect.X + innerRect.Width - innerRadius, innerRect.Y, innerRadius, innerRadius, 270, 90);
                                        innerPath.AddArc(innerRect.X + innerRect.Width - innerRadius, innerRect.Y + innerRect.Height - innerRadius, innerRadius, innerRadius, 0, 90);
                                        innerPath.AddArc(innerRect.X, innerRect.Y + innerRect.Height - innerRadius, innerRadius, innerRadius, 90, 90);
                                        innerPath.CloseAllFigures();
                                        graphics.DrawPath(highlightPen, innerPath);
                                    }
                                }
                            }
                        }
                    }
                    
                    // Convert to IPictureDisp for ribbon
                    return ConvertBitmapToIPicture(bitmap);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating DataSnipper-style icon: {ex.Message}");
                return null;
            }
        }
        
        private stdole.IPictureDisp ConvertBitmapToIPicture(System.Drawing.Bitmap bitmap)
        {
            try
            {
                // Save bitmap to memory stream
                using (var stream = new System.IO.MemoryStream())
                {
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                    stream.Position = 0;
                    
                    // Convert stream to byte array
                    byte[] imageBytes = stream.ToArray();
                    
                    // Create IPictureDisp from byte array using OLE functions
                    IntPtr hGlobal = Marshal.AllocHGlobal(imageBytes.Length);
                    try
                    {
                        Marshal.Copy(imageBytes, 0, hGlobal, imageBytes.Length);
                        
                        // Create IStream from global memory
                        if (CreateStreamOnHGlobal(hGlobal, false, out IStream stream2) == 0)
                        {
                            // Create IPicture from stream
                            Guid riid = new Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB"); // IID_IPicture
                            if (OleLoadPicture(stream2, imageBytes.Length, false, ref riid, out object picture) == 0)
                            {
                                return (stdole.IPictureDisp)picture;
                            }
                        }
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(hGlobal);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error converting bitmap to IPicture: {ex.Message}");
            }
            
            return null;
        }
        
        [ComImport, Guid("0000000c-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IStream
        {
            void Read([Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, uint cb, out uint pcbRead);
            void Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, uint cb, out uint pcbWritten);
            void Seek(long dlibMove, uint dwOrigin, out long plibNewPosition);
            void SetSize(long libNewSize);
            void CopyTo(IStream pstm, long cb, out long pcbRead, out long pcbWritten);
            void Commit(uint grfCommitFlags);
            void Revert();
            void LockRegion(long libOffset, long cb, uint dwLockType);
            void UnlockRegion(long libOffset, long cb, uint dwLockType);
            void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, uint grfStatFlag);
            void Clone(out IStream ppstm);
        }
        
        [DllImport("ole32.dll")]
        private static extern int CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out IStream ppstm);
        
        [DllImport("oleaut32.dll")]
        private static extern int OleLoadPicture(IStream lpstream, int lSize, bool fRunmode, ref Guid riid, out object lplpvObj);
        


        // Import Windows API for cleanup
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private void ExecuteCallback(string callbackName, System.Action action)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Executing {callbackName}");
                action();
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: {callbackName} completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro {callbackName} Error: {ex.Message}");
                MessageBox.Show($"Error in {callbackName}: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void ExecuteCallback(string callbackName, System.Func<System.Threading.Tasks.Task> asyncAction)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Executing {callbackName}");
                await asyncAction();
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: {callbackName} completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro {callbackName} Error: {ex.Message}");
                MessageBox.Show($"Error in {callbackName}: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegisterDSNames()
        {
            try
            {
                if (_application?.Names == null)
                {
                    System.Diagnostics.Debug.WriteLine("Cannot register DS names - Application or Names collection is null");
                    return;
                }

                System.Diagnostics.Debug.WriteLine("Registering DS defined names...");

                // Remove existing DS names first to avoid conflicts
                string[] dsNames = { "DS.TEXTS", "DS.SUMS", "DS.TABLE", "DS.VALIDATION", "DS.EXCEPTION" };
                foreach (var name in dsNames)
                {
                    try
                    {
                        var existingName = _application.Names.Item(name);
                        if (existingName != null)
                        {
                            existingName.Delete();
                            System.Diagnostics.Debug.WriteLine($"Removed existing name: {name}");
                        }
                    }
                    catch { /* Name doesn't exist, that's fine */ }
                }

                // Register DS names with proper error handling
                RegisterDSName("DS.TEXTS", "SnipperPro.Connect.TEXTS");
                RegisterDSName("DS.SUMS", "SnipperPro.Connect.SUMS");
                RegisterDSName("DS.TABLE", "SnipperPro.Connect.TABLE");
                RegisterDSName("DS.VALIDATION", "SnipperPro.Connect.VALIDATION");
                RegisterDSName("DS.EXCEPTION", "SnipperPro.Connect.EXCEPTION");
                RegisterDSName("DS.NAVIGATE", "SnipperPro.Connect.NAVIGATE");

                System.Diagnostics.Debug.WriteLine("DS defined names registration completed");
                
                // Test if our UDF functions are accessible
                TestUDFFunctions();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in RegisterDSNames: {ex.Message}");
            }
        }

        private void RegisterDSName(string shortName, string fullName)
        {
            try
            {
                // Try direct ProgId.Function format first (most reliable for UDFs)
                _application.Names.Add(shortName, $"=SnipperPro.Connect.{fullName.Split('.').Last()}");
                System.Diagnostics.Debug.WriteLine($"Successfully registered {shortName} -> SnipperPro.Connect.{fullName.Split('.').Last()}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Primary registration failed for {shortName}: {ex.Message}");
                
                // Fallback to full name
                try
                {
                    _application.Names.Add(shortName, $"={fullName}");
                    System.Diagnostics.Debug.WriteLine($"Fallback registration successful for {shortName}");
                }
                catch (Exception ex2)
                {
                    System.Diagnostics.Debug.WriteLine($"All registration methods failed for {shortName}: {ex2.Message}");
                }
            }
        }

        private void TestUDFFunctions()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Testing UDF function accessibility...");
                
                // Test direct function calls
                var testResult1 = VALIDATION("test-id");
                var testResult2 = EXCEPTION("test-id");
                
                System.Diagnostics.Debug.WriteLine($"Direct VALIDATION call result: {testResult1}");
                System.Diagnostics.Debug.WriteLine($"Direct EXCEPTION call result: {testResult2}");
                
                // Test if Excel can evaluate the formulas
                try
                {
                    var testCell = _application.ActiveSheet.Range["Z100"]; // Use a cell far away
                    testCell.Formula = "=SnipperPro.Connect.VALIDATION(\"test-id\")";
                    var cellValue = testCell.Value;
                    System.Diagnostics.Debug.WriteLine($"Excel evaluation of VALIDATION formula: {cellValue}");
                    testCell.Clear(); // Clean up
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Excel formula evaluation failed: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UDF function test failed: {ex.Message}");
            }
        }

        private void SetSnipMode(SnipMode snipMode)
        {
            _currentSnipMode = snipMode;
            
            // Enable snip mode in document viewer if it's open
            if (_documentViewer != null && !_documentViewer.IsDisposed)
            {
                _documentViewer.SetSnipMode(snipMode, true);
                _documentViewer.Show();
                _documentViewer.BringToFront();
            }
            else
            {
                MessageBox.Show($"{snipMode} Snip mode activated. Open a document to start snipping.", 
                    "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void InitializeDocumentViewer()
        {
            try
            {
                _documentViewer = new DocumentViewer(_snippEngine);
                _documentViewer.SnipAreaSelected += OnSnipAreaSelected;
                _documentViewer.SnipClicked += OnSnipClicked;
                Logger.Info("Document viewer initialized with full PDF support");
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to initialize document viewer: {ex.Message}", ex);
            }
        }

        private async void OnSnipAreaSelected(object sender, SnipAreaSelectedEventArgs e)
        {
            try
            {
                Logger.Info($"Processing {e.SnipMode} snip...");
                
                // Get the currently selected cell in Excel
                var activeCell = Application.ActiveCell;
                if (activeCell == null)
                {
                    MessageBox.Show("Please select a cell in Excel before snipping.", "Snipper Pro", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                string formula = "";
                string displayValue = "";
                string snipId = null;
                
                // Use the extracted data from the event args
                switch (e.SnipMode)
                {
                    case SnipMode.Text:
                        // With our new fallback system, we should always get some text
                        if (!string.IsNullOrEmpty(e.ExtractedText))
                        {
                            formula = DataSnipperFormulas.CreateTextFormula(e.DocumentPath, e.PageNumber,
                                e.ExtractedText, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height),
                                activeCell.Address, out snipId);
                            _documentViewer.AssignSnipIdToLastRecord(snipId);
                            displayValue = e.ExtractedText;
                        }
                        else
                        {
                            // Even if we get empty text, provide a meaningful message
                            displayValue = "[No text detected in image]";
                            formula = DataSnipperFormulas.CreateTextFormula(e.DocumentPath, e.PageNumber,
                                displayValue, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height),
                                activeCell.Address, out snipId);
                            _documentViewer.AssignSnipIdToLastRecord(snipId);
                        }
                        break;
                        
                    case SnipMode.Sum:
                        if (e.Success && e.ExtractedNumbers != null && e.ExtractedNumbers.Length > 0)
                        {
                            var numbers = new List<double>();
                            double sum = 0;
                            
                            foreach (var numStr in e.ExtractedNumbers)
                            {
                                if (double.TryParse(numStr.Replace(",", "").Replace("$", ""), out double num))
                                {
                                    numbers.Add(num);
                                    sum += num;
                                }
                            }
                            
                            if (numbers.Count > 0)
                            {
                                formula = DataSnipperFormulas.CreateSumFormula(e.DocumentPath, e.PageNumber,
                                    sum, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), numbers,
                                    activeCell.Address, out snipId);
                                _documentViewer.AssignSnipIdToLastRecord(snipId);
                                displayValue = sum.ToString("F2");
                            }
                            else
                            {
                                displayValue = "No numbers found";
                            }
                        }
                        else
                        {
                            displayValue = "No numbers extracted";
                        }
                        break;
                        
                    case SnipMode.Table:
                        if (e.Success && !string.IsNullOrEmpty(e.ExtractedText))
                        {
                            var parser = new TableParser();
                            var tableData = parser.ParseTable(e.ExtractedText);
                            if (!tableData.IsEmpty)
                            {
                                // Put formula in current cell (above table)
                                formula = DataSnipperFormulas.CreateTableFormula(e.DocumentPath, e.PageNumber,
                                    tableData, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height),
                                    activeCell.Address, out snipId);
                                _documentViewer.AssignSnipIdToLastRecord(snipId);
                                activeCell.Formula = formula;
                                
                                // Skip one row for spacing, then write the table data
                                var tableStartCell = activeCell.Offset[2, 0]; // 2 rows down for spacing
                                using (var helper = new ExcelHelper(Application))
                                {
                                    helper.WriteTableToRange(tableData, tableStartCell);
                                }
                                
                                displayValue = $"Table: {tableData.RowCount}×{tableData.ColumnCount}";
                            }
                            else
                            {
                                displayValue = "Table extraction failed";
                            }
                        }
                        else
                        {
                            displayValue = "Table extraction failed";
                        }
                        break;
                        
                    case SnipMode.Validation:
                        formula = DataSnipperFormulas.CreateValidationFormula(e.DocumentPath, e.PageNumber,
                            new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), activeCell.Address, out snipId);
                        _documentViewer.AssignSnipIdToLastRecord(snipId);
                        displayValue = "✓";
                        break;

                    case SnipMode.Exception:
                        formula = DataSnipperFormulas.CreateExceptionFormula(e.DocumentPath, e.PageNumber,
                            new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), activeCell.Address, out snipId);
                        _documentViewer.AssignSnipIdToLastRecord(snipId);
                        displayValue = "✗";
                        break;

                    case SnipMode.Image:
                        // Image snip already inserted by DocumentViewer.  Optionally record navigation info.
                        displayValue = "[Image]";
                        break;
                }

                if (!string.IsNullOrEmpty(snipId))
                {
                    var data = DataSnipperFormulas.GetSnipData(snipId);
                    if (data != null)
                        DataSnipperPersistence.Upsert(data);
                }
                
                // Insert the value into Excel based on snip type
                if (e.SnipMode == SnipMode.Text)
                {
                    activeCell.Value2 = displayValue;
                }
                else if (e.SnipMode == SnipMode.Sum)
                {
                    // For sum, put the actual number value
                    if (double.TryParse(displayValue, out double sumValue))
                        activeCell.Value2 = sumValue;
                    else
                        activeCell.Value2 = displayValue;
                }
                else if (e.SnipMode == SnipMode.Validation)
                {
                    activeCell.Value2 = "✓";
                }
                else if (e.SnipMode == SnipMode.Exception)
                {
                    activeCell.Value2 = "✗";
                }
                else if (e.SnipMode == SnipMode.Table)
                {
                    // For table, extract the table data to multiple cells
                    try
                    {
                        // Just put a placeholder for now - table extraction needs special handling
                        activeCell.Value2 = "[Table Data]";
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error setting table data: {ex.Message}");
                    }
                }

                // Create navigation link in adjacent cell
                        try
                        {
                    Excel.Range navCell = (Excel.Range)activeCell.Offset[0, 1];
                    
                    // Store snip data in DataSnipper system for navigation
                    var customBounds = new Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height);
                    
                    if (e.SnipMode == SnipMode.Text)
                    {
                        DataSnipperFormulas.CreateTextFormula(
                            e.DocumentPath, e.PageNumber, displayValue, customBounds, 
                            navCell.Address[false, false], out snipId);
                    }
                    else if (e.SnipMode == SnipMode.Sum)
                    {
                        var numbers = new List<double>();
                        if (double.TryParse(displayValue, out double sumValue))
                            numbers.Add(sumValue);
                        
                        DataSnipperFormulas.CreateSumFormula(
                            e.DocumentPath, e.PageNumber, sumValue, customBounds, numbers,
                            navCell.Address[false, false], out snipId);
                    }
                    else if (e.SnipMode == SnipMode.Validation)
                    {
                        DataSnipperFormulas.CreateValidationFormula(
                            e.DocumentPath, e.PageNumber, customBounds,
                            navCell.Address[false, false], out snipId);
                    }
                    else if (e.SnipMode == SnipMode.Exception)
                    {
                        DataSnipperFormulas.CreateExceptionFormula(
                            e.DocumentPath, e.PageNumber, customBounds,
                            navCell.Address[false, false], out snipId);
                    }
                    else if (e.SnipMode == SnipMode.Table)
                    {
                        DataSnipperFormulas.CreateTableFormula(
                            e.DocumentPath, e.PageNumber, null, customBounds,
                            navCell.Address[false, false], out snipId);
                    }
                    
                    // Create a simple clickable text instead of complex formula
                    navCell.Value2 = "→ Go to snip";
                    
                    // Make it look like a hyperlink
                    navCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    navCell.Font.Underline = true;
                    
                    // Store the snip ID in the cell comment for retrieval
                    navCell.AddComment($"SnipID:{snipId}");
                    navCell.Comment.Visible = false;
                    
                    System.Diagnostics.Debug.WriteLine($"Created navigation cell with snip ID: {snipId}");
                        }
                        catch (Exception ex)
                        {
                    System.Diagnostics.Debug.WriteLine($"Error creating navigation cell: {ex.Message}");
                }
                
                // Insert the value into Excel (except for table which handles its own cells)
                if (!string.IsNullOrEmpty(displayValue) && e.SnipMode != SnipMode.Table)
                {
                    // Always put the actual value in the cell for display
                    activeCell.Value2 = displayValue;
                    
                    // Add comment with source info and formula for reference
                    try
                    {
                        activeCell.ClearComments();
                        var commentText = $"Source: {Path.GetFileName(e.DocumentPath)}\nPage: {e.PageNumber}\nSnip Type: {e.SnipMode}";
                        if (!string.IsNullOrEmpty(formula))
                        {
                            commentText += $"\nFormula: {formula}";
                        }
                        var comment = activeCell.AddComment(commentText);
                        comment.Shape.TextFrame.AutoSize = true;
                    }
                    catch { }
                    
                    // Add visual indicator (border color)
                    var color = GetSnipColorCode(e.SnipMode);
                    activeCell.Borders.Color = color;
                    
                    Logger.Info($"{e.SnipMode} snip completed successfully - Value: {displayValue}");
                    
                    // Move to next cell automatically for continuous workflow
                    try
                    {
                        if (e.SnipMode == SnipMode.Table)
                        {
                            // For tables, move to after the table data (formula + space + table rows)
                            var parser = new TableParser();
                            var tableData = parser.ParseTable(e.ExtractedText);
                            if (!tableData.IsEmpty)
                            {
                                int headerRows = tableData.HasHeader && tableData.Headers != null ? 1 : 0;
                                int totalTableRows = tableData.Rows.Count + headerRows;
                                var nextCell = activeCell.Offset[2 + totalTableRows + 1, 0]; // formula + space + table + one more space
                                nextCell.Select();
                            }
                            else
                            {
                        var nextCell = activeCell.Offset[1, 0];
                        nextCell.Select();
                            }
                        }
                        else
                        {
                            // Move down one cell after each non-table snip
                            var nextCell = activeCell.Offset[1, 0];
                            nextCell.Select();
                        }
                    }
                    catch
                    {
                        // If we can't move down (last row), stay on current cell
                    }
                    
                    // Keep document viewer on top
                    _documentViewer.BringToFront();
                }
                else if (e.SnipMode == SnipMode.Table)
                {
                    // Table already handled above
                    _documentViewer.BringToFront();
                }
                else
                {
                    MessageBox.Show($"{e.SnipMode} snip failed - no data extracted", 
                        "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                
                // Highlight the region in the document viewer  
                _documentViewer.HighlightRegion(e.Bounds, GetSnipColor(e.SnipMode));
                
                // Disable snip mode after each snip so the user must re-activate it
                _documentViewer.SetSnipMode(e.SnipMode, false);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing snip: {ex.Message}", ex);
                MessageBox.Show($"Error processing snip: {ex.Message}", "Snipper Pro Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnSnipClicked(object sender, SnipClickedEventArgs e)
        {
            var snipData = DataSnipperFormulas.GetSnipData(e.Snip.SnipId);
            if (snipData == null || string.IsNullOrEmpty(snipData.CellAddress)) return;
            DataSnipperPersistence.Upsert(snipData);
            try
            {
                var cell = Application.Range[snipData.CellAddress];
                cell.Select();
            }
            catch { }
        }

        internal void CreateSnip(System.Drawing.Rectangle bounds, SnipKind kind, string extracted,
                                 string docPath, int page, Excel.Range targetCell)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Creating snip - Kind: {kind}, DocPath: {docPath}, Page: {page}");
                
            var snip = new Core.SnipOverlay
            {
                DocPath   = docPath,
                PageIndex = page,
                Bounds    = bounds,
                Kind      = kind,
                Extracted = extracted,
                SheetName = targetCell.Worksheet.Name,
                CellAddr  = targetCell.Address[false, false],
            };

                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Snip created with ID: {snip.Id}");
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Cell location: {snip.SheetName}!{snip.CellAddr}");

            // write value and formula into Excel
            targetCell.Value2  = extracted;
                var formula = $"=SNIP(\"{snip.Id}\")";
                targetCell.Formula = formula;
                
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Formula written: {formula}");

            Snips.Add(snip);
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Snip added to manager. Total snips: {Snips.All.Count}");
                
            _documentViewer?.Invalidate();  // redraw with new rectangle
                
                System.Diagnostics.Debug.WriteLine("Snipper Pro: CreateSnip completed successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: CreateSnip error: {ex.Message}");
                MessageBox.Show($"Create snip error: {ex.Message}", "Snipper Pro Debug", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnCellDoubleClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Snipper Pro: OnCellDoubleClick triggered");
                
            var f = Target.Formula as string;
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: Cell formula: '{f}'");
                
                // Check for navigation cell with snip ID in comment
                if (Target.Comment != null && Target.Comment.Text().StartsWith("SnipID:"))
                {
                    var commentText = Target.Comment.Text();
                    var snipId = commentText.Substring(7); // Remove "SnipID:" prefix
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Found navigation cell with snip ID: {snipId}");

            Cancel = true;
                    
                    // Use DataSnipper's navigation function
                    var navigated = DataSnipperFormulas.NavigateToSnip(snipId);
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Navigation result: {navigated}");
                    return;
                }
                
                if (f == null) 
                {
                    System.Diagnostics.Debug.WriteLine("Snipper Pro: No formula found and no navigation comment");
                    return;
                }

                // Check for DataSnipper format: =SnipperPro.Connect.TEXTS("guid")
                var dsMatch = System.Text.RegularExpressions.Regex.Match(f, @"SnipperPro\.Connect\.(TEXTS|SUMS|TABLE|VALIDATION|EXCEPTION)\(\""(?<id>[-A-F0-9]{36})\""");
                
                // Check for navigation formula: =IF(TRUE,"→ Go to snip",SnipperPro.Connect.NAVIGATE("guid"))
                var navMatch = System.Text.RegularExpressions.Regex.Match(f, @"SnipperPro\.Connect\.NAVIGATE\(\""(?<id>[-A-F0-9]{36})\""");
                
                if (dsMatch.Success)
                {
                    var snipId = dsMatch.Groups["id"].Value;
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Found DataSnipper formula with ID: {snipId}");
                    
                    Cancel = true;
                    
                    // Use DataSnipper's existing navigation function
                    var navigated = DataSnipperFormulas.NavigateToSnip(snipId);
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Navigation result: {navigated}");
                    return;
                }
                
                if (navMatch.Success)
                {
                    var snipId = navMatch.Groups["id"].Value;
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Found navigation formula with ID: {snipId}");
                    
                    Cancel = true;
                    
                    // Use DataSnipper's existing navigation function
                    var navigated = DataSnipperFormulas.NavigateToSnip(snipId);
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Navigation result: {navigated}");
                    return;
                }
                
                // Check for new SNIP format: =SNIP("guid") 
                var snipMatch = System.Text.RegularExpressions.Regex.Match(f, @"SNIP\(\""(?<id>[-A-F0-9]{36})\""");
                
                if (snipMatch.Success)
                {
                    var snipId = snipMatch.Groups["id"].Value;
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Found new SNIP format with ID: {snipId}");
                    
                    Cancel = true;
                    
                    // Try new system first
                    if (Guid.TryParse(snipId, out var id))
                    {
                        var snip = Snips.ById(id);
                        if (snip != null)
                        {
            var viewer = GetOrCreateViewer();
            viewer.LoadDocumentIfNeeded(snip.DocPath);
            viewer.ScrollToPage(snip.PageIndex);
            viewer.HighlightOnce(snip.Bounds);
                            System.Diagnostics.Debug.WriteLine("Snipper Pro: Navigation completed via new system");
                            return;
                        }
                    }
                    
                    // Fallback to DataSnipper system
                    var navigated = DataSnipperFormulas.NavigateToSnip(snipId);
                    System.Diagnostics.Debug.WriteLine($"Snipper Pro: Fallback navigation result: {navigated}");
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine("Snipper Pro: No recognized formula pattern found");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro: OnCellDoubleClick error: {ex.Message}");
            }
        }

        private void OnSheetChange(object Sh, Excel.Range Target)
        {
            if (Target.Count != 1) return;

            var addr  = Target.Address[false, false];
            var sheet = Target.Worksheet.Name;

            var snip = Snips.ByCell(sheet, addr);
            if (snip == null) return;

            // cell no longer holds the SNIP formula, so remove overlay
            Snips.Remove(snip.Id);
            _documentViewer?.Invalidate();
        }

        private DocumentViewer GetOrCreateViewer()
        {
            if (_documentViewer == null)
            {
                InitializeDocumentViewer();
            }
            return _documentViewer;
        }
        
        private System.Drawing.Color GetSnipColor(SnipMode snipMode)
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
        
        private int GetSnipColorCode(SnipMode snipMode)
        {
            return snipMode switch
            {
                SnipMode.Text => 16711680, // Blue in Excel color format
                SnipMode.Sum => 8388736,   // Purple
                SnipMode.Table => 8388736, // Purple
                SnipMode.Validation => 65280, // Green
                SnipMode.Exception => 255,    // Red
                _ => 8421504 // Gray
            };
        }

        // Excel UDFs for DataSnipper formulas
        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(1)]
        public object TEXTS(string snipId)
        {
            var data = DataSnipperFormulas.GetSnipData(snipId);
            return data != null ? data.ExtractedValue ?? string.Empty : "#N/A";
        }

        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(2)]
        public object SUMS(string snipId)
        {
            var data = DataSnipperFormulas.GetSnipData(snipId);
            if (data != null)
            {
                if (data.Numbers != null && data.Numbers.Count > 0)
                    return data.Numbers.Sum();
                if (double.TryParse(data.ExtractedValue, out var val))
                    return val;
                return data.ExtractedValue ?? string.Empty;
            }
            return "#N/A";
        }

        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(3)]
        public object TABLE(string snipId)
        {
            var data = DataSnipperFormulas.GetSnipData(snipId);
            if (data?.Table != null && data.Table.Rows.Count > 0)
            {
                var rows = data.Table.RowCount;
                var cols = data.Table.ColumnCount;
                var result = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        result[r, c] = data.Table.GetCell(r, c);
                    }
                }
                return result;
            }
            return "#N/A";
        }

        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(4)]
        public object VALIDATION(string snipId)
        {
            System.Diagnostics.Debug.WriteLine($"VALIDATION UDF called with snipId: {snipId}");
            var snipData = DataSnipperFormulas.GetSnipData(snipId);
            var result = snipData != null ? "✓" : "#N/A";
            System.Diagnostics.Debug.WriteLine($"VALIDATION UDF returning: {result}");
            return result;
        }

        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(5)]
        public object EXCEPTION(string snipId)
        {
            System.Diagnostics.Debug.WriteLine($"EXCEPTION UDF called with snipId: {snipId}");
            var snipData = DataSnipperFormulas.GetSnipData(snipId);
            var result = snipData != null ? "✗" : "#N/A";
            System.Diagnostics.Debug.WriteLine($"EXCEPTION UDF returning: {result}");
            return result;
        }

        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(6)]
        public object SNIP(string snipId)
        {
            System.Diagnostics.Debug.WriteLine($"SNIP UDF called with snipId: {snipId}");
            
            if (Guid.TryParse(snipId, out var id))
            {
                var snip = Snips.ById(id);
                if (snip != null)
                {
                    System.Diagnostics.Debug.WriteLine($"SNIP UDF returning: {snip.Extracted}");
                    return snip.Extracted ?? string.Empty;
                }
            }
            
            System.Diagnostics.Debug.WriteLine("SNIP UDF returning: #N/A");
            return "#N/A";
        }

        [ComVisible(true)]
        [System.Runtime.InteropServices.DispId(7)]
        public object NAVIGATE(string snipId)
        {
            System.Diagnostics.Debug.WriteLine($"NAVIGATE UDF called with snipId: {snipId}");
            
            try
            {
                // Call DataSnipper's navigation function
                var navigated = DataSnipperFormulas.NavigateToSnip(snipId);
                System.Diagnostics.Debug.WriteLine($"Navigation result: {navigated}");
                
                return "→ Go to snip";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Navigation error: {ex.Message}");
                return "Navigation failed";
            }
        }

        private void OnWorkbookBeforeSave(Excel.Workbook wb, bool saveAsUI, ref bool cancel)
        {
            DataSnipperFormulas.SaveSnips(wb);
            try
            {
                DataSnipperPersistence.Save(wb);
                DataSnipperPersistence.SaveSnipOverlays(wb);
            }
            catch { }
        }

        private void OnWorkbookOpen(Excel.Workbook wb)
        {
            DataSnipperFormulas.LoadSnips(wb);
            try
            {
                DataSnipperPersistence.Load(wb);
            }
            catch { }
        }
        #endregion

        private string GetFallbackRibbonXml()
        {
            System.Diagnostics.Debug.WriteLine("Snipper Pro: Using fallback ribbon XML");
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""SnipperProTab"" label=""SNIPPER PRO"">
        <group id=""ViewerGroup"" label=""Document Viewer"">
          <button id=""OpenViewerButton"" label=""Open Viewer"" size=""large"" onAction=""OnOpenViewer""
                  screentip=""Open document viewer""
                  supertip=""Open the document viewer to load and analyze documents."" />
          <button id=""MarkupButton"" label=""Markup"" size=""large"" onAction=""OnMarkupSnip""
                  screentip=""Toggle markup mode""
                  supertip=""Enable annotation and markup tools in the document viewer."" />
        </group>
        <group id=""SnipGroup"" label=""Snips"">
          <button id=""TextSnipButton"" label=""Text Snip"" size=""large"" onAction=""OnTextSnip""
                  getImage=""GetTextSnipIcon""
                  screentip=""Extract text from selected area"" 
                  supertip=""Use OCR to extract text from the selected area in the document viewer. Creates DS.TEXTS() formula."" />
          <button id=""SumSnipButton"" label=""Sum Snip"" size=""large"" onAction=""OnSumSnip""
                  getImage=""GetSumSnipIcon""
                  screentip=""Sum numbers from selected area""
                  supertip=""Extract and sum numerical values from the selected area. Creates DS.SUMS() formula."" />
          <button id=""TableSnipButton"" label=""Table Snip"" size=""large"" onAction=""OnTableSnip""
                  getImage=""GetTableSnipIcon""
                  screentip=""Extract table data""
                  supertip=""Extract structured table data from the selected area."" />
          <button id=""ValidationSnipButton"" label=""Validation"" size=""large"" onAction=""OnValidationSnip""
                  getImage=""GetValidationSnipIcon""
                  screentip=""Mark as validated""
                  supertip=""Mark the selected cell as validated with a checkmark."" />
          <button id=""ExceptionSnipButton"" label=""Exception"" size=""large"" onAction=""OnExceptionSnip""
                  getImage=""GetExceptionSnipIcon""
                  screentip=""Mark as exception""
                  supertip=""Mark the selected cell as an exception with an X mark."" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
} 