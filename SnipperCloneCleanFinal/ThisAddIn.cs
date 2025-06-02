using System;
using System.Collections.Generic;
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
        private DocumentViewerPane _documentViewer;
        private SnipMode _currentSnipMode = SnipMode.Text;

        public Excel.Application Application => _application;

        #region IDTExtensibility2 Implementation
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                lock (_lockObject)
                {
                    _application = (Excel.Application)application;
                    _snippEngine = new SnipEngine(_application);
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
                        Marshal.ReleaseComObject(_application);
                        _application = null;
                    }
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

        public void OnOpenViewer(IRibbonControl control)
        {
            ExecuteCallback("OnOpenViewer", () =>
            {
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Image files (*.png;*.jpg;*.jpeg;*.bmp;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.tiff|PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
                    openFileDialog.Title = "Select Document to View";
                    
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        if (_documentViewer == null || _documentViewer.IsDisposed)
                        {
                            InitializeDocumentViewer();
                        }
                        
                        if (_documentViewer.LoadDocument(openFileDialog.FileName))
                        {
                            _documentViewer.Show();
                            _documentViewer.BringToFront();
                            Logger.Info($"Document loaded: {openFileDialog.FileName}");
                        }
                        else
                        {
                            MessageBox.Show("Failed to load document. Please try an image file (PNG, JPG, BMP, TIFF).", 
                                "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                _documentViewer = DocumentViewerManager.GetOrCreateViewer();
                _documentViewer.SnipAreaSelected += OnSnipAreaSelected;
                Logger.Info("Document viewer initialized");
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
                
                switch (e.SnipMode)
                {
                    case SnipMode.Text:
                        var textResult = await _snippEngine.ProcessSnipAsync(e.SelectedImage, e.PageNumber, 
                            new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                        
                        if (textResult.Success)
                        {
                            formula = DataSnipperFormulas.CreateTextFormula(e.DocumentPath, e.PageNumber, 
                                textResult.Value, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                            displayValue = textResult.Value;
                        }
                        break;
                        
                    case SnipMode.Sum:
                        var sumResult = await _snippEngine.ProcessSnipAsync(e.SelectedImage, e.PageNumber, 
                            new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                        
                        if (sumResult.Success && double.TryParse(sumResult.Value, out double sumValue))
                        {
                            formula = DataSnipperFormulas.CreateSumFormula(e.DocumentPath, e.PageNumber, 
                                sumValue, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), new List<double> { sumValue });
                            displayValue = sumValue.ToString();
                        }
                        break;
                        
                    case SnipMode.Validation:
                        formula = DataSnipperFormulas.CreateValidationFormula(e.DocumentPath, e.PageNumber, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                        displayValue = "✓";
                        break;
                        
                    case SnipMode.Exception:
                        formula = DataSnipperFormulas.CreateExceptionFormula(e.DocumentPath, e.PageNumber, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                        displayValue = "✗";
                        break;
                }
                
                // Insert the formula into Excel
                if (!string.IsNullOrEmpty(formula))
                {
                    activeCell.Formula = formula;
                    
                    // For validation and exception, just show the symbol
                    if (e.SnipMode == SnipMode.Validation || e.SnipMode == SnipMode.Exception)
                    {
                        activeCell.Value = displayValue;
                    }
                    
                    // Add visual indicator (border color)
                    var color = GetSnipColorCode(e.SnipMode);
                    activeCell.Borders.Color = color;
                    
                    Logger.Info($"{e.SnipMode} snip completed successfully");
                }
                
                // Highlight the region in the document viewer
                _documentViewer.HighlightRegion(new System.Drawing.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), GetSnipColor(e.SnipMode));
                
                // Reset snip mode
                _documentViewer.SetSnipMode(e.SnipMode, false);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing snip: {ex.Message}", ex);
                MessageBox.Show($"Error processing snip: {ex.Message}", "Snipper Pro Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        #endregion

        private string GetFallbackRibbonXml()
        {
            System.Diagnostics.Debug.WriteLine("Snipper Pro: Using fallback ribbon XML");
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""SnipperProTab"" label=""SNIPPER PRO"">
        <group id=""SnipGroup"" label=""Snips"">
          <button id=""TextSnipButton"" label=""Text Snip"" size=""large"" onAction=""OnTextSnip""
                  screentip=""Extract text from selected area"" 
                  supertip=""Use OCR to extract text from the selected area in the document viewer. Creates DS.TEXTS() formula."" />
          <button id=""SumSnipButton"" label=""Sum Snip"" size=""large"" onAction=""OnSumSnip""
                  screentip=""Sum numbers from selected area""
                  supertip=""Extract and sum numerical values from the selected area. Creates DS.SUMS() formula."" />
          <button id=""TableSnipButton"" label=""Table Snip"" size=""large"" onAction=""OnTableSnip""
                  screentip=""Extract table data""
                  supertip=""Extract structured table data from the selected area."" />
          <button id=""ValidationSnipButton"" label=""Validation"" size=""large"" onAction=""OnValidationSnip""
                  screentip=""Mark as validated""
                  supertip=""Mark the selected cell as validated with a checkmark."" />
          <button id=""ExceptionSnipButton"" label=""Exception"" size=""large"" onAction=""OnExceptionSnip""
                  screentip=""Mark as exception""
                  supertip=""Mark the selected cell as an exception with an X mark."" />
        </group>
        <group id=""ViewerGroup"" label=""Document Viewer"">
          <button id=""OpenViewerButton"" label=""Open Viewer"" size=""large"" onAction=""OnOpenViewer""
                  screentip=""Open document viewer""
                  supertip=""Open the document viewer to load and analyze documents."" />
          <button id=""MarkupButton"" label=""Markup"" size=""large"" onAction=""OnMarkupSnip""
                  screentip=""Toggle markup mode""
                  supertip=""Enable annotation and markup tools in the document viewer."" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
} 