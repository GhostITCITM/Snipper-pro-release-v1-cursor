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
        private DocumentViewer _documentViewer;
        private SnipMode _currentSnipMode = SnipMode.Text;
        
        // Static instance for access from other classes
        public static ThisAddIn Instance { get; private set; }

        public Excel.Application Application => _application;
        public DocumentViewer DocumentViewer => _documentViewer;

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
                    
                    // Safely try to register Excel defined names for DS formulas
                    try
                    {
                        if (_application != null && _application.Names != null)
                        {
                            Application.Names.Add("DS.TEXTS", "=SnipperPro.Connect.TEXTS");
                            Application.Names.Add("DS.SUMS", "=SnipperPro.Connect.SUMS");
                            Application.Names.Add("DS.TABLE", "=SnipperPro.Connect.TABLE");
                            Application.Names.Add("DS.VALIDATION", "=SnipperPro.Connect.VALIDATION");
                            Application.Names.Add("DS.EXCEPTION", "=SnipperPro.Connect.EXCEPTION");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Snipper Pro: Could not register defined names: {ex.Message}");
                        // Continue without defined names - they're not critical for basic functionality
                    }
                    
                    // Safely try to register event handlers
                    try
                    {
                        if (_application != null)
                        {
                            ((Excel.AppEvents_Event)_application).SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;
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
                        ((Excel.AppEvents_Event)_application).SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
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
                
                // Use the extracted data from the event args
                switch (e.SnipMode)
                {
                    case SnipMode.Text:
                        if (e.Success && !string.IsNullOrEmpty(e.ExtractedText))
                        {
                            formula = DataSnipperFormulas.CreateTextFormula(e.DocumentPath, e.PageNumber, 
                                e.ExtractedText, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                            displayValue = e.ExtractedText;
                        }
                        else
                        {
                            displayValue = "OCR failed";
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
                                    sum, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height), numbers);
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
                                using (var helper = new ExcelHelper(Application))
                                {
                                    helper.WriteTableToRange(tableData, activeCell);
                                }

                                formula = DataSnipperFormulas.CreateTableFormula(e.DocumentPath, e.PageNumber,
                                    tableData, new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                                activeCell.Formula = formula;
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
                            new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                        displayValue = "✓";
                        break;
                        
                    case SnipMode.Exception:
                        formula = DataSnipperFormulas.CreateExceptionFormula(e.DocumentPath, e.PageNumber, 
                            new SnipperCloneCleanFinal.Core.Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height));
                        displayValue = "✗";
                        break;
                }
                
                // Insert the value into Excel (except for table which handles its own cells)
                if (!string.IsNullOrEmpty(displayValue) && e.SnipMode != SnipMode.Table)
                {
                    // For Validation and Exception snips, use the formula to enable double-click navigation
                    // For other snip types, use the value to avoid #NAME? errors
                    if (e.SnipMode == SnipMode.Validation || e.SnipMode == SnipMode.Exception)
                    {
                        // Convert the formula to use DS. prefix for proper Excel integration
                        if (!string.IsNullOrEmpty(formula))
                        {
                            var dsFormula = formula.Replace("=SnipperPro.Connect.", "=DS.");
                            activeCell.Formula = dsFormula;
                        }
                        else
                        {
                            activeCell.Value2 = displayValue; // Fallback if formula creation failed
                        }
                    }
                    else
                    {
                        // Always put the actual value in the cell instead of the formula to avoid #NAME? errors
                        activeCell.Value2 = displayValue;
                    }
                    
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
                        // Move down one cell after each snip
                        var nextCell = activeCell.Offset[1, 0];
                        nextCell.Select();
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
                
                // Don't reset snip mode - keep it active for continuous snipping
                // _documentViewer.SetSnipMode(e.SnipMode, false);
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

        // Excel UDFs for DataSnipper formulas
        [ComVisible(true)]
        public object TEXTS(string snipId)
        {
            var data = DataSnipperFormulas.GetSnipData(snipId);
            return data != null ? data.ExtractedValue ?? string.Empty : "#N/A";
        }

        [ComVisible(true)]
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
        public object VALIDATION(string snipId)
        {
            return DataSnipperFormulas.GetSnipData(snipId) != null ? "✓" : "#N/A";
        }

        [ComVisible(true)]
        public object EXCEPTION(string snipId)
        {
            return DataSnipperFormulas.GetSnipData(snipId) != null ? "✗" : "#N/A";
        }

        private void OnSheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            try
            {
                var formula = target.Formula as string;
                
                // Add debugging
                System.Diagnostics.Debug.WriteLine($"Double-click detected on cell {target.Address}");
                System.Diagnostics.Debug.WriteLine($"Cell formula: '{formula}'");
                System.Diagnostics.Debug.WriteLine($"Cell value: '{target.Value}'");
                
                if (string.IsNullOrWhiteSpace(formula)) 
                {
                    System.Diagnostics.Debug.WriteLine("No formula found in cell - exiting");
                    return;
                }

                var match = System.Text.RegularExpressions.Regex.Match(formula, @"=(?:DS\.|SnipperPro\.Connect\.)(?:TEXTS|SUMS|TABLE|VALIDATION|EXCEPTION)\(""(?<id>[^""]+)""\)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                
                System.Diagnostics.Debug.WriteLine($"Regex match success: {match.Success}");
                
                if (match.Success)
                {
                    var snipId = match.Groups["id"].Value;
                    System.Diagnostics.Debug.WriteLine($"Extracted snip ID: {snipId}");
                    
                    if (DataSnipperFormulas.NavigateToSnip(snipId))
                    {
                        System.Diagnostics.Debug.WriteLine("Navigation successful - cancelling Excel default action");
                        cancel = true;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("Navigation failed");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Formula didn't match DS pattern");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in double-click handler: {ex.Message}");
            }
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