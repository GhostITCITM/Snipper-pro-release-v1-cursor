using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Extensibility;
using System.Windows.Forms;
using System.Diagnostics;
using SnipperClone.Core;
using System.Threading.Tasks;

namespace SnipperClone
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789012")]
    [ProgId("SnipperClone.Connect")]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        // Field for log file path
        private static readonly string logFilePath = Path.Combine(Path.GetTempPath(), "SnipperCloneLog.txt");

        // Simple logging method
        private static void Log(string message)
        {
            try
            {
                File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                // If logging fails, write to Debug output
                System.Diagnostics.Debug.WriteLine($"Failed to write to log file: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Original log message: {message}");
            }
        }

        private Microsoft.Office.Interop.Excel.Application _application;
        private object _addInInstance;
        private SnipEngine _snippEngine;
        private DocumentViewer _documentViewer;
        private IRibbonUI _ribbon;

        public Connect()
        {
            Log("Connect: Constructor called.");
            // Constructor logic here
            Log("Connect: Constructor finished.");
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Log("Connect: OnConnection called.");
            try
            {
                _application = (Microsoft.Office.Interop.Excel.Application)application;
                _addInInstance = addInInst;
                Log("Connect: Application and AddInInstance assigned.");
                
                _snippEngine = new SnipEngine(_application);
                Log("Connect: SnipEngine initialized.");

                _application.SheetSelectionChange += OnSheetSelectionChange;
                _application.WorkbookOpen += OnWorkbookOpen;
                Log("Connect: Event handlers subscribed.");
                
                System.Diagnostics.Debug.WriteLine("SnipperClone: Successfully connected to Excel");
                Log("Connect: OnConnection finished successfully.");
            }
            catch (Exception ex)
            {
                Log($"Connect: OnConnection ERROR - {ex.ToString()}"); // Log full exception
                System.Diagnostics.Debug.WriteLine($"SnipperClone: Error during connection: {ex.Message}");
                MessageBox.Show($"Error initializing SnipperClone: {ex.Message}\n\nDetails: {ex.ToString()}", "SnipperClone Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_documentViewer != null && !_documentViewer.IsDisposed)
                {
                    _documentViewer.Close();
                    _documentViewer.Dispose();
                }

                if (_application != null)
                {
                    _application.SheetSelectionChange -= OnSheetSelectionChange;
                    _application.WorkbookOpen -= OnWorkbookOpen;
                }

                _snippEngine?.Dispose();
                _application = null;
                _addInInstance = null;
                _ribbon = null;
                
                System.Diagnostics.Debug.WriteLine("SnipperClone: Successfully disconnected from Excel");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SnipperClone: Error during disconnection: {ex.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        // Ribbon callback for loading
        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            Log("Connect: OnRibbonLoad called.");
            try
            {
                _ribbon = ribbonUI;
                System.Diagnostics.Debug.WriteLine("SnipperClone: Ribbon loaded successfully");
                Log("Connect: OnRibbonLoad finished successfully.");
            }
            catch (Exception ex)
            {
                Log($"Connect: OnRibbonLoad ERROR - {ex.ToString()}");
                System.Diagnostics.Debug.WriteLine($"SnipperClone: Error loading ribbon: {ex.Message}");
            }
        }

        public string GetCustomUI(string ribbonID)
        {
            Log($"Connect: GetCustomUI called for ribbonID: {ribbonID}");
            string ribbonXml = null;
            try
            {
                Log("Connect: Attempting to load SnipperRibbon.xml from embedded resource.");
                var assembly = Assembly.GetExecutingAssembly();
                // Make sure the resource name matches exactly, including namespace
                string resourceName = "SnipperClone.SnipperRibbon.xml"; 
                Log($"Connect: Using resource name: {resourceName}");

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        Log("Connect: Embedded resource stream found.");
                        using (var reader = new StreamReader(stream))
                        {
                            ribbonXml = reader.ReadToEnd();
                            Log("Connect: Successfully read XML from embedded resource.");
                            // Log first 100 chars of loaded XML for verification
                            Log($"Connect: Loaded XML (first 100 chars): {(ribbonXml.Length > 100 ? ribbonXml.Substring(0, 100) : ribbonXml)}");
                            return ribbonXml;
                        }
                    }
                    else
                    {
                        Log("Connect: Embedded resource stream NOT found. Falling back to GetRibbonXml().");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Connect: GetCustomUI ERROR loading embedded resource - {ex.ToString()}");
                Log("Connect: Falling back to GetRibbonXml() due to exception.");
            }

            // Fallback logic
            Log("Connect: Executing fallback GetRibbonXml().");
            ribbonXml = GetRibbonXml(); // This is your original hardcoded XML
            Log($"Connect: Fallback XML (first 100 chars): {(ribbonXml.Length > 100 ? ribbonXml.Substring(0, 100) : ribbonXml)}");
            return ribbonXml;
        }

        private string GetRibbonXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""SnipperCloneTab"" label=""DATASNIPPER"">
        <group id=""SnipGroup"" label=""Snips"">
          <button id=""TextSnipButton"" 
                  label=""Text Snip"" 
                  size=""large""
                  imageMso=""TextBox""
                  onAction=""OnTextSnip""
                  screentip=""Text Snip""
                  supertip=""Extract text from selected area using OCR"" />
          
          <button id=""SumSnipButton"" 
                  label=""Sum Snip"" 
                  size=""large""
                  imageMso=""FunctionSum""
                  onAction=""OnSumSnip""
                  screentip=""Sum Snip""
                  supertip=""Extract and sum numbers from selected area"" />
          
          <button id=""TableSnipButton"" 
                  label=""Table Snip"" 
                  size=""large""
                  imageMso=""Table""
                  onAction=""OnTableSnip""
                  screentip=""Table Snip""
                  supertip=""Extract table data from selected area"" />
          
          <separator id=""ValidationSeparator"" />
          
          <button id=""ValidationSnipButton"" 
                  label=""Validation"" 
                  size=""large""
                  imageMso=""AcceptInvitation""
                  onAction=""OnValidationSnip""
                  screentip=""Validation Snip""
                  supertip=""Mark area as validated with checkmark"" />
          
          <button id=""ExceptionSnipButton"" 
                  label=""Exception"" 
                  size=""large""
                  imageMso=""Cancel""
                  onAction=""OnExceptionSnip""
                  screentip=""Exception Snip""
                  supertip=""Mark area as exception with cross mark"" />
        </group>
        
        <group id=""ViewerGroup"" label=""Document"">
          <button id=""OpenViewerButton"" 
                  label=""Open Viewer"" 
                  size=""large""
                  imageMso=""FileOpen""
                  onAction=""OnOpenViewer""
                  screentip=""Document Viewer""
                  supertip=""Open the document viewer to import and analyze documents"" />
        </group>
        
        <group id=""NavigationGroup"" label=""Navigation"">
          <button id=""JumpBackButton"" 
                  label=""Jump Back"" 
                  size=""large""
                  imageMso=""GoToDialog""
                  onAction=""OnJumpBack""
                  screentip=""Jump Back""
                  supertip=""Navigate to the document location of the selected snip"" />
          
          <button id=""ClearModeButton"" 
                  label=""Clear Mode"" 
                  size=""normal""
                  imageMso=""Cancel""
                  onAction=""OnClearMode""
                  screentip=""Clear Mode""
                  supertip=""Clear the current snip mode"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // Ribbon button callbacks
        public void OnTextSnip(IRibbonControl control)
        {
            try
            {
                _snippEngine.SetMode(SnipMode.Text);
                ShowDocumentViewer();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Text Snip", ex);
            }
        }

        public void OnSumSnip(IRibbonControl control)
        {
            try
            {
                _snippEngine.SetMode(SnipMode.Sum);
                ShowDocumentViewer();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Sum Snip", ex);
            }
        }

        public void OnTableSnip(IRibbonControl control)
        {
            try
            {
                _snippEngine.SetMode(SnipMode.Table);
                ShowDocumentViewer();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Table Snip", ex);
            }
        }

        public void OnValidationSnip(IRibbonControl control)
        {
            try
            {
                _snippEngine.SetMode(SnipMode.Validation);
                ShowDocumentViewer();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Validation Snip", ex);
            }
        }

        public void OnExceptionSnip(IRibbonControl control)
        {
            try
            {
                _snippEngine.SetMode(SnipMode.Exception);
                ShowDocumentViewer();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Exception Snip", ex);
            }
        }

        public void OnOpenViewer(IRibbonControl control)
        {
            try
            {
                ShowDocumentViewer();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Open Viewer", ex);
            }
        }

        public void OnJumpBack(IRibbonControl control)
        {
            try
            {
                var selectedCell = _application.Selection as Range;
                if (selectedCell == null)
                {
                    MessageBox.Show("Please select a cell first.", "No Cell Selected", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var cellAddress = selectedCell.Address[false, false];
                var snipRecord = _snippEngine.FindSnipByCell(cellAddress);
                
                if (snipRecord == null)
                {
                    MessageBox.Show($"No snip found for cell {cellAddress}.", "No Snip Found", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Show document viewer and navigate to the snip location
                ShowDocumentViewer();
                
                if (_documentViewer != null)
                {
                    _documentViewer.NavigateToSnip(snipRecord);
                    _documentViewer.BringToFront();
                }
            }
            catch (Exception ex)
            {
                HandleRibbonError("Jump Back", ex);
            }
        }

        public void OnHighlightSnips(IRibbonControl control)
        {
            try
            {
                _snippEngine.HighlightAllSnips();
                MessageBox.Show("All snip cells have been highlighted.", "Snips Highlighted", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                HandleRibbonError("Highlight Snips", ex);
            }
        }

        public void OnClearHighlights(IRibbonControl control)
        {
            try
            {
                _snippEngine.ClearAllHighlights();
                MessageBox.Show("All snip highlights have been cleared.", "Highlights Cleared", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                HandleRibbonError("Clear Highlights", ex);
            }
        }

        public void OnDeleteSnip(IRibbonControl control)
        {
            try
            {
                var selectedCell = _application.Selection as Range;
                if (selectedCell == null)
                {
                    MessageBox.Show("Please select a cell first.", "No Cell Selected", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var cellAddress = selectedCell.Address[false, false];
                var snipRecord = _snippEngine.FindSnipByCell(cellAddress);
                
                if (snipRecord == null)
                {
                    MessageBox.Show($"No snip found for cell {cellAddress}.", "No Snip Found", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var result = MessageBox.Show(
                    $"Are you sure you want to delete the {snipRecord.Mode} snip from cell {cellAddress}?",
                    "Confirm Delete Snip",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    _snippEngine.DeleteSnip(cellAddress);
                    
                    // Clear the cell content if it's a validation or exception mark
                    if (snipRecord.Mode == SnipMode.Validation || snipRecord.Mode == SnipMode.Exception)
                    {
                        selectedCell.Clear();
                    }
                    
                    MessageBox.Show("Snip deleted successfully.", "Snip Deleted", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                HandleRibbonError("Delete Snip", ex);
            }
        }

        public void OnShowSnipInfo(IRibbonControl control)
        {
            try
            {
                var selectedCell = _application.Selection as Range;
                if (selectedCell == null)
                {
                    MessageBox.Show("Please select a cell first.", "No Cell Selected", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var cellAddress = selectedCell.Address[false, false];
                var snipRecord = _snippEngine.FindSnipByCell(cellAddress);
                
                if (snipRecord == null)
                {
                    MessageBox.Show($"No snip found for cell {cellAddress}.", "No Snip Found", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var info = $"Snip Information:\n\n" +
                          $"Cell: {snipRecord.CellAddress}\n" +
                          $"Type: {snipRecord.Mode}\n" +
                          $"Document: {snipRecord.DocumentName}\n" +
                          $"Page: {snipRecord.PageNumber}\n" +
                          $"Created: {snipRecord.CreatedAt:yyyy-MM-dd HH:mm:ss}\n" +
                          $"Rectangle: {snipRecord.Rectangle}\n" +
                          $"Extracted Text: {snipRecord.ExtractedText}";

                MessageBox.Show(info, "Snip Information", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                HandleRibbonError("Show Snip Info", ex);
            }
        }

        public void OnClearMode(IRibbonControl control)
        {
            try
            {
                _snippEngine.ClearMode();
            }
            catch (Exception ex)
            {
                HandleRibbonError("Clear Mode", ex);
            }
        }

        private void ShowDocumentViewer()
        {
            try
            {
                if (_documentViewer == null || _documentViewer.IsDisposed)
                {
                    _documentViewer = new DocumentViewer(_snippEngine);
                }

                if (!_documentViewer.Visible)
                {
                    _documentViewer.Show();
                }

                // Initialize WebView2 asynchronously
                var initTask = Task.Run(async () => 
                {
                    try
                    {
                        await _documentViewer.InitializeWebViewAsync();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error initializing WebView2: {ex}");
                        MessageBox.Show($"Error initializing WebView2: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });

                _documentViewer.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening document viewer: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnSheetSelectionChange(object sh, Range target)
        {
            try
            {
                _snippEngine.OnCellSelectionChanged(target);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in selection change: {ex.Message}");
            }
        }

        private void OnWorkbookOpen(Workbook wb)
        {
            try
            {
                _snippEngine.OnWorkbookOpened(wb);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in workbook open: {ex.Message}");
            }
        }

        private void HandleRibbonError(string action, Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"SnipperClone: Error in {action}: {ex.Message}");
            MessageBox.Show($"Error in {action}: {ex.Message}", "SnipperClone Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
} 